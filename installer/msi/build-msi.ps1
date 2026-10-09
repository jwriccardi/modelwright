<#
.SYNOPSIS
    Builds the per-user MSI, Modelwright-<version>-x64.msi, from an already built Modelwright64.xll.

.DESCRIPTION
    1. Finds the WiX v5 command-line tool (wix on PATH, or %USERPROFILE%\.dotnet\tools\wix.exe). Install it once with
         dotnet tool install --global wix --version 5.0.2
       WiX 5 is used on purpose: WiX 6 and later need the Open Source Maintenance Fee EULA to be accepted
       (installer/msi/README.md).
    2. Builds the custom actions (installer/msi/CustomActions, Release).
    3. Runs wix build on installer/msi/Modelwright.wxs.
    4. With -Validate: runs the ICE checks (wix msi validate), lists the MSI's files, registry entries and custom
       actions (from wix msi decompile) and checks the essentials, and extracts it with an administrative install
       (msiexec /a), which copies the files to a scratch folder without registering anything or running the
       custom actions, to prove the payload. Nothing is installed for the current user.

    Works in Windows PowerShell 5.1 and PowerShell 7+.

.PARAMETER XllPath
    The packed add-in. Default: src\Modelwright.AddIn\bin\Release\net48\publish\Modelwright64.xll.

.PARAMETER Version
    The release version (for example 0.1.0 or 0.1.0-rc.1). Default: the project version from Directory.Build.props.
    The MSI's ProductVersion is its numeric part; the file name keeps the whole version.

.PARAMETER OutputDirectory
    Where the MSI is written. Default: dist under the repository root.

.PARAMETER Validate
    Also validate, list and extract the MSI (see above).

.EXAMPLE
    dotnet build Modelwright.sln -c Release
    powershell -NoProfile -ExecutionPolicy Bypass -File installer/msi/build-msi.ps1 -Validate
#>
[CmdletBinding()]
param(
    [string]$XllPath,
    [string]$Version,
    [string]$OutputDirectory,
    [switch]$Validate
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
$repoRoot = (Resolve-Path (Join-Path $scriptDir '..\..')).Path
if (-not $XllPath) { $XllPath = Join-Path $repoRoot 'src\Modelwright.AddIn\bin\Release\net48\publish\Modelwright64.xll' }
if (-not $OutputDirectory) { $OutputDirectory = Join-Path $repoRoot 'dist' }

# The WiX tool.
$wix = $null
$onPath = Get-Command wix -CommandType Application -ErrorAction SilentlyContinue | Select-Object -First 1
if ($onPath) { $wix = $onPath.Source }
$toolsWix = Join-Path $env:USERPROFILE '.dotnet\tools\wix.exe'
if (-not $wix -and (Test-Path -LiteralPath $toolsWix)) { $wix = $toolsWix }
if (-not $wix) { throw 'The WiX tool is not installed. Install it with: dotnet tool install --global wix --version 5.0.2' }
$wixVersion = "$(& $wix --version)".Trim()
if ($wixVersion -notmatch '^5\.') { throw "WiX $wixVersion found at $wix; this build needs WiX 5 (dotnet tool install --global wix --version 5.0.2)." }
Write-Host "WiX $wixVersion ($wix)"

if (-not (Test-Path -LiteralPath $XllPath -PathType Leaf)) {
    throw "$XllPath does not exist. Build the add-in first: dotnet build Modelwright.sln -c Release"
}
$XllPath = (Resolve-Path -LiteralPath $XllPath).Path

if (-not $Version) {
    $Version = "$(& dotnet msbuild (Join-Path $repoRoot 'src\Modelwright.AddIn\Modelwright.AddIn.csproj') -getProperty:Version -p:Configuration=Release)".Trim()
    if ($LASTEXITCODE -ne 0) { throw 'Could not read the project version.' }
}
if ($Version -notmatch '^(\d+\.\d+\.\d+)(-[0-9A-Za-z.-]+)?$') { throw "Version '$Version' is not <major>.<minor>.<patch>[-suffix]." }
$productVersion = $Matches[1]

# Custom actions.
$caProject = Join-Path $scriptDir 'CustomActions\Modelwright.Installer.CustomActions.csproj'
& dotnet build $caProject -c Release --nologo
if ($LASTEXITCODE -ne 0) { throw 'The custom action build failed.' }
$caDll = Join-Path $scriptDir 'CustomActions\bin\x64\Release\net48\Modelwright.Installer.CustomActions.CA.dll'
if (-not (Test-Path -LiteralPath $caDll -PathType Leaf)) { throw "$caDll was not produced." }

# The MSI.
if (-not (Test-Path -LiteralPath $OutputDirectory -PathType Container)) { $null = New-Item -ItemType Directory -Path $OutputDirectory }
$msi = Join-Path (Resolve-Path -LiteralPath $OutputDirectory).Path "Modelwright-$Version-x64.msi"
$intermediate = Join-Path $scriptDir 'obj'
& $wix build (Join-Path $scriptDir 'Modelwright.wxs') -arch x64 -nologo `
    -d "Version=$productVersion" -d "XllSource=$XllPath" -d "CaDll=$caDll" -d "RepoRoot=$repoRoot" `
    -intermediateFolder $intermediate -o $msi
if ($LASTEXITCODE -ne 0) { throw 'wix build failed.' }
Write-Host "Built $msi (ProductVersion $productVersion)"

if ($Validate) {
    # Suppressed warnings, both deliberate: ICE61 is AllowSameVersionUpgrades (see Modelwright.wxs); ICE91 says
    # per-user files would not be copied to every profile in a per-machine install, and this package is per-user
    # only (Scope="perUser").
    Write-Host '--- ICE validation'
    & $wix msi validate -sice ICE61 -sice ICE91 -intermediateFolder (Join-Path $intermediate 'validate') $msi
    if ($LASTEXITCODE -ne 0) { throw 'wix msi validate reported errors.' }
    Write-Host 'ICE validation passed.'

    Write-Host '--- Contents (wix msi decompile)'
    $decompiled = Join-Path $intermediate 'decompiled.wxs'
    & $wix msi decompile $msi -o $decompiled -intermediateFolder (Join-Path $intermediate 'decompile')
    if ($LASTEXITCODE -ne 0) { throw 'wix msi decompile failed.' }
    [xml]$source = Get-Content -LiteralPath $decompiled -Raw
    $package = $source.SelectSingleNode("//*[local-name()='Package']")
    Write-Host "Package: $($package.GetAttribute('Name')) $($package.GetAttribute('Version')), scope $($package.GetAttribute('Scope')), UpgradeCode $($package.GetAttribute('UpgradeCode'))"
    $files = @($source.SelectNodes("//*[local-name()='File']") | ForEach-Object { $_.GetAttribute('Name') })
    foreach ($file in $files) { Write-Host "File: $file" }
    foreach ($reg in @($source.SelectNodes("//*[local-name()='RegistryValue' or local-name()='RegistryKey' or local-name()='RemoveRegistryKey']"))) {
        Write-Host "Registry ($($reg.LocalName)): $($reg.GetAttribute('Root'))\$($reg.GetAttribute('Key')) $($reg.GetAttribute('Name'))=$($reg.GetAttribute('Value'))"
    }
    $actions = @($source.SelectNodes("//*[local-name()='CustomAction']") | ForEach-Object { $_.GetAttribute('Id') })
    foreach ($action in $actions) { Write-Host "Custom action: $action" }
    foreach ($name in 'Modelwright64.xll', 'LICENSE.txt', 'THIRD_PARTY_NOTICES.md', 'INSTALL.txt') {
        if ($files -notcontains $name) { throw "The MSI does not contain $name." }
    }
    if ($files -match '32\.xll') { throw 'The MSI contains a 32-bit add-in.' }
    if ($package.GetAttribute('Scope') -ne 'perUser') { throw 'The MSI is not per-user.' }
    foreach ($name in 'MwCheckPrerequisites', 'MwScheduleRegistration', 'MwApplyRegistration', 'MwRollbackRegistration') {
        if ($actions -notcontains $name) { throw "The MSI is missing the custom action $name." }
    }

    Write-Host '--- Administrative extraction (msiexec /a: copies files only, no registration, no custom actions)'
    $extract = Join-Path $intermediate 'extract'
    if (Test-Path -LiteralPath $extract) { Remove-Item -LiteralPath $extract -Recurse -Force }
    $log = Join-Path $intermediate 'extract.log'
    $process = Start-Process -FilePath msiexec.exe -ArgumentList @('/a', "`"$msi`"", '/qn', "TARGETDIR=`"$extract`"", '/l*v', "`"$log`"") -Wait -PassThru
    if ($process.ExitCode -ne 0) { throw "msiexec /a failed with exit code $($process.ExitCode) (log: $log)." }
    $extractedXll = @(Get-ChildItem -LiteralPath $extract -Recurse -Filter 'Modelwright64.xll')
    if ($extractedXll.Count -ne 1) { throw 'The extracted MSI does not hold exactly one Modelwright64.xll.' }
    $expected = (Get-FileHash -LiteralPath $XllPath -Algorithm SHA256).Hash
    $actual = (Get-FileHash -LiteralPath $extractedXll[0].FullName -Algorithm SHA256).Hash
    if ($actual -ne $expected) { throw 'The extracted Modelwright64.xll differs from the built one.' }
    Get-ChildItem -LiteralPath $extract -Recurse -File | ForEach-Object { Write-Host "Extracted: $($_.FullName.Substring($extract.Length + 1))" }
    Write-Host "The extracted Modelwright64.xll matches the build (SHA-256 $actual)."
}
