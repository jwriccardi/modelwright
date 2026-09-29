<#
.SYNOPSIS
    Dependency license gate for the shipped projects.

.DESCRIPTION
    For every NuGet package (including transitive ones) used by the projects under src/, fails if:
      - the package is not in build/allowed-packages.json;
      - its .nuspec (in the NuGet global packages folder) is missing, or its license does not match the
        allowlist entry: the <license type="expression"> must equal the entry's "license", or, for legacy
        packages with no expression, the <licenseUrl> must equal the entry's "licenseUrl";
      - the license (nuspec or allowlist) is a denied one: (L|A)GPL, SSPL, BUSL, Commons-Clause or
        non-commercial (-NC-);
      - the entry is marked buildOnly but the package contributes runtime, native, resource or content
        files to the project (per obj/project.assets.json), i.e. something that could ship.
    Also fails if any src/**/*.csproj is missing from the solution (it would escape this check).
    Test projects (tests/) are exempt: they are never distributed.

    'dotnet list package' restores implicitly (SDK 10), so no separate restore is needed.
    Works in Windows PowerShell 5.1 and PowerShell 7+.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File build/check-licenses.ps1
#>
[CmdletBinding()]
param(
    [string]$Solution,
    [string]$Allowlist
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# $PSScriptRoot can be empty in Windows PowerShell 5.1 (e.g. with -File and a relative path).
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $Solution) { $Solution = Join-Path $scriptDir '..\ExcelModelingToolkit.sln' }
if (-not $Allowlist) { $Allowlist = Join-Path $scriptDir 'allowed-packages.json' }

# Licenses that must never ship: copyleft that would bind the add-in, source-available, non-commercial.
# Matched anywhere in the expression or URL, case-insensitively (-match).
$deniedLicensePattern = '\b(A|L)?GPL|SSPL|-NC-|BUSL|Commons-Clause'

# project.assets.json sections whose files are copied to the output or compiled into the project.
$shippingAssetSections = @('runtime', 'native', 'resource', 'runtimeTargets', 'contentFiles')

function Get-OptionalProperty($Object, [string]$Name) {
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return @() }
    return @($property.Value)
}

function Get-StringProperty($Object, [string]$Name) {
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property -or $null -eq $property.Value) { return $null }
    return ([string]$property.Value).Trim()
}

function Get-NuspecNode([xml]$Nuspec, [string]$Name) {
    return $Nuspec.SelectSingleNode("/*[local-name()='package']/*[local-name()='metadata']/*[local-name()='$Name']")
}

# Returns a list of problems with the package's license (empty if it matches the allowlist entry).
function Test-PackageLicense($Entry, [string]$Id, [string]$Version, [string[]]$PackageFolders) {
    $problems = New-Object System.Collections.Generic.List[string]
    $idLower = $Id.ToLowerInvariant()
    $nuspecPath = $null
    foreach ($folder in $PackageFolders) {
        $candidate = Join-Path $folder (Join-Path $idLower (Join-Path $Version.ToLowerInvariant() "$idLower.nuspec"))
        if (Test-Path -LiteralPath $candidate) { $nuspecPath = $candidate; break }
    }
    if ($null -eq $nuspecPath) {
        $problems.Add("nuspec not found in $($PackageFolders -join ', ')")
        return $problems
    }

    [xml]$nuspec = Get-Content -Raw -LiteralPath $nuspecPath
    $licenseNode = Get-NuspecNode $nuspec 'license'
    $licenseUrlNode = Get-NuspecNode $nuspec 'licenseUrl'
    $declaredLicense = Get-StringProperty $Entry 'license'
    $declaredUrl = Get-StringProperty $Entry 'licenseUrl'

    if (-not $declaredLicense) {
        $problems.Add('allowlist entry has no "license"')
    } elseif ($declaredLicense -match $deniedLicensePattern) {
        $problems.Add("allowlist license '$declaredLicense' is denied")
    }

    if ($null -ne $licenseNode) {
        $type = $licenseNode.GetAttribute('type')
        $actual = $licenseNode.InnerText.Trim()
        if ($type -ne 'expression') {
            $problems.Add("nuspec license is type '$type' ($actual); only expressions are checked - review it and extend this script")
        } elseif ($actual -match $deniedLicensePattern) {
            $problems.Add("nuspec license '$actual' is denied")
        } elseif ($declaredLicense -and -not [string]::Equals($actual, $declaredLicense, [StringComparison]::OrdinalIgnoreCase)) {
            $problems.Add("nuspec license is '$actual' but the allowlist says '$declaredLicense'")
        }
    } elseif ($null -ne $licenseUrlNode -and $licenseUrlNode.InnerText.Trim()) {
        $actualUrl = $licenseUrlNode.InnerText.Trim()
        if ($actualUrl -match $deniedLicensePattern) {
            $problems.Add("nuspec licenseUrl '$actualUrl' looks like a denied license")
        } elseif (-not $declaredUrl) {
            $problems.Add("nuspec has no license expression, only licenseUrl '$actualUrl'; review it and add it to the entry as ""licenseUrl""")
        } elseif (-not [string]::Equals($actualUrl, $declaredUrl, [StringComparison]::Ordinal)) {
            $problems.Add("nuspec licenseUrl is '$actualUrl' but the allowlist says '$declaredUrl'")
        }
    } else {
        $problems.Add("nuspec ($nuspecPath) declares no license")
    }

    return $problems
}

# Returns the files the package contributes to shipping asset sections (ignoring '_._' placeholders).
function Get-ShippingAssets($Assets, [string]$Id, [string]$Version) {
    $files = New-Object System.Collections.Generic.List[string]
    foreach ($target in $Assets.targets.PSObject.Properties) {
        foreach ($library in $target.Value.PSObject.Properties) {
            if ($library.Name -ne "$Id/$Version") { continue } # -ne is case-insensitive
            foreach ($section in $shippingAssetSections) {
                foreach ($sectionValue in @(Get-OptionalProperty $library.Value $section)) {
                    foreach ($file in $sectionValue.PSObject.Properties) {
                        if ((Split-Path -Leaf $file.Name) -ne '_._') { $files.Add("$($target.Name): $section/$($file.Name)") }
                    }
                }
            }
        }
    }
    return $files
}

$Solution = (Resolve-Path $Solution).Path
$repoRoot = Split-Path -Parent $Solution
$srcDir = Join-Path $repoRoot 'src'
$srcPrefix = ($srcDir -replace '\\', '/').TrimEnd('/') + '/'

$failures = New-Object System.Collections.Generic.List[string]

# Every shipped project must be in the solution, or 'dotnet list package' below never sees it.
$solutionProjects = @{}
foreach ($line in (Get-Content -LiteralPath $Solution)) {
    if ($line -match '^Project\("\{[^}]+\}"\)\s*=\s*"[^"]*"\s*,\s*"([^"]+\.csproj)"') {
        $solutionProjects[[IO.Path]::GetFullPath((Join-Path $repoRoot $Matches[1]))] = $true
    }
}
foreach ($csproj in @(Get-ChildItem -LiteralPath $srcDir -Recurse -File -Filter '*.csproj')) {
    if (-not $solutionProjects.ContainsKey($csproj.FullName)) { # hashtable keys are case-insensitive
        $failures.Add("$($csproj.FullName) is not in $Solution, so its packages are not checked")
        Write-Host ("  FAIL {0} - not in the solution" -f $csproj.FullName)
    }
}

$allowed = @{}
foreach ($entry in (Get-Content -Raw -Path $Allowlist | ConvertFrom-Json).packages) {
    $allowed[$entry.id.ToLowerInvariant()] = $entry
}

$output = & dotnet list $Solution package --include-transitive --format json
if ($LASTEXITCODE -ne 0) {
    Write-Host ($output -join [Environment]::NewLine)
    Write-Host "check-licenses: 'dotnet list package' failed (exit $LASTEXITCODE)."
    exit 1
}
$report = ($output -join [Environment]::NewLine) | ConvertFrom-Json

$problems = @(Get-OptionalProperty $report 'problems')
if ($problems.Count -gt 0) {
    $problems | ForEach-Object { Write-Host "check-licenses: problem: $($_.text)" }
    exit 1
}

$seen = @{}
$licenseResults = @{}
$shippedProjects = 0

foreach ($project in @(Get-OptionalProperty $report 'projects')) {
    $path = $project.path -replace '\\', '/'
    if (-not $path.StartsWith($srcPrefix, [StringComparison]::OrdinalIgnoreCase)) { continue }
    $shippedProjects++
    $projectName = [IO.Path]::GetFileNameWithoutExtension($path)

    # Written by the implicit restore above; locates the packages and lists the assets each one contributes.
    $assetsPath = Join-Path (Split-Path -Parent $project.path) 'obj\project.assets.json'
    if (-not (Test-Path -LiteralPath $assetsPath)) {
        $failures.Add("$projectName has no $assetsPath")
        Write-Host ("  FAIL [{0}] - {1} not found" -f $projectName, $assetsPath)
        continue
    }
    $assets = Get-Content -Raw -LiteralPath $assetsPath | ConvertFrom-Json
    $packageFolders = @($assets.packageFolders.PSObject.Properties | ForEach-Object { $_.Name })

    foreach ($framework in @(Get-OptionalProperty $project 'frameworks')) {
        $packages = @(Get-OptionalProperty $framework 'topLevelPackages') + @(Get-OptionalProperty $framework 'transitivePackages')
        foreach ($package in $packages) {
            $key = "$projectName|$($package.id)|$($package.resolvedVersion)"
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = $true
            $label = "{0} {1} [{2}]" -f $package.id, $package.resolvedVersion, $projectName

            $entry = $allowed[$package.id.ToLowerInvariant()]
            if ($null -eq $entry) {
                $failures.Add("$($package.id) $($package.resolvedVersion) ($projectName, $($framework.framework)) is not in build/allowed-packages.json")
                Write-Host "  FAIL $label - not allowlisted"
                continue
            }

            $packageKey = "$($package.id)|$($package.resolvedVersion)"
            if (-not $licenseResults.ContainsKey($packageKey)) {
                $licenseResults[$packageKey] = @(Test-PackageLicense $entry $package.id $package.resolvedVersion $packageFolders)
            }
            $packageProblems = @($licenseResults[$packageKey])

            $buildOnly = (Get-StringProperty $entry 'buildOnly') -eq 'True'
            if ($buildOnly) {
                foreach ($file in @(Get-ShippingAssets $assets $package.id $package.resolvedVersion)) {
                    $packageProblems += "marked buildOnly but contributes $file"
                }
            }

            if ($packageProblems.Count -gt 0) {
                foreach ($problem in $packageProblems) {
                    $failures.Add("$($package.id) $($package.resolvedVersion) ($projectName): $problem")
                    Write-Host "  FAIL $label - $problem"
                }
                continue
            }

            $kind = if ($buildOnly) { 'build-only' } else { 'shipped' }
            Write-Host ("  ok   {0} - {1}, {2}" -f $label, $entry.license, $kind)
        }
    }
}

if ($shippedProjects -eq 0) {
    Write-Host "check-licenses: no projects under $srcPrefix were found in $Solution."
    exit 1
}

if ($failures.Count -gt 0) {
    Write-Host ''
    Write-Host "check-licenses: FAILED ($($failures.Count) problem(s)):"
    $failures | ForEach-Object { Write-Host "  - $_" }
    exit 1
}

Write-Host "check-licenses: passed ($shippedProjects shipped project(s), $($seen.Count) package reference(s))."
exit 0
