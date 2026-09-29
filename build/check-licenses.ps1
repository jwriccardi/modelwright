<#
.SYNOPSIS
    Dependency license gate for the shipped projects.

.DESCRIPTION
    Lists every NuGet package (including transitive ones) used by the projects under src/ and fails
    if any package is not in build/allowed-packages.json, or if an allowlisted license is GPL/AGPL.
    Test projects (tests/) are exempt: they are never distributed.

    Requires a prior `dotnet restore`. Works in Windows PowerShell 5.1 and PowerShell 7+.

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

# Licenses that must never ship (copyleft that would bind the add-in).
$deniedLicensePattern = '^(A)?GPL'

function Get-OptionalProperty($Object, [string]$Name) {
    $property = $Object.PSObject.Properties[$Name]
    if ($null -eq $property) { return @() }
    return @($property.Value)
}

$Solution = (Resolve-Path $Solution).Path
$repoRoot = Split-Path -Parent $Solution
$srcPrefix = ((Join-Path $repoRoot 'src') -replace '\\', '/').TrimEnd('/') + '/'

$allowed = @{}
foreach ($entry in (Get-Content -Raw -Path $Allowlist | ConvertFrom-Json).packages) {
    $allowed[$entry.id.ToLowerInvariant()] = $entry
}

$output = & dotnet list $Solution package --include-transitive --format json
if ($LASTEXITCODE -ne 0) {
    Write-Host ($output -join [Environment]::NewLine)
    Write-Host "check-licenses: 'dotnet list package' failed (exit $LASTEXITCODE). Run 'dotnet restore' first."
    exit 1
}
$report = ($output -join [Environment]::NewLine) | ConvertFrom-Json

$problems = @(Get-OptionalProperty $report 'problems')
if ($problems.Count -gt 0) {
    $problems | ForEach-Object { Write-Host "check-licenses: problem: $($_.text)" }
    exit 1
}

$failures = New-Object System.Collections.Generic.List[string]
$seen = @{}
$shippedProjects = 0

foreach ($project in @(Get-OptionalProperty $report 'projects')) {
    $path = $project.path -replace '\\', '/'
    if (-not $path.StartsWith($srcPrefix, [StringComparison]::OrdinalIgnoreCase)) { continue }
    $shippedProjects++
    $projectName = [IO.Path]::GetFileNameWithoutExtension($path)

    foreach ($framework in @(Get-OptionalProperty $project 'frameworks')) {
        $packages = @(Get-OptionalProperty $framework 'topLevelPackages') + @(Get-OptionalProperty $framework 'transitivePackages')
        foreach ($package in $packages) {
            $key = "$projectName|$($package.id)|$($package.resolvedVersion)"
            if ($seen.ContainsKey($key)) { continue }
            $seen[$key] = $true

            $entry = $allowed[$package.id.ToLowerInvariant()]
            if ($null -eq $entry) {
                $failures.Add("$($package.id) $($package.resolvedVersion) ($projectName, $($framework.framework)) is not in build/allowed-packages.json")
                Write-Host ("  FAIL {0} {1} [{2}] - not allowlisted" -f $package.id, $package.resolvedVersion, $projectName)
                continue
            }
            if ($entry.license -match $deniedLicensePattern) {
                $failures.Add("$($package.id) is licensed '$($entry.license)', which is not allowed in shipped code")
                Write-Host ("  FAIL {0} {1} [{2}] - denied license {3}" -f $package.id, $package.resolvedVersion, $projectName, $entry.license)
                continue
            }
            $kind = if ($entry.buildOnly) { 'build-only' } else { 'shipped' }
            Write-Host ("  ok   {0} {1} [{2}] - {3}, {4}" -f $package.id, $package.resolvedVersion, $projectName, $entry.license, $kind)
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
