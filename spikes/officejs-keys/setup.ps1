<#
.SYNOPSIS
  One-time setup for the K1 keyboard-shortcut spike (owner-run only).

.DESCRIPTION
  1) Installs the office-addin-dev-certs localhost dev certificate (via npx, no local
     install). This WILL show a Windows "Do you want to allow this app to install a
     certificate..." / certificate trust prompt - accept it, the cert is scoped to
     localhost and is the standard Office-Addin dev workflow cert.
  2) Registers the three spike manifests for sideloading by writing string values under
     HKCU:\Software\Microsoft\Office\16.0\WEF\Developer, where:
       - the VALUE NAME is the add-in's manifest <Id> GUID
       - the VALUE DATA is the full path to that manifest's .xml file
     This is Microsoft's documented manual-sideload convention for Windows desktop Office
     (see "Sideload Office Add-ins on Windows for testing" -
     https://learn.microsoft.com/en-us/office/dev/add-ins/testing/troubleshoot-manifest#verify-the-manifest-is-valid
     and https://learn.microsoft.com/en-us/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-add-in-deployment,
     and it is exactly what office-addin-dev-settings / office-addin-debugging write on
     `npx office-addin-debugging start` under the hood on Windows - same registry key,
     same GUID-name/path-value convention).

.NOTES
  Run this yourself; it is NOT executed automatically. Excel should ideally be closed
  while you run this (registry changes take effect the next time Excel starts / the
  Developer add-ins list is refreshed), then reopen Excel afterward.

  Does NOT launch Excel. Does NOT start the server (run serve.ps1 separately).
#>

$ErrorActionPreference = "Stop"

$root = $PSScriptRoot
$manifestA = Join-Path $root "manifests\manifest-a.xml"
$manifestB = Join-Path $root "manifests\manifest-b.xml"
$manifestC = Join-Path $root "manifests\manifest-c.xml"

foreach ($m in @($manifestA, $manifestB, $manifestC)) {
    if (-not (Test-Path $m)) {
        throw "Manifest not found: $m"
    }
}

Write-Host "== Step 1: install office-addin-dev-certs localhost certificate ==" -ForegroundColor Cyan
Write-Host "This will prompt you (Windows certificate trust dialog / npx package confirmation)." -ForegroundColor Yellow
npx --yes office-addin-dev-certs install
if ($LASTEXITCODE -ne 0) {
    throw "office-addin-dev-certs install failed with exit code $LASTEXITCODE"
}

Write-Host ""
Write-Host "== Step 2: register manifests under HKCU:\Software\Microsoft\Office\16.0\WEF\Developer ==" -ForegroundColor Cyan

$wefKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
if (-not (Test-Path $wefKey)) {
    New-Item -Path $wefKey -Force | Out-Null
    Write-Host "Created registry key: $wefKey"
}

function Register-Manifest {
    param(
        [string]$ManifestPath,
        [string]$Guid,
        [string]$FriendlyName
    )
    Set-ItemProperty -Path $wefKey -Name $Guid -Value $ManifestPath -Type String -Force
    Write-Host "Registered $FriendlyName -> $Guid = $ManifestPath"
}

Register-Manifest -ManifestPath $manifestA -Guid "91c5a88b-1521-47fb-ac5a-8c381f0eb854" -FriendlyName "EMT K1 Named Keys (A)"
Register-Manifest -ManifestPath $manifestB -Guid "4d1cd616-e36b-42f3-ad1f-8467d78bb100" -FriendlyName "EMT K1 Literal Keys (B)"
Register-Manifest -ManifestPath $manifestC -Guid "181291ff-36c5-47e3-8980-2c68d40d950a" -FriendlyName "EMT K1 Single (C)"

Write-Host ""
Write-Host "Done. Next steps:" -ForegroundColor Green
Write-Host "  1. Start the server in its own terminal:  .\serve.ps1"
Write-Host "  2. Restart Excel (fully close, then reopen) so it re-reads the Developer registry key."
Write-Host "  3. Home tab > Add-ins > My Add-ins > Developer Add-ins - the three EMT K1 add-ins should be listed."
Write-Host "  4. When finished, run .\cleanup.ps1 to remove the registry entries (and optionally the dev cert)."
