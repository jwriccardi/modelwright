<#
.SYNOPSIS
  Removes the K1 spike's sideload registration (owner-run only).

.DESCRIPTION
  Removes the three EMT K1 registry values under
  HKCU:\Software\Microsoft\Office\16.0\WEF\Developer that setup.ps1 created.
  Optionally also uninstalls the office-addin-dev-certs localhost certificate
  (pass -UninstallCert), which will show its own Windows/npx prompt(s).

.PARAMETER UninstallCert
  Also run `npx office-addin-dev-certs uninstall`. Off by default since the cert may be
  useful for other Office-Addin dev work on this machine.

.EXAMPLE
  .\cleanup.ps1
  .\cleanup.ps1 -UninstallCert
#>

param(
    [switch]$UninstallCert
)

$ErrorActionPreference = "Stop"

$wefKey = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
$guids = @(
    "91c5a88b-1521-47fb-ac5a-8c381f0eb854",  # EMT K1 Named Keys (A)
    "4d1cd616-e36b-42f3-ad1f-8467d78bb100",  # EMT K1 Literal Keys (B)
    "181291ff-36c5-47e3-8980-2c68d40d950a"   # EMT K1 Single (C)
)

if (Test-Path $wefKey) {
    foreach ($g in $guids) {
        $existing = Get-ItemProperty -Path $wefKey -Name $g -ErrorAction SilentlyContinue
        if ($null -ne $existing) {
            Remove-ItemProperty -Path $wefKey -Name $g -Force
            Write-Host "Removed registry value: $g"
        } else {
            Write-Host "Not present (already removed): $g"
        }
    }
} else {
    Write-Host "Registry key $wefKey does not exist - nothing to remove."
}

if ($UninstallCert) {
    Write-Host ""
    Write-Host "Uninstalling office-addin-dev-certs localhost certificate ..." -ForegroundColor Cyan
    npx --yes office-addin-dev-certs uninstall
}

Write-Host ""
Write-Host "Cleanup done. Restart Excel to fully drop the three EMT K1 add-ins from the Developer list." -ForegroundColor Green
