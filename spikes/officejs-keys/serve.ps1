<#
.SYNOPSIS
  Starts the K1 spike's static HTTPS server on https://localhost:3100 (owner-run only).

.DESCRIPTION
  Runs server.js (Node's built-in https module, zero npm dependencies) which serves the
  ./web directory. Requires the office-addin-dev-certs localhost certificate to already be
  installed (run setup.ps1 first) at:
    %USERPROFILE%\.office-addin-dev-certs\localhost.crt
    %USERPROFILE%\.office-addin-dev-certs\localhost.key

  Run this in its own terminal window and leave it running while you test in Excel.
  Stop it with Ctrl+C when done.
#>

$ErrorActionPreference = "Stop"

$certDir = Join-Path $env:USERPROFILE ".office-addin-dev-certs"
$cert = Join-Path $certDir "localhost.crt"
$key = Join-Path $certDir "localhost.key"

if (-not (Test-Path $cert) -or -not (Test-Path $key)) {
    Write-Host "Dev certificate not found at $cert / $key." -ForegroundColor Red
    Write-Host "Run .\setup.ps1 first." -ForegroundColor Red
    exit 1
}

$root = $PSScriptRoot
Write-Host "Starting EMT K1 spike server on https://localhost:3100 ..." -ForegroundColor Cyan
Write-Host "Press Ctrl+C to stop." -ForegroundColor Cyan
node (Join-Path $root "server.js")
