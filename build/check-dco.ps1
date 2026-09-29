<#
.SYNOPSIS
    Developer Certificate of Origin gate: every non-merge commit in <Base>..HEAD must be signed off.

.DESCRIPTION
    Fails (exit 1) and prints the offending commits if any non-merge commit reachable from HEAD but not
    from -Base lacks a "Signed-off-by:" trailer (see CONTRIBUTING.md). Needs the history of both refs
    (in CI: a full-depth checkout). Works in Windows PowerShell 5.1 and PowerShell 7+.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File build/check-dco.ps1 -Base origin/main
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)]
    [string]$Base
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

& git rev-parse --verify --quiet "$Base^{commit}" | Out-Null
if ($LASTEXITCODE -ne 0) {
    Write-Host "check-dco: base ref '$Base' not found (fetch it, or use a full-depth checkout)."
    exit 1
}

$commits = @(& git rev-list --no-merges "$Base..HEAD")
if ($LASTEXITCODE -ne 0) {
    Write-Host "check-dco: 'git rev-list $Base..HEAD' failed (exit $LASTEXITCODE)."
    exit 1
}

$unsigned = New-Object System.Collections.Generic.List[string]
foreach ($sha in $commits) {
    # Quoted: an unquoted comma would make PowerShell pass an array.
    $signers = @(& git show -s '--format=%(trailers:key=Signed-off-by,valueonly)' $sha) | Where-Object { $_.Trim() }
    if ($LASTEXITCODE -ne 0) {
        Write-Host "check-dco: 'git show $sha' failed (exit $LASTEXITCODE)."
        exit 1
    }
    if (@($signers).Count -eq 0) {
        $unsigned.Add((& git show -s '--format=%H %s' $sha))
    }
}

if ($unsigned.Count -gt 0) {
    Write-Host "check-dco: FAILED - $($unsigned.Count) of $($commits.Count) commit(s) in $Base..HEAD have no Signed-off-by trailer:"
    $unsigned | ForEach-Object { Write-Host "  $_" }
    Write-Host "Sign them off with 'git rebase --signoff $Base' and force-push (see CONTRIBUTING.md)."
    exit 1
}

Write-Host "check-dco: passed ($($commits.Count) commit(s) in $Base..HEAD, all signed off)."
exit 0
