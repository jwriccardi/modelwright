<#
.SYNOPSIS
    Tests install/install.ps1 and install/uninstall.ps1 in Constrained Language mode, against a scratch registry
    key and a scratch folder. Plain PowerShell assertions (Pester is not on the CI image by default).

.DESCRIPTION
    The session is switched to ConstrainedLanguage first, and the scripts inherit it, so any .NET method call,
    Add-Type or COM use in them fails here, as it would under AppLocker script rules.
    Nothing outside the scratch locations is changed: the scripts get -RegistryRoot (a throwaway key under
    HKCU:\Software), -AddInsFolder and -SourceFolder (under %TEMP%), -ExcelBitness (no Excel needed) and
    -ExcelProcessName (a name that is not running, so a running Excel does not matter). The real Excel
    Options key is compared before and after to prove it was not touched. Both are removed at the end.

    Exit code: 0 if every check passed, 1 otherwise.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File tests/install/install-scripts.Tests.ps1
#>
$ExecutionContext.SessionState.LanguageMode = 'ConstrainedLanguage'

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
$repoRoot = Split-Path -Parent (Split-Path -Parent $scriptDir)
$install = Join-Path $repoRoot 'install\install.ps1'
$uninstall = Join-Path $repoRoot 'install\uninstall.ps1'

$runId = Get-Random -Minimum 100000 -Maximum 999999
$scratch = Join-Path $env:TEMP "modelwright-install-test-$runId"
$regRoot = "HKCU:\Software\Modelwright-InstallTest-$runId"
$officeRoot = Join-Path $regRoot 'Office'
$optionsKey = Join-Path $officeRoot '16.0\Excel\Options'
$managerKey = Join-Path $officeRoot '16.0\Excel\Add-in Manager'
$source = Join-Path $scratch 'release'
$addIns = Join-Path $scratch 'AddIns'
$realOptionsKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
$notRunning = "NoSuchProcess$runId"

$script:failures = 0

function Assert-True([bool]$Condition, [string]$Message) {
    if ($Condition) {
        Write-Host "PASS  $Message"
    } else {
        Write-Host "FAIL  $Message" -ForegroundColor Red
        $script:failures++
    }
}

function Assert-Equal($Actual, $Expected, [string]$Message) {
    if ($Actual -ceq $Expected) {
        Write-Host "PASS  $Message"
    } else {
        Write-Host "FAIL  $Message" -ForegroundColor Red
        Write-Host "      expected: $Expected"
        Write-Host "      actual:   $Actual"
        $script:failures++
    }
}

# "NAME=value" lines for every value of a key, sorted; "<missing>" if the key does not exist.
function Get-Snapshot([string]$Key) {
    if (-not (Test-Path -LiteralPath $Key)) { return '<missing>' }
    $lines = @(foreach ($name in @((Get-Item -LiteralPath $Key).Property)) {
        "$name=$(Get-ItemPropertyValue -LiteralPath $Key -Name $name)"
    })
    return (@($lines | Sort-Object) -join "`n")
}

# Runs a script with the scratch locations; returns its output (Write-Host included) as one string.
# The exit code is left in $LASTEXITCODE.
function Invoke-Script([string]$Path, [hashtable]$Arguments) {
    $Arguments.RegistryRoot = $officeRoot
    $Arguments.AddInsFolder = $addIns
    $Arguments.NoPause = $true
    if (-not $Arguments.ContainsKey('ExcelProcessName')) { $Arguments.ExcelProcessName = $notRunning }
    $output = & $Path @Arguments 6>&1 2>&1 | Out-String
    Write-Verbose $output
    return $output
}

function Reset-Options([hashtable]$Values) {
    if (Test-Path -LiteralPath $optionsKey) { Remove-Item -LiteralPath $optionsKey -Recurse }
    $null = New-Item -Path $optionsKey -Force
    foreach ($name in $Values.Keys) {
        $null = New-ItemProperty -LiteralPath $optionsKey -Name $name -Value $Values[$name] -PropertyType String
    }
}

$realBefore = Get-Snapshot $realOptionsKey
try {
    $null = New-Item -ItemType Directory -Path $source -Force
    Set-Content -LiteralPath (Join-Path $source 'Modelwright64.xll') -Value 'fake 64-bit add-in'
    Set-Content -LiteralPath (Join-Path $source 'Modelwright32.xll') -Value 'fake 32-bit add-in'
    # Mark the 64-bit file as downloaded from the internet (Mark-of-the-Web), as a browser would.
    Set-Content -LiteralPath (Join-Path $source 'Modelwright64.xll') -Stream 'Zone.Identifier' -Value "[ZoneTransfer]`r`nZoneId=3"
    $target64 = Join-Path $addIns 'Modelwright64.xll'
    $target32 = Join-Path $addIns 'Modelwright32.xll'
    $value64 = "/R `"$target64`""
    $value32 = "/R `"$target32`""

    Write-Host '--- Constrained Language mode'
    Assert-Equal "$($ExecutionContext.SessionState.LanguageMode)" 'ConstrainedLanguage' 'the test session is in ConstrainedLanguage'
    $probe = Join-Path $scratch 'probe.ps1'
    Set-Content -LiteralPath $probe -Value '$ExecutionContext.SessionState.LanguageMode'
    Assert-Equal "$(& $probe)" 'ConstrainedLanguage' 'a script run from the test session is in ConstrainedLanguage too'

    Write-Host '--- install.ps1 refuses while Excel is running'
    $running = (Get-Process -Id $PID).ProcessName
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source; ExcelProcessName = $running }
    Assert-Equal $LASTEXITCODE 2 'exit code 2'
    Assert-True ($out -match 'Excel is running') 'says Excel is running'
    Assert-Equal (Get-Snapshot $optionsKey) '<missing>' 'registry untouched'
    Assert-True (-not (Test-Path -LiteralPath $addIns)) 'add-ins folder untouched'

    Write-Host '--- install.ps1 -WhatIf on a clean profile changes nothing'
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source; WhatIf = $true }
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    Assert-True ($out -match 'What if') 'prints What if lines'
    Assert-Equal (Get-Snapshot $optionsKey) '<missing>' 'registry untouched'
    Assert-True (-not (Test-Path -LiteralPath $addIns)) 'add-ins folder not created'

    Write-Host '--- install.ps1 on a clean profile'
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    Assert-Equal (Get-Snapshot $optionsKey) "OPEN=$value64" 'registered as OPEN with the /R "path" form'
    Assert-Equal (Get-Content -LiteralPath $target64) 'fake 64-bit add-in' 'copied the 64-bit add-in'
    Assert-True ($null -eq (Get-Item -LiteralPath $target64 -Stream 'Zone.Identifier' -ErrorAction SilentlyContinue)) 'unblocked the copy (no Zone.Identifier)'
    Assert-True ($out -match 'Modelwright tab') 'tells the user the next step'

    Write-Host '--- install.ps1 next to other add-ins, replacing an old ModelingToolkit build'
    Reset-Options @{
        OPEN  = '"C:\Program Files\Other\First.xlam"'
        OPEN1 = '/R "C:\dev\bin\Release\net48\publish\ModelingToolkit64-packed.xll"'
        OPEN2 = '"C:\Other\Second.xlam"'
    }
    $null = New-ItemProperty -LiteralPath $optionsKey -Name 'Options' -Value 0 -PropertyType DWord
    $null = New-Item -Path $managerKey -Force
    $null = New-ItemProperty -LiteralPath $managerKey -Name 'C:\dev\publish\Modelwright64.xll' -Value '' -PropertyType String
    $null = New-ItemProperty -LiteralPath $managerKey -Name 'C:\Other\Unrelated.xll' -Value '' -PropertyType String
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    $expected = (@('OPEN="C:\Program Files\Other\First.xlam"', "OPEN1=$value64", 'OPEN2="C:\Other\Second.xlam"', 'Options=0') | Sort-Object) -join "`n"
    Assert-Equal (Get-Snapshot $optionsKey) $expected 'old entry replaced in place; other values kept'
    Assert-Equal (Get-Snapshot $managerKey) 'C:\Other\Unrelated.xll=' 'stale Modelwright Add-in Manager entry removed, others kept'

    Write-Host '--- install.ps1 again is idempotent'
    $before = Get-Snapshot $optionsKey
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    Assert-True ($out -match 'Already registered') 'says it is already registered'
    Assert-Equal (Get-Snapshot $optionsKey) $before 'registry unchanged'

    Write-Host '--- install.ps1 appends after other add-ins'
    Reset-Options @{ OPEN = '"C:\Other\First.xlam"'; OPEN1 = '"C:\Other\Second.xlam"' }
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    $expected = (@('OPEN="C:\Other\First.xlam"', 'OPEN1="C:\Other\Second.xlam"', "OPEN2=$value64") | Sort-Object) -join "`n"
    Assert-Equal (Get-Snapshot $optionsKey) $expected 'registered as the next free OPENn'

    Write-Host '--- install.ps1 switching to 32-bit replaces the 64-bit entry'
    $out = Invoke-Script $install @{ ExcelBitness = '32'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    $expected = (@('OPEN="C:\Other\First.xlam"', 'OPEN1="C:\Other\Second.xlam"', "OPEN2=$value32") | Sort-Object) -join "`n"
    Assert-Equal (Get-Snapshot $optionsKey) $expected 'one Modelwright entry, now the 32-bit add-in'
    Assert-Equal (Get-Content -LiteralPath $target32) 'fake 32-bit add-in' 'copied the 32-bit add-in'

    Write-Host '--- install.ps1 checks SHA256SUMS.txt'
    $sums = Join-Path $source 'SHA256SUMS.txt'
    Set-Content -LiteralPath $sums -Value ('0' * 64 + '  Modelwright64.xll')
    $before = Get-Snapshot $optionsKey
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 6 'a wrong checksum: exit code 6'
    Assert-Equal (Get-Snapshot $optionsKey) $before 'a wrong checksum: registry unchanged'
    # Get-FileHash (script code in 5.1) is unavailable in a hand-constrained session; certutil is part of Windows.
    $hash = @(& certutil.exe -hashfile (Join-Path $source 'Modelwright64.xll') SHA256)[1] -replace '\s', ''
    Set-Content -LiteralPath $sums -Value "$hash  Modelwright64.xll"
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 0 'a matching checksum: exit code 0'
    Remove-Item -LiteralPath $sums

    Write-Host '--- install.ps1 with the add-in file missing'
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = (Join-Path $scratch 'empty') }
    Assert-Equal $LASTEXITCODE 4 'exit code 4'

    Write-Host '--- install.ps1 bitness detection (read-only registry and EXCEL.EXE reads; no Excel on CI)'
    $out = Invoke-Script $install @{ SourceFolder = $source; WhatIf = $true }
    Assert-True ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq 3) "detection ends cleanly (exit $LASTEXITCODE)"
    if ($LASTEXITCODE -eq 0) {
        Assert-True ($out -match 'Excel is (32|64)-bit') 'reports the detected bitness'
    } else {
        Assert-True ($out -match 'Could not tell') 'explains that it could not tell'
    }

    Write-Host '--- uninstall.ps1 refuses while Excel is running'
    $before = Get-Snapshot $optionsKey
    $out = Invoke-Script $uninstall @{ ExcelProcessName = $running }
    Assert-Equal $LASTEXITCODE 2 'exit code 2'
    Assert-Equal (Get-Snapshot $optionsKey) $before 'registry unchanged'

    Write-Host '--- uninstall.ps1 -WhatIf changes nothing'
    Reset-Options @{ OPEN = '"C:\Other\First.xlam"'; OPEN1 = $value64; OPEN2 = '"C:\Other\Second.xlam"' }
    $before = Get-Snapshot $optionsKey
    $out = Invoke-Script $uninstall @{ RemoveFile = $true; WhatIf = $true }
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    Assert-True ($out -match 'What if') 'prints What if lines'
    Assert-Equal (Get-Snapshot $optionsKey) $before 'registry unchanged'
    Assert-True (Test-Path -LiteralPath $target64) 'add-in file kept'

    Write-Host '--- uninstall.ps1 removes the entry and renumbers the rest'
    $out = Invoke-Script $uninstall @{}
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    $expected = (@('OPEN="C:\Other\First.xlam"', 'OPEN1="C:\Other\Second.xlam"') | Sort-Object) -join "`n"
    Assert-Equal (Get-Snapshot $optionsKey) $expected 'OPEN, OPEN1 left, no gap'
    Assert-True (Test-Path -LiteralPath $target64) 'add-in file kept without -RemoveFile'

    Write-Host '--- uninstall.ps1 closes gaps left by others'
    Reset-Options @{ OPEN = $value64; OPEN2 = '"C:\Other\First.xlam"'; OPEN5 = '"C:\Other\Second.xlam"' }
    $out = Invoke-Script $uninstall @{ RemoveFile = $true }
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    Assert-Equal (Get-Snapshot $optionsKey) $expected 'renumbered OPEN, OPEN1 in the old order'
    Assert-True (-not (Test-Path -LiteralPath $target64)) '-RemoveFile deleted Modelwright64.xll'
    Assert-True (-not (Test-Path -LiteralPath $target32)) '-RemoveFile deleted Modelwright32.xll'

    Write-Host '--- uninstall.ps1 again has nothing to do'
    $out = Invoke-Script $uninstall @{}
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    Assert-True ($out -match 'nothing to remove') 'says there was nothing to remove'
    Assert-Equal (Get-Snapshot $optionsKey) $expected 'registry unchanged'

    Write-Host '--- the real Excel Options key'
    Assert-Equal (Get-Snapshot $realOptionsKey) $realBefore 'unchanged by the tests'
} catch {
    Write-Host "FAIL  unexpected error: $($_.Exception.Message) at line $($_.InvocationInfo.ScriptLineNumber)" -ForegroundColor Red
    $script:failures++
} finally {
    if (Test-Path -LiteralPath $regRoot) { Remove-Item -LiteralPath $regRoot -Recurse }
    if (Test-Path -LiteralPath $scratch) { Remove-Item -LiteralPath $scratch -Recurse -Force }
}

Write-Host ''
if ($script:failures -gt 0) {
    Write-Host "$($script:failures) check(s) failed." -ForegroundColor Red
    exit 1
}
Write-Host 'All install-script checks passed.' -ForegroundColor Green
exit 0
