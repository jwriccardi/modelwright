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
    Options and Add-in Manager keys are compared before and after to prove they were not touched. The scratch
    key and folder are removed at the end.

    The CI runner has no Office, and its PSModulePath makes Get-FileHash fail under hand-set ConstrainedLanguage.
    Both are reproduced on any machine: -ExcelExePath points detection at a missing file (or at powershell.exe as
    a stand-in), and Get-FileHash / certutil.exe are shadowed by test aliases and functions.
    Every failed check prints the last script run's command line, exit code and full output.

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
# Stand-ins for HKCU's and HKLM's Software\Microsoft\Office\Excel\Addins (the COM add-ins install.ps1 looks through).
$comUserKey = Join-Path $regRoot 'ComAddIns\User'
$comMachineKey = Join-Path $regRoot 'ComAddIns\Machine'
$macabacusNote = 'Macabacus is installed too. Both add-ins use the same keyboard shortcuts.'
$source = Join-Path $scratch 'release'
$addIns = Join-Path $scratch 'AddIns'
$realOptionsKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
$realManagerKey = 'HKCU:\Software\Microsoft\Office\16.0\Excel\Add-in Manager'
$notRunning = "NoSuchProcess$runId"

$script:failures = 0
$script:lastRun = $null

# On a failed check, print the exit code and the full output of the last script run, so CI logs show why.
function Write-LastRun {
    if (-not $script:lastRun) { return }
    Write-Host "      last run: $($script:lastRun.Command)"
    Write-Host "      exit code: $($script:lastRun.ExitCode)"
    Write-Host '      output (stdout, stderr, Write-Host):'
    foreach ($line in @($script:lastRun.Output -split "`r?`n")) { Write-Host "      | $line" }
}

function Assert-True([bool]$Condition, [string]$Message) {
    if ($Condition) {
        Write-Host "PASS  $Message"
    } else {
        Write-Host "FAIL  $Message" -ForegroundColor Red
        Write-LastRun
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
        Write-LastRun
        $script:failures++
    }
}

# SHA-256 via certutil.exe, parsed as install.ps1 does: the first line that is 64 hex digits without spaces.
function Get-CertUtilSha256([string]$Path) {
    foreach ($line in @(& certutil.exe -hashfile $Path SHA256)) {
        $hex = "$line" -replace '\s', ''
        if ($hex -match '^[0-9A-Fa-f]{64}$') { return $hex.ToUpperInvariant() }
    }
    return $null
}

# "NAME=value" lines for every value of a key, sorted; "<missing>" if the key does not exist.
function Get-Snapshot([string]$Key) {
    if (-not (Test-Path -LiteralPath $Key)) { return '<missing>' }
    $lines = @(foreach ($name in @((Get-Item -LiteralPath $Key).Property)) {
        "$name=$(Get-ItemPropertyValue -LiteralPath $Key -Name $name)"
    })
    return (@($lines | Sort-Object) -join "`n")
}

# Runs a script with the scratch locations; returns all its output (Write-Host, verbose and errors included) as
# one string. The exit code is left in $LASTEXITCODE; both are kept in $script:lastRun for failure reports.
function Invoke-Script([string]$Path, [hashtable]$Arguments) {
    if (-not $Arguments.ContainsKey('RegistryRoot')) { $Arguments.RegistryRoot = $officeRoot }
    $Arguments.AddInsFolder = $addIns
    $Arguments.NoPause = $true
    if (-not $Arguments.ContainsKey('ExcelProcessName')) { $Arguments.ExcelProcessName = $notRunning }
    if ($Path -eq $install) { $Arguments.ComAddInKeys = @($comUserKey, $comMachineKey) }
    $output = & $Path @Arguments *>&1 | Out-String -Width 4096
    $exitCode = $LASTEXITCODE
    $argText = @(foreach ($key in @($Arguments.Keys | Sort-Object)) { "-$key $($Arguments[$key])" }) -join ' '
    $script:lastRun = @{ Command = "$(Split-Path -Leaf $Path) $argText"; ExitCode = $exitCode; Output = $output }
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

# Recreates the scratch COM add-in keys with these add-ins: "User\Name" or "Machine\Name" = LoadBehavior.
function Reset-ComAddIns([hashtable]$AddIns) {
    foreach ($key in @($comUserKey, $comMachineKey)) {
        if (Test-Path -LiteralPath $key) { Remove-Item -LiteralPath $key -Recurse }
    }
    foreach ($name in $AddIns.Keys) {
        $root = if ($name -like 'User\*') { $comUserKey } else { $comMachineKey }
        $key = Join-Path $root ($name -replace '^(User|Machine)\\', '')
        $null = New-Item -Path $key -Force
        $null = New-ItemProperty -LiteralPath $key -Name 'LoadBehavior' -Value $AddIns[$name] -PropertyType DWord
        $null = New-ItemProperty -LiteralPath $key -Name 'FriendlyName' -Value 'stand-in' -PropertyType String
    }
}

# Recreates the Add-in Manager key with these value names (empty string values, as Excel writes them).
function Reset-Manager([string[]]$Names) {
    if (Test-Path -LiteralPath $managerKey) { Remove-Item -LiteralPath $managerKey -Recurse }
    $null = New-Item -Path $managerKey -Force
    foreach ($name in $Names) {
        $null = New-ItemProperty -LiteralPath $managerKey -Name $name -Value '' -PropertyType String
    }
}

$realBefore = Get-Snapshot $realOptionsKey
$realManagerBefore = Get-Snapshot $realManagerKey
try {
    $null = New-Item -ItemType Directory -Path $source -Force
    Set-Content -LiteralPath (Join-Path $source 'Modelwright64.xll') -Value 'fake 64-bit add-in'
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

    Write-Host '--- install.ps1 without Macabacus says nothing about it'
    Assert-True ($out -notmatch 'Macabacus') 'no Macabacus note (no COM add-in keys)'
    Reset-ComAddIns @{ 'User\Other.Connect' = 3; 'Machine\Macabacus' = 2 }
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    Assert-True ($out -notmatch 'Macabacus is installed') 'no note for another add-in, nor for Macabacus not set to load (LoadBehavior 2)'

    Write-Host '--- install.ps1 with Macabacus installed for this user notes the shared shortcuts'
    Reset-ComAddIns @{ 'User\Macabacus' = 3 }
    $before = Get-Snapshot $optionsKey
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 0 'exit code 0 (the note changes nothing)'
    Assert-True ($out.Contains($macabacusNote)) 'prints the Macabacus note'
    Assert-True ($out.Contains('click Modelwright > Shortcuts to switch Modelwright''s off')) 'the note says how to keep Macabacus''s'
    Assert-True ($out.IndexOf('Done.') -lt $out.IndexOf($macabacusNote)) 'the note comes after the install succeeded'
    Assert-Equal (Get-Snapshot $optionsKey) $before 'registry the same as without Macabacus'

    Write-Host '--- install.ps1 with Macabacus installed for all users (another ProgId) notes it too'
    Reset-ComAddIns @{ 'Machine\Macabacus.Excel.vsto' = 3 }
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    Assert-True ($out.Contains($macabacusNote)) 'prints the Macabacus note'

    Write-Host '--- install.ps1 -WhatIf with Macabacus installed prints no note'
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source; WhatIf = $true }
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    Assert-True ($out -notmatch 'Macabacus is installed') 'no note when nothing was installed'
    Reset-ComAddIns @{}

    Write-Host '--- install.ps1 appends after other add-ins'
    Reset-Options @{ OPEN = '"C:\Other\First.xlam"'; OPEN1 = '"C:\Other\Second.xlam"' }
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    $expected = (@('OPEN="C:\Other\First.xlam"', 'OPEN1="C:\Other\Second.xlam"', "OPEN2=$value64") | Sort-Object) -join "`n"
    Assert-Equal (Get-Snapshot $optionsKey) $expected 'registered as the next free OPENn'

    Write-Host '--- install.ps1 -ExcelBitness 32 is refused (Modelwright is 64-bit only)'
    $before = Get-Snapshot $optionsKey
    $out = Invoke-Script $install @{ ExcelBitness = '32'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 8 'exit code 8'
    Assert-True ($out -match 'Modelwright is 64-bit only; your Excel is 32-bit') 'says Modelwright is 64-bit only'
    Assert-Equal (Get-Snapshot $optionsKey) $before 'registry unchanged'
    Assert-True (-not (Test-Path -LiteralPath $target32)) 'no 32-bit add-in copied'

    Write-Host '--- install.ps1 replaces a 32-bit entry left by an earlier build'
    Reset-Options @{ OPEN = '"C:\Other\First.xlam"'; OPEN1 = $value32 }
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    $expected = (@('OPEN="C:\Other\First.xlam"', "OPEN1=$value64") | Sort-Object) -join "`n"
    Assert-Equal (Get-Snapshot $optionsKey) $expected 'the 32-bit entry is replaced by the 64-bit one in its slot'

    # On exit Excel rewrites entries for files in its own AddIns folder as a bare name, and may keep the
    # full-path entry too.
    Write-Host '--- install.ps1 replaces a bare entry (as Excel rewrites it) in its slot and drops a full-path duplicate'
    Reset-Options @{
        OPEN  = '"C:\Other\First.xlam"'
        OPEN1 = '/R "Modelwright64.xll"'
        OPEN2 = '"C:\Other\Second.xlam"'
        OPEN3 = $value64
        OPEN4 = '/R "C:\Other\MyModelwright64.xll"'
        OPEN5 = '"C:\Other\Modelwright64.xll.bak"'
    }
    Reset-Manager @('Modelwright64.xll', 'MODELWRIGHT32.XLL', 'ModelingToolkit64-packed.xll', 'C:\Other\Unrelated.xll')
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    $expected = (@('OPEN="C:\Other\First.xlam"', "OPEN1=$value64", 'OPEN2="C:\Other\Second.xlam"',
        'OPEN3=/R "C:\Other\MyModelwright64.xll"', 'OPEN4="C:\Other\Modelwright64.xll.bak"') | Sort-Object) -join "`n"
    Assert-Equal (Get-Snapshot $optionsKey) $expected 'bare entry replaced in its slot, duplicate removed, look-alikes kept, renumbered'
    Assert-Equal (Get-Snapshot $managerKey) 'C:\Other\Unrelated.xll=' 'bare Add-in Manager entries (any case, old name too) removed, others kept'

    Write-Host '--- install.ps1 with only a bare entry of the right add-in still writes the full path'
    Reset-Options @{ OPEN = '/R "Modelwright64.xll"'; OPEN1 = '"C:\Other\First.xlam"' }
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    $expected = (@("OPEN=$value64", 'OPEN1="C:\Other\First.xlam"') | Sort-Object) -join "`n"
    Assert-Equal (Get-Snapshot $optionsKey) $expected 'bare entry replaced by the full path in the same slot'

    Write-Host '--- install.ps1 replaces a bare old ModelingToolkit entry'
    Reset-Options @{ OPEN = '"C:\Other\First.xlam"'; OPEN1 = '/R "ModelingToolkit64-packed.xll"'; OPEN2 = '"C:\Other\Second.xlam"' }
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    $expected = (@('OPEN="C:\Other\First.xlam"', "OPEN1=$value64", 'OPEN2="C:\Other\Second.xlam"') | Sort-Object) -join "`n"
    Assert-Equal (Get-Snapshot $optionsKey) $expected 'bare ModelingToolkit entry replaced in its slot'

    # In a hand-constrained Windows PowerShell 5.1 session, Get-FileHash (script code there) is either missing
    # (when PowerShell 7's Microsoft.PowerShell.Utility is first on PSModulePath, as on a developer machine) or
    # fails on its first .NET call (as on the CI runner). Both are forced here by shadowing it: an alias wins over
    # a function, and a function over a cmdlet or an .exe, and the scripts run in a child scope that sees them.
    Write-Host '--- install.ps1 checks SHA256SUMS.txt (Get-FileHash fails, as on CI: certutil.exe is used)'
    $sums = Join-Path $source 'SHA256SUMS.txt'
    $hash = Get-CertUtilSha256 (Join-Path $source 'Modelwright64.xll')
    Assert-True ($null -ne $hash) "the test can hash with certutil.exe ($hash)"
    Set-Item -Path 'function:Get-FileHash' -Value { throw 'Get-FileHash is disabled by the test' }
    Set-Content -LiteralPath $sums -Value ('0' * 64 + '  Modelwright64.xll')
    $before = Get-Snapshot $optionsKey
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source; Verbose = $true }
    Assert-Equal $LASTEXITCODE 6 'a wrong checksum: exit code 6'
    Assert-True ($out -match 'does not match SHA256SUMS') 'a wrong checksum: says it does not match'
    Assert-True ($out -match 'using certutil') 'a wrong checksum: fell back to certutil.exe'
    Assert-Equal (Get-Snapshot $optionsKey) $before 'a wrong checksum: registry unchanged'
    Set-Content -LiteralPath $sums -Value "$($hash.ToLowerInvariant())  Modelwright64.xll"
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 0 'a matching (lower-case) checksum: exit code 0'

    Write-Host '--- install.ps1 checks SHA256SUMS.txt (Get-FileHash missing, as on a developer machine)'
    Remove-Item -LiteralPath 'function:Get-FileHash'
    Set-Alias -Name Get-FileHash -Value "NoSuchCommand$runId" -Scope Script
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 0 'a matching checksum: exit code 0'

    Write-Host '--- install.ps1 parses certutil.exe output by shape, not wording'
    # Localised wording and the hex printed in spaced pairs (older Windows).
    # The fake reads $fakeCertUtilOutput through PowerShell's dynamic scoping (install.ps1 does not define it).
    $fakeCertUtilOutput = @('Hachage SHA256 de Modelwright64.xll :', ($hash -replace '(..)', '$1 ').Trim(),
        "CertUtil: -hashfile La commande s'est terminee correctement.")
    Set-Item -Path 'function:certutil.exe' -Value { $fakeCertUtilOutput }
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 0 'localised, spaced certutil output: exit code 0'
    Set-Item -Path 'function:certutil.exe' -Value { 'CertUtil: unexpected output'; 'no hash here' }
    $before = Get-Snapshot $optionsKey
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 7 'certutil.exe output without a hash: exit code 7'
    Assert-True ($out -match 'Could not check Modelwright64\.xll against SHA256SUMS') 'says it could not check the file'
    Assert-Equal (Get-Snapshot $optionsKey) $before 'could not check: registry unchanged'
    Set-Alias -Name certutil.exe -Value "NoSuchCommand$runId" -Scope Script
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = $source }
    Assert-Equal $LASTEXITCODE 7 'certutil.exe missing: exit code 7'
    Remove-Item -LiteralPath 'alias:certutil.exe', 'function:certutil.exe', 'alias:Get-FileHash'
    Remove-Item -LiteralPath $sums

    Write-Host '--- install.ps1 with the add-in file missing'
    $out = Invoke-Script $install @{ ExcelBitness = '64'; SourceFolder = (Join-Path $scratch 'empty') }
    Assert-Equal $LASTEXITCODE 4 'exit code 4'

    # As on the CI runner (no Office): no Office keys under -RegistryRoot, and -ExcelExePath blocks the lookup of
    # this machine's EXCEL.EXE.
    Write-Host '--- install.ps1 bitness detection without Excel'
    $noOfficeRoot = Join-Path $regRoot 'NoOffice'
    $before = Get-Snapshot $optionsKey
    $out = Invoke-Script $install @{ SourceFolder = $source; RegistryRoot = $noOfficeRoot; ExcelExePath = (Join-Path $scratch 'NoOffice\EXCEL.EXE') }
    Assert-Equal $LASTEXITCODE 3 'EXCEL.EXE missing: exit code 3'
    Assert-True ($out -match 'Could not find or read EXCEL\.EXE') 'says it could not find EXCEL.EXE'
    Assert-True ($out -match 'Could not tell whether Excel is 32-bit or 64-bit') 'explains that it could not tell'
    Assert-True ($out -match '-ExcelBitness 64') 'explains how to pass -ExcelBitness'
    Assert-True (-not (Test-Path -LiteralPath $noOfficeRoot)) 'EXCEL.EXE missing: nothing written under -RegistryRoot'
    Assert-Equal (Get-Snapshot $optionsKey) $before 'EXCEL.EXE missing: registry unchanged'
    $out = Invoke-Script $install @{ SourceFolder = $source; RegistryRoot = $noOfficeRoot; ExcelExePath = (Join-Path $source 'Modelwright64.xll') }
    Assert-Equal $LASTEXITCODE 3 'a file that is not an .exe: exit code 3'

    Write-Host '--- install.ps1 bitness detection from an .exe header (PowerShell itself stands in for EXCEL.EXE)'
    if ($env:PROCESSOR_ARCHITECTURE -eq 'AMD64') {
        $out = Invoke-Script $install @{ SourceFolder = $source; RegistryRoot = $noOfficeRoot; WhatIf = $true; ExcelExePath = (Join-Path $PSHOME 'powershell.exe') }
        Assert-Equal $LASTEXITCODE 0 'a 64-bit .exe: exit code 0'
        Assert-True ($out -match 'Excel is 64-bit: installing Modelwright64\.xll') 'a 64-bit .exe: picks Modelwright64.xll'
        $wow64 = Join-Path $env:SystemRoot 'SysWOW64\WindowsPowerShell\v1.0\powershell.exe'
        if (Test-Path -LiteralPath $wow64) {
            $out = Invoke-Script $install @{ SourceFolder = $source; RegistryRoot = $noOfficeRoot; ExcelExePath = $wow64 }
            Assert-Equal $LASTEXITCODE 8 'a 32-bit .exe: exit code 8'
            Assert-True ($out -match 'Modelwright is 64-bit only; your Excel is 32-bit') 'a 32-bit .exe: says 32-bit Excel is not supported'
            Assert-True (-not (Test-Path -LiteralPath $noOfficeRoot)) 'a 32-bit .exe: nothing written under -RegistryRoot'
        }
    } else {
        Write-Host "SKIP  needs a 64-bit (AMD64) PowerShell, this is $env:PROCESSOR_ARCHITECTURE"
    }

    Write-Host '--- install.ps1 bitness detection on this machine (read-only registry and EXCEL.EXE reads)'
    $out = Invoke-Script $install @{ SourceFolder = $source; WhatIf = $true }
    Assert-True ($LASTEXITCODE -eq 0 -or $LASTEXITCODE -eq 3 -or $LASTEXITCODE -eq 8) "detection ends with 0 (64-bit), 8 (32-bit) or 3 (no Excel), not an error (exit $LASTEXITCODE)"
    if ($LASTEXITCODE -eq 0) {
        Assert-True ($out -match 'Excel is 64-bit') 'reports the detected bitness'
    } elseif ($LASTEXITCODE -eq 8) {
        Assert-True ($out -match 'Modelwright is 64-bit only') 'explains that 32-bit Excel is not supported'
    } else {
        Assert-True ($out -match 'Could not tell whether Excel is 32-bit or 64-bit') 'explains that it could not tell'
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
    Set-Content -LiteralPath $target32 -Value 'fake 32-bit add-in left by an earlier build'
    $out = Invoke-Script $uninstall @{ RemoveFile = $true }
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    Assert-Equal (Get-Snapshot $optionsKey) $expected 'renumbered OPEN, OPEN1 in the old order'
    Assert-True (-not (Test-Path -LiteralPath $target64)) '-RemoveFile deleted Modelwright64.xll'
    Assert-True (-not (Test-Path -LiteralPath $target32)) '-RemoveFile deleted a Modelwright32.xll left by an earlier build'

    Write-Host '--- uninstall.ps1 again has nothing to do'
    $out = Invoke-Script $uninstall @{}
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    Assert-True ($out -match 'nothing to remove') 'says there was nothing to remove'
    Assert-Equal (Get-Snapshot $optionsKey) $expected 'registry unchanged'

    # The 2026-10-09 bug: Excel had rewritten the entry as /R "Modelwright64.xll" (and kept the full path too);
    # only the full path was removed, the file was deleted, and Excel stalled on "Cannot find add-in".
    Write-Host '--- uninstall.ps1 removes bare and full-path entries together and renumbers'
    Reset-Options @{
        OPEN  = '"C:\Other\First.xlam"'
        OPEN1 = '/R "Modelwright64.xll"'
        OPEN2 = '"C:\Other\Second.xlam"'
        OPEN3 = $value64
        OPEN4 = '/R "modelwright32.XLL"'
        OPEN6 = '"C:\Other\Third.xlam"'
        OPEN7 = '/R "C:\Other\MyModelwright64.xll"'
    }
    Reset-Manager @('Modelwright64.xll', $target32, 'C:\Other\Unrelated.xll')
    $out = Invoke-Script $uninstall @{}
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    $expected = (@('OPEN="C:\Other\First.xlam"', 'OPEN1="C:\Other\Second.xlam"', 'OPEN2="C:\Other\Third.xlam"',
        'OPEN3=/R "C:\Other\MyModelwright64.xll"') | Sort-Object) -join "`n"
    Assert-Equal (Get-Snapshot $optionsKey) $expected 'all three Modelwright entries removed, the rest OPEN..OPEN3 in order'
    Assert-Equal (Get-Snapshot $managerKey) 'C:\Other\Unrelated.xll=' 'bare and full-path Add-in Manager entries removed, others kept'

    Write-Host '--- uninstall.ps1 removes a bare old ModelingToolkit entry'
    Reset-Options @{ OPEN = '/R "ModelingToolkit64-packed.xll"'; OPEN1 = '"C:\Other\First.xlam"' }
    Reset-Manager @('ModelingToolkit64-packed.xll', 'C:\Other\Unrelated.xll')
    $out = Invoke-Script $uninstall @{}
    Assert-Equal $LASTEXITCODE 0 'exit code 0'
    Assert-Equal (Get-Snapshot $optionsKey) 'OPEN="C:\Other\First.xlam"' 'old entry removed, the other moved to OPEN'
    Assert-Equal (Get-Snapshot $managerKey) 'C:\Other\Unrelated.xll=' 'bare old Add-in Manager entry removed'

    Write-Host '--- the real Excel Options and Add-in Manager keys'
    Assert-Equal (Get-Snapshot $realOptionsKey) $realBefore 'Options unchanged by the tests'
    Assert-Equal (Get-Snapshot $realManagerKey) $realManagerBefore 'Add-in Manager unchanged by the tests'
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
