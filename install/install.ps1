<#
.SYNOPSIS
    Installs the Modelwright Excel add-in for the current user. No administrator rights needed.

.DESCRIPTION
    1. Stops if Excel is running (Excel rewrites its add-in list when it closes, which would undo the install).
    2. Works out whether Excel is 32-bit or 64-bit, from the header of EXCEL.EXE (found through the registry),
       falling back to the Click-to-Run "Platform" setting. -ExcelBitness overrides the detection.
    3. Checks that .NET Framework 4.8 or later is installed.
    4. Copies Modelwright64.xll or Modelwright32.xll from this script's folder (or -SourceFolder) to
       %APPDATA%\Microsoft\AddIns, after checking it against SHA256SUMS.txt when that file is present, and
       unblocks the copy (removes the "downloaded from the internet" mark that makes Excel refuse it).
    5. Registers it in HKCU\Software\Microsoft\Office\16.0\Excel\Options as the next free OPEN/OPENn value
       (/R "<path>"), the same thing Excel's own Add-ins dialog does. An existing Modelwright entry (or an old
       ModelingToolkit build) is replaced in place; if the right entry is already there, nothing is changed.
    Your Modelwright settings (%APPDATA%\Modelwright) are never touched.

    Works in Windows PowerShell 5.1 and in Constrained Language mode (only built-in cmdlets are used).

    Exit codes: 0 installed (or already installed); 1 unexpected error; 2 Excel is running;
    3 Excel's bitness could not be worked out or is not supported; 4 the add-in file is missing;
    5 .NET Framework 4.8 is missing; 6 the add-in file does not match SHA256SUMS.txt.

.PARAMETER ExcelBitness
    32 or 64: skip the detection and install that add-in.

.PARAMETER OfficeVersion
    The Office registry version. 16.0 covers Office 2016 and every later version, including Microsoft 365.

.PARAMETER SourceFolder
    The folder holding Modelwright64.xll / Modelwright32.xll. Default: this script's folder.

.PARAMETER AddInsFolder
    Where the add-in is copied. Default: %APPDATA%\Microsoft\AddIns (Excel's own per-user add-ins folder).

.PARAMETER RegistryRoot
    The registry key that holds <OfficeVersion>\Excel. Default: HKCU:\Software\Microsoft\Office. For testing.

.PARAMETER ExcelProcessName
    The process that must not be running. Default: EXCEL. For testing.

.PARAMETER NoPause
    Do not wait for Enter before closing (the default waits, so a "Run with PowerShell" window stays open).

.EXAMPLE
    Right-click install.ps1 > Run with PowerShell.

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\install.ps1 -WhatIf -Verbose
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateSet('32', '64')]
    [string]$ExcelBitness,

    [ValidatePattern('^\d+\.0$')]
    [string]$OfficeVersion = '16.0',

    [string]$SourceFolder,

    [string]$AddInsFolder,

    [string]$RegistryRoot = 'HKCU:\Software\Microsoft\Office',

    [string]$ExcelProcessName = 'EXCEL',

    [switch]$NoPause
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# $PSScriptRoot can be empty in Windows PowerShell 5.1 with some launch styles.
$scriptDir = if ($PSScriptRoot) { $PSScriptRoot } else { Split-Path -Parent $MyInvocation.MyCommand.Definition }
if (-not $SourceFolder) { $SourceFolder = $scriptDir }
if (-not $AddInsFolder) { $AddInsFolder = Join-Path $env:APPDATA 'Microsoft\AddIns' }

# Our entries: Modelwright32.xll / Modelwright64.xll, and the builds from before the rename (ModelingToolkit*.xll).
$ourAddInPattern = '\\(Modelwright(32|64)|ModelingToolkit[^\\"]*)\.xll"?\s*$'

# -WhatIf: $PSCmdlet.ShouldProcess is a method call, which Constrained Language mode blocks, so it is done here.
function Test-ShouldChange([string]$Target, [string]$Action) {
    if ($WhatIfPreference) {
        Write-Host "What if: $Action on `"$Target`"."
        return $false
    }
    return $true
}

# SHA-256 of a file as hex. Get-FileHash is script code in Windows PowerShell 5.1 and is unavailable when the
# session is constrained by hand (not by AppLocker), so certutil.exe (part of Windows) is the fallback.
function Get-Sha256([string]$Path) {
    if (Get-Command Get-FileHash -ErrorAction SilentlyContinue) {
        return (Get-FileHash -LiteralPath $Path -Algorithm SHA256).Hash
    }
    $lines = @(& certutil.exe -hashfile $Path SHA256)
    if ($LASTEXITCODE -ne 0 -or $lines.Count -lt 2) { throw "certutil could not hash $Path" }
    return ($lines[1] -replace '\s', '')
}

function Exit-Script([int]$Code) {
    if (-not $NoPause) {
        try { $null = Read-Host 'Press Enter to close' } catch { }
    }
    exit $Code
}

function Stop-WithError([int]$Code, [string]$Message) {
    Write-Host ''
    Write-Host $Message -ForegroundColor Red
    Exit-Script $Code
}

# Returns 32 or 64 from EXCEL.EXE's PE header (Machine field), or $null if the file cannot be read.
function Get-ExeBitness([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $null }
    $read = @{ LiteralPath = $Path; TotalCount = 4096 }
    if ($PSVersionTable.PSVersion.Major -ge 6) { $read.AsByteStream = $true } else { $read.Encoding = 'Byte' }
    $bytes = @(Get-Content @read)
    if ($bytes.Count -lt 64 -or $bytes[0] -ne 0x4D -or $bytes[1] -ne 0x5A) { return $null } # 'MZ'
    $peOffset = [int]$bytes[0x3C] + [int]$bytes[0x3D] * 256 + [int]$bytes[0x3E] * 65536 + [int]$bytes[0x3F] * 16777216
    if ($peOffset + 6 -gt $bytes.Count) { return $null }
    if ($bytes[$peOffset] -ne 0x50 -or $bytes[$peOffset + 1] -ne 0x45 -or $bytes[$peOffset + 2] -ne 0 -or $bytes[$peOffset + 3] -ne 0) {
        return $null # 'PE\0\0'
    }
    $machine = [int]$bytes[$peOffset + 4] + [int]$bytes[$peOffset + 5] * 256
    Write-Verbose ('EXCEL.EXE machine type: 0x{0:X4}' -f $machine)
    switch ($machine) {
        0x8664 { return '64' }  # x64
        0xAA64 {                # Arm64: 64-bit Excel on Arm loads x64 add-ins (untested)
            Write-Warning 'Excel for Arm detected. Modelwright installs its 64-bit add-in, which has not been tested on Arm.'
            return '64'
        }
        0x014C { return '32' }  # x86
        default { return $null }
    }
}

function Get-RegistryString([string]$Key, [string]$Name) {
    if (-not (Test-Path -LiteralPath $Key)) { return $null }
    $item = Get-ItemProperty -LiteralPath $Key -Name $Name -ErrorAction SilentlyContinue
    if ($null -eq $item) { return $null }
    return [string]$item.$Name
}

function Get-ExcelBitness {
    $candidates = @()
    foreach ($appPaths in @('HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe',
                            'HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\App Paths\excel.exe',
                            'HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe')) {
        $path = Get-RegistryString $appPaths '(default)'
        if ($path) { $candidates += $path.Trim().Trim('"') }
    }
    $clickToRun = 'HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration'
    $installPath = Get-RegistryString $clickToRun 'InstallationPath'
    if ($installPath) { $candidates += Join-Path $installPath 'root\Office16\EXCEL.EXE' }

    foreach ($exe in $candidates) {
        Write-Verbose "Checking $exe"
        $bitness = Get-ExeBitness $exe
        if ($bitness) {
            Write-Verbose "Excel is $bitness-bit (from the header of $exe)"
            return $bitness
        }
    }

    switch (Get-RegistryString $clickToRun 'Platform') {
        'x64' { Write-Verbose 'Excel is 64-bit (from the Click-to-Run Platform setting)'; return '64' }
        'x86' { Write-Verbose 'Excel is 32-bit (from the Click-to-Run Platform setting)'; return '32' }
    }
    return $null
}

# The OPEN, OPEN1, OPEN2... values of the Options key, sorted by number.
function Get-OpenEntries([string]$Key) {
    $entries = @()
    if (-not (Test-Path -LiteralPath $Key)) { return $entries }
    foreach ($name in @((Get-Item -LiteralPath $Key).Property)) {
        if ($name -match '^OPEN(\d*)$') {
            $index = if ($Matches[1]) { [int]$Matches[1] } else { 0 }
            $entries += @{ Name = $name; Index = $index; Value = [string](Get-ItemPropertyValue -LiteralPath $Key -Name $name) }
        }
    }
    return @($entries | Sort-Object { $_.Index })
}

# Rewrites the OPEN values as OPEN, OPEN1, ... (no gaps) holding $Values in order. Unchanged values are left alone.
function Write-OpenEntries([string]$Key, [object[]]$Existing, [string[]]$Values) {
    $current = @{}
    foreach ($entry in $Existing) { $current[$entry.Name.ToUpperInvariant()] = $entry }
    $wanted = @{}
    for ($i = 0; $i -lt $Values.Count; $i++) {
        $name = if ($i -eq 0) { 'OPEN' } else { "OPEN$i" }
        $wanted[$name] = $true
        $old = $current[$name]
        if ($null -eq $old -or $old.Value -cne $Values[$i]) {
            Write-Verbose "Setting $name = $($Values[$i])"
            if ($null -ne $old -and $old.Name -cne $name) { Remove-ItemProperty -LiteralPath $Key -Name $old.Name }
            $null = New-ItemProperty -LiteralPath $Key -Name $name -Value $Values[$i] -PropertyType String -Force
        }
    }
    foreach ($entry in $Existing) {
        if (-not $wanted[$entry.Name.ToUpperInvariant()]) {
            Write-Verbose "Removing $($entry.Name) (was $($entry.Value))"
            Remove-ItemProperty -LiteralPath $Key -Name $entry.Name
        }
    }
}

try {
    $excelKey = Join-Path $RegistryRoot "$OfficeVersion\Excel"
    $optionsKey = Join-Path $excelKey 'Options'
    $managerKey = Join-Path $excelKey 'Add-in Manager'

    # 1. Excel must be closed.
    if (@(Get-Process -Name $ExcelProcessName -ErrorAction SilentlyContinue).Count -gt 0) {
        Stop-WithError 2 ('Excel is running. Close Excel completely (check the taskbar, and Task Manager for a ' +
            'background EXCEL.EXE), then run this again.')
    }

    # 2. Bitness.
    if ($ExcelBitness) {
        Write-Verbose "Excel bitness given: $ExcelBitness-bit"
    } else {
        $ExcelBitness = Get-ExcelBitness
        if (-not $ExcelBitness) {
            Stop-WithError 3 ('Could not tell whether Excel is 32-bit or 64-bit. In Excel, open File > Account > ' +
                'About Excel; the first line ends in "32-bit" or "64-bit". Then run this again with ' +
                '-ExcelBitness 64 (or 32).')
        }
    }
    $xllName = "Modelwright$ExcelBitness.xll"
    Write-Host "Excel is $ExcelBitness-bit: installing $xllName."

    # 3. .NET Framework 4.8 (Release 528040 or later) is part of Windows 10 1903+ and every Windows 11.
    $netRelease = Get-RegistryString 'HKLM:\SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full' 'Release'
    Write-Verbose ".NET Framework 4.x release: $netRelease"
    if (-not $netRelease -or [int]$netRelease -lt 528040) {
        Stop-WithError 5 ('Modelwright needs .NET Framework 4.8 or later, which is not installed. Install it from ' +
            'https://dotnet.microsoft.com/download/dotnet-framework (or ask IT), then run this again.')
    }

    # 4. Copy and unblock.
    $source = Join-Path $SourceFolder $xllName
    if (-not (Test-Path -LiteralPath $source -PathType Leaf)) {
        Stop-WithError 4 "$xllName is not in $SourceFolder. Extract the whole Modelwright zip, then run install.ps1 from the extracted folder."
    }
    $sums = Join-Path $SourceFolder 'SHA256SUMS.txt'
    if (Test-Path -LiteralPath $sums -PathType Leaf) {
        $expected = $null
        foreach ($line in @(Get-Content -LiteralPath $sums)) {
            if ($line -match '^([0-9A-Fa-f]{64})\s+\*?(.+?)\s*$' -and $Matches[2] -eq $xllName) { $expected = $Matches[1] }
        }
        if ($expected) {
            $actual = Get-Sha256 $source
            if ($actual -ne $expected) {
                Stop-WithError 6 "$xllName does not match SHA256SUMS.txt (it may be damaged or altered). Download the release again."
            }
            Write-Verbose "$xllName matches SHA256SUMS.txt"
        }
    }

    $target = Join-Path $AddInsFolder $xllName
    if (-not (Test-Path -LiteralPath $AddInsFolder -PathType Container)) {
        if (Test-ShouldChange $AddInsFolder 'Create folder') {
            $null = New-Item -ItemType Directory -Path $AddInsFolder
        }
    }
    if (Test-ShouldChange $target "Copy $xllName and unblock it") {
        Copy-Item -LiteralPath $source -Destination $target -Force
        Unblock-File -LiteralPath $target
        Write-Host "Copied $xllName to $AddInsFolder"
    }

    # 5. Register: replace our old entry in place, or add the next OPENn.
    $value = "/R `"$target`""
    $entries = @(Get-OpenEntries $optionsKey)
    $ours = @($entries | Where-Object { $_.Value -match $ourAddInPattern })
    if ($ours.Count -eq 1 -and $ours[0].Value -eq $value) {
        Write-Host "Already registered with Excel ($($ours[0].Name)); left as it is."
    } else {
        $values = @()
        $slot = -1
        foreach ($entry in $entries) {
            if ($entry.Value -match $ourAddInPattern) {
                if ($slot -lt 0) { $slot = $values.Count; $values += $value }
                Write-Verbose "Replacing $($entry.Name) = $($entry.Value)"
            } else {
                $values += $entry.Value
            }
        }
        if ($slot -lt 0) { $slot = $values.Count; $values += $value }
        $slotName = if ($slot -eq 0) { 'OPEN' } else { "OPEN$slot" }
        if (Test-ShouldChange "$optionsKey\$slotName" "Register $value") {
            if (-not (Test-Path -LiteralPath $optionsKey)) {
                # The key does not exist, so -Force (which creates missing parents) cannot wipe anything.
                $null = New-Item -Path $optionsKey -Force
            }
            Write-OpenEntries $optionsKey $entries $values
            Write-Host "Registered with Excel as $slotName."
        }
    }

    # Excel's list of add-ins that are known but not ticked: drop stale Modelwright entries.
    if (Test-Path -LiteralPath $managerKey) {
        foreach ($name in @((Get-Item -LiteralPath $managerKey).Property)) {
            if ($name -match $ourAddInPattern -and (Test-ShouldChange "$managerKey\$name" 'Remove stale entry')) {
                Remove-ItemProperty -LiteralPath $managerKey -Name $name
                Write-Verbose "Removed Add-in Manager entry $name"
            }
        }
    }

    Write-Host ''
    if ($WhatIfPreference) {
        Write-Host 'WhatIf: nothing was changed.'
    } else {
        Write-Host 'Done. Start Excel: the Modelwright tab appears on the ribbon.' -ForegroundColor Green
    }
    Exit-Script 0
} catch {
    Stop-WithError 1 "Install failed: $($_.Exception.Message)"
}
