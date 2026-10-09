<#
.SYNOPSIS
    Removes the Modelwright Excel add-in for the current user. No administrator rights needed.

.DESCRIPTION
    1. Stops if Excel is running (Excel rewrites its add-in list when it closes, which would undo this).
    2. Removes the Modelwright entry (Modelwright32.xll / Modelwright64.xll, or an old ModelingToolkit build)
       from the OPEN/OPENn values in HKCU\Software\Microsoft\Office\16.0\Excel\Options, renumbering the other
       add-ins' entries as OPEN, OPEN1, ... with no gaps, and from Excel's list of known add-ins (Add-in Manager).
    3. With -RemoveFile, also deletes Modelwright32.xll / Modelwright64.xll from %APPDATA%\Microsoft\AddIns.
    Your Modelwright settings (%APPDATA%\Modelwright) are never touched; delete that folder yourself if you want.

    Works in Windows PowerShell 5.1 and in Constrained Language mode (only built-in cmdlets are used).

    Exit codes: 0 removed (or nothing to remove); 1 unexpected error; 2 Excel is running.

.PARAMETER RemoveFile
    Also delete the add-in files from the add-ins folder.

.PARAMETER OfficeVersion
    The Office registry version. 16.0 covers Office 2016 and every later version, including Microsoft 365.

.PARAMETER AddInsFolder
    The folder install.ps1 copied the add-in to. Default: %APPDATA%\Microsoft\AddIns.

.PARAMETER RegistryRoot
    The registry key that holds <OfficeVersion>\Excel. Default: HKCU:\Software\Microsoft\Office. For testing.

.PARAMETER ExcelProcessName
    The process that must not be running. Default: EXCEL. For testing.

.PARAMETER NoPause
    Do not wait for Enter before closing (the default waits, so a "Run with PowerShell" window stays open).

.EXAMPLE
    powershell -NoProfile -ExecutionPolicy Bypass -File .\uninstall.ps1 -RemoveFile
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [switch]$RemoveFile,

    [ValidatePattern('^\d+\.0$')]
    [string]$OfficeVersion = '16.0',

    [string]$AddInsFolder,

    [string]$RegistryRoot = 'HKCU:\Software\Microsoft\Office',

    [string]$ExcelProcessName = 'EXCEL',

    [switch]$NoPause
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

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

try {
    $excelKey = Join-Path $RegistryRoot "$OfficeVersion\Excel"
    $optionsKey = Join-Path $excelKey 'Options'
    $managerKey = Join-Path $excelKey 'Add-in Manager'

    if (@(Get-Process -Name $ExcelProcessName -ErrorAction SilentlyContinue).Count -gt 0) {
        Stop-WithError 2 ('Excel is running. Close Excel completely (check the taskbar, and Task Manager for a ' +
            'background EXCEL.EXE), then run this again.')
    }

    $changed = $false

    # The OPEN values, sorted by number; ours are dropped and the rest are rewritten as OPEN, OPEN1, ... in order.
    if (Test-Path -LiteralPath $optionsKey) {
        $entries = @()
        foreach ($name in @((Get-Item -LiteralPath $optionsKey).Property)) {
            if ($name -match '^OPEN(\d*)$') {
                $index = if ($Matches[1]) { [int]$Matches[1] } else { 0 }
                $entries += @{ Name = $name; Index = $index; Value = [string](Get-ItemPropertyValue -LiteralPath $optionsKey -Name $name) }
            }
        }
        $entries = @($entries | Sort-Object { $_.Index })
        $ours = @($entries | Where-Object { $_.Value -match $ourAddInPattern })
        $keep = @($entries | Where-Object { $_.Value -notmatch $ourAddInPattern })
        $what = 'Unregister ' + (@($ours | ForEach-Object { "$($_.Name) = $($_.Value)" }) -join '; ') + ' and renumber the rest'
        if ($ours.Count -gt 0 -and (Test-ShouldChange $optionsKey $what)) {
            # Rewrite without gaps: slot i gets the i-th kept value; values that are already right are left alone.
            $current = @{}
            foreach ($entry in $entries) { $current[$entry.Name.ToUpperInvariant()] = $entry }
            $wanted = @{}
            for ($i = 0; $i -lt $keep.Count; $i++) {
                $name = if ($i -eq 0) { 'OPEN' } else { "OPEN$i" }
                $wanted[$name] = $true
                $old = $current[$name]
                if ($null -eq $old -or $old.Value -cne $keep[$i].Value) {
                    Write-Verbose "Setting $name = $($keep[$i].Value)"
                    if ($null -ne $old -and $old.Name -cne $name) { Remove-ItemProperty -LiteralPath $optionsKey -Name $old.Name }
                    $null = New-ItemProperty -LiteralPath $optionsKey -Name $name -Value $keep[$i].Value -PropertyType String -Force
                }
            }
            foreach ($entry in $entries) {
                if (-not $wanted[$entry.Name.ToUpperInvariant()]) {
                    Write-Verbose "Removing $($entry.Name)"
                    Remove-ItemProperty -LiteralPath $optionsKey -Name $entry.Name
                }
            }
            foreach ($entry in $ours) { Write-Host "Unregistered $($entry.Name): $($entry.Value)" }
            $changed = $true
        }
    }

    # Excel's list of add-ins that are known but not ticked.
    if (Test-Path -LiteralPath $managerKey) {
        foreach ($name in @((Get-Item -LiteralPath $managerKey).Property)) {
            if ($name -match $ourAddInPattern -and (Test-ShouldChange "$managerKey\$name" 'Remove entry')) {
                Remove-ItemProperty -LiteralPath $managerKey -Name $name
                Write-Host "Removed $name from Excel's add-in list."
                $changed = $true
            }
        }
    }

    if ($RemoveFile) {
        foreach ($xllName in @('Modelwright64.xll', 'Modelwright32.xll')) {
            $path = Join-Path $AddInsFolder $xllName
            if ((Test-Path -LiteralPath $path -PathType Leaf) -and (Test-ShouldChange $path 'Delete file')) {
                Remove-Item -LiteralPath $path
                Write-Host "Deleted $path"
                $changed = $true
            }
        }
    }

    Write-Host ''
    if ($WhatIfPreference) {
        Write-Host 'WhatIf: nothing was changed.'
    } elseif ($changed) {
        Write-Host 'Done. Modelwright will not load the next time Excel starts. Your settings in %APPDATA%\Modelwright were kept.' -ForegroundColor Green
    } else {
        Write-Host 'Modelwright was not registered with Excel; nothing to remove.'
    }
    Exit-Script 0
} catch {
    Stop-WithError 1 "Uninstall failed: $($_.Exception.Message)"
}
