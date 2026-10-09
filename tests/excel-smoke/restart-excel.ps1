# Restarts Excel so a rebuilt add-in can load: Excel keys its add-in list by file name and keeps a loaded .xll locked
# for its lifetime, so a new build at any path is never loaded until Excel restarts. Refuses if any workbook has
# unsaved changes (nothing is ever saved here). Windows PowerShell 5.1.
#   powershell -ExecutionPolicy Bypass -File tests/excel-smoke/restart-excel.ps1 [-Stop | -Start] [-ExitWaitSeconds 15] [-KillIfStuck]
# -Stop only closes Excel and -Start only starts it: build between the two, since Excel loads the installed
# add-in at startup and locks it before a build could replace it. Without a switch it does both.
# Safety: only the Excel instance COM hands out (its process taken from Application.Hwnd) is closed, with Quit; if that
# process has not exited within -ExitWaitSeconds the script stops with a message and kills nothing. Any other Excel
# process is left alone (its unsaved work too).
param([switch]$Stop, [switch]$Start, [int]$ExitWaitSeconds = 15, [switch]$KillIfStuck)
$ErrorActionPreference = 'Stop'
Add-Type @"
using System; using System.Runtime.InteropServices;
public static class RestartExcelNative {
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  public static uint ProcessOf(IntPtr h) { uint p; GetWindowThreadProcessId(h, out p); return p; }
}
"@
if (-not $Stop -and -not $Start) { $Stop = $true; $Start = $true }

# The process of an Excel Application object (its main window's), or 0.
function ProcessOfExcel($app) {
  try { return [int][RestartExcelNative]::ProcessOf([IntPtr]([int64]$app.Hwnd)) } catch { return 0 }
}

$xl = $null
try { $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application') } catch { }
if ($Stop -and $xl) {
  $excelPid = ProcessOfExcel $xl
  if ($excelPid -eq 0) { throw "ABORT: could not find the process of the running Excel (Application.Hwnd). Excel was not closed." }
  $unsaved = @($xl.Workbooks | Where-Object { -not $_.Saved } | ForEach-Object { $_.Name })
  if ($unsaved.Count -gt 0) { throw "ABORT: unsaved changes in: $($unsaved -join ', ') (Excel process $excelPid). Excel was not closed." }
  $others = @(Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.Id -ne $excelPid } | ForEach-Object { $_.Id })
  if ($others.Count -gt 0) { "Other Excel processes are left alone: $($others -join ', ')" }
  $process = Get-Process -Id $excelPid
  "Closing Excel process $excelPid (open, all saved: $(@($xl.Workbooks | ForEach-Object { $_.Name }) -join ', '))"
  $xl.DisplayAlerts = $false
  $xl.Quit()
  [Runtime.InteropServices.Marshal]::ReleaseComObject($xl) | Out-Null
  $xl = $null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  if (-not $process.WaitForExit($ExitWaitSeconds * 1000)) {
    if ($KillIfStuck) {
      # Only this instance, and only after every workbook in it was checked as saved.
      "Excel process $excelPid did not exit within $ExitWaitSeconds s of Quit; stopping that process (-KillIfStuck; all its workbooks were saved)"
      Stop-Process -Id $excelPid -Force
      $process.WaitForExit(10000) | Out-Null
    }
    else {
      throw "ABORT: Excel process $excelPid did not exit within $ExitWaitSeconds s of Quit (a dialog may be open, or something still holds it). It was not killed: close it yourself, then run with -Start, or run again with -KillIfStuck."
    }
  }
  "Excel process $excelPid exited."
}
elseif ($Stop) { "No running Excel to close (none is registered for COM)." }
if (-not $Start) { "Excel is closed."; return }

$running = @(Get-Process EXCEL -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
if ($running.Count -gt 0) { "Note: Excel processes already running: $($running -join ', '); excel.exe may open its window in one of them." }
$started = Start-Process excel.exe -PassThru
"Started excel.exe: process $($started.Id)"
$xl = $null
for ($i = 0; $i -lt 60; $i++) {
  Start-Sleep -Milliseconds 500
  if ($started.HasExited) { throw "ABORT: excel.exe (process $($started.Id)) exited at once: it handed over to a running Excel ($($running -join ', ')), which this script does not attach to." }
  try {
    $candidate = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
    if ((ProcessOfExcel $candidate) -eq $started.Id -and $candidate.Ready) { $xl = $candidate; break }
  } catch { }
}
if (-not $xl) { throw "ABORT: the new Excel (process $($started.Id)) did not become the running Excel COM hands out within 30 s." }
"Excel restarted: version $($xl.Version) build $($xl.Build) pid $($started.Id); workbooks: $(@($xl.Workbooks | ForEach-Object { $_.Name }) -join ', ')"
