# Excel smoke test: formatting cycle + undo ordering, driven by REAL keystrokes into the running Excel.
# Run with Windows PowerShell 5.1 (needs Marshal.GetActiveObject), with Excel open:
#   powershell -ExecutionPolicy Bypass -File tests/excel-smoke/undo-smoke.ps1 [-Xll <path to packed xll>] [-Bitness 64|32]
# -Bitness picks the default -Xll (Modelwright64.xll or Modelwright32.xll): the one matching Excel's bitness.
# Safety: works only in a new scratch workbook (closed without saving at the end); checks that Excel is the
# foreground window before EVERY keystroke and aborts otherwise. Don't touch the keyboard while it runs.
# The focus trick taps Shift: an Alt tap would turn on ribbon KeyTips and send the next key to the ribbon.
# Note: it (re)installs the given xll in Excel's add-in list, replacing any other Modelwright64.xll or
# Modelwright32.xll (or pre-rename ModelingToolkit64-packed.xll).
param([string]$Xll = '', [ValidateSet('64', '32')][string]$Bitness = '64')
$ErrorActionPreference = 'Stop'
if (-not $Xll) { $Xll = Join-Path $PSScriptRoot "..\..\src\Modelwright.AddIn\bin\Release\net48\publish\Modelwright$Bitness.xll" }
Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System; using System.Runtime.InteropServices;
public static class W {
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr h, int c);
  [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
}
"@

$fixXll = (Resolve-Path $Xll).Path

# Our add-in in Excel's list: this build's name, or the name builds had before the rename to Modelwright (D12), so an
# old build left installed is swapped out as well (both would claim the same shortcuts).
function Is-OurXll([string]$fullName) { return ($fullName -like '*Modelwright64.xll') -or ($fullName -like '*Modelwright32.xll') -or ($fullName -like '*ModelingToolkit64-packed.xll') }
$log    = Join-Path $env:LOCALAPPDATA 'Modelwright\log.txt'

function Retry([scriptblock]$b) {
  for ($i = 0; $i -lt 40; $i++) { try { return & $b } catch { Start-Sleep -Milliseconds 150 } }
  throw "COM call kept failing"
}

$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
# This Excel's process (from Application.Hwnd), not just any EXCEL process: another instance may be running.
[uint32]$excelPid = 0; [W]::GetWindowThreadProcessId([IntPtr]([int64]$xl.Hwnd), [ref]$excelPid) | Out-Null
"Excel version $($xl.Version) build $($xl.Build)"

# --- Swap add-in builds (same as Add-ins dialog) ---
foreach ($a in @($xl.AddIns)) {
  if ((Is-OurXll $a.FullName) -and $a.FullName -ne $fixXll -and $a.Installed) {
    "Unloading $($a.FullName)"; $a.Installed = $false
  }
}
# --- Scratch workbook (AddIns.Add needs an open workbook) ---
$wb = Retry { $xl.Workbooks.Add() }
$added = $xl.AddIns.Add($fixXll)
if (-not $added.Installed) { $added.Installed = $true }
Start-Sleep -Milliseconds 1500
"Loaded: $($added.FullName) installed=$($added.Installed)"
$ws = $wb.Worksheets.Item(1)
$ws.Range('A1').Value2 = 1234.5; $ws.Range('A2').Value2 = -987; $ws.Range('A3').Value2 = 0
$ws.Range('A1:A3').Select() | Out-Null

function Focus-Excel {
  $h = [IntPtr]([int64]$xl.ActiveWindow.Hwnd)
  [W]::keybd_event(0x10, 0, 0, [UIntPtr]::Zero); [W]::keybd_event(0x10, 0, 2, [UIntPtr]::Zero)   # Shift tap (harmless) releases foreground lock; Alt would trigger ribbon KeyTips
  [W]::SetForegroundWindow($h) | Out-Null
  Start-Sleep -Milliseconds 300
}
function Assert-ExcelFront {
  $fg = [W]::GetForegroundWindow(); $p = 0; [W]::GetWindowThreadProcessId($fg, [ref]$p) | Out-Null
  if ($p -ne $excelPid) { throw "ABORT: foreground window belongs to pid $p, not Excel ($excelPid). No keys sent." }
}
function Send([string]$keys, [string]$label) {
  Assert-ExcelFront
  [System.Windows.Forms.SendKeys]::SendWait($keys)
  Start-Sleep -Milliseconds 700
  $f = Retry { $ws.Range('A1').NumberFormat }
  $d = Retry { $ws.Range('D1').Value2 }
  $sb = Retry { $xl.StatusBar }
  "{0,-28} A1 fmt=[{1}]  D1=[{2}]  status=[{3}]" -f $label, $f, $d, $sb
}

Focus-Excel
"Start                        A1 fmt=[$($ws.Range('A1').NumberFormat)]"
Send '^+1' 'Ctrl+Shift+1 #1'
Send '^+1' 'Ctrl+Shift+1 #2'
Retry { $ws.Range('D1').Select() } | Out-Null
Focus-Excel
Send '5~' 'type 5 in D1 + Enter'
Send '^z' 'Ctrl+Z #1 (expect 5 gone)'
Send '^z' 'Ctrl+Z #2 (expect A1 item 1)'
Send '^z' 'Ctrl+Z #3 (expect General)'

$finalFmt = Retry { $ws.Range('A1').NumberFormat }
$finalD1  = Retry { $ws.Range('D1').Value2 }
$pass = ($finalFmt -eq 'General') -and ($null -eq $finalD1)
if ($pass) { "RESULT: PASS (A1 back to General, D1 empty)" } else { "RESULT: FAIL (A1 fmt=[$finalFmt], D1=[$finalD1])" }

"--- log lines since start ---"
Get-Content $log | Select-Object -Last 14
foreach ($n in @($wb.Name)) { try { $xl.Workbooks.Item($n).Close($false); "Closed scratch $n without saving" } catch {} }
if (-not $pass) { exit 1 }
