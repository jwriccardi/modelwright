# Excel smoke test: Trace In / Last Audited Cell, driven by REAL keystrokes into the running Excel.
# Run with Windows PowerShell 5.1 (needs Marshal.GetActiveObject), with Excel open:
#   powershell -ExecutionPolicy Bypass -File tests/excel-smoke/trace-smoke.ps1 [-Xll <path to packed xll>]
# It (re)builds the fixture first (build-trace-fixture.ps1: %TEMP%\emt-trace-fixture), opens TraceFixture.xlsx and
# checks through COM, after every key, which workbook, sheet and cell is active, and whether the Trace In window is
# open. Covered: cross-sheet and same-sheet navigation, names, a hidden sheet (Goto refused, no error), F2 passing
# through to Excel with the window open, expanding into the closed external workbook (Trace In opens it), Left back
# out of it, Enter (stay), Esc (back to the audited cell), Last Audited Cell through four levels, and native undo
# surviving a trace that stays in the workbook (type in B1, trace + navigate + Esc, Ctrl+Z: B1 is empty again).
# Timings (open, and each Up/Down step) are read back from the diagnostics log and reported against the targets
# (300 ms, 100 ms); they are reported, not failed on.
# Safety: works only in the fixture workbooks (closed without saving at the end); checks that Excel is the
# foreground window before EVERY keystroke and aborts otherwise. Don't touch the keyboard while it runs.
# The focus trick taps Shift: an Alt tap would turn on ribbon KeyTips and send the next key to the ribbon.
# Note: it (re)installs the given xll in Excel's add-in list, replacing any other ModelingToolkit64-packed.xll.
param(
    [string]$Xll = (Join-Path $PSScriptRoot '..\..\src\ExcelModelingToolkit.AddIn\bin\Release\net48\publish\ModelingToolkit64-packed.xll'),
    [string]$FixtureDir = (Join-Path $env:TEMP 'emt-trace-fixture')
)
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System; using System.Runtime.InteropServices; using System.Text;
public static class W {
  [DllImport("user32.dll")] public static extern IntPtr GetForegroundWindow();
  [DllImport("user32.dll")] public static extern bool SetForegroundWindow(IntPtr h);
  [DllImport("user32.dll")] public static extern void keybd_event(byte vk, byte scan, uint flags, UIntPtr extra);
  [DllImport("user32.dll")] public static extern uint GetWindowThreadProcessId(IntPtr h, out uint pid);
  [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h);
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetWindowText(IntPtr h, StringBuilder s, int n);
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc f, IntPtr l);
  // True if process pid has a visible top-level window titled exactly title.
  public static bool HasWindow(uint pid, string title) {
    bool found = false;
    EnumWindows((h, l) => {
      uint p; GetWindowThreadProcessId(h, out p);
      if (p == pid && IsWindowVisible(h)) {
        var s = new StringBuilder(256); GetWindowText(h, s, 256);
        if (s.ToString() == title) { found = true; return false; }
      }
      return true;
    }, IntPtr.Zero);
    return found;
  }
}
"@

$fixXll = (Resolve-Path $Xll).Path
$log    = Join-Path $env:LOCALAPPDATA 'ModelingToolkit\log.txt'
$logStart = if (Test-Path $log) { (Get-Item $log).Length } else { 0 }

function Retry([scriptblock]$b) {
  for ($i = 0; $i -lt 40; $i++) { try { return & $b } catch { Start-Sleep -Milliseconds 150 } }
  throw "COM call kept failing"
}

$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$excelPid = (Get-Process EXCEL | Select-Object -First 1).Id
"Excel version $($xl.Version) build $($xl.Build)"

# --- Fixture: rebuilt, so the external workbook is closed and the trace must open it ---
& (Join-Path $PSScriptRoot 'build-trace-fixture.ps1') -Excel $xl -OutDir $FixtureDir
$mainPath = Join-Path $FixtureDir 'TraceFixture.xlsx'
$extName = 'TraceExternal.xlsx'
foreach ($open in @($xl.Workbooks)) { if ($open.Name -eq $extName) { throw "ABORT: $extName is open; close it first." } }

$alerts = $xl.DisplayAlerts
$xl.DisplayAlerts = $false
$wb = Retry { $xl.Workbooks.Open($mainPath, 0) }   # UpdateLinks 0: no prompt about the external link
$xl.DisplayAlerts = $alerts
$iterationBefore = $xl.Iteration
$xl.Iteration = $true                             # the fixture's circular reference, without Excel's warning
$mainName = $wb.Name

# --- Swap add-in builds (same as Add-ins dialog) ---
foreach ($a in @($xl.AddIns)) {
  if ($a.FullName -like '*ModelingToolkit64-packed.xll' -and $a.FullName -ne $fixXll -and $a.Installed) {
    "Unloading $($a.FullName)"; $a.Installed = $false
  }
}
$added = $xl.AddIns.Add($fixXll)
if (-not $added.Installed) { $added.Installed = $true }
Start-Sleep -Milliseconds 1500
"Loaded: $($added.FullName) installed=$($added.Installed)"
$calc = $wb.Worksheets.Item('Calc')
Retry { $calc.Activate() } | Out-Null

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
function Select-Cell([string]$address) {
  Retry { $calc.Activate() } | Out-Null
  Retry { $calc.Range($address).Select() } | Out-Null
  Focus-Excel
}
function Window-Open { [W]::HasWindow([uint32]$excelPid, 'Trace In') }

$failures = New-Object System.Collections.Generic.List[string]
# Sends keys, waits, then checks the active workbook/sheet/cell and (if given) whether the window is open.
function Step([string]$keys, [string]$label, [string]$book, [string]$sheet, [string]$cell, $window = $null, [int]$waitMs = 700) {
  Assert-ExcelFront
  [System.Windows.Forms.SendKeys]::SendWait($keys)
  $deadline = (Get-Date).AddMilliseconds([Math]::Max($waitMs, 700))
  do {
    Start-Sleep -Milliseconds 200
    $b = Retry { $xl.ActiveWorkbook.Name }
    $s = Retry { $xl.ActiveSheet.Name }
    $c = Retry { $xl.ActiveCell.Address($false, $false) }
    $w = Window-Open
    $ok = ($b -eq $book) -and ($s -eq $sheet) -and ($c -eq $cell) -and (($null -eq $window) -or ($w -eq $window))
  } while (-not $ok -and (Get-Date) -lt $deadline)
  $sb = Retry { $xl.StatusBar }
  $mark = if ($ok) { 'ok  ' } else { 'FAIL' }
  "{0} {1,-44} active=[{2}]{3}!{4} window={5} status=[{6}]" -f $mark, $label, $b, $s, $c, $w, $sb
  if (-not $ok) { $failures.Add("$label (expected [$book]$sheet!$cell window=$window, got [$b]$s!$c window=$w)") }
}

try {
  # --- A. Native undo survives a trace that stays in the workbook ---
  Select-Cell 'B1'
  Step '777~'   'type 777 in B1 + Enter'         $mainName 'Calc' 'B2' $false
  Step '^+{[}'  'Ctrl+Shift+[ on B2'             $mainName 'Calc' 'B2' $true -waitMs 3000
  Step '{DOWN}' 'Down: Inputs!B2'                $mainName 'Inputs' 'B2' $true
  Step '{DOWN}' 'Down: Growth -> Inputs!B4'      $mainName 'Inputs' 'B4' $true
  Step '{ESC}'  'Esc: back to B2, window closed' $mainName 'Calc' 'B2' $false
  Step '^z'     'Ctrl+Z (expect B1 empty)'       $mainName 'Calc' 'B1' $false
  $b1 = Retry { $calc.Range('B1').Value2 }
  if ($null -ne $b1) { $failures.Add("native undo after a trace: B1 still holds [$b1]") ; "FAIL B1 still holds [$b1]" } else { "ok   B1 is empty again: native undo survived the trace" }

  # --- B. The main trace: every reference of Calc!A1, F2, then into the closed external workbook and back ---
  Select-Cell 'A1'
  Step '^+{[}'  'Ctrl+Shift+[ on A1'             $mainName 'Calc' 'A1' $true -waitMs 3000
  foreach ($target in @('B2', 'B3', 'B4', 'B5', 'B6', 'B7', 'B8', 'B10', 'B11', 'B13')) {
    Step '{DOWN}' "Down: $target"                $mainName 'Calc' $target $true
  }
  $formulaBefore = Retry { $calc.Range('B13').Formula }
  Step '{F2}'   'F2: edit B13, window stays'     $mainName 'Calc' 'B13' $true
  Step '{ESC}'  'Esc in edit mode goes to Excel' $mainName 'Calc' 'B13' $true
  $formulaAfter = Retry { $calc.Range('B13').Formula }
  if ($formulaAfter -ne $formulaBefore) { $failures.Add("F2/Esc changed B13: [$formulaBefore] -> [$formulaAfter]") }
  Step '{UP}'    'Up: B11'                       $mainName 'Calc' 'B11' $true
  Step '{RIGHT}' 'Right: expand B11 (opens ext)' $mainName 'Calc' 'B11' $true -waitMs 15000
  $extOpen = $false; foreach ($open in @($xl.Workbooks)) { if ($open.Name -eq $extName) { $extOpen = $true } }
  if (-not $extOpen) { $failures.Add("expanding B11 did not open $extName") ; "FAIL $extName is not open" } else { "ok   $extName was opened by the trace" }
  Step '{DOWN}'  'Down: [External]Rates!B3'      $extName 'Rates' 'B3' $true -waitMs 3000
  Step '{RIGHT}' 'Right: expand Rates!B3'        $extName 'Rates' 'B3' $true
  Step '{DOWN}'  'Down: Rates!B2'                $extName 'Rates' 'B2' $true
  Step '{LEFT}'  'Left: up to Rates!B3'          $extName 'Rates' 'B3' $true
  Step '{LEFT}'  'Left: collapse Rates!B3'       $extName 'Rates' 'B3' $true
  Step '{LEFT}'  'Left: up to Calc!B11'          $mainName 'Calc' 'B11' $true -waitMs 3000
  Step '{ESC}'   'Esc: back to A1, closed'       $mainName 'Calc' 'A1' $false

  # --- C. A hidden sheet and a constant name: the selection moves in the tree, Excel stays put ---
  Select-Cell 'B5'
  Step '^+{[}'  'Ctrl+Shift+[ on B5'             $mainName 'Calc' 'B5' $true -waitMs 3000
  Step '{DOWN}' 'Down: Hidden!A1 (refused)'      $mainName 'Calc' 'B5' $true
  Step '{DOWN}' 'Down: TaxRate (no cells)'       $mainName 'Calc' 'B5' $true
  Step '{ESC}'  'Esc: back to B5'                $mainName 'Calc' 'B5' $false

  # --- D. A circular reference ---
  Select-Cell 'B8'
  Step '^+{[}'   'Ctrl+Shift+[ on B8'            $mainName 'Calc' 'B8' $true -waitMs 3000
  Step '{DOWN}'  'Down: B9'                      $mainName 'Calc' 'B9' $true
  Step '{RIGHT}' 'Right: expand B9'              $mainName 'Calc' 'B9' $true
  Step '{DOWN}'  'Down: B8 (circular)'           $mainName 'Calc' 'B8' $true
  Step '{ESC}'   'Esc: back to B8'               $mainName 'Calc' 'B8' $false

  # --- E. Enter stays; Last Audited Cell goes back through the audits (A1, B8, B5, A1) ---
  Select-Cell 'A1'
  Step '^+{[}'  'Ctrl+Shift+[ on A1'             $mainName 'Calc' 'A1' $true -waitMs 3000
  Step '{DOWN}' 'Down: B2'                       $mainName 'Calc' 'B2' $true
  Step '~'      'Enter: close, stay on B2'       $mainName 'Calc' 'B2' $false
  Step '^+\'    'Ctrl+Shift+\ #1: A1'            $mainName 'Calc' 'A1' $false
  Step '^+\'    'Ctrl+Shift+\ #2: B8'            $mainName 'Calc' 'B8' $false
  Step '^+\'    'Ctrl+Shift+\ #3: B5'            $mainName 'Calc' 'B5' $false
  Step '^+\'    'Ctrl+Shift+\ #4: A1'            $mainName 'Calc' 'A1' $false
}
catch {
  $failures.Add("aborted: $($_.Exception.Message)")
  "ABORTED: $($_.Exception.Message)"
}

# --- Diagnostics log: the trace lines, and the timings against the Phase 4 targets ---
"--- trace log lines since start ---"
$lines = @()
if (Test-Path $log) {
  $all = [IO.File]::ReadAllBytes($log)
  $from = if ($all.Length -ge $logStart) { $logStart } else { 0 }   # the log rolled over
  $lines = [Text.Encoding]::UTF8.GetString($all, $from, $all.Length - $from) -split "`r?`n" |
    Where-Object { $_ -match "`t(Trace\w*|LastAuditedCell)`t" }
}
$lines
function Max-Ms($pattern) {
  $values = $lines | Where-Object { $_ -match $pattern } | ForEach-Object { if ($_ -match "`tms=([0-9.]+)") { [double]$Matches[1] } }
  if ($values) { ($values | Measure-Object -Maximum).Maximum } else { $null }
}
$openMax = Max-Ms "`tTraceOpen`t.*`tok"
$stepMax = Max-Ms "`tTraceNavigate`tkey=(Up|Down)`tsource=key`t.*`tmove=Moved`t"
"PERF: Trace In open max {0} ms (target <= 300; the first open in a session includes one-time loading)" -f $openMax
"PERF: Up/Down step max {0} ms from key press (target <= 100)" -f $stepMax
if (-not ($lines | Where-Object { $_ -match "`tTraceHook`tuninstalled" })) { $failures.Add('no TraceHook uninstalled line: the key hook may still be installed') }

# --- Close what this test opened, without saving ---
try { Retry { $xl.Iteration = $iterationBefore } | Out-Null } catch {}
foreach ($n in @($extName, $mainName)) {
  foreach ($open in @($xl.Workbooks)) { if ($open.Name -eq $n) { try { $open.Close($false); "Closed $n without saving" } catch {} } }
}

if ($failures.Count -eq 0) { "RESULT: PASS" } else { "RESULT: FAIL"; $failures | ForEach-Object { "  - $_" }; exit 1 }
