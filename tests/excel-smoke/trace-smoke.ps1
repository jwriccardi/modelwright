# Excel smoke test: Trace In / Last Audited Cell, driven by REAL keystrokes into the running Excel.
# Run with Windows PowerShell 5.1 (needs Marshal.GetActiveObject), with Excel open:
#   powershell -ExecutionPolicy Bypass -File tests/excel-smoke/trace-smoke.ps1 [-Xll <path to packed xll>]
# It (re)builds the fixture first (build-trace-fixture.ps1: %TEMP%\emt-trace-fixture), opens EMT_TraceMain.xlsx and
# checks through COM, after every key, which workbook, sheet and cell is active, and whether the Trace In window is
# open. Covered: cross-sheet and same-sheet navigation, names, a hidden sheet (Goto refused, no error), F2 passing
# through to Excel with the window open, expanding into the closed external workbook (Trace In opens it), Left back
# out of it, Enter (stay), Esc (back to the audited cell), Last Audited Cell through four levels, ROW() in OFFSET
# evaluated for its own cell, a name whose target is in the closed external workbook, and native undo surviving a
# trace that stays in the workbook (type in B1, trace + navigate + Esc, Ctrl+Z: B1 is empty again).
# Timings (open, and each Up/Down step) are read back from the diagnostics log and reported against the targets
# (300 ms, 100 ms); they are reported, not failed on. With the diagnostics log off (diagnosticsLog: false in
# settings.json) those log checks are skipped and the output says so; settings.json is never touched.
#
# Safety (it runs against your live Excel; don't touch the keyboard while it runs):
# - Before EVERY key it checks that the foreground window is a workbook window (XLMAIN) of the Excel it drives (the
#   process is taken from Application.Hwnd), that it is Excel's active window, and that the active workbook is one of
#   the fixture workbooks in the fixture folder; otherwise it aborts without sending the key.
# - The first step that does not end where expected aborts the run: no further key is sent. Ctrl+Z is sent only once
#   COM shows the 777 typed in the fixture's B1 and Excel's own Undo is available (so the add-in's formatting undo
#   cannot take the key).
# - It records Excel's DisplayAlerts and Iteration settings and the installed state of every
#   *ModelingToolkit64-packed.xll in the add-in list, and puts all of them back in `finally` (each on its own), however
#   the run ends: the add-in build under test is uninstalled again unless it was installed before, and the builds that
#   were installed are installed again. Only the fixture workbooks (in the fixture folder) are closed, without saving.
# The focus trick taps Shift (only when Excel is not already in front): an Alt tap would turn on ribbon KeyTips and
# send the next key to the ribbon.
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
  [DllImport("user32.dll", CharSet = CharSet.Unicode)] public static extern int GetClassName(IntPtr h, StringBuilder s, int n);
  public delegate bool EnumProc(IntPtr h, IntPtr l);
  [DllImport("user32.dll")] public static extern bool EnumWindows(EnumProc f, IntPtr l);
  public static uint ProcessOf(IntPtr h) { uint p; GetWindowThreadProcessId(h, out p); return p; }
  public static string ClassOf(IntPtr h) { var s = new StringBuilder(64); GetClassName(h, s, 64); return s.ToString(); }
  // True if process pid has a visible top-level window titled exactly title.
  public static bool HasWindow(uint pid, string title) {
    bool found = false;
    EnumWindows((h, l) => {
      if (ProcessOf(h) == pid && IsWindowVisible(h)) {
        var s = new StringBuilder(256); GetWindowText(h, s, 256);
        if (s.ToString() == title) { found = true; return false; }
      }
      return true;
    }, IntPtr.Zero);
    return found;
  }
}
"@
if (-not ('EmtPath' -as [type])) {
Add-Type @"
using System; using System.Runtime.InteropServices; using System.Text;
public static class EmtPath {
  [DllImport("kernel32.dll", CharSet = CharSet.Unicode)] static extern uint GetLongPathName(string s, StringBuilder l, uint n);
  // The full path, 8.3 short names (C:\Users\JOHNRI~1, as %TEMP% may be) expanded in the part that exists, without a
  // trailing separator; the path as given if it is not a local path.
  public static string Long(string path) {
    try {
      var full = System.IO.Path.GetFullPath(path).TrimEnd('\\');
      var tail = string.Empty;
      var buffer = new StringBuilder(32768);
      while (true) {
        var length = GetLongPathName(full, buffer, (uint)buffer.Capacity);
        if (length > 0 && length < buffer.Capacity) { return buffer.ToString().TrimEnd('\\') + tail; }
        var parent = System.IO.Path.GetDirectoryName(full);
        if (string.IsNullOrEmpty(parent)) { return full + tail; }
        tail = "\\" + System.IO.Path.GetFileName(full) + tail;
        full = parent.TrimEnd('\\');
      }
    } catch (Exception) { return path.TrimEnd('\\'); }
  }
  // The folder of a workbook's FullName, as Long gives it; empty for an unsaved workbook (no folder).
  public static string FolderOf(string fullName) {
    var cut = fullName.LastIndexOfAny(new[] { '\\', '/' });
    return cut < 0 ? string.Empty : Long(fullName.Substring(0, cut));
  }
}
"@
}

$fixXll = (Resolve-Path $Xll).Path
$fixtureFolder = [EmtPath]::Long($FixtureDir)
$mainName = 'EMT_TraceMain.xlsx'
$extName = 'EMT_TraceExternal.xlsx'
$mainPath = Join-Path $fixtureFolder $mainName
$log = Join-Path $env:LOCALAPPDATA 'ModelingToolkit\log.txt'
$logStart = if (Test-Path $log) { (Get-Item $log).Length } else { 0 }
$low32 = [int64]4294967295

function Retry([scriptblock]$b) {
  for ($i = 0; $i -lt 40; $i++) { try { return & $b } catch { Start-Sleep -Milliseconds 150 } }
  throw "COM call kept failing"
}

$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$excelPid = [W]::ProcessOf([IntPtr]([int64]$xl.Hwnd))
if ($excelPid -eq 0) { throw "ABORT: could not find the process of the Excel window (Application.Hwnd). Nothing was changed." }
"Excel version $($xl.Version) build $($xl.Build), process $excelPid"

# A fixture workbook: one of the two names, opened from the fixture folder.
function Is-Fixture($book) {
  if ($null -eq $book) { return $false }
  $folder = [EmtPath]::FolderOf([string]$book.FullName)
  return (@($mainName, $extName) -contains [string]$book.Name) -and ($folder -ieq $fixtureFolder)
}

# --- Record everything this run changes, so `finally` can put it back ---
$alertsBefore = $xl.DisplayAlerts
$iterationBefore = $null
try { $iterationBefore = $xl.Iteration } catch { }   # needs an open workbook; read once the fixture is open if not
$addinsBefore = @()
foreach ($a in @($xl.AddIns)) {
  if ($a.FullName -like '*ModelingToolkit64-packed.xll') {
    $addinsBefore += New-Object PSObject -Property @{ FullName = [string]$a.FullName; Installed = [bool]$a.Installed }
  }
}
$testWasInstalled = @($addinsBefore | Where-Object { $_.FullName -ieq $fixXll -and $_.Installed }).Count -gt 0
"Add-in list before: " + (($addinsBefore | ForEach-Object { "$($_.FullName) installed=$($_.Installed)" }) -join '; ')

$failures = New-Object System.Collections.Generic.List[string]

function Focus-Excel {
  $h = [IntPtr]([int64](Retry { $xl.ActiveWindow.Hwnd }))
  if ([W]::GetForegroundWindow() -eq $h) { return }
  if ([W]::ProcessOf([W]::GetForegroundWindow()) -ne $excelPid) {
    # Shift tap (harmless) releases the foreground lock; Alt would trigger ribbon KeyTips.
    [W]::keybd_event(0x10, 0, 0, [UIntPtr]::Zero); [W]::keybd_event(0x10, 0, 2, [UIntPtr]::Zero)
  }
  [W]::SetForegroundWindow($h) | Out-Null
  Start-Sleep -Milliseconds 300
}

# Checked before EVERY key: the key can only reach a fixture workbook's window of this Excel.
function Assert-SafeToSend([string]$what) {
  $fg = [W]::GetForegroundWindow()
  $p = [W]::ProcessOf($fg)
  if ($p -ne $excelPid) { throw "ABORT before [$what]: the foreground window belongs to process $p, not this Excel ($excelPid). No key sent." }
  $class = [W]::ClassOf($fg)
  if ($class -ne 'XLMAIN') { throw "ABORT before [$what]: the foreground window is not an Excel workbook window (class $class). No key sent." }
  $active = [int64](Retry { $xl.ActiveWindow.Hwnd })
  if (($active -band $low32) -ne ($fg.ToInt64() -band $low32)) { throw "ABORT before [$what]: the foreground window is not Excel's active window. No key sent." }
  $book = Retry { $xl.ActiveWorkbook }
  if (-not (Is-Fixture $book)) { throw "ABORT before [$what]: the active workbook is $($book.FullName), not a fixture workbook. No key sent." }
}

function Window-Open { [W]::HasWindow([uint32]$excelPid, 'Trace In') }

# Sends keys, waits, then checks the active workbook/sheet/cell and (if given) whether the window is open. The first
# step that does not end there aborts the run (throws): no further key is sent.
function Step([string]$keys, [string]$label, [string]$book, [string]$sheet, [string]$cell, $window = $null, [int]$waitMs = 700) {
  Assert-SafeToSend $label
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
  "{0} {1,-48} active=[{2}]{3}!{4} window={5} status=[{6}]" -f $(if ($ok) { 'ok  ' } else { 'FAIL' }), $label, $b, $s, $c, $w, $sb
  if (-not $ok) { throw "step [$label] failed: expected [$book]$sheet!$cell window=$window, got [$b]$s!$c window=$w" }
}

try {
  # --- Fixture: rebuilt, so the external workbook is closed and the trace must open it (the builder aborts if a
  # fixture-named workbook is open from anywhere but the fixture folder) ---
  & (Join-Path $PSScriptRoot 'build-trace-fixture.ps1') -Excel $xl -OutDir $fixtureFolder
  foreach ($open in @($xl.Workbooks)) {
    if (@($mainName, $extName) -contains $open.Name) { throw "ABORT: $($open.Name) is still open ($($open.FullName))." }
  }

  $xl.DisplayAlerts = $false
  $wb = Retry { $xl.Workbooks.Open($mainPath, 0) }   # UpdateLinks 0: no prompt about the external link
  $xl.DisplayAlerts = $alertsBefore
  if (-not (Is-Fixture $wb)) { throw "ABORT: opened $($wb.FullName), not the fixture." }
  if ($null -eq $iterationBefore) { $iterationBefore = $xl.Iteration }
  $xl.Iteration = $true                             # the fixture's circular reference, without Excel's warning

  # --- Swap add-in builds (same as the Add-ins dialog); put back in `finally` ---
  foreach ($a in @($xl.AddIns)) {
    if ($a.FullName -like '*ModelingToolkit64-packed.xll' -and $a.FullName -ne $fixXll -and $a.Installed) {
      "Unloading $($a.FullName) for the run"; $a.Installed = $false
    }
  }
  $added = $xl.AddIns.Add($fixXll)
  if (-not $added.Installed) { $added.Installed = $true }
  Start-Sleep -Milliseconds 1500
  "Loaded: $($added.FullName) installed=$($added.Installed)"
  $calc = $wb.Worksheets.Item('Calc')

  function Select-Cell([string]$address) {
    Retry { $wb.Activate() } | Out-Null
    Retry { $calc.Activate() } | Out-Null
    Retry { $calc.Range($address).Select() } | Out-Null
    Focus-Excel
  }

  # --- A. Native undo survives a trace that stays in the workbook ---
  Select-Cell 'B1'
  $b1 = Retry { $calc.Range('B1').Value2 }
  if ($null -ne $b1) { throw "ABORT: the fixture's Calc!B1 is not empty ([$b1])." }
  Step '777~'   'type 777 in B1 + Enter'         $mainName 'Calc' 'B2' $false
  $b1 = Retry { $calc.Range('B1').Value2 }
  if ($b1 -ne 777) { throw "ABORT: 777 did not land in the fixture's Calc!B1 (it holds [$b1]); Ctrl+Z will not be sent." }
  Step '^+{[}'  'Ctrl+Shift+[ on B2'             $mainName 'Calc' 'B2' $true -waitMs 3000
  Step '{DOWN}' 'Down: Inputs!B2'                $mainName 'Inputs' 'B2' $true
  Step '{DOWN}' 'Down: Growth -> Inputs!B4'      $mainName 'Inputs' 'B4' $true
  Step '{ESC}'  'Esc: back to B2, window closed' $mainName 'Calc' 'B2' $false
  # If the trace had lost Excel's undo history, Ctrl+Z could go to the add-in's own formatting undo (which may act on
  # another workbook): send it only while Excel's Undo is available.
  $canUndo = Retry { $xl.CommandBars.GetEnabledMso('Undo') }
  if (-not $canUndo) {
    $failures.Add('native undo after a trace: Excel can no longer undo (its history was lost); Ctrl+Z not sent')
    "FAIL Excel's Undo is unavailable after the trace; Ctrl+Z not sent"
  }
  else {
    Step '^z'     'Ctrl+Z (expect B1 empty)'     $mainName 'Calc' 'B1' $false
    $b1 = Retry { $calc.Range('B1').Value2 }
    if ($null -ne $b1) { $failures.Add("native undo after a trace: B1 still holds [$b1]"); "FAIL B1 still holds [$b1]" } else { "ok   B1 is empty again: native undo survived the trace" }
  }

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
  if ($formulaAfter -ne $formulaBefore) { throw "ABORT: F2/Esc changed B13: [$formulaBefore] -> [$formulaAfter]" }
  Step '{UP}'    'Up: B11'                       $mainName 'Calc' 'B11' $true
  Step '{RIGHT}' 'Right: expand B11 (opens ext)' $mainName 'Calc' 'B11' $true -waitMs 15000
  $extOpen = $false; foreach ($open in @($xl.Workbooks)) { if ($open.Name -eq $extName -and (Is-Fixture $open)) { $extOpen = $true } }
  if (-not $extOpen) { throw "ABORT: expanding B11 did not open the fixture's $extName" }
  "ok   $extName was opened by the trace"
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

  # --- F. ROW() in a computed reference is the formula cell's: OFFSET($A$1,ROW()-1,0) in B14 is A14, not A1 ---
  Select-Cell 'B14'
  Step '^+{[}'  'Ctrl+Shift+[ on B14'            $mainName 'Calc' 'B14' $true -waitMs 3000
  Step '{DOWN}' 'Down: OFFSET(...ROW()...) -> A14' $mainName 'Calc' 'A14' $true
  Step '{ESC}'  'Esc: back to B14'               $mainName 'Calc' 'B14' $false

  # --- G. A name whose target is in a closed workbook can be followed (Trace In opens the workbook) ---
  foreach ($open in @($xl.Workbooks)) { if ($open.Name -eq $extName -and (Is-Fixture $open)) { $open.Close($false); "Closed $extName (fixture) so ExtRate points into a closed workbook" } }
  Select-Cell 'B15'
  Step '^+{[}'  'Ctrl+Shift+[ on B15 (opens ext)' $mainName 'Calc' 'B15' $true -waitMs 15000
  Step '{DOWN}' 'Down: ExtRate -> [External]Rates!B2' $extName 'Rates' 'B2' $true -waitMs 3000
  Step '{ESC}'  'Esc: back to B15'               $mainName 'Calc' 'B15' $false -waitMs 3000
}
catch {
  $failures.Add("aborted: $($_.Exception.Message)")
  "ABORTED: $($_.Exception.Message)"
}
finally {
  "--- putting Excel back ---"
  # 1. The add-in builds: the test build out unless it was installed before, then the builds that were installed.
  if (-not $testWasInstalled) {
    try {
      foreach ($a in @($xl.AddIns)) { if ($a.FullName -ieq $fixXll -and $a.Installed) { $a.Installed = $false; "Uninstalled the test build $fixXll" } }
    }
    catch { $failures.Add("could not uninstall the test build: $($_.Exception.Message)"); "WARNING: could not uninstall the test build $fixXll`: $($_.Exception.Message)" }
  }
  foreach ($before in $addinsBefore) {
    if (-not $before.Installed) { continue }
    try {
      foreach ($a in @($xl.AddIns)) { if ($a.FullName -ieq $before.FullName -and -not $a.Installed) { $a.Installed = $true; "Re-installed $($before.FullName)" } }
    }
    catch { $failures.Add("could not re-install $($before.FullName): $($_.Exception.Message)"); "WARNING: could not re-install $($before.FullName): $($_.Exception.Message)" }
  }

  # 2. Iteration (set while a workbook is open: the fixture still is).
  if ($null -ne $iterationBefore) {
    try { if ((Retry { $xl.Iteration }) -ne $iterationBefore) { Retry { $xl.Iteration = $iterationBefore } | Out-Null; "Iteration back to $iterationBefore" } }
    catch { $failures.Add("could not put Iteration back: $($_.Exception.Message)"); "WARNING: could not put Excel's Iteration back to $iterationBefore`: $($_.Exception.Message)" }
  }

  # 3. The fixture workbooks (only those in the fixture folder), without saving.
  foreach ($n in @($extName, $mainName)) {
    try {
      foreach ($open in @($xl.Workbooks)) { if ($open.Name -eq $n -and (Is-Fixture $open)) { $open.Close($false); "Closed $n without saving" } }
    }
    catch { "WARNING: could not close $n`: $($_.Exception.Message)" }
  }

  # 4. DisplayAlerts.
  try { $xl.DisplayAlerts = $alertsBefore } catch { $failures.Add("could not put DisplayAlerts back: $($_.Exception.Message)"); "WARNING: could not put Excel's DisplayAlerts back to $alertsBefore`: $($_.Exception.Message)" }
}

# --- Diagnostics log: the trace lines, and the timings against the Phase 4 targets (only if the log is on) ---
Start-Sleep -Milliseconds 1500   # the key hook stays up to 1 s after a close on Esc/Enter (until the key is released)
$lines = @()
if (Test-Path $log) {
  $all = [IO.File]::ReadAllBytes($log)
  $from = if ($all.Length -ge $logStart) { $logStart } else { 0 }   # the log rolled over
  $lines = @([Text.Encoding]::UTF8.GetString($all, $from, $all.Length - $from) -split "`r?`n" |
    Where-Object { $_ -match "`t(Trace\w*|LastAuditedCell)`t" })
}
if (-not ($lines | Where-Object { $_ -match "`tTraceOpen`t" })) {
  "SKIP: no Trace In lines in the diagnostics log since the start ($log): the log is off (diagnosticsLog: false in settings.json) or elsewhere. The timing report and the key-hook-removed check were skipped; the key checks above do not need the log."
}
else {
  "--- trace log lines since start ---"
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
}

if ($failures.Count -eq 0) { "RESULT: PASS" } else { "RESULT: FAIL"; $failures | ForEach-Object { "  - $_" }; exit 1 }
