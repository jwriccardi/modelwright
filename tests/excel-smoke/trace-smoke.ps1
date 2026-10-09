# Excel smoke test: Trace In / Last Audited Cell, driven by REAL keystrokes into the running Excel.
# Run with Windows PowerShell 5.1 (needs Marshal.GetActiveObject), with Excel open:
#   powershell -ExecutionPolicy Bypass -File tests/excel-smoke/trace-smoke.ps1 [-Xll <path to packed xll>] [-Bitness 64|32]
# -Bitness picks the default -Xll (Modelwright64.xll or Modelwright32.xll): the one matching Excel's bitness.
# It (re)builds the fixture first (build-trace-fixture.ps1: %TEMP%\emt-trace-fixture), opens EMT_TraceMain.xlsx and
# checks through COM, after every key, which workbook, sheet and cell is active, and whether the Trace In window is
# open. Covered: cross-sheet and same-sheet navigation, names, a hidden sheet (Goto refused, no error), F2 on a
# reference row (back to the audited cell in Point mode; Esc cancels, the window stays), expanding into the closed
# external workbook (Trace In opens it), Left back out of it, Enter (stay), Esc (back to the audited cell), Last
# Audited Cell through four levels, ROW() in OFFSET evaluated for its own cell, a name whose target is in the closed
# external workbook, native undo surviving a trace that stays in the workbook (type in B1, trace + navigate + Esc,
# Ctrl+Z: B1 is empty again), and F2 edits of a reference, where the add-in's keys go to the reference's target with
# Excel's Go To dialog so Point mode moves from it: B16 =A14+A2 (Down to A2, F2, Down in Point mode, Enter: B16 is
# =A14+A3, Excel is back on B16 with the window open, and Ctrl+Z restores =A14+A2); B2 =Inputs!B2*(1+Growth) (another
# sheet: Down makes it Inputs!B3, Ctrl+Z restores it); B11 into the external workbook the trace opens (Go To across
# windows is unreliable in Point mode, so the keys switch to that workbook's window with Ctrl+Tab, then Go To
# 'Rates'!B3 within it: F2 + Enter without moving makes it Excel's Point-mode form [External]Rates!$B$3 and Ctrl+Z
# restores it; traced again, Down makes it [External]Rates!$B$4, Ctrl+Z restores it; traced again, F2 + Esc leaves
# B11 as it was and Excel back on B11 with the window open).
# Timings (open, and each Up/Down step) are read back from the diagnostics log and reported against the targets
# (300 ms, 100 ms); they are reported, not failed on. With the diagnostics log off (diagnosticsLog: false in
# settings.json) those log checks are skipped and the output says so; settings.json is never touched.
#
# Safety (it runs against your live Excel; don't touch the keyboard while it runs):
# - Before EVERY key it checks that the foreground window is a workbook window (XLMAIN) of the Excel it drives (the
#   process is taken from Application.Hwnd), that it is Excel's active window, and that the active workbook is one of
#   the fixture workbooks in the fixture folder; otherwise it aborts without sending the key. While Excel is editing
#   (COM is refused then) the check is the window title: a fixture workbook's, or Excel's Go To dialog (titled "Go To",
#   or in another language a dialog window of this Excel's, class bosa_sdm_*), which the add-in's F2 keys open and close
#   in Point mode: a key sent then is queued behind them.
# - The first step that does not end where expected aborts the run: no further key is sent. Ctrl+Z is sent only once
#   COM shows the 777 typed in the fixture's B1 and Excel's own Undo is available (so the add-in's formatting undo
#   cannot take the key).
# - It records Excel's DisplayAlerts and Iteration settings and the installed state of every Modelwright64.xll or
#   Modelwright32.xll (or pre-rename ModelingToolkit64-packed.xll) in the add-in list, and puts all of them back in `finally` (each on its
#   own), however the run ends: the add-in build under test is uninstalled again unless it was installed before, and
#   the builds that were installed are installed again. Only the fixture workbooks (in the fixture folder) are closed, without saving.
#   If Excel is still editing a cell then (a run that aborted in Point mode), Esc is sent first (at most three, each
#   after the same foreground check).
# The focus trick taps Shift (only when Excel is not already in front): an Alt tap would turn on ribbon KeyTips and
# send the next key to the ribbon.
param(
    [string]$Xll = '',
    [ValidateSet('64', '32')][string]$Bitness = '64',
    [string]$FixtureDir = (Join-Path $env:TEMP 'emt-trace-fixture')
)
$ErrorActionPreference = 'Stop'
if (-not $Xll) { $Xll = Join-Path $PSScriptRoot "..\..\src\Modelwright.AddIn\bin\Release\net48\publish\Modelwright$Bitness.xll" }
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
  public static string TitleOf(IntPtr h) { var s = new StringBuilder(256); GetWindowText(h, s, 256); return s.ToString(); }
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

# Our add-in in Excel's list: this build's name, or the name builds had before the rename to Modelwright (D12), so an
# old build left installed is swapped out as well (both would claim the same shortcuts).
function Is-OurXll([string]$fullName) { return ($fullName -like '*Modelwright64.xll') -or ($fullName -like '*Modelwright32.xll') -or ($fullName -like '*ModelingToolkit64-packed.xll') }
$fixtureFolder = [EmtPath]::Long($FixtureDir)
$mainName = 'EMT_TraceMain.xlsx'
$extName = 'EMT_TraceExternal.xlsx'
$mainPath = Join-Path $fixtureFolder $mainName
$log = Join-Path $env:LOCALAPPDATA 'Modelwright\log.txt'
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
  if (Is-OurXll $a.FullName) {
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
  $title = [W]::TitleOf($fg)
  # The add-in's F2 keys open Excel's Go To dialog in Point mode and close it again (one SendInput batch): a key sent
  # while it shows is queued behind those keys and reaches Excel after the dialog has closed.
  $goTo = (Is-GoToDialog $fg) -and (Edit-Mode)
  if ($class -ne 'XLMAIN' -and -not $goTo) { throw "ABORT before [$what]: the foreground window is not an Excel workbook window (class $class, [$title]). No key sent." }
  if (Edit-Mode) {
    # Excel rejects every COM call while a cell is being edited, so the check is the window title: "<book> - Excel".
    if ($goTo) { return }
    if (-not (($title -like "$mainName*") -or ($title -like "$extName*"))) { throw "ABORT before [$what]: Excel is editing, and the foreground window is [$title], not a fixture workbook. No key sent." }
    return
  }
  $active = [int64](Retry { $xl.ActiveWindow.Hwnd })
  if (($active -band $low32) -ne ($fg.ToInt64() -band $low32)) { throw "ABORT before [$what]: the foreground window is not Excel's active window. No key sent." }
  $book = Retry { $xl.ActiveWorkbook }
  if (-not (Is-Fixture $book)) { throw "ABORT before [$what]: the active workbook is $($book.FullName), not a fixture workbook. No key sent." }
}

function Window-Open { [W]::HasWindow([uint32]$excelPid, 'Trace In') }

# True for Excel's Go To dialog: a window of this Excel's that is not a workbook window, titled "Go To" (English), or
# of the class of Excel's dialogs (bosa_sdm_*) in any language.
function Is-GoToDialog([IntPtr]$h) {
  if ($h -eq [IntPtr]::Zero -or [W]::ProcessOf($h) -ne $excelPid) { return $false }
  $class = [W]::ClassOf($h)
  return ($class -ne 'XLMAIN') -and (([W]::TitleOf($h) -eq 'Go To') -or ($class -like 'bosa_sdm_*'))
}

# True while Excel is editing a cell: Application.Ready is false then (or the call is refused while Excel is busy).
function Edit-Mode { try { return -not [bool]$xl.Ready } catch { return $true } }

# The add-in's diagnostics log lines since the run started that match a pattern (none when the log is off).
function Log-Lines([string]$pattern) {
  if (-not (Test-Path $log)) { return @() }
  $all = [IO.File]::ReadAllBytes($log)
  $from = if ($all.Length -ge $logStart) { $logStart } else { 0 }
  return @([Text.Encoding]::UTF8.GetString($all, $from, $all.Length - $from) -split "`r?`n" | Where-Object { $_ -match $pattern })
}

# After F2 on a reference row (Step -Editing): waits until the add-in's keys have all arrived (a TraceSynth line newer
# than the $synthBefore there were, "complete") and Excel's Go To dialog has closed, so the next key goes to Point
# mode; with the diagnostics log on, also that the F2 used the Go To step with $goTo (the text it typed, with nothing
# after it: "A2", "'Inputs'!B2", "'Rates'!B3") and, for a target in another workbook (-SwitchWindow), that it switched
# to that workbook's window with Ctrl+Tab first. The log off: only the dialog is waited for.
function Wait-EditKeys([int]$synthBefore, [string]$goTo, [string]$label, [switch]$SwitchWindow) {
  $logOn = @(Log-Lines "`tTraceOpen`t").Count -gt 0
  $deadline = (Get-Date).AddSeconds(4)
  do {
    Start-Sleep -Milliseconds 200
    $synth = @(Log-Lines "`tTraceSynth`t")
    $dialog = Is-GoToDialog ([W]::GetForegroundWindow())
    $arrived = (-not $logOn) -or ($synth.Count -gt $synthBefore)
  } while (($dialog -or -not $arrived) -and (Get-Date) -lt $deadline)
  if ($dialog) { throw "ABORT [$label]: Excel's Go To dialog is still open (the reference was not accepted?)." }
  # Excel is still closing the Go To dialog and switching windows when the last key has arrived: a key sent within
  # ~150 ms of it lands in the edited cell's window instead of the pointed one. A person cannot press that fast.
  Start-Sleep -Milliseconds 600
  if (-not (Edit-Mode)) { throw "ABORT [$label]: Excel is not editing (Application.Ready is true)." }
  if (-not $logOn) { "SKIP: no diagnostics log (it is off): only Application.Ready and the Go To dialog were checked"; return }
  if (-not $arrived) { throw "ABORT [$label]: no TraceSynth line: the add-in sent no keys." }
  if (-not ($synth[-1] -match "`tcomplete`t")) { throw "ABORT [$label]: the keys F2 sent did not all arrive: $($synth[-1])" }
  "ok   the add-in's keys all arrived: $($synth[-1])"
  $edit = @(Log-Lines "`tTraceEditReference`t")
  $switchText = if ($SwitchWindow) { 'switch window; ' } else { '' }
  if ($edit.Count -eq 0 -or -not $edit[-1].EndsWith("`tok; ${switchText}go to $goTo")) {
    throw "ABORT [$label]: F2 did not go to [$goTo]$(if ($SwitchWindow) { ' after Ctrl+Tab' }) in Point mode: $(if ($edit.Count) { $edit[-1] } else { 'no TraceEditReference line' })"
  }
  "ok   F2 went to $goTo$(if ($SwitchWindow) { ' (after Ctrl+Tab to its window)' }) in Point mode"
}

# Sends keys, waits, then checks the active workbook/sheet/cell (an empty cell: any) and (if given) whether the window
# is open. The first step that does not end there aborts the run (throws): no further key is sent.
function Step([string]$keys, [string]$label, [string]$book, [string]$sheet, [string]$cell, $window = $null, [int]$waitMs = 700, [switch]$Editing) {
  Assert-SafeToSend $label
  [System.Windows.Forms.SendKeys]::SendWait($keys)
  $deadline = (Get-Date).AddMilliseconds([Math]::Max($waitMs, 700))
  if ($Editing) {
    # Expected to leave Excel editing a cell of $book: COM is rejected then, so only Win32 is consulted.
    do {
      Start-Sleep -Milliseconds 200
      $fg = [W]::GetForegroundWindow(); $title = [W]::TitleOf($fg)
      $w = Window-Open
      $ok = (Edit-Mode) -and ([W]::ProcessOf($fg) -eq $excelPid) -and ($title -like "$book*") -and (($null -eq $window) -or ($w -eq $window))
    } while (-not $ok -and (Get-Date) -lt $deadline)
    "{0} {1,-48} editing in [{2}] window={3}" -f $(if ($ok) { 'ok  ' } else { 'FAIL' }), $label, $title, $w
    if (-not $ok) { throw "step [$label] failed: expected Excel editing in [$book] window=$window, got editing=$(Edit-Mode) title=[$title] window=$w" }
    return
  }
  do {
    Start-Sleep -Milliseconds 200
    $b = Retry { $xl.ActiveWorkbook.Name }
    $s = Retry { $xl.ActiveSheet.Name }
    $c = Retry { $xl.ActiveCell.Address($false, $false) }
    $w = Window-Open
    $ok = ($b -eq $book) -and ($s -eq $sheet) -and (($cell -eq '') -or ($c -eq $cell)) -and (($null -eq $window) -or ($w -eq $window))
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
    if ((Is-OurXll $a.FullName) -and $a.FullName -ne $fixXll -and $a.Installed) {
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
  Step '777~'   'type 777 in B1 + Enter'         $mainName 'Calc' '' $false   # where Enter leaves the cursor depends on the user's "move after Enter" setting
  $b1 = Retry { $calc.Range('B1').Value2 }
  if ($b1 -ne 777) { throw "ABORT: 777 did not land in the fixture's Calc!B1 (it holds [$b1]); Ctrl+Z will not be sent." }
  Select-Cell 'B2'
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
  # F2 on the B13 row edits that reference in A1's formula: back on A1, in Point mode; Esc cancels (Excel's) and
  # leaves Excel on A1 with the window open.
  $formulaBefore = Retry { $calc.Range('A1').Formula }
  $synthBefore = @(Log-Lines "`tTraceSynth`t").Count
  Step '{F2}'   'F2: A1, editing B13 in its formula' $mainName 'Calc' 'A1' $true -waitMs 1500 -Editing
  Wait-EditKeys $synthBefore 'B13' 'F2 on the B13 row'
  Step '{ESC}'  'Esc in Point mode goes to Excel' $mainName 'Calc' 'A1' $true -waitMs 1500
  $formulaAfter = Retry { $calc.Range('A1').Formula }
  if ($formulaAfter -ne $formulaBefore) { throw "ABORT: F2/Esc changed A1: [$formulaBefore] -> [$formulaAfter]" }
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

  # --- H. F2 edits a reference in Point mode: B16 =A14+A2, the A2 row, F2, Down (A3), Enter; then Ctrl+Z ---
  $b16 = Retry { $calc.Range('B16').Formula }
  if ($b16 -ne '=A14+A2') { throw "ABORT: the fixture's Calc!B16 is [$b16], not =A14+A2." }
  Select-Cell 'B16'
  Step '^+{[}'  'Ctrl+Shift+[ on B16'            $mainName 'Calc' 'B16' $true -waitMs 3000
  Step '{DOWN}' 'Down: A14'                      $mainName 'Calc' 'A14' $true
  Step '{DOWN}' 'Down: A2'                       $mainName 'Calc' 'A2' $true
  $synthBefore = @(Log-Lines "`tTraceSynth`t").Count
  Step '{F2}'   'F2: back on B16, editing A2'    $mainName 'Calc' 'B16' $true -waitMs 1500 -Editing
  Wait-EditKeys $synthBefore 'A2' 'F2 on the A2 row'
  "ok   Excel is editing B16 (Application.Ready is false)"
  Step '{DOWN}' 'Down in Point mode: A2 -> A3'   $mainName 'Calc' '' $true -Editing
  Step '~'      'Enter: commit, back on B16'     $mainName 'Calc' 'B16' $true -waitMs 3000
  if (Edit-Mode) { throw "ABORT: Excel is still editing after Enter." }
  $b16 = Retry { $calc.Range('B16').Formula }
  if ($b16 -ne '=A14+A3') { throw "ABORT: after the F2 edit B16 is [$b16], not =A14+A3." }
  "ok   B16 is now =A14+A3, Excel is back on B16 and the window is open"
  $ended = @(Log-Lines "`tTraceEditEnd`t")
  if ($ended.Count -gt 0) {
    if (-not ($ended[-1] -match "`tchanged=true`t")) { $failures.Add("F2 edit: the log does not show the edit's end as a change: $($ended[-1])") }
    else { "ok   the tree was rebuilt: $($ended[-1])" }
  }
  $canUndo = Retry { $xl.CommandBars.GetEnabledMso('Undo') }
  if (-not $canUndo) {
    $failures.Add('undo after an F2 edit: Excel cannot undo it; Ctrl+Z not sent')
    "FAIL Excel's Undo is unavailable after the F2 edit; Ctrl+Z not sent"
  }
  else {
    Step '^z'   'Ctrl+Z: undo the F2 edit'       $mainName 'Calc' 'B16' $true
    $b16 = Retry { $calc.Range('B16').Formula }
    if ($b16 -ne '=A14+A2') { $failures.Add("undo after an F2 edit: B16 is [$b16], not =A14+A2"); "FAIL B16 is [$b16] after Ctrl+Z" }
    else { "ok   Ctrl+Z restored =A14+A2: Excel's undo of the F2 edit survived" }
  }
  Step '{ESC}'  'Esc: back to B16, closed'       $mainName 'Calc' 'B16' $false

  # --- I. F2 on a reference to another sheet: B2 =Inputs!B2*(1+Growth), the Inputs!B2 row, F2 (Go To 'Inputs'!B2),
  # Down (Inputs!B3), Enter; then Ctrl+Z ---
  $b2Before = Retry { $calc.Range('B2').Formula }
  if ($b2Before -ne '=Inputs!B2*(1+Growth)') { throw "ABORT: the fixture's Calc!B2 is [$b2Before], not =Inputs!B2*(1+Growth)." }
  Select-Cell 'B2'
  Step '^+{[}'  'Ctrl+Shift+[ on B2'             $mainName 'Calc' 'B2' $true -waitMs 3000
  Step '{DOWN}' 'Down: Inputs!B2'                $mainName 'Inputs' 'B2' $true
  $synthBefore = @(Log-Lines "`tTraceSynth`t").Count
  Step '{F2}'   'F2: back on B2, editing Inputs!B2' $mainName '' '' $true -waitMs 1500 -Editing
  Wait-EditKeys $synthBefore "'Inputs'!B2" 'F2 on the Inputs!B2 row'
  Step '{DOWN}' 'Down in Point mode: Inputs!B3'  $mainName '' '' $true -Editing
  Step '~'      'Enter: commit, back on B2'      $mainName 'Calc' 'B2' $true -waitMs 3000
  if (Edit-Mode) { throw "ABORT: Excel is still editing after Enter." }
  $b2 = Retry { $calc.Range('B2').Formula }
  if (($b2 -notlike '*Inputs!B3*') -or ($b2 -like '*Inputs!B2*')) { throw "ABORT: after the F2 edit B2 is [$b2], not Inputs!B3 in place of Inputs!B2." }
  "ok   B2 is now $b2"
  $canUndo = Retry { $xl.CommandBars.GetEnabledMso('Undo') }
  if (-not $canUndo) {
    $failures.Add('undo after a cross-sheet F2 edit: Excel cannot undo it; Ctrl+Z not sent')
    "FAIL Excel's Undo is unavailable after the cross-sheet F2 edit; Ctrl+Z not sent"
  }
  else {
    Step '^z'   'Ctrl+Z: undo the cross-sheet edit' $mainName 'Calc' 'B2' $true
    $b2 = Retry { $calc.Range('B2').Formula }
    if ($b2 -ne $b2Before) { $failures.Add("undo after a cross-sheet F2 edit: B2 is [$b2], not $b2Before"); "FAIL B2 is [$b2] after Ctrl+Z" }
    else { "ok   Ctrl+Z restored $b2Before" }
  }
  Step '{ESC}'  'Esc: back to B2, closed'        $mainName 'Calc' 'B2' $false

  # --- J. F2 on a reference into the external workbook, which the trace opens: B11 =[External]Rates!B3*2 traced with
  # the workbook closed (its formula then holds the path, gone once it is open). Go To into another workbook's window
  # is unreliable in Point mode, so the add-in goes to Rates!B3 and back to B11 (the external window is then Excel's
  # previously active one), and its keys switch there with Ctrl+Tab, then Go To 'Rates'!B3 within that window. F2 +
  # Enter without moving writes Excel's Point-mode form for another workbook, absolute ([External]Rates!$B$3); Ctrl+Z
  # restores it. Then, traced again, F2, Down ($B$4), Enter; then Ctrl+Z. Then, traced again, F2, Esc (cancel, from
  # the external window): B11 unchanged, Excel back on B11 with the window open ---
  foreach ($open in @($xl.Workbooks)) { if ($open.Name -eq $extName -and (Is-Fixture $open)) { $open.Close($false); "Closed $extName (fixture) so tracing B11 opens it" } }
  Select-Cell 'B11'
  Step '^+{[}'  'Ctrl+Shift+[ on B11 (opens ext)' $mainName 'Calc' 'B11' $true -waitMs 15000
  $extOpen = $false; foreach ($open in @($xl.Workbooks)) { if ($open.Name -eq $extName -and (Is-Fixture $open)) { $extOpen = $true } }
  if (-not $extOpen) { throw "ABORT: tracing B11 did not open the fixture's $extName" }
  $b11Before = Retry { $calc.Range('B11').Formula }
  if ($b11Before -ne "=[$extName]Rates!B3*2") { throw "ABORT: with $extName open the fixture's Calc!B11 is [$b11Before], not =[$extName]Rates!B3*2." }
  Step '{DOWN}' 'Down: [External]Rates!B3'       $extName 'Rates' 'B3' $true -waitMs 3000
  $synthBefore = @(Log-Lines "`tTraceSynth`t").Count
  Step '{F2}'   'F2: back on B11, editing Rates!B3' '' '' '' $true -waitMs 1500 -Editing
  Wait-EditKeys $synthBefore "'Rates'!B3" 'F2 on the Rates!B3 row' -SwitchWindow
  Step '~'      'Enter without moving: back on B11' $mainName 'Calc' 'B11' $true -waitMs 3000
  if (Edit-Mode) { throw "ABORT: Excel is still editing after Enter." }
  $b11 = Retry { $calc.Range('B11').Formula }
  # Point mode writes a reference into another workbook absolute (verified in Excel 2026-10-09).
  $b11Pointed = "=[$extName]Rates!`$B`$3*2"
  if ($b11 -ne $b11Pointed) { throw "ABORT: after F2 + Enter without moving B11 is [$b11], not $b11Pointed." }
  "ok   F2 + Enter without moving made B11 $b11 (Point mode's absolute form)"
  $canUndo = Retry { $xl.CommandBars.GetEnabledMso('Undo') }
  if (-not $canUndo) {
    $failures.Add('undo after an external F2 + Enter: Excel cannot undo it; Ctrl+Z not sent')
    "FAIL Excel's Undo is unavailable after the external F2 + Enter; Ctrl+Z not sent"
  }
  else {
    Step '^z'   'Ctrl+Z: undo the F2 + Enter'    $mainName 'Calc' 'B11' $true
    $b11 = Retry { $calc.Range('B11').Formula }
    if ($b11 -ne $b11Before) { $failures.Add("undo after an external F2 + Enter: B11 is [$b11], not $b11Before"); "FAIL B11 is [$b11] after Ctrl+Z" }
    else { "ok   Ctrl+Z restored $b11Before" }
  }
  # Traced again from B11 for the next edit, so the tree and the formula are known to agree.
  Step '{ESC}'  'Esc: back to B11, closed'       $mainName 'Calc' 'B11' $false
  $b11Before = Retry { $calc.Range('B11').Formula }
  Step '^+{[}'  'Ctrl+Shift+[ on B11 again'      $mainName 'Calc' 'B11' $true -waitMs 3000
  Step '{DOWN}' 'Down: [External]Rates!B3'       $extName 'Rates' 'B3' $true -waitMs 3000
  $synthBefore = @(Log-Lines "`tTraceSynth`t").Count
  Step '{F2}'   'F2 again: editing Rates!B3'     '' '' '' $true -waitMs 1500 -Editing
  Wait-EditKeys $synthBefore "'Rates'!B3" 'F2 again on the Rates!B3 row' -SwitchWindow
  Step '{DOWN}' 'Down in Point mode: Rates!B4'   '' '' '' $true -Editing
  Step '~'      'Enter: commit, back on B11'     $mainName 'Calc' 'B11' $true -waitMs 3000
  if (Edit-Mode) { throw "ABORT: Excel is still editing after Enter." }
  $b11 = Retry { $calc.Range('B11').Formula }
  $b11Moved = "=[$extName]Rates!`$B`$4*2"
  if ($b11 -ne $b11Moved) { throw "ABORT: after the F2 edit B11 is [$b11], not $b11Moved." }
  "ok   B11 is now $b11"
  $canUndo = Retry { $xl.CommandBars.GetEnabledMso('Undo') }
  if (-not $canUndo) {
    $failures.Add('undo after an external F2 edit: Excel cannot undo it; Ctrl+Z not sent')
    "FAIL Excel's Undo is unavailable after the external F2 edit; Ctrl+Z not sent"
  }
  else {
    Step '^z'   'Ctrl+Z: undo the external edit' $mainName 'Calc' 'B11' $true
    $b11 = Retry { $calc.Range('B11').Formula }
    if ($b11 -ne $b11Before) { $failures.Add("undo after an external F2 edit: B11 is [$b11], not $b11Before"); "FAIL B11 is [$b11] after Ctrl+Z" }
    else { "ok   Ctrl+Z restored $b11Before" }
  }
  Step '{ESC}'  'Esc: back to B11, closed'       $mainName 'Calc' 'B11' $false

  # Cancel: Esc in Point mode while the external window is the active one ends the edit; the session goes back to B11.
  $b11Before = Retry { $calc.Range('B11').Formula }
  if ($b11Before -ne "=[$extName]Rates!B3*2") { "SKIP F2 + Esc on B11: B11 is [$b11Before] (the undo above failed)" }
  else {
    Step '^+{[}'  'Ctrl+Shift+[ on B11 a third time' $mainName 'Calc' 'B11' $true -waitMs 3000
    Step '{DOWN}' 'Down: [External]Rates!B3'       $extName 'Rates' 'B3' $true -waitMs 3000
    $synthBefore = @(Log-Lines "`tTraceSynth`t").Count
    Step '{F2}'   'F2 a third time: editing Rates!B3' '' '' '' $true -waitMs 1500 -Editing
    Wait-EditKeys $synthBefore "'Rates'!B3" 'F2 a third time on the Rates!B3 row' -SwitchWindow
    Step '{ESC}'  'Esc in Point mode: cancel, back on B11' $mainName 'Calc' 'B11' $true -waitMs 3000
    if (Edit-Mode) { throw "ABORT: Excel is still editing after Esc." }
    $b11 = Retry { $calc.Range('B11').Formula }
    if ($b11 -ne $b11Before) { throw "ABORT: F2 + Esc changed B11: [$b11Before] -> [$b11]" }
    "ok   F2 + Esc left B11 $b11, Excel back on B11 with the window open"
    Step '{ESC}'  'Esc: back to B11, closed'       $mainName 'Calc' 'B11' $false
  }
}
catch {
  $failures.Add("aborted: $($_.Exception.Message)")
  "ABORTED: $($_.Exception.Message)"
  $_.InvocationInfo.PositionMessage
}
finally {
  "--- putting Excel back ---"
  # 0. A run that aborted while Excel was editing (Point mode): Esc, only to a fixture window of this Excel (or the Go
  # To dialog the add-in opened in Point mode: one Esc closes it, the next ends the edit); at most three.
  try {
    for ($i = 0; ($i -lt 3) -and (Edit-Mode); $i++) {
      Assert-SafeToSend 'Esc to end the cell edit'
      [System.Windows.Forms.SendKeys]::SendWait('{ESC}')
      Start-Sleep -Milliseconds 500
      "Sent Esc: Excel was still editing a cell"
    }
  }
  catch { "WARNING: Excel may still be editing a cell: $($_.Exception.Message)" }

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
  "PERF: Trace In open max {0} ms (target <= 300; the add-in warms up 1.5 s after it loads, so a first open sooner pays the one-time loading)" -f $openMax
  $warmup = @($lines | Where-Object { $_ -match "`tTraceWarmup`t" })
  "PERF: warm-up: " + $(if ($warmup.Count) { $warmup[-1] } else { 'no TraceWarmup line since the start (it runs once, 1.5 s after the add-in loads)' })
  "PERF: Up/Down step max {0} ms from key press (target <= 100)" -f $stepMax
  if (-not ($lines | Where-Object { $_ -match "`tTraceHook`tuninstalled" })) { $failures.Add('no TraceHook uninstalled line: the key hook may still be installed') }
}

if ($failures.Count -eq 0) { "RESULT: PASS" } else { "RESULT: FAIL"; $failures | ForEach-Object { "  - $_" }; exit 1 }
