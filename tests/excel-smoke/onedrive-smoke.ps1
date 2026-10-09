# Excel smoke test: Trace In across workbooks stored on OneDrive / SharePoint, driven by real keystrokes.
# Excel rewrites a workbook saved in a OneDrive sync folder to its https URL, and writes links to it, once it is
# closed, as ='https://.../[OD_Source.xlsx]Rates'!B1. The trace must open such a workbook from the URL.
#
# Fixture: two workbooks the owner allowed in the business OneDrive (created 2026-10-08, never deleted):
#   <OneDrive - Pegasus Technology Group LLC>\Modelwright-scratch\OD_Main.xlsx    Calc!A1 =[OD_Source.xlsx]Rates!B1*2
#                                                                                 Calc!A2 =[OD_Source.xlsx]Rates!B2+1
#   <OneDrive - Pegasus Technology Group LLC>\Modelwright-scratch\OD_Source.xlsx  Rates!B1 0.07, Rates!B2 =B1*2
# Run with Windows PowerShell 5.1, Excel open with the add-in loaded, nobody at the keyboard:
#   powershell -ExecutionPolicy Bypass -File tests/excel-smoke/onedrive-smoke.ps1
# Safety: keys go only to an Excel window of this process whose title starts with OD_ (the fixture); the first failed
# step aborts; everything opened is closed without saving; nothing is ever saved or deleted.
param([string]$ScratchDir = (Join-Path $env:USERPROFILE 'OneDrive - Pegasus Technology Group LLC\Modelwright-scratch'))
$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type @"
using System; using System.Runtime.InteropServices; using System.Text;
public static class W2 {
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
  public static bool HasWindow(uint pid, string title) {
    bool found = false;
    EnumWindows((h, l) => { if (ProcessOf(h) == pid && IsWindowVisible(h) && TitleOf(h) == title) { found = true; return false; } return true; }, IntPtr.Zero);
    return found;
  }
}
"@

$mainName = 'OD_Main.xlsx'; $srcName = 'OD_Source.xlsx'
$mainPath = Join-Path $ScratchDir $mainName
if (-not (Test-Path $mainPath) -or -not (Test-Path (Join-Path $ScratchDir $srcName))) { throw "ABORT: fixture workbooks not found in $ScratchDir." }
$log = Join-Path $env:LOCALAPPDATA 'Modelwright\log.txt'
$logStart = if (Test-Path $log) { (Get-Item $log).Length } else { 0 }

function Retry([scriptblock]$b) {
  for ($i = 0; $i -lt 40; $i++) { try { return & $b } catch { Start-Sleep -Milliseconds 150 } }
  throw "COM call kept failing"
}

$xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
$excelPid = [W2]::ProcessOf([IntPtr]([int64]$xl.Hwnd))
"Excel version $($xl.Version) build $($xl.Build), process $excelPid"
$loaded = @($xl.AddIns | Where-Object { $_.FullName -like '*Modelwright64.xll' -and $_.Installed }).Count -gt 0
if (-not $loaded) { throw "ABORT: the add-in is not installed in this Excel; run trace-smoke.ps1 first (it installs the build)." }

function Is-Fixture($book) { return ($null -ne $book) -and (@($mainName, $srcName) -contains [string]$book.Name) -and ([string]$book.FullName -like '*Modelwright-scratch*') }

function Focus-Excel {
  $h = [IntPtr]([int64](Retry { $xl.ActiveWindow.Hwnd }))
  if ([W2]::GetForegroundWindow() -eq $h) { return }
  if ([W2]::ProcessOf([W2]::GetForegroundWindow()) -ne $excelPid) { [W2]::keybd_event(0x10, 0, 0, [UIntPtr]::Zero); [W2]::keybd_event(0x10, 0, 2, [UIntPtr]::Zero) }
  [W2]::SetForegroundWindow($h) | Out-Null
  Start-Sleep -Milliseconds 300
}
function Assert-SafeToSend([string]$what) {
  $fg = [W2]::GetForegroundWindow()
  if ([W2]::ProcessOf($fg) -ne $excelPid) { throw "ABORT before [$what]: the foreground window belongs to another process. No key sent." }
  if ([W2]::ClassOf($fg) -ne 'XLMAIN') { throw "ABORT before [$what]: the foreground window is not an Excel workbook window. No key sent." }
  $title = [W2]::TitleOf($fg)
  if (-not ($title -like 'OD_*')) { throw "ABORT before [$what]: the foreground window is [$title], not a fixture workbook. No key sent." }
  if (-not (Is-Fixture (Retry { $xl.ActiveWorkbook }))) { throw "ABORT before [$what]: the active workbook is not a fixture workbook. No key sent." }
}
function Window-Open { [W2]::HasWindow([uint32]$excelPid, 'Trace In') }
function Step([string]$keys, [string]$label, [string]$book, [string]$sheet, [string]$cell, $window = $null, [int]$waitMs = 700) {
  Assert-SafeToSend $label
  [System.Windows.Forms.SendKeys]::SendWait($keys)
  $deadline = (Get-Date).AddMilliseconds([Math]::Max($waitMs, 700))
  do {
    Start-Sleep -Milliseconds 200
    $b = Retry { $xl.ActiveWorkbook.Name }; $s = Retry { $xl.ActiveSheet.Name }; $c = Retry { $xl.ActiveCell.Address($false, $false) }
    $w = Window-Open
    $ok = ($b -eq $book) -and ($s -eq $sheet) -and (($cell -eq '') -or ($c -eq $cell)) -and (($null -eq $window) -or ($w -eq $window))
  } while (-not $ok -and (Get-Date) -lt $deadline)
  $sb = Retry { $xl.StatusBar }
  "{0} {1,-44} active=[{2}]{3}!{4} window={5} status=[{6}]" -f $(if ($ok) { 'ok  ' } else { 'FAIL' }), $label, $b, $s, $c, $w, $sb
  if (-not $ok) { throw "step [$label] failed: expected [$book]$sheet!$cell window=$window, got [$b]$s!$c window=$w" }
}

$failures = New-Object System.Collections.Generic.List[string]
$alertsBefore = $xl.DisplayAlerts
try {
  foreach ($open in @($xl.Workbooks)) { if (@($mainName, $srcName) -contains $open.Name) { throw "ABORT: $($open.Name) is already open ($($open.FullName)); close it first (not saved by this script)." } }
  $xl.DisplayAlerts = $false
  $wb = Retry { $xl.Workbooks.Open($mainPath, 0) }    # UpdateLinks 0
  $xl.DisplayAlerts = $alertsBefore
  if (-not (Is-Fixture $wb)) { throw "ABORT: opened $($wb.FullName), not the fixture." }
  "Main opened: FullName=[$($wb.FullName)]"
  $calc = $wb.Worksheets.Item('Calc')
  "A1 formula (source closed): " + $calc.Range('A1').Formula
  if ($calc.Range('A1').Formula -notlike "='https://*") { throw "ABORT: the fixture's A1 does not hold an https link; the source must be closed and the link rewritten." }

  # --- 1. Source CLOSED: the trace must open it from the https URL ---
  $calc.Activate(); $calc.Range('A1').Select() | Out-Null; Focus-Excel
  Step '^+{[}'  'Ctrl+Shift+[ on A1 (opens OneDrive source)' $mainName 'Calc' 'A1' $true -waitMs 20000
  $srcOpen = $false; foreach ($open in @($xl.Workbooks)) { if ($open.Name -eq $srcName) { $srcOpen = $true; "Source opened by the trace: FullName=[$($open.FullName)] readOnly=$($open.ReadOnly)" } }
  if (-not $srcOpen) { throw "ABORT: the trace did not open $srcName from the https link." }
  Step '{DOWN}' 'Down: [OD_Source]Rates!B1'               $srcName 'Rates' 'B1' $true -waitMs 3000
  Step '{ESC}'  'Esc: back to OD_Main!Calc!A1'            $mainName 'Calc' 'A1' $false

  # --- 2. Source OPEN: links are [OD_Source.xlsx]Rates!B2; expand into the source's own formula ---
  "A2 formula (source open): " + $calc.Range('A2').Formula
  $calc.Range('A2').Select() | Out-Null; Focus-Excel
  Step '^+{[}'  'Ctrl+Shift+[ on A2'                      $mainName 'Calc' 'A2' $true -waitMs 3000
  Step '{DOWN}' 'Down: [OD_Source]Rates!B2'               $srcName 'Rates' 'B2' $true -waitMs 3000
  Step '{RIGHT}' 'Right: expand Rates!B2 (=B1*2)'         $srcName 'Rates' 'B2' $true
  Step '{DOWN}' 'Down: Rates!B1'                          $srcName 'Rates' 'B1' $true
  Step '{LEFT}' 'Left: up to Rates!B2'                    $srcName 'Rates' 'B2' $true
  Step '{ESC}'  'Esc: back to OD_Main!Calc!A2'            $mainName 'Calc' 'A2' $false
}
catch {
  $failures.Add("aborted: $($_.Exception.Message)")
  "ABORTED: $($_.Exception.Message)"
  $_.InvocationInfo.PositionMessage
}
finally {
  "--- putting Excel back ---"
  try { $xl.DisplayAlerts = $alertsBefore } catch { }
  foreach ($n in @($srcName, $mainName)) {
    foreach ($open in @($xl.Workbooks)) { if ($open.Name -eq $n -and (Is-Fixture $open)) { try { $open.Close($false); "Closed $n without saving" } catch { } } }
  }
  "--- trace log lines since start ---"
  if (Test-Path $log) {
    $all = [IO.File]::ReadAllBytes($log); $from = if ($all.Length -ge $logStart) { $logStart } else { 0 }
    [Text.Encoding]::UTF8.GetString($all, $from, $all.Length - $from) -split "`r?`n" | Where-Object { $_ -match "`tTrace(Open|WorkbookOpen|Navigate|Close)`t" }
  }
  if ($failures.Count -eq 0) { "RESULT: PASS" } else { "RESULT: FAIL"; $failures | ForEach-Object { "  - $_" } }
}
if ($failures.Count -gt 0) { exit 1 }
