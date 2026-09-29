# Excel-DNA decision spike (K2 / K3 / K4)

**Throwaway code.** It exists only to answer spikes K2, K3 and K4 in [`docs/PLAN.md` §5](../../docs/PLAN.md). It is not the product scaffold. Do not build on it.

## Build

Requirements: .NET SDK (tested with 10.0.401). Visual Studio is not needed; the .NET Framework 4.8 reference assemblies come from NuGet.

```powershell
cd spikes\exceldna
dotnet build -c Release
```

Output. Load the **packed 64-bit** file:

```
spikes\exceldna\bin\Release\net48\publish\EmtSpike-AddIn64-packed.xll
```

(`bin\Release\net48\EmtSpike-AddIn64.xll` is the unpacked version. It needs `EmtSpike.dll` and the `.dna` file beside it.)

## Load / unload

**Before testing:** disable Macabacus (or change its keys). Otherwise, whichever add-in calls `OnKey` last owns the key.

- **Load:** File > Options > Add-ins > Manage: **Excel Add-ins** > Go… > **Browse…** > select `EmtSpike-AddIn64-packed.xll` > OK. Make sure it is ticked in the list.
  - You should see an **EMT Spike** ribbon tab and a status-bar message "EMT spike loaded (… ms)".
- **Unload:** same dialog; untick **EMT Spike**. This runs `AutoClose`, which unregisters every key and closes the hidden helper workbook.
  - To remove it from the list for good, untick it, then Browse to it again and delete/move the file (Excel then offers to remove the entry).
- If Excel is closed with the add-in ticked, it loads again next time Excel starts.

## Log

- Every event is appended as one JSON line to `%LOCALAPPDATA%\EmtSpike\log.jsonl` (ribbon: **Open log folder**).
- `autoOpen` records the Excel version and build, bitness, the xll path and the load time.
- `register` records one line per key: `method` is `xlcOnKey`, `COM OnKey` (the fallback) or `failed`, plus any error text.
- `cmd` records `{cmd, key, elapsedMs, selectionCellCount, method, detail}` each time a command fires. `cmdError` is written when a command fails. Commands never throw; failures also go to the status bar.
- The status bar shows e.g. `EMT spike: Ctrl+' FontColorCycle fired (3.2 ms)`. It is not reset afterwards.

## K3 — exact Macabacus keys and speed

`elapsedMs` covers the time from command entry until the COM write completes. Each cycle reads only the active cell and then writes the whole selection with one COM call. The next-item rule: if the current value matches item *k*, apply *k+1* (wrapping round); otherwise apply the first item. The log records the format read back (`detail.current`). Excel may normalise format codes, which would make a match fail.

| Key | OnKey | Command | Action |
|---|---|---|---|
| Ctrl+Shift+1 | `^+1` | NumberCycle | `_(#,##0_)_%;…` → 0.0 → 0.00 → `#,##0;(#,##0);"–";@` |
| Ctrl+Shift+2 | `^+2` | DateCycle | `yyyy"A"` → `mm-dd-yyyy` → `yyyy-mm-dd` → `mmm-yy` |
| Ctrl+Shift+4 | `^+4` | CurrencyCycle | `[$$]` 0 dp → 2 dp |
| Ctrl+Shift+5 | `^+5` | PercentCycle | 0.0% → 0% → 0.00% |
| Ctrl+Shift+8 | `^+8` | MultipleCycle | `0.0"x"` → `0.00"x"` |
| Ctrl+' | `^'` | FontColorCycle | blue → green → purple → red → white → black |
| Ctrl+Shift+K | `^+k` and `^+K` | FillCycle | (201,218,248) → (210,242,255) → (244,204,204) → (252,229,205) → (28,69,135) → no fill |
| Ctrl+; | `^;` | BlueBlackToggle | font blue ↔ black |
| Ctrl+, / Ctrl+. | `^,` / `^.` | DecimalsIncrease / Decrease | log only |
| Ctrl+Shift+[ | `^+{[}` and `^{{}` | ProPrecedents | opens trace window A (WPF), see K4 |
| Ctrl+Alt+[ | `^%{[}` | ShowAllPrecedents | log only |
| Ctrl+Shift+\ | `^+\` and `^\|` | LastAuditedCell | Goto the root of the last trace |

`^+K` and `^{{}` are hedges. Each alternative binding has its own macro, so the log's `key` field shows which binding actually fired. If `^+k` and `^+K` turn out to be the same key to Excel, the one registered second wins.

**Coverage set** (log + status bar only, no action). These keys answer "can every Macabacus punctuation or named key be bound?":

| Key | OnKey | Macabacus command |
|---|---|---|
| Alt+Shift+; | `%+;` | RatioCycle |
| Alt+Shift+, / . | `%+,` / `%+.` | ShiftDecimalLeft / Right |
| Ctrl+Alt+Shift+' | `^%+'` | BorderColorCycle |
| Ctrl+Alt+. | `^%.` | AutoColorCycle |
| Ctrl+Alt+' | `^%'` | CommentFormula |
| Ctrl+Shift+] | `^+{]}` | TraceOut |
| Ctrl+Alt+Shift+[ / ] | `^%+{[}` / `^%+{]}` | AutoTracePrecedents / Dependents |
| Ctrl+Alt+] | `^%{]}` | ShowAllDependents |
| Ctrl+Alt+\ | `^%\` | ClearArrows |
| Ctrl+Alt+= / - | `^%=` / `^%-` | ZoomIn / ZoomOut |
| Ctrl+Alt+Shift+, / . | `^%+,` / `^%+.` | GoToMin / GoToMax |
| Alt+Shift+= / Ctrl+Alt+Shift+= | `%+=` / `^%+=` | ExpandAllRows / Columns |
| Alt+Shift+- / Ctrl+Alt+Shift+- | `%+-` / `^%+-` | CollapseAllRows / Columns |
| Ctrl+Alt+Shift+Up | `^%+{UP}` | TopBorder |
| Alt+Shift+PgUp | `%+{PGUP}` | RowHeightCycle |
| Ctrl+Alt+Shift+Ins | `^%+{INSERT}` | InsertColumn |
| Ctrl+F2 | `^{F2}` | AnchorFormulaCycle |
| Alt+F12 | `%{F12}` | QuickSaveAs |
| Ctrl+Alt+Home | `^%{HOME}` | FirstSheet |
| Ctrl+Shift+Y | `^+y` | BinaryCycle |

**Spike-only keys.** These sit on Ctrl+Alt+Shift+F-keys to stay clear of Macabacus's Ctrl+Alt+Shift+O/T/6/7:

| Key | Command |
|---|---|
| Ctrl+Alt+Shift+F1 | Override: re-register every key and log each result (also on the ribbon) |
| Ctrl+Alt+Shift+F2 | Trace window B (WinForms) |
| Ctrl+Alt+Shift+F3 | Trace window C (non-activating + keyboard hook) |
| Ctrl+Alt+Shift+F5…F12 | K2 variants (below) |

**Test.**
1. Load the add-in. Check that every `register` line has `method: "xlcOnKey"`.
2. Select about 1,000 cells and press each key 30 times.
3. From the `cmd` lines, compute p95 `elapsedMs` per command. Pass: p95 ≤ 50 ms and 100% of presses fire.

## K2 — native undo

Each variant writes these undo snapshots: `before`, `after` (still inside the command), `afterQueued` (the next `QueueAsMacro`) and `after500ms`. Each snapshot is logged as `evt: "undo"` with `{phase, variant, enabled, count, items[], source, errors}`:
- `enabled` comes from `CommandBars.GetEnabledMso("Undo")`.
- `count` and `items` come from `CommandBars("Standard").Controls("&Undo")`, falling back to `FindControl(Id:=128)`.

| Key | Ribbon | Variant |
|---|---|---|
| Ctrl+Alt+Shift+F5 | 1 COM Bold (control) | `Selection.Font.Bold = !bold` via COM |
| Ctrl+Alt+Shift+F6 | 2 ExecuteMso Bold | `CommandBars.ExecuteMso("Bold")` directly in the command |
| Ctrl+Alt+Shift+F7 | 3 ExecuteMso Bold (deferred) | Same, but in a later `QueueAsMacro`. Adds `deferred:beforeExecute/afterExecute` snapshots |
| Ctrl+Alt+Shift+F8 | 4 ExecuteMso PercentStyle | `ExecuteMso("PercentStyle")`. Logs an error if the idMso is invalid |
| Ctrl+Alt+Shift+F9 | 5 Copy + PasteSpecial formats | Hidden add-in workbook `A1` (pre-formatted 0.0%) → `.Copy()` → `Selection.PasteSpecial(xlPasteFormats)` → `CutCopyMode=False`. Adds `afterCopy` and `afterPasteSpecial` snapshots |
| Ctrl+Alt+Shift+F10 | 6 Copy + ExecuteMso PasteFormatting | Same copy, then `ExecuteMso` of `PasteFormatting` or `PasteFormats` (logs which idMsos exist and which ran) |
| Ctrl+Alt+Shift+F11 | 7 xlcFormatNumber (C API) | `XlCall.Excel(xlcFormatNumber, "0.0%")` |
| Ctrl+Alt+Shift+F12 | Log undo snapshot | Snapshot only, no action |

The hidden helper workbook is created once, via `QueueAsMacro` from `AutoOpen`. Its creation is logged with its own `hiddenWb:beforeCreate/afterCreate` snapshots, so creation never happens inside a test.

**Test.** For each variant:
1. Type a value in a cell and press Enter.
2. Press Ctrl+Alt+Shift+F12 and check that `count ≥ 1`.
3. Run the variant.
4. Press Ctrl+Z twice. Pass: both the format and the typed value are undone.
5. Compare with the logged `items[]`.

## K4 — trace window keeps keyboard focus?

1. Ribbon **Create K4 fixtures**. This writes `EMT_Fixture_B.xlsx` (sheet `Data`, C3:C5) and `EMT_Fixture_A.xlsx` to `%LOCALAPPDATA%\EmtSpike`, leaves both open and selects `EMT_Fixture_A.xlsx` Sheet1!A1. That cell's formula refers to Sheet2, to the other workbook (`[EMT_Fixture_B.xlsx]Data!C3`), to its own sheet, to a range, and to a **hidden** sheet (`Hidden!A1`, so the Goto should fail and be logged).
2. Open a trace window on A1:

| Variant | Open with | How it works |
|---|---|---|
| **A — WPF, focused** | Ctrl+Shift+[ (or `^{{}`), ribbon **Trace (WPF)** / **Trace (WPF, re-activate)** | WPF `Window`, owner = `Application.Hwnd` through `WindowInteropHelper`, `ElementHost.EnableModelessKeyboardInterop`. The ListBox has focus. |
| **B — WinForms, focused** | Ctrl+Alt+Shift+F2, ribbon **Trace (WinForms)** / **Trace (WinForms, re-activate)** | Modeless `Form`, owner = a `NativeWindow` wrapping the Excel HWND. |
| **C — non-activating + hook (Macabacus-style)** | Ctrl+Alt+Shift+F3, ribbon **Trace (hook)** | WPF window with `WS_EX_NOACTIVATE` and `ShowActivated=false`, so **Excel keeps focus**. A thread-local `WH_KEYBOARD` hook on Excel's main thread swallows *unmodified* Up/Down/Left/Right/Enter/Esc and routes them to the list. Every other key passes through. **F2** passes through and switches arrows to Excel until Enter/Esc ends the edit (Enter also re-reads the formula). The hook is removed when the window closes and in `AutoClose`. |

3. Keys in the window: **Up/Down** select a reference, then `QueueAsMacro` activates its workbook window if needed and runs `Application.Goto`. **Enter** closes and keeps the selection. **Esc** goes back to the root cell and closes.
4. The **Re-activate after Goto** checkbox (A and B only) calls `Activate()` and focuses the list after each Goto.

What is logged:
- `k4nav`:
  - `elapsedMs` from the key press to the end of the Goto
  - `gotoError` (e.g. the hidden sheet)
  - `activeBook`
  - `afterGoto`: `{foregroundIsOurs, win32FocusIsOurs, frameworkFocus}`
  - `afterReactivate`, when the checkbox is on
- `k4navLate`: the same focus check about 250 ms later, because Excel may take focus back after the macro returns.
- Variant C also logs every intercepted key (`hookKey`: swallowed / F2 editing / edit ended), with the foreground and focus HWNDs at that moment. These entries show whether the hook still receives keys after a Goto into the other workbook.

**Pass (PLAN K4).** Up/Down stay in the trace (window or hook) both for another sheet and for the other workbook.

Known limits of variant C:
- If you start typing into a cell *without* F2, the arrows are still swallowed.
- Left/Right are swallowed but only logged; there is no expand/collapse.

## Files

| File | Purpose |
|---|---|
| `EmtSpike.csproj` | net48, ExcelDna.AddIn 1.9.0, WPF + WinForms, 64-bit packed xll only |
| `AddIn.cs` | AutoOpen/AutoClose, key tables, `xlcOnKey` registration with COM `OnKey` fallback, runtime coverage macros |
| `Commands.cs` | `[ExcelCommand]` entry points, the `Run` wrapper (timing, log, status bar), ribbon dispatch |
| `Cycles.cs` | number/date/currency/percent/multiple/font/fill cycles |
| `K2.cs`, `UndoProbe.cs` | undo variants and the native undo-list reader |
| `Trace.cs` | reference regex and resolution, the shared navigation session, keyboard hook, P/Invoke |
| `TraceWindowWpf.cs`, `TraceForm.cs` | trace windows A/C and B |
| `Fixtures.cs` | K4 fixture workbooks |
| `Ribbon.cs` | "EMT Spike" tab |
| `Log.cs` | JSON-lines logger and small COM helpers |
