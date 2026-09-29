# Decision spike results

*Run by the owner on 2026-09-28: Microsoft 365 Excel x64, build 16.0.20326, Current Channel. Evidence: `%LOCALAPPDATA%\EmtSpike\log.jsonl` (284 lines at 16:09).*

| Spike | Status |
|---|---|
| A1 — Load the add-in | ✅ Passed. Loaded in 251 ms; the ribbon tab appeared. |
| **K3 — Exact keys and speed** | ✅ **Passed** for every core key (details below). A few coverage keys still need confirming. |
| K2 — Native undo | ✅ **Breakthrough (K2b).** A built-in command run **outside macro context** gives full native multi-level undo *and* keeps earlier history, even in volatile workbooks. **K2c:** any COM write wipes the history, so arbitrary formats get Macabacus parity at best with Excel-DNA alone. A hybrid with Office.js (K2d) could beat it. |
| K4 — Trace window focus | ✅ **Passed with variants B (WinForms, focused) and C (hook, Excel keeps focus)**, including cross-workbook navigation. Variant A (WPF) is rejected. The F2 check on C is pending. |
| K1 — Office.js named keys | 🟡 Run 1: **no shortcuts registered at all**, not even the controls. Retest (K1b) isolates whether invalid keys cause the whole file to be rejected. |

## K3 details

**Registration.** All **53** keys registered through `xlcOnKey` without an error. The COM `OnKey` fallback was never needed.

**Core v1 keys.** Each fired and changed the formatting as specified. The owner reports "passed with flying colors."

| Key | Command | Presses | Median ms | Max ms |
|---|---|---|---|---|
| Ctrl+Shift+1 | Number cycle | 73 | 6.5 | 20.7 |
| Ctrl+Shift+2 | Date cycle | 14 | 9.4 | 13.5 |
| Ctrl+Shift+4 | Currency cycle | 41 | 6.5 | 8.7 |
| Ctrl+Shift+5 | Percent cycle | 5 | 1.2 | 1.6 |
| **Ctrl+'** | Font color cycle | 6 | 0.9 | 36.3 |
| Ctrl+Shift+K | Fill cycle | 13 | 1.0 | 18.5 |
| **Ctrl+;** | Blue-black toggle | 4 | 0.8 | 7.7 |
| **Ctrl+, / Ctrl+.** | Decimals (log only) | 25 / 19 | 0.0 | 0.0 |

**Latency on a large selection:** 83 presses on a 4,770-cell selection.
- Median **7.1 ms**, p95 **9.8 ms**, max 20.7 ms.
- The target was p95 ≤ 50 ms, so it **passes by 5×**.
- The time is measured from when the command starts until the COM write finishes. How long Windows takes to deliver the key press isn't included, but no lag was noticed.

**Coverage keys that fired:**
- Alt+Shift+; , .
- Alt+Shift+=
- Ctrl+Alt+Shift+, and .
- Ctrl+Alt+Shift+Up
- Alt+Shift+PgUp
- Ctrl+F2
- Alt+F12
- Ctrl+Alt+Home
- Ctrl+Shift+Y

**Registered but never fired. Either they weren't pressed, or they don't fire (needs the owner to confirm):**
- **Ctrl+Shift+8** (Multiple cycle)
- **Ctrl+Alt + punctuation:** `'` `.` `=` `-` `\` `[` `]`
  - These include **Show All Precedents (Ctrl+Alt+[)** and **Clear Arrows (Ctrl+Alt+\\)**, both rated 5/5 by Macabacus.
- Ctrl+Alt+Shift + `'` `[` `]` `=` `-` `Ins`
- Alt+Shift+-
- Ctrl+Shift+] and Ctrl+Shift+\\ (pressed only in A4)
- Ctrl+Shift+[ (pressed only in A4)

**One possible cause.** On Windows, Ctrl+Alt is treated as AltGr. So Ctrl+Alt+punctuation may be turned into a character before Excel's `OnKey` sees it, and those keys might need the keyboard hook instead. The fact that Ctrl+Alt+Shift+, and . fired makes the picture mixed. **Retest** with the step-by-step list in the runbook addendum (below).

## K3 retest (19:15)

| Key | Result |
|---|---|
| Ctrl+Shift+8 (Multiple) | ✅ fired ×6, median ~1.4 ms (it just hadn't been pressed earlier) |
| **Ctrl+Alt+[** (Show All Precedents) | ✅ fired |
| **`Ctrl+Alt+\`** (Clear Arrows) | ✅ fired |
| Ctrl+Alt+] (Show All Dependents) | ✅ fired in the 19:23 recheck |
| Ctrl+Alt+' (Comment Formula) | ✅ fired in the 19:23 recheck |
| Ctrl+Alt+., Ctrl+Alt+=, Ctrl+Alt+-, Ctrl+Alt+Shift+[, Ctrl+Alt+Shift+', Alt+Shift+- | No log entry from the 19:15 pass. The owner saw status-bar messages, but the spike never clears the status bar, so those were probably leftovers. The same pass also missed Ctrl+Alt+] and ', which then fired on recheck. So these are **very likely fine**. They're not v1 keys; confirm them when convenient. |

So the AltGr theory is **disproved for Ctrl+Alt+[ and `Ctrl+Alt+\`**. Ctrl+Alt+punctuation *can* be bound through `OnKey`.

**K3 verdict: PASS.** Every v1 key and every Ctrl+Alt+punctuation key tested so far binds through `xlcOnKey`. Latency is about 7 ms median.

**Lessons for the product:**
- Clear or time out status-bar feedback, so a stale message can't pass for a successful key press.
- Add an in-product "key test" diagnostic that lists every binding and whether it last fired.


## K2 details (19:26–19:33)

**The probe.** It reads `GetEnabledMso("Undo")` before and after each action. The owner then pressed Ctrl+Z twice after typing `1` and `2`.

| Key | Method | Undo enabled before → after | Owner's Ctrl+Z observation |
|---|---|---|---|
| F5 | COM `Font.Bold` (control) | T → **F** | as expected (wiped) |
| F6 | `ExecuteMso("Bold")` directly | T → **T** | 1st Ctrl+Z undid the format*; **2nd did nothing** (the earlier typing was lost) |
| F7 | `ExecuteMso("Bold")` via `QueueAsMacro` | T → **F** (flips at `deferred:afterExecute`) | wiped |
| F8 | `ExecuteMso("PercentStyle")` | T → **T** | same as F6 |
| F9 | Copy + COM `PasteSpecial(xlPasteFormats)` (rerun at 19:32) | T → **F** (flips at `afterPasteSpecial`) | Ctrl+Z did **not** remove the percent format (wiped) |
| F10 | Copy + `ExecuteMso("PasteFormatting")` | T → **T** | same as F6. The `PasteFormats` idMso doesn't exist; `PasteFormatting` does. |
| F11 | `xlcFormatNumber("0.0%")` (C API) | T → **T** | same as F6 |

\*The owner's "second Ctrl+Z did nothing" implies the first one did something. Confirmation is pending.

**What we can conclude so far:**
- **Running any add-in command loses Excel's earlier undo history.**
- **Built-in commands (`ExecuteMso`) and the C API (`xlcFormatNumber`) put our own change onto Excel's native stack as a normal entry.** So Ctrl+Z undoes it natively, with Excel's own Undo list and redo.
- COM writes, including COM `PasteSpecial`, add nothing.
- Deferring through `QueueAsMacro` breaks it: F7 is wiped while F6 is kept.

**Accumulation test (19:45–19:46): ✅ it accumulates.**
- **On a hard-coded cell:** F6 (built-in Bold) ×3 toggled bold three times, and **Ctrl+Z ×3 undid all three, one step at a time**. That is native multi-level undo for our own actions, with no custom undo stack.
- **On a `=RAND()*1000000` cell:** F6 ×3 toggled bold. But the **first Ctrl+Z unbolded and recalculated, and after that nothing more could be undone.**
  - The add-in logged nothing during the undo, and it has no calculation event handlers.
  - **Control test (native Ctrl+B ×3, then Ctrl+Z ×3 on the RAND cell, add-in still loaded):** ✅ all three undo, and each one recalculates. So Excel itself handles volatile cells fine. The problem is specific to undo entries **created inside an add-in macro**.
  - **T1, manual calculation, RAND cell:** only the 1st Ctrl+Z works, and the volatile cells are marked dirty (stale-value strikethrough). **So recalculation isn't the trigger.** Undoing an entry created by a macro dirties the volatile cells, and that ends the undo history.
  - **T2, volatile formula elsewhere, plain cell formatted:** only the 1st Ctrl+Z works. **Any volatile formula in the workbook triggers it.**

**Revised K2 finding.** Built-in commands run from an `OnKey` macro give:
- **full native multi-level undo in workbooks without volatile functions**;
- **only one level in workbooks with them** (OFFSET, INDIRECT, TODAY, RAND… are common in models).

**Next probe (K2b):** trigger the built-in command **outside a macro context**. A keyboard hook catches the key, and `ExecuteMso` runs from Excel's message loop via a posted callback. The hope is that Excel then records it like a user Ctrl+B. The ribbon-callback path (no `QueueAsMacro`) is included for comparison.

*Earlier question, kept for the record:* do three presses give three undo levels?
- If yes, native undo covers our actions *and* anything the user does afterwards. The only loss is history from before the user's first add-in keystroke.
- **Test:** F6 three times, then Ctrl+Z three times.

**Candidate routes if it accumulates:**

| Format | Route | Caveat |
|---|---|---|
| Number formats | `xlcFormatNumber` | No clipboard involved. Cleanest. |
| Font and fill RGB | copy from a pre-formatted cell in the hidden add-in workbook + `ExecuteMso("PasteFormatting")` | Overwrites the clipboard. It also pastes *all* formats unless the template cell copies the target's other properties. |
| Font and fill | C API `xlcFormatFont` / `xlcPatterns` | These take **palette indexes (1–56), not RGB**. Untested. |


## K2b: outside macro context (21:17–21:23, build `bin\K2b\`)

Excel was on Automatic calculation, with `=RAND()` present in the workbook.

| Variant | How it's triggered | Owner's Ctrl+Z observation |
|---|---|---|
| K2.8 | **Thread keyboard hook** (Ctrl+Alt+Shift+Z) → `BeginInvoke` to a hidden WinForms control → `ExecuteMso("Bold")` | ✅ **Bold ×3, then Ctrl+Z ×3: all three undone** |
| K2.10 | Ribbon callback runs `ExecuteMso("Bold")` directly (no `QueueAsMacro`) | ✅ **all three undone** |
| K2.9 | Hook → copy the hidden template cell (0.0%) → `ExecuteMso("PasteFormatting")` | ✅ Typed `1`, `2`, applied the format, Ctrl+Z ×2: **the 1st undid our format change, the 2nd undid the typing of "2"**. So the **earlier history is kept.** **Confirmed on rerun (22:02–22:04, build K4b): the 1st Ctrl+Z removed the percent format, the 2nd removed the "2".** K2.8 and K2.10 also reproduced: bold ×3 without recalculation, then Ctrl+Z ×3 all toggle, and each undo recalculates. That recalculation is normal Excel behavior; the native Ctrl+B control does the same. |

**Conclusion.**
- **Macro context is what breaks undo.** Built-in commands dispatched from Excel's normal message loop are recorded exactly like user actions: multi-level, and no history lost, even with volatile formulas.
- **Architecture consequence:** catch keys with a thread keyboard hook (which also allows the exact Macabacus keys), post the work to the message loop, and apply formatting through built-in commands.

**Side finding.** Creating the hidden template workbook at startup wiped undo (`hiddenWb:afterCreate enabled=False`). That's harmless at startup, when there's nothing to undo yet, but the product should create any template workbook only at load.

**Still open (K2c):**
- (a) Does a COM write from the message-loop context (e.g. `Font.Color`) wipe the history, and is it undoable?
- (b) Can a COM write to the *hidden template* (to clone the target's formats before pasting) be done without wiping the history?
- (c) Paste Formatting copies *all* formats and overwrites the clipboard. On a mixed selection it makes everything uniform.


## K2c: COM writes outside macro context (21:38–21:41, build stamp `K2c (2026-09-28)`)

| Variant | Undo enabled before → after | Owner's observation |
|---|---|---|
| K2.11: hook → COM `Selection.Font.Color` | T → **F** | Ctrl+Z did nothing: the history was wiped and the change wasn't undoable |
| K2.12: hook → COM writes that clone the target's formats onto the **hidden template** → copy → `ExecuteMso("PasteFormatting")` | T → **F at `afterTemplateWrite`** → T after the paste | One Ctrl+Z (the paste), then nothing: **the history was wiped by the write to the hidden workbook** |
| Mixed-selection check (C) | not run | The fills were set with the COM fill cycle (which also wipes), and no Ctrl+Alt+Shift+X press was logged. Superseded by the finding below. |

**Rule established.** **Any COM write to any cell, even in a hidden add-in workbook and even outside macro context, erases Excel's undo history.** Only built-in commands (`ExecuteMso`) dispatched outside macro context are recorded like user actions.

**Consequences for arbitrary formats** (exact number-format codes and exact RGB colors):
- **Built-in commands exist only for fixed formats:** Bold, Percent Style, Comma Style and similar.
- **Template + Paste Formatting:**
  - The templates must be pre-built at load time, because writing them later wipes history.
  - Paste Formatting copies **all** formats, so a number-format cycle would reset the target's font color, fill and borders. **Rejected.**
- So with Excel-DNA alone, the best we can do for arbitrary formats is **Macabacus parity**: our own multi-level undo stack for our changes, with Excel's earlier history lost at the first keystroke.

**Candidate that might beat parity: a hybrid (K2d, needs approval).**
- The Excel-DNA thread hook catches the exact keys and forwards the command over a local channel to an Office.js add-in (shared runtime) in the same Excel instance.
- Office.js applies `numberFormat` / `font.color` / `fill.color`, which are native-undoable per ExcelApi 1.20.
- The forwarding step makes no workbook change, so nothing gets wiped.
- Open questions:
  - Does a WebView2 page served over https accept a `ws://localhost` connection?
  - What latency does the round trip add?
  - Is the Office.js runtime always loaded?


## K4, run 1 (21:50–21:54, build K2c)

**Variant A (WPF, focused; Ctrl+Shift+[):**
- **The arrow keys drove the list.** About 40 navigation events were logged.
- **Across sheets in the same workbook:** Goto took 2–25 ms, and after each Goto the window **kept** foreground and keyboard focus (`foregroundIsOurs=true`).
- **Into `EMT_Fixture_B`:** Goto took 65–116 ms. Excel brought B's own top-level window to the front (Excel is SDI, one window per workbook), and our window, owned by A's window, went **behind** it. From then on the arrow keys went to the grid. That matches what the owner saw.
- **Hidden sheet:** `Application.Goto` fails with "Unable to get the Goto property" (expected). The product needs a hidden-sheet badge, or to unhide temporarily.

**Variants B (WinForms) and C (hook): not really tested.** When they were opened, the active cell was in EMT_Fixture_B (`Data!C3`, a constant `10`), left there by the earlier navigation. The formula had no references, so the lists were empty. That's a flaw in the test design.

**Fixes in build K4b:**
- After each Goto, if Excel's active window changed, the trace window is **re-owned** to the new workbook window (`GWLP_HWNDPARENT`) and raised to the top. The hook variant is raised *without* activation.
- Focused variants take focus back automatically after a workbook switch.
- Trace refuses to open on a cell with no formula, and says so in the status bar.


## K4, run 2 (22:11–22:12, build K4b)

| Variant | Result |
|---|---|
| A: WPF, focused | ❌ Opened twice; **no navigation events at all**, so the keys never reached the WPF list. Hosting WPF on Excel's thread without a WPF dispatcher loop is unreliable. **Rejected.** |
| **B: WinForms, focused** | ✅ Full up/down passes, 3 sessions. Same workbook: 4–36 ms per step. **Jump into another workbook:** 41–94 ms. The window is re-owned to the new workbook window and focus is taken back (`react->fg`) every time. Enter and Esc behave correctly. |
| **C: window doesn't take focus, plus thread keyboard hook** | ✅ Full passes. **The hook kept receiving keys while EMT_Fixture_B's window was in front.** It's thread-local, so any Excel window will do. Re-owning kept the window visible. Same workbook: 6–19 ms per step; cross-workbook: 78–108 ms. Enter closed it and kept the selection. The F2 pass-through hasn't been exercised yet. |
| Hidden sheet row (all variants) | `Application.Goto` fails, as expected. The product shows a badge (optionally, "Unhide rows & columns"). |

**Decision.**
- **Variant C is the target design for Trace In:** Macabacus's own model, and the only one that allows F2 editing in Point mode with the window open.
- **Variant B is the fallback**, if the hook proves fragile alongside other add-ins (Macabacus documents conflicts with Workshare and Anaplan).
- **UI toolkit:** variant C's window *is* WPF, and it works, because its keys come from the hook rather than from WPF input. **WPF is fine for rendering; only WPF keyboard focus is unreliable on Excel's thread.** So:
  - C can use WPF (richer tree UI) or WinForms.
  - Fallback B must use WinForms.


## K1, run 1 (22:33, Office.js, Excel 16.0.20326.20158, PC)

All three add-ins loaded:
- `SharedRuntime 1.1 = true`
- `KeyboardShortcuts 1.1 = true`
- actions associated

**But `Office.actions.getShortcuts()` returned `{}` for every add-in, and no key fired, not even the control keys** (Ctrl+Alt+Shift+L/K/J, Ctrl+Alt+Up).

`areShortcutsInUse` and `replaceShortcuts` failed with "invalid format" for every call. That's inconclusive: the calls referenced actions Office had never registered, and batches mixed invalid strings.

**The test-design flaw:** even the "Single" file contained one candidate (`Ctrl+Shift+BracketLeft`). So "one invalid key rejects the whole file" and "the file never loaded" couldn't be told apart.

**K1b retest:**
- "Single" now serves a **control-only** file (`shortcuts-c2.json`) and the manifest version is bumped (1.0.0.1).
- Every key string gets its own `areShortcutsInUse` check.
- The *registered* control action is re-keyed with `replaceShortcuts` to each candidate. `Ctrl+Alt+Shift+M` is the positive control. Afterwards everything is reverted.
- The server logs every file Office fetches.

## Side findings
- **Undo list can't be read.** Reading Excel's undo *list* through `CommandBars("Standard").Controls("&Undo")` or `FindControl(128)` fails with E_FAIL on this build. `GetEnabledMso("Undo")` works (it returned `false` at startup, as expected). So K2 relies on the Undo-enabled flag plus the owner's Ctrl+Z observations.
- **Helper workbook.** The hidden helper workbook opened as "Book2", which confirms the startup side effect noted in the spike README.

## Runbook addendum: K3 retest (2 min, do before A3)
With the add-in loaded, select any cell. Press each key below **once**, and check that the status bar says "EMT spike: … fired":

1. Ctrl+Shift+8
2. Ctrl+Alt+[
3. Ctrl+Alt+]
4. Ctrl+Alt+\
5. Ctrl+Alt+'
6. Ctrl+Alt+.
7. Ctrl+Alt+=
8. Ctrl+Alt+-
9. Ctrl+Alt+Shift+[
10. Ctrl+Alt+Shift+'
11. Alt+Shift+-

✍️ Which ones did **not** update the status bar? ______
