# Decision spike results

*Run by the owner on 2026-09-28: Microsoft 365 Excel x64, build 16.0.20326, Current Channel. Evidence: `%LOCALAPPDATA%\EmtSpike\log.jsonl` (284 lines at 16:09).*

| Spike | Status |
|---|---|
| A1 — Load the add-in | ✅ Passed. Loaded in 251 ms; the ribbon tab appeared. |
| **K3 — Exact keys and speed** | ✅ **Passed** for every core key (details below). A few coverage keys still need confirming. |
| K2 — Native undo | 🟡 Partly done (details below) |
| K4 — Trace window focus | Pending (runbook A4) |
| K1 — Office.js named keys | Pending (runbook Part B) |

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

**Still open: does it accumulate?** Do three presses give three undo levels?
- If yes, native undo covers our actions *and* anything the user does afterwards. The only loss is history from before the user's first add-in keystroke.
- **Test:** F6 three times, then Ctrl+Z three times.

**Candidate routes if it accumulates:**

| Format | Route | Caveat |
|---|---|---|
| Number formats | `xlcFormatNumber` | No clipboard involved. Cleanest. |
| Font and fill RGB | copy from a pre-formatted cell in the hidden add-in workbook + `ExecuteMso("PasteFormatting")` | Overwrites the clipboard. It also pastes *all* formats unless the template cell copies the target's other properties. |
| Font and fill | C API `xlcFormatFont` / `xlcPatterns` | These take **palette indexes (1–56), not RGB**. Untested. |

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
