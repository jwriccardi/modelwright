# Decision spike results

*Run by the owner on 2026-09-28: Microsoft 365 Excel x64, build 16.0.20326, Current Channel. Evidence: `%LOCALAPPDATA%\EmtSpike\log.jsonl` (284 lines at 16:09).*

| Spike | Status |
|---|---|
| A1 — Load the add-in | ✅ Passed. Loaded in 251 ms; the ribbon tab appeared. |
| **K3 — Exact keys and speed** | ✅ **Passed** for every core key (details below). A few coverage keys still need confirming. |
| K2 — Native undo | Pending (runbook A3) |
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
