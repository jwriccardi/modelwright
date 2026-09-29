# Decision spikes K1–K4: owner runbook

**Time:** about 40 minutes.
**What you record:** almost nothing. The Excel-DNA add-in logs everything to `%LOCALAPPDATA%\EmtSpike\log.jsonl`, and Claude reads that file. You only note what you *see*, in the ✍️ boxes. One word each is fine.

> **Safety:** everything here is throwaway test code.
> - Part A (Excel-DNA) is loaded and unloaded through Excel's own Add-ins dialog.
> - Part B (Office.js) adds three registry values and a localhost development certificate. `cleanup.ps1` removes both.
> - Nothing touches your real workbooks. **Save and close your work before starting.**

---

## 0. Prep (2 min)
1. Save and close all your workbooks.
2. **Turn Macabacus off**, because it uses the same keys: **File › Options › Add-ins**, set **Manage: COM Add-ins**, click **Go…**, and **untick "Macabacus COM Add-in"**. Click OK.
3. Close Excel completely and reopen it with a **blank workbook**.

---

## Part A — Excel-DNA test add-in (K2, K3, K4), ~25 min

### A1. Load it
1. **File › Options › Add-ins**, set **Manage: Excel Add-ins**, click **Go… › Browse…**
2. Pick `C:\Users\JohnRiccardi\source\repos\excel-modeling-toolkit\spikes\exceldna\bin\Release\net48\publish\EmtSpike-AddIn64-packed.xll`, then click OK.
3. You should see a new **EMT Spike** ribbon tab. A hidden helper workbook is also created, which is expected.

✍️ Did the EMT Spike tab appear? ______

### A2. K3: do the exact Macabacus keys work, and how fast?
1. In the blank workbook, type some numbers in **A1:A5**, some negative, and select A1:A5.
2. Press each key below **three or four times**. Watch the cell formats change and the **status bar** at the bottom-left, which says "EMT spike: … fired (x ms)".

| Key | Expected |
|---|---|
| Ctrl+Shift+1 | Number formats cycle |
| Ctrl+Shift+2 | Date formats |
| Ctrl+Shift+4 | Currency |
| Ctrl+Shift+5 | Percent |
| Ctrl+Shift+8 | Multiple (x) |
| **Ctrl+'** | Font color: blue → green → purple → red → white → black |
| Ctrl+Shift+K | Fill: light blue → cyan → pink → peach → navy → none |
| **Ctrl+;** | Font blue ↔ black |
| **Ctrl+,** and **Ctrl+.** | Status bar only |

3. **Coverage keys:** these have no visible effect; they only put a message in the status bar. Press each once:
   - Alt+Shift+;
   - Alt+Shift+, and Alt+Shift+.
   - Ctrl+Alt+Shift+'
   - Ctrl+Alt+. and Ctrl+Alt+'
   - Ctrl+Shift+]
   - Ctrl+Alt+Shift+[ and Ctrl+Alt+Shift+]
   - Ctrl+Alt+[ and Ctrl+Alt+] and Ctrl+Alt+\
   - Ctrl+Alt+= and Ctrl+Alt+-
   - Ctrl+Alt+Shift+, and Ctrl+Alt+Shift+.
   - Alt+Shift+= and Alt+Shift+-
   - Ctrl+Alt+Shift+Up
   - Alt+Shift+PgUp
   - Ctrl+F2
   - Alt+F12
   - Ctrl+Alt+Home
   - Ctrl+Shift+Y
4. **Speed on a big selection:** select **A1:J100** and press **Ctrl+Shift+1 thirty times**.

✍️ Did any key do nothing, or do something *else* such as a native Excel action? Which ones? ______
✍️ Did anything feel laggy? ______

### A3. K2: can formatting keep Excel's own Undo history?
For each of the 7 keys below, repeat these steps:
1. Click an empty cell, **type `1` and press Enter**. Type **`2`** in the next cell and press Enter. (This gives Excel something to undo.)
2. Click back on the `1` cell.
3. Press the **test key**.
4. Press **Ctrl+Alt+Shift+F12** (logs the undo state).
5. Press **Ctrl+Z twice**, and note what got undone.

| Test key | What it tries | ✍️ What did Ctrl+Z ×2 undo? (e.g. "format + the 2", "nothing", "only the format") |
|---|---|---|
| Ctrl+Alt+Shift+F5 | Plain code, bold (the control: expected to wipe undo) | |
| Ctrl+Alt+Shift+F6 | Built-in Bold command | |
| Ctrl+Alt+Shift+F7 | Built-in Bold command, delayed | |
| Ctrl+Alt+Shift+F8 | Built-in Percent Style | |
| Ctrl+Alt+Shift+F9 | Copy format + Paste Special Formats | |
| Ctrl+Alt+Shift+F10 | Copy + built-in "Paste Formatting" | |
| Ctrl+Alt+Shift+F11 | Old Excel 4 format command | |

### A4. K4: can the trace window keep the keyboard?
1. On the **EMT Spike** tab, click **Create K4 fixtures**. Two small workbooks open, and **EMT_Fixture_A › Sheet1!A1** is selected. A1's formula points to Sheet2, a range, another workbook and a hidden sheet.
2. **Variant A:** with A1 selected, press **Ctrl+Shift+[**. A trace window opens. Then:
   - Press **Down** repeatedly through every row, then **Up** back to the top. Excel should jump to each precedent, including into **EMT_Fixture_B**.
   - Press **Enter**. The window closes, and Excel should stay on that cell.
   - Reopen it, move **Down** twice, then press **Esc**. The window closes, and Excel should return to A1.
3. **Variant A, re-activate version:** click **Trace (WPF, re-activate)** on the ribbon and repeat step 2.
4. **Variant B:** press **Ctrl+Alt+Shift+F2**, repeat step 2, then do the same with **Trace (WinForms, re-activate)**.
5. **Variant C (Macabacus-style):** press **Ctrl+Alt+Shift+F3** and repeat step 2. Then reopen it and also try this:
   - Press **F2**. Excel should enter edit mode on A1's formula, with the window still open.
   - Press the arrow keys. They should now move *inside Excel* (Point or Edit mode).
   - Press **Esc** to leave edit mode.
   - Press **Down**. The arrow keys should drive the window again.

| Variant | ✍️ Did Up/Down *always* move in the window, including after jumping to the other workbook? | ✍️ Enter / Esc behave as described? | ✍️ Anything odd (flicker, focus lost, window hidden behind Excel)? |
|---|---|---|---|
| A (WPF) | | | |
| A (re-activate) | | | |
| B (WinForms) | | | |
| B (re-activate) | | | |
| C (hook) | | | F2 worked? |

### A5. Unload it
1. **File › Options › Add-ins**, set **Manage: Excel Add-ins**, click **Go…**, **untick EmtSpike**, then click OK.
2. Close the fixture workbooks **without saving**.

---

## Part B — Office.js key test (K1), ~15 min

1. **Close Excel.**
2. Open **PowerShell** and run:
   ```powershell
   cd C:\Users\JohnRiccardi\source\repos\excel-modeling-toolkit\spikes\officejs-keys
   .\setup.ps1
   ```
   Accept the Windows **certificate** prompt. It's a localhost-only development certificate.
3. Open a **second** PowerShell window and **leave it running**:
   ```powershell
   cd C:\Users\JohnRiccardi\source\repos\excel-modeling-toolkit\spikes\officejs-keys
   .\serve.ps1
   ```
4. Open Excel with a blank workbook. On the **Home** tab you should see three buttons: **Show K1 Named**, **Show K1 Literal** and **Show K1 Single**.
   - If you don't see them: **Home › Add-ins › More Add-ins › My Add-ins**, and look under **Developer Add-ins**.
5. **Test the three add-ins one at a time.** Open one pane, click a cell in the grid, then press each key in its list. If Office shows a **"which action?" conflict dialog**, choose the **EMT** option.
   - The pane's log shows every key that **fired**.
   - At the top it also shows what Office actually registered (`getShortcuts`). That is the key evidence.

| Add-in | Keys to press |
|---|---|
| **K1 Single** | Ctrl+Alt+Shift+L (control), Ctrl+Shift+[ |
| **K1 Named Keys** | Ctrl+Alt+Shift+K and Ctrl+Alt+Up (controls), then Ctrl+Shift+[, Ctrl+[, Ctrl+], Ctrl+', Ctrl+;, Ctrl+,, Ctrl+., Ctrl+F2 |
| **K1 Literal Keys** | Ctrl+Alt+Shift+J (control), then Ctrl+Shift+[, Ctrl+[, Ctrl+', Ctrl+;, Ctrl+,, Ctrl+. |

6. **Take a screenshot of each pane** (Win+Shift+S) after testing it. Scroll so both the log and the `getShortcuts` section are visible, or take two shots.
7. **Clean up:** close Excel, stop the server (Ctrl+C in its window), then run:
   ```powershell
   .\cleanup.ps1
   ```
   To also remove the development certificate, add `-UninstallCert`.

✍️ Did the **control** keys fire in each pane? ______
✍️ Did any punctuation key fire? Which? ______

---

## C. Finish
1. **Turn Macabacus back on:** **File › Options › Add-ins**, set **Manage: COM Add-ins**, click **Go…**, tick **Macabacus**, then click OK.
2. Tell Claude you're done. Paste your ✍️ answers and the K1 screenshot paths. Claude reads `%LOCALAPPDATA%\EmtSpike\log.jsonl` directly.
