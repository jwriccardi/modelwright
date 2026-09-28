# Open questions

## A. What to check in the installed copy of Macabacus

Public documentation doesn't settle these points. Screenshots or a short screen recording of each would close them. Each item notes where its answer goes in the spec.

**Formatting cycles**

1. **Number cycle defaults.** Settings › Configure › Excel › Format › Numbers. Capture the list for each cycle (General, Percent, Multiple, Currency, Date, Binary) with its format codes. *→ Sets our default cycle contents.*
2. **Font color cycle defaults.** Settings › … › Colors. Capture the RGB values in the Font Color Cycle, Fill Color Cycle and AutoColor scheme. *→ Sets our default palettes.*
3. **What resets the cycle position?** Select a cell, press Ctrl+Shift+1 twice, then:
   - (a) move to another cell and press it again;
   - (b) come back to the first cell and press it again;
   - (c) press it on a cell that already has the 3rd format.

   Does the cycle restart at item 1 each time? *→ Confirms the "cycle state" rule in the PLAN.*
4. **Multi-cell selections with mixed formats.** Which cell decides the "current" format? *→ Cycle engine spec.*
5. **Undo.** After a Ctrl+Shift+1 cycle, does Ctrl+Z work? Is it Excel's own undo or Macabacus's? How many levels? *→ Tells us how high the bar is.*

**Pro Precedents (Ctrl+Shift+[)**

6. **Default behavior**, on a formula like `=SUM(Sheet2!A1:A10)+B5*Assumptions!Growth`:
   - How are named ranges shown?
   - Can the range node be expanded into its 10 cells?
   - Does moving with Up/Down switch sheets?
7. **OK vs Cancel vs Esc.** Where does the cursor end up after each?
8. **"Last Audited Cell."** Is Ctrl+Shift+\ the key? How deep is its history?
9. **Hard cases.** What does the dialog show for:
   - INDIRECT / OFFSET;
   - a reference to a closed external workbook;
   - a very large range (A1:A10000);
   - a cell on a hidden sheet?
10. **Evaluate functions & groups (Ctrl+E).** A screenshot of a nested IF/SUM formula expanded this way.
11. **Priorities.** Which of these do you use *every day*, and which could wait for v2?

## B. Decisions for the owner

| # | Decision | Recommendation | Status |
|---|---|---|---|
| D1 | Platform and architecture | Office.js + TypeScript, targeting Windows, Mac and the web (ADR-0001) | Proposed |
| D2 | License | MIT, with DCO sign-off for contributions (research/03) | Proposed |
| D3 | Copyright holder | Individual vs Pegasus Technology Group LLC — **owner decides** | Open |
| D4 | Build fresh vs fork XLerate | See PLAN §"Build vs fork" | Open |
| D5 | Number format cycles in v1 | 5 cycles on one engine: Number, Percent, Multiple, Currency, Date | Proposed |
| D6 | Default keymap | Macabacus-compatible where Office.js allows it; see PLAN | Proposed, to be tested in the spike |
| D7 | Product name | `excel-modeling-toolkit` is a working name | Open |
