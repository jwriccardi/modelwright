# Open questions

## A. What to check in the installed copy of Macabacus

Public documentation doesn't settle these points. Screenshots or a short screen recording of each would close them. Each item notes where its answer goes in the spec.

**Formatting cycles**

1. ◐ *General Number cycle captured (research/06). Still needed: Percent, Currency, Multiple, Date and Binary.* **Number cycle defaults.** Settings › Configure › Excel › Format › Numbers. Capture the list for each cycle (General, Percent, Multiple, Currency, Date, Binary) with its format codes. *→ Sets our default cycle contents.*
2. ✅ *Captured (research/06).* **Font color cycle defaults.** Settings › … › Colors. Capture the RGB values in the Font Color Cycle, Fill Color Cycle and AutoColor scheme. *→ Sets our default palettes.*
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
8. ◐ *Key confirmed as Ctrl+Shift+\. History depth still open.* **"Last Audited Cell."** Is Ctrl+Shift+\ the key? How deep is its history?
9. **Hard cases.** What does the dialog show for:
   - INDIRECT / OFFSET;
   - a reference to a closed external workbook;
   - a very large range (A1:A10000);
   - a cell on a hidden sheet?
10. ✅ *Captured from the help PDF (research/07).* **Evaluate functions & groups (Ctrl+E).** A screenshot of a nested IF/SUM formula expanded this way.
11. **Priorities.** Which of these do you use *every day*, and which could wait for v2?

## B. Decisions for the owner

| # | Decision | Recommendation | Status |
|---|---|---|---|
| D0 | Exact Macabacus keys required? | — | **Decided 2026-09-28: yes, critical** |
| D1 | Platform and architecture | — | **Decided 2026-09-28: Excel-DNA (C#, net48), Windows desktop first. ADR-0002 accepted.** |
| D2 | License | MIT, with DCO sign-off for contributions (research/03) | **Adopted for the scaffold 2026-09-29.** The owner can still change it before the repo goes public. |
| D3 | Copyright holder | — | **Decided 2026-09-28: Pegasus Technology Group LLC** |
| D4 | Build fresh vs fork XLerate | Build fresh; XLerate is a design reference only (it's TypeScript) | Proposed |
| D5 | Cycles in v1 | Number, Date, Currency, Percent, Multiple, Font, Fill (7). Blue-black and decimals are stretch goals | Proposed |
| D6 | Keymap | Macabacus defaults exactly, and users can remap them (PLAN §4.2) | Proposed |
| D7 | Product name | `excel-modeling-toolkit` is a working name | Open |
| D8 | Mac / web follow-up | Office.js, which K1c showed can bind the exact punctuation keys. **Owner idea (2026-09-29): a web edition that just does the color cycles.** | Future (PLAN, v2 roadmap) |
| D9 | Approve decision spikes K1–K4 (throwaway code) | — | **Approved 2026-09-28** |
| D11 | Cross-workbook trace | — | **Decided 2026-09-28: essential.** This is why Excel-DNA was kept over Office.js. |
| D10 | Undo target | — | **Decided 2026-09-28: Macabacus parity in v1. The Excel-DNA + Office.js hybrid for full native undo is a v2 candidate.** |

## C. Development-machine note
Macabacus is installed on the owner's machine, and it binds the same keys. Whichever add-in registers a key last owns it. **During spikes and development, disable Macabacus** (File › Options › Add-ins › COM Add-ins) or remap its keys in its Shortcut Manager.
