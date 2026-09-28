# 07 — Macabacus Trace In / Trace Out: observed behavior

*Source: "Precedents & Dependents — Macabacus Help Center" (PDF supplied by the owner, 2026-09-28). The PDF itself is kept locally in `reference/macabacus/` and is not published. What follows is a factual summary written for our spec.*

## The dialog (from the help center screenshot)

- **Title:** "Trace In".
- **Formula header:**
  - Shows the whole audited formula.
  - Each reference is colored the way Excel colors references in edit mode.
  - A **wrap toggle** sits at the right.
  - Example formula: `=(B2+C2/D2)+IF(E2>0,F2+G2,SUM(H2:J2))*(K2+L2-ABS(M2))+PRODUCT(N2:T2,U2)+V2`
- **Tree columns:** **Precedents | Argument | Value**.
- **Tree for the example formula, with "Evaluate functions & groups" on:**

```
⌂ A2                                          $15,066
  > (x) (B2+C2/D2)                             5.5        ← parenthesized group node
  v ƒx  IF(...)                                14.9       ← function node, expanded
      > (x) E2>0              logical_test     True
      > (x) F2+G2             [value_if_true]  14.9
      > ƒx  SUM(...)          [value_if_false] 20
  > (x) (K2+L2-ABS(M2))                        (4)
  > ƒx  PRODUCT(...)                           15,120
    ▦  V2                                      (cell leaf)
```

- **Footer:** a gear menu (options), OK, and Cancel.

**What this means for our tree model:**
- With evaluation on, the tree follows the **formula's own structure**. Nodes are:
  - (a) parenthesized groups, labelled `(x)`;
  - (b) function calls, labelled `ƒx NAME(...)`, whose children are that function's **arguments**;
  - (c) plain cell or range references, which are leaf nodes that can be expanded.
- The **Argument column** shows the parameter name *in the way Excel's own tooltips write it*: `logical_test`, `[value_if_true]`, `[value_if_false]`. So we need a **function signature table**: argument names for Excel's functions, starting with the ~100 most common in financial models.
- **Every node has a value**, including intermediate values of subexpressions. We compute these with `Worksheet.Evaluate(subexpression)`, evaluated in the context of the audited cell's sheet.
- Evaluation "will also allow you to trace to the result of a formula, for example, the result of INDEX". So a function that returns a reference (INDEX, OFFSET, INDIRECT, CHOOSE) is resolved to its **target range**, which you can then navigate to.
- When evaluation is **off**, the tree shows only the referenced cells and ranges: the "classic" view, and our v1 baseline.

## Keyboard

| Key | Action |
|---|---|
| Up / Down (or mouse) | Select a node → **Excel navigates** to that range: off-screen, on another sheet, or in another workbook |
| Right | Expand the node, tracing one level deeper |
| Left | Go back up a level |
| F2 | Edit the audited formula in Excel. Macabacus goes to **Point mode** when it can, so arrow keys change the reference. F2 again gives Edit mode. **The dialog stays open.** |
| Ctrl+E | Toggle "Evaluate functions & groups" |
| Ctrl+Up/Down/Left/Right | Move the dialog |
| Ctrl+Home / Ctrl+End | Snap the dialog to the top-left / bottom-right of the screen (also recovers a dialog left off-screen) |
| Shift+Up/Down/Left/Right | Resize the dialog |

**The whole round trip needs no mouse:** "Open the Trace In dialog, navigate multiple levels, and close it."

## Options in the gear menu
- **Evaluate Functions and Groups** (Ctrl+E), described above.
- **Highlight Navigated Cells:** shades the rows and columns crossing the selected range. **Macabacus notes that this clears Excel's Undo stack**, because it changes cell formatting.
  - Our version should draw the highlight *without* changing any cells, for example with a transparent overlay window or a temporary shape. Then undo is untouched. This is a candidate for v2.
- **Wrap Formula Text:** needed for formulas that contain Alt+Enter line breaks.
- Position memory: **the dialog remembers its last position.**

## Trace Out
- It mirrors Trace In but goes in the opposite direction.
- It "relies on Excel for this data". So Macabacus gets dependents from Excel's own dependency tracking (`Range.Dependents`, or ShowDependents with NavigateArrow). This is why it can't find dependents that go through OFFSET and some other volatile functions.
- **By contrast, Trace In parses the formula itself**, which is why it has the parser limitations below.

## Documented limitations (useful as our minimum bar)
- Structured table references must be **fully qualified** (`DeptSales[Sales]`, not `[Sales]`) and in the same workbook.
- Merged cells may not work.
- Many unused range names can break the parser.
- Other add-ins (Workshare, Anaplan) can make **Up/Down move the worksheet cursor instead**.

## Key inference: how the dialog handles the keyboard
- Two facts point the same way:
  - F2 puts *Excel* into Point or Edit mode while the dialog is open.
  - Other add-ins can make the arrow keys reach the *worksheet*.
- Together they strongly suggest that **the Trace In dialog doesn't take keyboard focus.** Excel keeps it, and Macabacus intercepts Up/Down/Left/Right/Enter/Esc with a **keyboard hook** while the dialog is open. Add-ins that install their own hooks can get in first.
- **Spike K4 variant C** tests this model directly: a window that doesn't take focus, plus a keyboard hook on Excel's thread.
  - If it works, it's the only design that also supports the F2-in-Point-mode behavior.

## Show All Precedents / Dependents, and AutoTrace (v2)
- **Show All Precedents** draws Excel's trace arrows for **every selected cell**; Excel's own button does only the active cell. Pressing the key again before changing the selection clears the arrows.
- **AutoTrace** switches off for selections over 20 cells, to avoid freezing Excel.
