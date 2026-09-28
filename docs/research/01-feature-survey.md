# 01 — Feature survey: Macabacus and competitors

*Researched 2026-09-28. Sources are cited inline.*

**Tags used below:**
- **[UNVERIFIED]**: could not be confirmed from public sources.
- **[INFERRED]**: reasoned from indirect evidence.

We can check both kinds against the Macabacus copy installed on the project owner's machine (see [`../open-questions.md`](../open-questions.md)).

Macabacus's own help center requires a login. The sources used instead were:
- the public HelpScout mirror (`macabacus-help-center1.helpscoutdocs.com`);
- official Macabacus PDFs;
- an LSEG/Barclays deployment guide that includes screenshots of the settings dialogs.

## Source key

| Key | Source |
|---|---|
| KB-PDF | Macabacus Keyboard Shortcuts, Dec 2025 — https://macabacus.com/assets/2025/12/Macabacus_Keyboard-Shortcuts.pdf |
| QSG | Macabacus Quick Start Guide (CFI) — https://macabacus.com/assets/2024/09/Macabacus-Quick-Start-Guide.pdf |
| FMT-PDF | Format a Financial Model with Macabacus — https://macabacus.com/assets/2024/09/Format-a-Financial-Model-with-Macabacus-Summary.pdf |
| AUD-PDF | Audit a Financial Model with Macabacus — https://macabacus.com/assets/2024/09/Audit-a-Financial-Model-with-Macabacus-Summary.pdf |
| LSEG | LSEG/Barclays Macabacus Quick Start Guide — https://www.lseg.com/content/dam/marketing/en_us/documents/learning-centre/macabacus-start-guide.pdf |
| HC-Cycles | https://macabacus-help-center1.helpscoutdocs.com/article/771-format-cycles |
| HC-Colors | https://macabacus-help-center1.helpscoutdocs.com/article/776-format-colors |
| HC-Numbers | https://macabacus-help-center1.helpscoutdocs.com/article/777-number-formats |
| HC-Keyboard | https://macabacus-help-center1.helpscoutdocs.com/article/917-keyboard |
| HC-Trace | https://macabacus-help-center1.helpscoutdocs.com/article/842-precedents-dependents |
| HC-SysReq | https://macabacus-help-center1.helpscoutdocs.com/article/1038-system-requirements |
| HC-Lite | https://macabacus-help-center1.helpscoutdocs.com/article/1042-macabacus-lite |
| FDS | FactSet Hot Keys PDF — https://media.felix.fe.training/2019/10/Factset-keyboard-shortcuts-for-Excel.pdf |

---

## 0. How Macabacus cycles work

This applies to every cycle (HC-Cycles, HC-Colors, QSG, LSEG).

- **What a cycle is.** An ordered list of formats, colors or styles. Macabacus recommends at most about 6 items, ordered from most to least used. Each press of the cycle's shortcut applies the next item to the whole selection.
- **Always starts from item 1.** Macabacus ignores the cell's current format: "formatting cycles always start at the beginning, allowing you to reach the third style by pressing the shortcut three times, regardless of existing cell formatting."
  - So Macabacus must remember the cycle position between presses. [INFERRED] The position resets when the selection changes or another command runs.
  - XLerate (open source) does the opposite. It works out which preset the cell already has and applies the next one, or applies the first preset if nothing matches. **This is a design choice we need to make** (see the plan).
- **Wrap-around.** "When the cycle ends, pressing the shortcut again starts it over."
- **Where cycles are edited.**
  - Settings › Configure › Excel › Format, with sub-pages Colors, Numbers, Fonts, Cell Size, Styles, Fast Format and Other.
  - Color cycles pick from a shared **Color Palette** (Settings › General › Color Palettes).
  - Settings can be exported, imported, shared and reset, and enterprise admins can push them to users.
- **Custom cycles.** User-defined cycles of whole cell styles, bound to Ctrl+Alt+1…8.
- **Shortcut Manager.**
  - Every action can be rebound or disabled.
  - An "Override" button forces Macabacus to win shortcut conflicts with other add-ins.
  - HC-Keyboard says conflicts with Bloomberg, FactSet and CapIQ are common, and are decided by the unpredictable order in which add-ins load.
- **Built-in Excel shortcuts Macabacus deliberately overrides:** Ctrl+Shift+1/2/4/5/8, Ctrl+; (insert date), Ctrl+' (copy formula from above) and Ctrl+Shift+[ ] (select precedents).

## 1. Number format cycling

### Macabacus defaults

Keys are from KB-PDF, FMT-PDF and LSEG.

| Cycle | Default key |
|---|---|
| General Number | Ctrl+Shift+1 |
| Date | Ctrl+Shift+2 |
| Local Currency | Ctrl+Shift+4 |
| Foreign Currency | Ctrl+Alt+Shift+4 |
| Percent | Ctrl+Shift+5 |
| Multiple | Ctrl+Shift+8 |
| Binary (Yes/No, On/Off) | Ctrl+Shift+Y |
| Increase / Decrease decimals | Ctrl+, / Ctrl+. |

**What the default formats look like (QSG):**
- Negatives in parentheses.
- Zero shown as an en-dash.
- Text padded so it lines up with numbers.
- Examples: `1,111` / `(1,111)` / `–`; `$1,111`; `35%`; `8x`.

**Default General Number Cycle codes (LSEG p.6 screenshot):**

```
Comma 0 Dec Lg Align   _(#,##0_)_%;(#,##0)_%;_("–"_)_%;_(@_)_%
Comma 1 Dec Lg Align   (same pattern, #,##0.0)
Comma 2 Dec Lg Align   (same pattern, #,##0.00)
Comma 0 Dec No Align   #,##0;(#,##0);"–";@
```

The trailing `_%` pads each value by the width of a `%` sign. That keeps plain numbers aligned with percentages in the same column.

- **Settings for each format:** name, format code, live preview, and left or right alignment.
- **Options for the whole cycle:** "Align numbers right" and "Italicize percentages".

**Edge cases (HC-Numbers):**
- **Format limit.** A workbook can hold only a limited number of distinct number formats. At the limit, applying a new one fails with "invalid number format". FactSet includes a "Clean Up Number Formats" command for this.
- **Locale.** Use `[$$]` instead of a bare `$` so the dollar sign displays correctly under any Windows locale.

### Competitors
- **FactSet "SmartCycles":** keys almost identical to Macabacus's, plus "My SmartCycle 1–5" (Ctrl+Alt+1–5).
- **TTS Turbo Macros:** one Ctrl+Shift+N cycle through number, currency, %, multiple and general. https://trainingthestreet.com/macros/
- **WST Macros:** dollars, %, multiples, basis points and real-estate formats. https://wallst.training/macros/
- **Excel built-in:** fixed formats with no cycling. Ctrl+Shift+1 = Number with 2 decimals, 4 = Currency, 5 = Percent, and so on.
- **XLerate (MIT):**
  - Keys: Ctrl+Shift+1 number, 2 cell formats, 3 date, 4 text styles.
  - Detects the current preset and applies the next.
  - Stores presets **per workbook**.

## 2. Font color cycling and AutoColor

### Macabacus manual color controls
- **Font Color Cycle: Ctrl+'.** 6 slots. The factory defaults are [UNVERIFIED]; the LSEG screenshot starts black, blue, ….
- **Blue-Black toggle: Ctrl+;.** Switches between the AutoColor "input" color (typically blue) and the default font color (typically black).
- **Default colors (HC-Colors):**
  - Font = black.
  - Border = black.
  - Shading = light gray, used by Alternate Row Shading (Ctrl+Alt+R).

### Macabacus AutoColor

AutoColor colors each cell's font by what the cell contains. Selection / Sheet / Workbook are **Ctrl+Alt+A / S / Q**.

| Category | Typical color |
|---|---|
| Inputs (numeric constants) | blue |
| Partial inputs (formula containing a hardcoded number, e.g. `=A1+12.34`) | orange (CFI screenshot) / unset (LSEG screenshot) |
| Formulas (references on the same sheet) | black |
| Links to other worksheets | green |
| Links to other workbooks | purple |
| Hyperlinks | orange / red |
| External data (a user-editable list of functions: BDH, BDP, BDS, CIQ, RDP.Data…) | cyan |

- **Partial-input exceptions:** the numbers 0, 1, 100, 1000 and 1,000,000 don't count as hardcodes.
- **Settings seen in the screenshot:**
  - "AutoColor on entry" is off by default. Macabacus warns it can slow things down and interfere with Undo.
  - Dates are colored.
  - Text is not colored.
- **"No AutoColors" list:** colors that AutoColor never overwrites, such as red check cells.

**Scope note.** AutoColor is **not** one of our four features, but it builds on the same color model. It is the natural next feature after launch, so the color-scheme data model should allow for it (see PLAN).

### Competitors
- **Common convention (BIWS, TTS, Breakdown):**
  - blue = hardcode;
  - black = formula;
  - green = link to another sheet;
  - red = link to another file.
  - https://breakingintowallstreet.com/kb/excel/how-to-color-code-in-excel/
- **TTS:** font color cycle Ctrl+Shift+C goes blue → green → red → automatic.
- **Breakdown (MIT):** AutoColor that also marks `WEBSERVICE` and errors other than #N/A in red, and leaves text, dates, hyperlinks and blanks alone.
- **Excel Campus Formatting Shortcuts (VBA):** applies formats with Paste Special through SendKeys, specifically so Excel's Undo keeps working. This shows how much undo matters to users. https://www.excelcampus.com/tools/formatting-shortcuts-add-in-help/

## 3. Fill (background) color cycling

- **Macabacus: Fill Color Cycle, Ctrl+Shift+K.**
  - The QSG example goes peach → orange → light blue → navy → light gray.
  - The LSEG screenshot shows about 7 slots, the last apparently "No fill".
  - Macabacus calls it "useful for markup".
- **FactSet:** Cell Back Color SmartCycle, Alt+Shift+K.
- **TTS:** Ctrl+Shift+V goes light yellow → light turquoise → gray → **none**.
- **Spec takeaway:** include an explicit **No fill** slot so the user can cycle back to a clean cell.

## 4. Trace precedents (and dependents)

### Macabacus

| Command | Default key | What it does |
|---|---|---|
| Pro Precedents ("Trace In") | Ctrl+Shift+[ | Opens the tracing dialog |
| Pro Dependents ("Trace Out") | Ctrl+Shift+] | Same dialog, reverse direction |
| Show All Precedents / Dependents | Ctrl+Alt+[ / ] | Draws Excel's arrows for every cell in the selection |
| Clear Arrows | Ctrl+Alt+\ | Removes the arrows |
| Last Audited Cell | Ctrl+Shift+\ [INFERRED] | Jumps back to the cell you audited |
| AutoTrace | ribbon | Arrows follow the active cell (turned off for selections over 20 cells) |

**The Pro Precedents dialog** (HC-Trace, LSEG p.7–8):
- **Header:** the audited formula, with references colored as in Excel's edit mode.
- **Tree:** two columns, **Precedents | Value**. The root is the audited cell; each child is one reference in the formula, with its current value.
- **Keyboard only:**
  - **Up/Down** moves through the tree, and **Excel immediately selects and scrolls to that range**. This works on other sheets and even in other workbooks.
  - **Right** expands a node to show that cell's own precedents. **Left** collapses it.
  - **F2** edits the audited formula while the dialog stays open.
  - **Ctrl+E** turns "Evaluate functions & groups" on or off.
  - Ctrl+Arrow moves the dialog, Shift+Arrow resizes it, and Ctrl+Home rescues it after a monitor change.
- **Gear options:**
  - Evaluate functions & groups: splits out each function or parenthesized group as its own node, showing its intermediate value.
  - Evaluate function arguments.
  - Unhide rows & columns.
  - **Open linked workbooks.**
  - Highlight navigated cells: shades the row and column of the cell you land on.
  - Wrap formula text.
- **OK vs Cancel:** [UNVERIFIED] Probably OK stays on the cell you navigated to and Cancel returns to the audited cell.

**Limitations Macabacus documents (HC-Trace):**
- Structured table references must be fully qualified and in the same workbook.
- Merged cells may not parse.
- Many unused names can break the parser (so Macabacus uses its own formula parser).
- Dependents can't be found through volatile functions such as OFFSET.
- Arrow keys fail when other add-ins intercept them (Workshare and Anaplan are named).

### Excel built-in
- **Ctrl+[ / Ctrl+Shift+{** select direct / all precedents, **on the same sheet only**. **F5 then Enter** jumps back.
- **Trace Precedents arrows** go one level at a time. A reference to another sheet shows as a dashed arrow to a sheet icon; double-clicking it opens a Go To list.
- **Inquire › Cell Relationship:** an interactive diagram. Windows only, and only in some Office editions.

### Competitors
- **Arixcel Explorer** (Windows COM add-in) is the best in class and a good model for the UX:
  - Ctrl+Q explores; pressing it again drills into nested formulas.
  - Explorer windows can be stacked.
  - **Enter** closes and keeps the selection; **Esc** closes and restores it.
  - **Ctrl+Backspace** returns to the original cell, with a history of about 100 steps.
  - It evaluates the targets of VLOOKUP, OFFSET, INDEX and INDIRECT, and **bolds the IF/CHOOSE branch actually taken**.
  - https://www.arixcel.com/explorer/functions/explore
- **FactSet:** Smart Precedents Ctrl+Shift+{, Return To Audited Cell Ctrl+Shift+|.
- **Accelerate Excel "Explore Formula":** a free tree view with + or Right to expand, covering other sheets and workbooks. https://www.accelerate-excel.com/blog/excel-trace-precedents-explore-formula
- **PerfectXL add-in** (Office.js, Marketplace): breaks a formula into parts with intermediate values and spotlights precedents and dependents.
- **QuickCel, TTS (Ctrl+Shift+T), WST, Kutools, FormulaDesk:** each has some kind of trace panel or live arrows; details are in the research transcript.

## 5. Comparison table

| Product | Number cycle | Font cycle | Fill cycle | AutoColor | Precedent tree | Jump back | Architecture | License / price |
|---|---|---|---|---|---|---|---|---|
| Macabacus | ✅ many | Ctrl+' | Ctrl+Shift+K | ✅ | ✅ best-known | ✅ | **VSTO/.NET COM, Windows only** | Closed, $200–360/yr |
| Macabacus Lite | subset | subset | ? | ? | "auditing" | ? | **Office.js** (Windows, Mac, web) | Closed, subscription |
| FactSet Spreadsheet Tools | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | Windows plug-in | Bundled with FactSet |
| TTS Turbo Macros | ✅ | ✅ | ✅ | ✅ | ✅ | ? | Office.js (Marketplace) | Subscription |
| WST Macros | ✅ | ✅ | ✅ | ✅ | ✅ | ? | VBA, Windows | $125/yr |
| Arixcel | ? | ✅ | ✅ | map view | ✅✅ | ✅ | Windows COM | Quote |
| XLerate | ✅ | via AutoColor | ? | ✅ | ✅ | ? | **Office.js + TS** | **MIT** |
| Breakdown | – | – | – | ✅ | ✅ | ✅ | VBA .xlam (Windows and Mac) | **MIT** |

Macabacus says outright that it is Windows-only because Office for Mac doesn't support "sophisticated COM add-ins", and calls Microsoft's web API "woefully underdeveloped" (HC-SysReq, https://macabacus.com/faqs). Much of that criticism predates the 2025 Office.js additions: undo support, ExcelApiDesktop 1.1, and precedent APIs that work across sheets. See [02](02-architecture-options.md).
