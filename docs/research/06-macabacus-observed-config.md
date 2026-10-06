# 06 — Macabacus configuration observed on the owner's machine

*Captured 2026-09-28 from the owner's installed Macabacus: Settings screenshots, plus the Shortcut Manager export `macabacus-shortcut-list.xlsx`.*

**The source files are local only**, in `reference/macabacus/`, which is git-ignored. They contain Macabacus's UI and help text, which we must not publish. This doc records only the **facts**: codes, RGB values and key assignments.

RGB values were read from the swatch pixels in the screenshots, not estimated by eye.

**What this means for our defaults.** These are the owner's *current* settings. Some may be customized rather than Macabacus factory defaults. The target users are Macabacus users, so we adopt them as **our factory defaults**.

## Number formats

### General Number Cycle (Ctrl+Shift+1)

| # | Name | Code |
|---|---|---|
| 1 | Comma 0 Dec Lg Align | `_(#,##0_)_%;(#,##0)_%;_("–"_)_%;_(@_)_%` |
| 2 | Comma 1 Dec Lg Align | `_(#,##0.0_)_%;(#,##0.0)_%;_("–"_)_%;_(@_)_%` |
| 3 | Comma 2 Dec Lg Align | `_(#,##0.00_)_%;(#,##0.00)_%;_("–"_)_%;_(@_)_%` |
| 4 | Comma 0 Dec No Align | `#,##0;(#,##0);"–";@` |

- Both cycle-wide options are **off**: "Align numbers right" and "Italicize percentages".
- **Still needed:** screenshots of the Percent, Currency, Multiple and Date cycle lists (open-questions A1).
- Macabacus's descriptions give examples of what those cycles cover:
  - Percent: `35%`, `150bps`, `L+350`.
  - Multiple: `8.2x`.
  - Date: `9/30/14`, `Sep 30, 2014`.
  - Binary: `Yes/No`, `Y/N`, `On/Off`.
  - Ratio: `1.2345:1`, `6/10`.

## Color cycles

| Cycle | Order of items (RGB) |
|---|---|
| **Font Colors** (Ctrl+') | Blue (0,0,255) → Green (0,128,0) → Purple (128,0,128) → Red (255,0,0) → White (255,255,255) → Black (0,0,0) |
| **Fill Colors** (Ctrl+Shift+K) | Light blue (201,218,248) → Light cyan (210,242,255) → Light pink (244,204,204) → Peach (252,229,205) → Navy (28,69,135) → **No fill** |
| Border Colors (Ctrl+Alt+Shift+') | Black (0,0,0) → White (255,255,255) → Gray (128,128,128) → Dark red (204,0,0) |
| Chart Colors (Ctrl+Alt+C) | Blue → Green → Purple → Red (same RGBs as the font colors) |

**Notes:**
- The font cycle **starts at blue**, not black. Black is last.
- The fill cycle has a **navy** slot, which pairs with the White slot in the font cycle.
- The fill cycle ends with **No fill**. This confirms the spec decision to include it.

## Other color sets

| Set | Values (RGB) |
|---|---|
| Recolor Colors | (0,0,255), (0,0,0), (0,128,0), (128,0,128), (255,102,0) |
| No AutoColors (colors AutoColor never overwrites) | (255,0,0), (255,255,255) |
| Chart Series | (91,155,213), (237,125,49), (165,165,165), (255,192,0), (68,114,196), (112,173,71). These are Office's default theme accents. |
| Default font color | (0,0,0) |
| Default border color | (0,0,0) |
| Default shading | (220,220,220) |

## AutoColor scheme (future feature; Blue-Black toggle uses Inputs and the default font color)

| Category | Color |
|---|---|
| Inputs | Blue (0,0,255) |
| Partial inputs | *none (unset)* |
| Formulas | Black (0,0,0) |
| Worksheet links | Green (0,128,0) |
| Workbook links | Purple (128,0,128) |
| Hyperlinks | Orange (255,102,0) |
| External data | *none (unset)* |

**Behavior settings:**
- AutoColor on entry: **off**.
- AutoColor dates: **on**.
- AutoColor text: **off**.

**External data functions:** `RDP.Data`, `BDH`, `BDP`, `BDS`, `CIQ` (the list scrolls, so there may be more).

## Default color palette ("Default (15)")

Only the first 9 of 15 were visible:

| Color | RGB |
|---|---|
| Dark Blue | (0,44,89) |
| Bright Blue | (0,115,230) |
| Dark Turquoise | (2,188,245) |
| Light Turquoise | (144,224,255) |
| Dark Purple | (99,91,255) |
| Bright Purple | (169,96,238) |
| Pink | (255,89,150) |
| Red | (255,51,61) |
| Yellow | (255,203,87) |

The palette is a shared picker source. None of the cycle colors above come from it.

## Things these files settle
- **Last Audited Cell is `Ctrl+Shift+\`** (confirmed). Trace Out is `Ctrl+Shift+]`. AutoTrace Precedents/Dependents are `Ctrl+Alt+Shift+[` / `]`. Clear Arrows is `Ctrl+Alt+\`.
- **Help text for Trace In:** it "will attempt to open any external workbooks containing precedent cells". So opening closed linked workbooks is default behavior, not an option.
- **Many more punctuation and named keys** are in use: Alt+Shift+; , . = -, Ctrl+Alt+= -, Ctrl+Alt+Shift+, . = - ', Ctrl+F2, Alt+F12, Shift+F12, Ctrl+Alt+Home/End/PgUp/PgDn, and more. None of these could be bound in Office.js. This further confirms ADR-0002.
- **Keys our spike used that clash with Macabacus:**
  - Ctrl+Alt+Shift+O is Reopen.
  - Ctrl+Alt+Shift+T is Untranspose.
  - Ctrl+Alt+Shift+6/7 are Footnote Cycle / Inside Borders.

  The spike has moved its own test commands to Ctrl+Alt+Shift+F-keys. The product's "Override" command needs a key that Macabacus doesn't use.


## Factory defaults (Macabacus 9.9.5 settings export, 2026-10-06)

The owner supplied a settings export from a **fresh installation**. It is kept local-only at `reference/macabacus/MacabacusSettings-factory-9.9.5.xml`, which is git-ignored.

**What it confirms and settles:**
- **These match the owner's screenshots, so the earlier values were already factory defaults:**
  - the General Number cycle;
  - the Font and Fill cycles (Fill ends with an empty entry, meaning No fill);
  - the AutoColor scheme;
  - default colors, recolor colors and chart series.
- **Number-format cycles**, adopted verbatim as our defaults (the en dash is U+2013):

| Cycle (key) | Items |
|---|---|
| Local Currency (Ctrl+Shift+4) | USD 0, 1 and 2 Dec Lg Align (`_([$$]#,##0_)_%;([$$]#,##0)_%;_("–"_)_%;_(@_)_%` …), USD 0 Dec No Align, EUR 0 Dec Lg Align (`[$€-2]`), GBP (`[$£-809]`), YEN (`[$¥-2]`) |
| Percent (Ctrl+Shift+5) | Percent Aligned Neg Pct `_(#,##0.0%_);(#,##0.0%);_("–"_)_%;_(@_)_%`, Percent Unaligned, Hard Percent Aligned and Unaligned (a literal `"%"`), SOFR + `"S"+0_)_%;"S"-0_)_%;"S"+0_)_%`, LIBOR + `L+0_)_%;L-0_)_%;L+0_)_%` |
| Multiple (Ctrl+Shift+8) | Mult 1 and 2 Decimal Aligned Neg Pct `_(0.0x_)_)_';_((0.0x)_'_';_("–"_)_%;_(@_)_%`, Mult 1 and 2 Decimal Unaligned `0.0x;(0.0x);"–"` |
| Date (Ctrl+Shift+2) | `m/d/yyyy;@`, `mmmm d, yyyy;@`, **Date Actual Year `0000\A`**, **Date Estimated Year `0000\E`**. The year formats are *number* formats applied to a plain year value such as 2025, not to a date serial. |
| Binary (Ctrl+Shift+Y) | Yes/No `"Yes";"ERROR";"No";"ERROR"`, Y/N, On/Off, True/False |
| Ratio (Alt+Shift+;) | Exchange Ratio `0.0\:1_);(0.0)\:1_);0.0\:1_);@_)`, Fraction 1, 2 and 3 (`# ?/?` …), Halves `# ?/2`, Thirds `# ?/3` |

- **Other factory facts:**
  - Undo/Redo is enabled, with a **maximum of 2,000** steps.
  - Factory Excel shortcuts differ from the owner's install in a few non-v1 places. For example, Paste Number Formats is factory Ctrl+Alt+U; the owner's copy uses Ctrl+Alt+F, and factory Ctrl+Alt+F is Super Find. So the owner has customized some keys.
  - Macabacus also binds **Copy (Ctrl+C) and Cut (Ctrl+X)** by default.
- **AutoColor `DataFunctions`:** RDP.Data, BDH, BDP, BDS, CIQ, FDS, SNL.

## Full Macabacus keymap (132 commands)

"Our scope" tags the target release. Blank means not currently planned. Every row is a candidate for future scope, and **all keys stay reserved** so a later feature can use its Macabacus key.

| Command | Keystroke | Category | Macabacus utility (1-5) | Our scope |
|---|---|---|---|---|
| General Number Cycle | Ctrl+Shift+1 | Numbers | 5 | **v1** |
| Currency Cycle | Ctrl+Shift+4 | Numbers | 5 | **v1** |
| Percent Cycle | Ctrl+Shift+5 | Numbers | 5 | **v1** |
| Multiple Cycle | Ctrl+Shift+8 | Numbers | 4 | **v1** |
| Date Cycle | Ctrl+Shift+2 | Numbers | 4 | **v1** |
| Binary Cycle | Ctrl+Shift+Y | Numbers | 3 | v2 candidate |
| Ratio Cycle | Alt+Shift+; | Numbers | 2 | v2 candidate |
| Increase Decimals | Ctrl+, | Numbers | 5 | stretch |
| Decrease Decimals | Ctrl+. | Numbers | 5 | stretch |
| Shift Decimal Left | Alt+Shift+, | Numbers | 4 |  |
| Shift Decimal Right | Alt+Shift+. | Numbers | 4 |  |
| Blue-Black Toggle | Ctrl+; | Colors | 5 | stretch |
| Font Color Cycle | Ctrl+' | Colors | 5 | **v1** |
| Fill Color Cycle | Ctrl+Shift+K | Colors | 4 | **v1** |
| Border Color Cycle | Ctrl+Alt+Shift+' | Colors | 4 |  |
| Chart Color Cycle | Ctrl+Alt+C | Colors | 3 |  |
| Pinstripes Cycle | Ctrl+Alt+R | Colors | 3 |  |
| AutoColor Cycle | Ctrl+Alt+. | Colors | 3 |  |
| AutoColor Selection | Ctrl+Alt+A | Colors | 4 | v2 candidate |
| AutoColor Sheet | Ctrl+Alt+S | Colors | 3 | v2 candidate |
| AutoColor Workbook | Ctrl+Alt+Q | Colors | 2 | v2 candidate |
| Horizontal Cycle | Ctrl+Shift+H | Alignment | 5 |  |
| Vertical Cycle | Ctrl+Shift+V | Alignment | 2 |  |
| Left Indent Cycle | Ctrl+Shift+I | Alignment | 3 |  |
| Right Indent Cycle | Ctrl+Alt+Shift+I | Alignment | 2 |  |
| Top Border | Ctrl+Alt+Shift+Up | Borders | 2 |  |
| Bottom Border | Ctrl+Alt+Shift+Down | Borders | 2 |  |
| Left Border | Ctrl+Alt+Shift+Left | Borders | 1 |  |
| Right Border | Ctrl+Alt+Shift+Right | Borders | 1 |  |
| Outside Borders | Ctrl+Shift+7 | Borders | 3 |  |
| Inside Borders | Ctrl+Alt+Shift+7 | Borders | 3 |  |
| No Border | Ctrl+Shift+- | Borders | 3 |  |
| Font Style Cycle | Ctrl+Alt+O | Fonts | 3 |  |
| Font Size Cycle | Alt+Shift+G | Fonts | 3 |  |
| Increase Font | Ctrl+Shift+F | Fonts | 4 |  |
| Decrease Font | Ctrl+Shift+G | Fonts | 4 |  |
| Increase Table Size | Ctrl+Alt+Shift+F | Fonts | 2 |  |
| Decrease Table Size | Ctrl+Alt+Shift+G | Fonts | 2 |  |
| Paintbrush Capture | Ctrl+Alt+Shift+C | Paintbrush | 4 |  |
| Paintbrush Cycle | Ctrl+Alt+Shift+P | Paintbrush | 4 |  |
| Underline Cycle | Ctrl+Shift+U | Other Formatting | 3 |  |
| Case Cycle | Alt+Shift+C | Other Formatting | 2 |  |
| List Cycle | Ctrl+Alt+Shift+L | Other Formatting | 2 |  |
| Leader Dots | Alt+Shift+L | Other Formatting | 2 |  |
| Footnote Cycle | Ctrl+Alt+Shift+6 | Other Formatting | 2 |  |
| Footnote Toggle | Ctrl+Shift+6 | Other Formatting | 2 |  |
| Wrap Text | Ctrl+Shift+W | Other Formatting | 3 |  |
| Sum Bar | Ctrl+Shift+M | Other Formatting | 2 |  |
| Custom Cycle 1 | Ctrl+Alt+1 | Custom Cycles | 5 | v2 candidate |
| Custom Cycle 2 | Ctrl+Alt+2 | Custom Cycles | 5 | v2 candidate |
| Custom Cycle 3 | Ctrl+Alt+3 | Custom Cycles | 4 |  |
| Custom Cycle 4 | Ctrl+Alt+4 | Custom Cycles | 4 |  |
| Custom Cycle 5 | Ctrl+Alt+5 | Custom Cycles | 3 |  |
| Custom Cycle 6 | Ctrl+Alt+6 | Custom Cycles | 3 |  |
| Custom Cycle 7 | Ctrl+Alt+7 | Custom Cycles | 2 |  |
| Custom Cycle 8 | Ctrl+Alt+8 | Custom Cycles | 2 |  |
| Fast Fill Right | Ctrl+Shift+R | Modeling | 5 |  |
| Fast Fill Down | Ctrl+Shift+D | Modeling | 5 |  |
| Error Wrap | Ctrl+Shift+E | Modeling | 5 |  |
| Simplify Formula | Ctrl+Alt+Shift+Q | Modeling | 4 |  |
| Comment Formula | Ctrl+Alt+' | Modeling | 3 |  |
| Clean Cells | Ctrl+Shift+L | Modeling | 3 |  |
| Anchor Formula Cycle | Ctrl+F2 | Modeling | 3 |  |
| Flatten | Ctrl+Shift+3 | Modeling | 3 |  |
| Flip Sign | Ctrl+Shift+N | Modeling | 2 |  |
| Untranspose | Ctrl+Alt+Shift+T | Modeling | 2 |  |
| Wrap Parentheses | Ctrl+Alt+9 | Modeling | 2 |  |
| Paste Insert | Ctrl+Alt+I | Copy & Paste | 5 |  |
| Paste Number Formats | Ctrl+Alt+F | Copy & Paste | 2 |  |
| Paste Links | Ctrl+Alt+L | Copy & Paste | 2 |  |
| Paste Exact | Ctrl+Alt+E | Copy & Paste | 3 |  |
| Paste Duplicate | Ctrl+Alt+D | Copy & Paste | 5 |  |
| Paste Transpose | Ctrl+Alt+T | Copy & Paste | 4 |  |
| Trace In | Ctrl+Shift+[ | Auditing | 5 | **v1** |
| Trace Out | Ctrl+Shift+] | Auditing | 5 | v2 candidate |
| Last Audited Cell | Ctrl+Shift+\ | Auditing | 4 | **v1** |
| AutoTrace Precedents | Ctrl+Alt+Shift+[ | Auditing | 3 |  |
| AutoTrace Dependents | Ctrl+Alt+Shift+] | Auditing | 3 |  |
| Show All Precedents | Ctrl+Alt+[ | Auditing | 5 | v2 candidate |
| Show All Dependents | Ctrl+Alt+] | Auditing | 5 | v2 candidate |
| Clear Arrows | Ctrl+Alt+\ | Auditing | 5 | v2 candidate |
| Uniformulas | Ctrl+Q | Auditing | 4 |  |
| Zoom In | Ctrl+Alt+= | View | 4 |  |
| Zoom Out | Ctrl+Alt+- | View | 4 |  |
| Toggle Gridlines | Ctrl+Alt+G | View | 3 |  |
| Hide Page Breaks | Ctrl+Alt+B | View | 3 |  |
| Smart Print Area | Ctrl+Alt+Shift+B | View | 2 |  |
| Maximize View | Ctrl+Alt+Shift+W | View | 4 |  |
| Add Watch | Ctrl+Alt+W | Navigation | 4 |  |
| Next Watch | Ctrl+Alt+N | Navigation | 4 |  |
| Previous Watch | Ctrl+Alt+P | Navigation | 4 |  |
| Remove Watch | Ctrl+Alt+X | Navigation | 4 |  |
| First Sheet | Ctrl+Alt+Home | Navigation | 3 |  |
| Last Sheet | Ctrl+Alt+End | Navigation | 3 |  |
| Next Sheet Loop | Ctrl+Alt+PgDn | Navigation | 3 |  |
| Previous Sheet Loop | Ctrl+Alt+PgUp | Navigation | 3 |  |
| Activate Sheet | Ctrl+Shift+T | Navigation | 3 |  |
| Move Sheet Left | Ctrl+Alt+Shift+N | Navigation | 3 |  |
| Move Sheet Right | Ctrl+Alt+Shift+M | Navigation | 3 |  |
| Go To Min | Ctrl+Alt+Shift+, | Navigation | 2 |  |
| Go To Max | Ctrl+Alt+Shift+. | Navigation | 2 |  |
| Row Height Cycle | Alt+Shift+PgUp | Rows & Columns | 4 |  |
| Column Width Cycle | Ctrl+Alt+Shift+PgUp | Rows & Columns | 4 |  |
| AutoFit Height | Alt+Shift+PgDn | Rows & Columns | 3 |  |
| AutoFit Width | Ctrl+Alt+Shift+PgDn | Rows & Columns | 3 |  |
| Insert Row | Alt+Shift+Ins | Rows & Columns | 4 |  |
| Insert Column | Ctrl+Alt+Shift+Ins | Rows & Columns | 4 |  |
| Delete Row | Alt+Shift+Del | Rows & Columns | 4 |  |
| Delete Column | Ctrl+Alt+Shift+Del | Rows & Columns | 4 |  |
| Group Row | — | Rows & Columns | 3 |  |
| Group Column | — | Rows & Columns | 3 |  |
| Ungroup Row | — | Rows & Columns | 3 |  |
| Ungroup Column | — | Rows & Columns | 3 |  |
| Hide Row | Alt+Shift+Home | Rows & Columns | 3 |  |
| Hide Column | Ctrl+Alt+Shift+Home | Rows & Columns | 3 |  |
| Unhide Row | Alt+Shift+End | Rows & Columns | 3 |  |
| Unhide Column | Ctrl+Alt+Shift+End | Rows & Columns | 3 |  |
| Expand All Rows | Alt+Shift+= | Rows & Columns | 3 |  |
| Expand All Columns | Ctrl+Alt+Shift+= | Rows & Columns | 3 |  |
| Collapse All Rows | Alt+Shift+- | Rows & Columns | 3 |  |
| Collapse All Columns | Ctrl+Alt+Shift+- | Rows & Columns | 3 |  |
| Proper Hide | Ctrl+Alt+H | Rows & Columns | 2 |  |
| Copy Row/Column Info | Ctrl+Alt+J | Rows & Columns | 3 |  |
| Paste Row/Column Info | Ctrl+Alt+K | Rows & Columns | 3 |  |
| Quick Export To PowerPoint | Ctrl+Alt+Left | Export | 5 |  |
| Quick Export To Word | Ctrl+Alt+Right | Export | 5 |  |
| Quick Save | Ctrl+Shift+S | Utilities & Other | 5 |  |
| Quick Save All | Ctrl+Alt+Shift+S | Utilities & Other | 2 |  |
| Quick Save As | Alt+F12 | Utilities & Other | 4 |  |
| Quick Save Up | Shift+F12 | Utilities & Other | 4 |  |
| Reopen | Ctrl+Alt+Shift+O | Utilities & Other | 2 |  |
| Delete Comments & Notes | Ctrl+Alt+Shift+D | Utilities & Other | 3 |  |
