# 02 — Add-in architecture options

*Researched 2026-09-28. The two claims the recommendation depends on were checked directly against Microsoft Learn on 2026-09-28; those are marked ✔︎ below.*

## What we require

| # | Requirement | Weight |
|---|---|---|
| R1 | Runs in Excel on Windows, Mac and the web. The owner prefers all three, will accept Windows + Mac, and will accept Windows-only only if that makes a *really big* difference. | High |
| R2 | Fast keyboard shortcuts. Finance modelers live on the keyboard. | High |
| R3 | **Undo still works** after a format cycle. | High |
| R4 | Trace precedents across sheets, and navigate to them and back. | High |
| R5 | Easy to install for someone who isn't a developer. | Medium |
| R6 | Friendly to open-source contributors: common language, CI, text-based source. | Medium |
| R7 | Microsoft will keep supporting it long-term. | Medium |

## Candidates

### A. Office Add-in (Office.js, TypeScript) ← **recommended**

**Undo is preserved ✔︎.**
- Changes made through the Excel JavaScript API now go onto the user's undo stack.
- `Excel.run({ mergeUndoGroup: true }, …)` groups several calls into a single undo step.
- The only APIs that clear the stack are on a published list: protection, sheet copy/delete, `Style.*` edits on desktop, and similar. **Range `numberFormat`, `font.color` and `fill.color` are not on that list.**
- Doc: https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-undo-capabilities
- Announcement: https://devblogs.microsoft.com/microsoft365dev/excel-announces-undo-support-for-3rd-party-add-ins/
- Available from ExcelApi 1.20: Windows Version 2509 or later, Mac 16.100 or later. **Perpetual and LTSC Office (including 2024) don't get it.**

**Keyboard shortcuts ✔︎.**
- Need SharedRuntime 1.1. Supported on Excel for Windows 2102 or later, Mac 16.55 or later, and the web.
- A shortcut is one or more modifiers (Ctrl/Cmd, Alt/Option, Shift) plus exactly one key.
- Shift can't be the only modifier.
- **Allowed keys: A–Z, 0–9 and `- _ +`** (Microsoft's own samples also use arrow keys).
- **So Ctrl+[, Ctrl+', Ctrl+; and Ctrl+\ can't be registered.** Those are Macabacus's signature keys.
- Built-in Excel shortcuts *can* be overridden, such as Ctrl+Shift+1. The first time a clashing shortcut is pressed, Office shows the user a dialog to choose which action wins. The choice is saved per user and per platform, and can be reset with "Reset Office Add-ins shortcut preferences".
- Users can remap shortcuts at runtime with `Office.actions.replaceShortcuts` (KeyboardShortcuts 1.1).
- On the web, shortcuts don't fire while the task pane has focus, and some browser combos can't be overridden: Ctrl+C/V/X/N/T/W, Ctrl+Shift+N/T/W/S/X, Ctrl+PgUp/PgDn, and others.
- Doc: https://learn.microsoft.com/en-us/office/dev/add-ins/design/keyboard-shortcuts
- There are reports of shortcuts being slow or unreliable on the web: https://github.com/OfficeDev/office-js/issues/3091, https://github.com/OfficeDev/office-js/issues/2270

**Precedents.**
- `Range.getDirectPrecedents()` (ExcelApi 1.12) and `getPrecedents()` (1.14) return a `WorkbookRangeAreas` grouped by sheet. **They cross sheets**, which COM's `DirectPrecedents` does not.
- **They do not cross workbooks**, and an add-in can't read or select cells in another open workbook.
  - We can parse external references out of the formula text and list them.
  - We **cannot** show their live values or jump to them.
- ExcelApiDesktop 1.1 (desktop only) adds `showPrecedents` arrows, `Worksheet.evaluate`, and window/scroll APIs. These help with navigate-and-return.
- https://learn.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-ranges-precedents-dependents
- https://learn.microsoft.com/en-us/javascript/api/requirement-sets/excel/excel-api-desktop-1-1-requirement-set

**Performance.**
- Every `context.sync()` is a round trip between processes; on the web it's a network round trip.
- A format cycle is 1–2 syncs, which should be fine. **Actual per-keypress latency is not measured yet**, so it's a spike item.

**Distribution.**
- Needs HTTPS hosting for the web assets. GitHub Pages works.
- Microsoft Marketplace (formerly AppSource) gives one-click install. Listing is free, through Partner Center, with 3–5 days of validation.
- Until then:
  - Microsoft 365 admins can deploy it to their organization.
  - Individuals can sideload it: through a network-share catalog on Windows, or "Upload My Add-in" on the web. Microsoft describes sideloading as "not for production".
- **There's no binary**, so no code signing, antivirus false positives or Mark-of-the-Web blocking.

**Manifest.** Start with the XML "add-in only" manifest, which works on every platform. The unified JSON manifest requires Windows 2501 or later / Mac 16.103 or later, and doesn't work on perpetual Office. Migrate later.

**Microsoft's outlook.** This is Microsoft's strategic platform, with active API releases through 2025–26.

### B. Excel-DNA (C#/.NET, packaged as an .xll) ← **fallback / optional Windows companion**

- **License and runtime.** Excel-DNA is zlib-licensed. Stable 1.9 supports .NET Framework 4.7.2 or later and .NET 6–10. 1.10 previews add Native AOT.
- **Shortcuts.** `xlcOnKey` / `Application.OnKey` can bind **any** combo, including Ctrl+[ and Ctrl+', and override built-in shortcuts silently.
- **Precedents.** Full COM object model, including cross-workbook navigation and opening linked workbooks.
- **Undo is cleared.** Any macro that changes the sheet (VBA, VSTO, Excel-DNA or C API) empties Excel's undo stack. The workaround is a home-grown undo stack; Macabacus does this and calls it "complex and costly". https://groups.google.com/g/exceldna/c/_NQvVgzpgdI , https://macabacus.com/docs/excel/undo-redo
- **Windows only.**
- **Install friction:**
  - Excel blocks an .xll downloaded from the internet (Mark of the Web) until the user unblocks the file or puts it in a Trusted Location.
  - Authenticode signing is recommended.
  - Packed .xll files have triggered antivirus false positives.
  - https://support.microsoft.com/en-us/topic/excel-is-blocking-untrusted-xll-add-ins-by-default-1e3752e2-1177-4444-a807-7b700266a6fb

### C. VSTO (.NET Framework COM add-in) — rejected
This is what Macabacus uses. It's stuck on .NET Framework 4.8, is in maintenance only, and runs only on Windows. It needs a VBA stub or a keyboard hook for shortcuts, and it also clears undo. It has no advantage over Excel-DNA for a new project.

### D. Native C/C++ XLL — rejected
It's the fastest option, but it also clears undo, runs only on Windows, costs far more to develop, and few contributors could work on it.

### E. VBA .xlam — rejected for the product, fine for throwaway prototypes
- Runs on Windows and Mac, and binds any key.
- Clears undo (`OnUndo` gives only one level).
- A downloaded .xlam is blocked by Mark of the Web, and signing doesn't get around it.
- The source is binary, and the UI toolkit is weak for a tree view.
- Breakdown (MIT) shows it can be done, but it isn't a good base for a maintained open-source product.

### F. Python (PyXLL / xlwings) — rejected
- PyXLL needs a paid license for every end user.
- Classic xlwings goes through COM, clears undo, and needs Python installed.
- xlwings Lite is really just Office.js.

## Comparison matrix

| | Office.js | Excel-DNA | VSTO | C++ XLL | VBA | Python |
|---|---|---|---|---|---|---|
| R1 Platforms | **Windows, Mac, web** | Windows | Windows | Windows | Windows, Mac | Windows |
| R2 Shortcuts | Good; restricted key set; one-time conflict prompt | **Any key, silent** | Awkward | Any | Any | Varies |
| R3 Undo | **Preserved** (M365 2509+) | Cleared | Cleared | Cleared | Cleared | Cleared |
| R4 Cross-sheet trace | **Native API** | Parser + COM | Parser + COM | Parser | NavigateArrow hack | COM |
| R4 Cross-workbook | ❌ list only | **✅ full** | ✅ | ✅ | ✅ | ✅ |
| R5 Install | Marketplace one-click / sideload | .xll + unblock + signing | MSI/ClickOnce | .xll + unblock | .xlam + unblock | Python |
| R6 Open-source friendly | **Excellent** (TS/npm) | Good (C#) | Fair | Poor | Poor | Mixed |
| R7 Outlook | **Strategic** | Stable (C API/COM kept for back-compat) | Maintenance | Legacy | Legacy | n/a |

## Recommendation

**Office.js + TypeScript is the primary and, for now, only codebase.**

It is the only option that meets R1 (all three platforms) and R3 (undo) together. Before September 2025, undo would have been the argument *against* Office.js. It is now the argument *for* it: every other option needs a custom undo engine for the three simple features.

**What we accept by choosing it** (each is tested in the Phase 1 spike):

1. **Macabacus's exact keys aren't available.**
   - Ctrl+Shift+[ (trace), Ctrl+' (font cycle) and Ctrl+; (blue-black) can't be registered.
   - We ship a default keymap built from letters and numbers, and users can remap it.
   - Ctrl+Shift+1/4/5/8 and Ctrl+Shift+K *can* be registered, with a one-time conflict prompt.
2. **Precedents in other workbooks can be listed but not opened or navigated to.**
3. **Undo requires Microsoft 365 builds of 2509 or later.** On perpetual Office and LTSC, the formatting commands will still work but will clear undo, and we'll show a one-time notice.
4. **Per-keypress latency is unmeasured.** Pass/fail thresholds are defined in the PLAN spike.

**When to switch to Excel-DNA instead.** Add an optional Windows-only Excel-DNA companion (a second codebase) only if the spike shows either of these:
- Office.js shortcut latency or reliability is unacceptable;
- cross-workbook trace with navigation is a must-have for the owner.

The unified manifest's `alternates` element can declare the .xll as the preferred version on Windows. We are **not** planning this now.

## Formula parsing (for display, function grouping and external references)

| Library | Language | License | Notes |
|---|---|---|---|
| excel-formula-tokenizer + excel-formula-ast | TS/JS | MIT | Small; a port of Bachtal's tokenizer. Good enough for splitting a formula into references and groups. |
| Formualizer | Rust → WASM | MIT/Apache-2.0 | Newer. Full parser plus evaluator. |
| fast-formula-parser | JS | MIT | Chevrotain-based, but unmaintained since about 2020. |
| HyperFormula | TS | **GPLv3 or commercial** | **Avoid.** Its license would conflict with ours. |
| XLParser | C# | MPL-2.0 | Only relevant if we take the Excel-DNA path. |
| ClosedXML.Parser | C# | MIT | Only relevant if we take the Excel-DNA path. |
