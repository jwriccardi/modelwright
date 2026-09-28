# 05 — Exact Macabacus keys and undo

*Researched 2026-09-28, after the owner ruled that **exact Macabacus keyboard shortcuts are critical**: "we don't want to ask users to learn new keyboard shortcuts".*

## 1. Which architectures can bind which keys

| Macabacus key | What it does | `OnKey` string | Office.js | Excel-DNA / VBA (Win) | VBA (Mac) |
|---|---|---|---|---|---|
| Ctrl+Shift+1/2/4/5/8 | Number cycles | `^+1` … | ✅ (one-time conflict prompt) | ✅ | ✅ with Control [UNVERIFIED] |
| Ctrl+Shift+K | Fill cycle | `^+k` | ✅ | ✅ | ✅ with Control [UNVERIFIED] |
| **Ctrl+'** | Font color cycle | `^'` | ❌ | ✅ | ? |
| **Ctrl+;** | Blue-black toggle | `^;` | ❌ | ✅ | ? |
| **Ctrl+Shift+[ / ]** | Pro Precedents / Dependents | `^+{[}` (hedge: `^{{}`) | ❌ | ✅ | ? |
| **Ctrl+Alt+[ / ]** | Show all arrows | `^%{[}` | ❌ | ✅ | ? |
| **Ctrl+, / Ctrl+.** | Decimals | `^,` / `^.` | ❌ | ✅ | ? |

### Office.js: ruled out for exact keys
- Microsoft's extended-manifest schema forces every key string to match `^[A-Za-z0-9-_+]+$`. We checked the schema directly (https://developer.microsoft.com/json-schemas/office-js/extended-manifest.schema.json).
- **One untested loophole:** Microsoft's own sample uses the named keys `Up`/`Down`. Names like `BracketLeft`, `Quote` or `Semicolon` would pass the regex.
  - We found no evidence that any such name works. None of the ~30 public `shortcuts.json` files we checked use one.
  - Testing it takes about 10 minutes (spike K1). **If it works, Office.js is back in contention.** It would give Windows + Mac with native undo.

### Excel-DNA and VBA on Windows
- Both register keys through `Application.OnKey` / `xlcOnKey`. That binds everything in the table above.
  - Bracket keys must be written in braces (`^{[}`); `"^["` throws an error.
  - Excel-DNA pattern: `XlCall.Excel(XlCall.xlcOnKey, "^+{[}", "TraceIn")` in `AutoOpen`, then cleared in `AutoClose`.
    - https://github.com/Excel-DNA/docs.excel-dna.net/blob/gh-pages/keyboard-shortcut.md
    - https://learn.microsoft.com/en-us/office/vba/api/excel.application.onkey
- **Limits.** These are the same limits Macabacus has; its own help center documents them.
  - **Add-ins that bind the same key: whoever calls `OnKey` last wins.** Macabacus's "Override" button just calls it again. CapIQ reportedly re-binds keys periodically. https://macabacus-help-center1.helpscoutdocs.com/article/917-keyboard
  - Shortcuts don't work while a cell is being edited.
  - Keyboard layouts where `[` needs AltGr (German, French) have no physical Ctrl+[ key [INFERRED].
- **Optional hardening:** a thread-level keyboard hook (`SetWindowsHookEx(WH_KEYBOARD)`) that uses virtual-key codes. It fixes the AltGr layouts and add-ins that steal keys, but Microsoft discourages it. Keep it out of v1.

### VBA on Mac
- `OnKey` works with **Control (`^`)**, but **not Cmd**, which modern Mac Excel can't detect.
- Whether punctuation keys fire is **unverified**.
- Macabacus itself doesn't run on Mac at all. Mac users have no Macabacus muscle memory; they'd only have Windows habits.
- Sources:
  - https://learn.microsoft.com/en-us/answers/questions/c1478637-9ee9-42d9-97f1-341fccec2811/applicationonkey-in-excel-2016-for-mac
  - https://sysmod.wordpress.com/2022/03/23/excel-for-mac-vba-onkey-macros-and-macos-monterey/

### Excel on the web
**No architecture can bind these keys on the web.** Web add-ins are Office.js only, so they have the same key restriction.

## 2. Undo

### What Macabacus does today
Macabacus users already live without Excel's native undo stack.
- Its help center says: "executing any code that changes a spreadsheet clears Excel's Undo/Redo stacks."
- So it keeps **its own undo/redo stacks**, driven by "native Excel shortcuts (Ctrl+Z and Ctrl+Y) and Quick Access Toolbar buttons".
  - It covers formatting changes only, with a cap on cells and properties.
  - It can't undo row or column inserts.
  - It slows formatting on large selections.
- https://macabacus-help-center1.helpscoutdocs.com/article/916-undo-redo

### Options outside Office.js
Ranked from best to worst:

1. **`CommandBars.ExecuteMso` with built-in format commands.** If Excel treats these as user commands, the native stack might survive. **Unverified in either direction.** Test it in 30 minutes (spike K2). Try it both called directly and deferred with `Application.OnTime`.
2. **A custom multi-level formatting-undo stack, as Macabacus does.**
   - Before each change, snapshot the number format, font color and fill of the affected cells.
   - Rebind Ctrl+Z and Ctrl+Y (`OnKey`) to our stack.
   - When our stack is empty, pass through to Excel's Undo.
   - This matches what Macabacus users experience now, **so it meets the "no relearning" bar.**
3. **`Application.OnUndo`.** Gives one undo level only; Ctrl+Z runs our procedure. This is the minimum fallback. https://learn.microsoft.com/en-us/office/vba/api/excel.application.onundo
4. **SendKeys Paste-Special-Formats trick** (Excel Campus). It keeps the native stack but is timing-dependent and fragile. Rejected.

### Office.js undo on the web
Native undo for add-ins is confirmed only for **Windows and Mac desktop** (ExcelApiDesktop 1.1 / ExcelApi 1.20). Microsoft says web support "will also be available on the web soon, based on customer demand". https://devblogs.microsoft.com/microsoft365dev/ignite-2025-whats-new-for-office-add-ins/

This corrects [02](02-architecture-options.md), which implied web support.

## 3. A possible bridge (research idea, not planned)

Excel-DNA binds the exact keys. Its handler changes **nothing** in the workbook; it only sends a message over localhost (for example a WebSocket) to an Office.js add-in running in a shared runtime. The Office.js add-in then applies the format with `Excel.run`, which gives native undo.

- Macros that don't change the workbook shouldn't clear undo [UNVERIFIED].
- This would give exact keys **and** native undo on Windows. The costs are two runtimes, localhost security questions, and more moving parts.
- It isn't possible on Mac: VBA there is sandboxed and has no sockets.
- **Keep it as a later experiment. v1 should not depend on it.**

## 4. What this means for the architecture

- **Windows desktop:**
  - **Excel-DNA** binds every Macabacus key and gives full COM access, including navigating into other workbooks, as Macabacus does.
  - Undo reaches Macabacus parity through option 1 or 2 above.
- **Mac:**
  - Only VBA can bind Control-based keys, and punctuation is unverified.
  - Mac users have no Macabacus habits to preserve.
  - So Mac is a *secondary* target, with different trade-offs.
- **Web:** exact keys are impossible there.

The owner's platform preference was "all three > Windows + Mac > Windows only if it makes a really big difference". Requiring exact keys **is** that really big difference, so **Windows desktop via Excel-DNA** is the primary target (ADR-0002).
