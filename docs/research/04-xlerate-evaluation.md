# 04 — XLerate evaluation: fork, contribute or build fresh?

*Evaluated 2026-09-28 against XLerate at HEAD v1.5.1 (2026-08-31). Repository: https://github.com/omegarhovega/XLerate.*

**About the project.**
- MIT, "Copyright (c) 2024 omegarhovega".
- An Office.js + TypeScript add-in covering format cycles, auto-color and a trace precedents dialog.
- Hosted on GitHub Pages and installed by sideloading.
- One contributor (199 commits), 24 stars, and no outside PRs ever merged.

## Findings

### Structure
- **Two products share one repo.**
  - A legacy VBA add-in: about 8k lines, plus a committed `.xlam` binary.
  - The Office.js add-in under `XLerate/`.
- **Clean layering in the Office.js code.**
  - `core/`: pure logic, about 2.9k LOC.
  - `adapters/`: the only code that calls `Excel.run`.
  - `services/`.
  - `taskpane/`: UI, about 2.6k LOC.
  - The boundaries are enforced by ESLint rules and dependency-cruiser. This is a good pattern to copy.
- **Tooling.**
  - Built with webpack and babel, from the Yo Office template.
  - The base `tsconfig` targets ES5, has **strict mode off**, and still lists IE11 in `browserslist`.
  - Strict mode applies only to `core/`, `adapters/` and `services/`.

### Manifest
- XML add-in-only manifest with a shared runtime (`lifetime="long"`).
- **It declares no requirement sets.** It checks the API version at runtime instead.

### Shortcuts
- Ctrl+Shift+1/2/3/4 (number, cell style, date, text style), Ctrl+Shift+R and Ctrl+Shift+0.
- No keys specific to Mac.
- **Trace has no shortcut** (open issue #8).

### Format cycles
- **Presets.** Three number formats. "Cell format" presets bundle fill, font and borders together.
- **No standalone font-color or fill-color cycle.**
- **Next-item rule.** Most cycles detect the cell's current format and apply the next one. The text-style cycle instead remembers its position in memory.
- **Settings are stored per workbook** (`document.settings`), not per user.
- **No `mergeUndoGroup`.** Each action relies on a single batch to make one undo step.
- **Performance weak spots.**
  - It reads cell properties for the whole selection, with no cap.
  - The number-format cycle writes cell by cell instead of assigning the whole range at once.

### Trace
- **Data.** `getDirectPrecedents` / `getDirectDependents`, loaded one level at a time as nodes are opened.
- **Large ranges.** Summarized above 50 cells and shown in pages of 100.
- **Limits.** Node and depth caps, duplicate cells merged, and a priority scheduler that can cancel work.
- **Parsing.** A hand-written tokenizer for clickable references, covering:
  - names;
  - structured table references;
  - external references (marked as not navigable).
- **UI.**
  - An **Office dialog window, not a task pane**, built in plain DOM.
  - Keyboard behavior is essentially Macabacus's: Up/Down select in Excel, Right/Left expand and collapse, Enter stays, Esc goes back.
  - No multi-step history.

### Tests and CI
- Vitest with about 234 tests covering `core/` and service contracts, run against fakes.
- **No UI tests.**
- CI runs typecheck, lint, dependency-cruiser, tests, manifest validation and the build, then deploys to Pages.

## Verdict: **build fresh, and borrow MIT code with attribution**

- **Why not fork.**
  - We would inherit the VBA tree and binary, the ES5, non-strict UI layer, and a manifest that declares no requirement sets.
  - Every product identity would need replacing.
  - We want a docked task pane, so the dialog UI would be rewritten anyway.
  - It has no color cycles, and its keys collide with Excel's built-ins.
- **Why not contribute upstream.**
  - It is a one-person project that has never merged an outside PR, and it works from a private spec.
  - Our UX differs: a task-pane trace, color cycles and per-user settings.

**Code worth borrowing**, keeping the MIT notice in `THIRD_PARTY_NOTICES.md`:

1. **Trace back end.**
   - Adapters: `adapters/tracePort*.ts`.
   - Core: `core/traceGraph.ts`, `tracePolicy.ts`, `traceRangePaging.ts`, `traceAreaMaterializer.ts`, `traceUtils.ts`.
   - Services: `services/traceSession.service.ts`, `traceScheduler.ts`.
   - Their tests.
2. **Formula references and tree state.**
   - `core/formulaReferences.ts` with its tests.
   - `core/traceDialogState.ts` (pure tree state).
   - The key-handling pattern in `taskpane/traceDialog.ts:395-452`.
3. **Cross-host color matching.**
   - `core/cellFormatCycle.ts` `doesFillMatch` / `normalizeColor`. These handle desktop reporting `pattern=null` and the web reporting solid white as "None".
   - The fill-mutation ordering in `adapters/excelPortLive.ts`.
4. **Knowledge, not code.**
   - The Office.js gotchas in their `CLAUDE.md`, including that `settings.saveAsync` breaks the undo chain and how focus returns via `workbook.focus()`.
   - Their `sideload-checklist.md` as a template for manual tests.

**Things that change our plan:**
- **Dialog vs task pane.** XLerate chose a dialog for trace, possibly because of the task-pane focus problem. Spike S5 must test **both** a task pane and a dialog for keyboard focus and `range.select()` behavior.
- **Undo.** Saving settings must not happen on the formatting hot path (the `saveAsync` gotcha).
- **Manifest.** Declare requirement sets: SharedRuntime 1.1, ExcelApi 1.12+, and DialogApi 1.2 if a dialog is used.
