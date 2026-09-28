# Work plan — Excel Modeling Toolkit

> **Status: PENDING APPROVAL.** Planning only. No code has been written yet.
> Research: [`research/01-feature-survey.md`](research/01-feature-survey.md), [`research/02-architecture-options.md`](research/02-architecture-options.md), [`research/03-licensing.md`](research/03-licensing.md), [`research/04-xlerate-evaluation.md`](research/04-xlerate-evaluation.md).
> Decisions: [`decisions/0001-platform-architecture.md`](decisions/0001-platform-architecture.md). Open items: [`open-questions.md`](open-questions.md).

## 1. Requirements summary

An open-source Excel add-in for financial modelers that does four things:

1. **Number format cycling.** A shortcut steps the selection through an ordered, user-editable list of number formats.
2. **Font color cycling.** The same mechanism for font color.
3. **Fill color cycling.** The same mechanism for fill color, including a "No fill" slot.
4. **Smart trace precedents.** A keyboard-driven tree of the active cell's precedents, with values, that works across sheets. Moving through the tree selects each cell in Excel, and you can jump back to where you started.

**Constraints:**
- Runs on Excel for Windows, Mac and the web (owner's preference: all three > Windows + Mac > Windows only).
- Ctrl+Z must undo every formatting action.
- Settings are per user and can be exported.
- Will be open-sourced later.

**Out of v1, but the design should leave room for:** AutoColor, trace dependents, "evaluate functions & groups", custom style cycles, and cross-workbook navigation.

## 2. RALPLAN-DR summary

**Principles**
1. The keyboard comes first. Every feature can be used without the mouse.
2. Never break the user's undo.
3. One codebase for every platform, unless a measured blocker forces otherwise.
4. Keep logic separate from Excel. Cycle and precedent logic is pure TypeScript that can be unit-tested without Excel.
5. Configuration is data. Cycles, palettes and keymaps are JSON that users can export and share.

**Decision drivers**
1. Support all three platforms (R1).
2. Undo still works (R3).
3. Keyboard latency and reliability (R2).

**Options considered.** Full comparison in research/02.

| Option | Verdict |
|---|---|
| **A. Office.js + TypeScript** | **Chosen.** The only option that meets drivers 1 and 2. Driver 3 is tested in the Phase 1 spike. |
| B. Excel-DNA (C#) | Fallback or optional Windows companion. Best keys and cross-workbook support, but Windows only and clears undo. |
| C–F. VSTO / C++ XLL / VBA / Python | Rejected: see research/02. |

## 3. Build vs fork

*See [`research/04-xlerate-evaluation.md`](research/04-xlerate-evaluation.md).*

XLerate (MIT, Office.js + TS, one maintainer) already has format cycles, auto-color and a trace dialog.

**Recommendation: build fresh, and borrow MIT code with attribution.**

- **Why not fork:**
  - It carries a legacy VBA tree and a committed binary.
  - Its UI layer is ES5 and non-strict.
  - Its manifest declares no requirement sets.
  - It has no color cycles.
  - Its trace UI is a dialog, not a docked pane.
- **Why not contribute upstream:** it has never merged an outside PR, and it works from a private spec.
- **What we borrow** (keeping their MIT notice in `THIRD_PARTY_NOTICES.md`):
  - the trace back end (precedent loading, range paging, scheduler, tree state);
  - the formula-reference tokenizer;
  - the cross-host color-matching helpers.

  This removes an estimated one-third of the work in Phase 4.

## 4. Proposed design (no code yet)

### 4.1 Module layout

```
src/
  core/        pure TS, no Office.js: cycle engine, format/color matching,
               formula tokenizing and reference extraction, precedent-tree model,
               settings schema and migrations
  excel/       thin Office.js adapters (read state, apply format, get precedents, select)
  commands/    shortcut and ribbon handlers (Office.actions.associate)
  taskpane/    trace pane and settings UI
manifest/      XML add-in-only manifest + shortcuts.json
test/          unit tests (core), adapter tests with office-addin-mock, fixture workbooks
docs/
```

### 4.2 How the cycles work (features 1–3)

**Cycle definition.**
- `{ id, name, kind: "numberFormat" | "fontColor" | "fillColor", items: [...] }`.
- Number-format items carry `{ name, code }`.
- Color items carry `{ name, rgb | "none" | "automatic" }`.

**Choosing the next item (proposed hybrid rule):**
1. If the previous command was *this cycle on the same selection address*, move to the next position (position + 1, wrapping).
2. Otherwise, read the active cell's current value. If it equals item *k*, apply item *k+1*. If it matches nothing, apply item 1.

On a clean cell this behaves like Macabacus ("press 3× to reach item 3"). Like XLerate, it also continues from wherever a cell already is. **Check against Macabacus** (open-questions A3).

**Other rules:**
- **Mixed selections.** The active cell decides the "current" value. The result is applied to the whole selection.
- **Undo.** One key press is one `Excel.run({ mergeUndoGroup: true })`, which is one undo step. If ExcelApi 1.20 is missing, the command still works, and a one-time notice says undo isn't available on this Excel build.
  - Never save settings on the formatting path: `document.settings.saveAsync` breaks the undo chain (XLerate finding).
- **Selection size.** Read only the active cell's format, and write the whole selection in one range assignment, never cell by cell.
- **Number-format limit.** Catch the "invalid number format" error and show a clear message.
- **Default contents.**
  - Number formats use Macabacus-style alignment conventions (see research/01 §1), with `[$$]` for currency.
  - Colors follow the blue/black/green/red convention.
  - Fill includes a **No fill** slot.
  - Final values will be taken from the installed Macabacus (open-questions A1–A2).

**v1 cycles (proposal D5):**
- Number, Percent, Multiple, Currency and Date. These are one engine with five sets of data.
- One font cycle and one fill cycle.

### 4.3 Default keymap (to be tested in the spike)

The keys must fit Office.js's allowed set. Keys can be set per platform (`windows` / `mac` / `web`), and users can remap them at runtime (KeyboardShortcuts 1.1).

| Action | Windows / Mac | Web | Notes |
|---|---|---|---|
| Number cycle | Ctrl+Shift+1 | same | Same as Macabacus; overrides Excel's Number format (one-time conflict prompt) |
| Date cycle | Ctrl+Shift+2 | same | Same as Macabacus |
| Currency cycle | Ctrl+Shift+4 | same | Same as Macabacus |
| Percent cycle | Ctrl+Shift+5 | same | Same as Macabacus |
| Multiple cycle | Ctrl+Shift+8 | same | Same as Macabacus |
| Font color cycle | Ctrl+Shift+C *(candidate)* | Ctrl+Alt+Shift+C | Macabacus's Ctrl+' isn't allowed. The web variant avoids the browser DevTools shortcut. |
| Fill color cycle | Ctrl+Shift+K | same *(verify)* | Same as Macabacus |
| Trace precedents | Ctrl+Shift+P *(candidate)* | Ctrl+Alt+Shift+P | Macabacus's Ctrl+Shift+[ isn't allowed. The web variant avoids Edge InPrivate / Firefox private-window shortcuts. |
| Return to audited cell | TBD in spike | TBD | |

**Mac.** Cmd+Shift+3/4/5 are macOS screenshot keys, so the Mac number cycles use **Ctrl**, not Cmd (verify).

### 4.4 Smart trace precedents (feature 4, v1)

**Invoking it.**
- A shortcut or ribbon button on the active cell opens the task pane.
- If the cell has no formula, show a message instead.

**Pane layout.**
- Header: the audited cell's address and its formula.
- Tree rows: the reference as written, the resolved address, the value, and an ƒ badge if that precedent is itself a formula.

**Where the data comes from.**
- The **formula parser** gives the order and labels. Rows are listed in the order they're written, and named ranges and table references keep their names.
- **`getDirectPrecedents()`** gives the addresses across sheets and validates the parse.
- Children load when a node is expanded (lazy loading).

**Keys.**

| Key | Action |
|---|---|
| Up / Down | Select the node's range in Excel (switching sheet if needed) |
| Right | Expand |
| Left | Collapse, or go to the parent |
| Enter | Close and stay on the current cell |
| Esc | Close and return to the audited cell |
| "Return to audited cell" command | Walks back through a history of up to 20 audits |

**Edge cases.**

| Case | v1 behavior |
|---|---|
| Multi-cell range | One node, e.g. "A1:A100 (100 cells)". Expands in pages of 50. |
| Circular path | Marked ↻ and can't be expanded |
| External-workbook reference | Listed with a 🔗 badge; **can't be navigated to** (Office.js limit) |
| Hidden sheet or rows | Marked with a badge; selecting it doesn't throw an error |
| INDIRECT / OFFSET | Whatever `getDirectPrecedents` returns, labelled "dynamic" |

**Not in v1:** dependents, evaluating functions and groups, arrows, cross-workbook navigation.

### 4.5 Settings
- Stored per user in the add-in's storage (`OfficeRuntime.storage` or `localStorage`; chosen in spike S7).
- Versioned JSON schema.
- **Export and import JSON** in the settings pane. This is also how users share settings, as Macabacus users do.
- A settings UI to edit, reorder and preview cycle items.

## 5. Phases and acceptance criteria

### Phase 0 — Decisions (now)
- D1–D7 in `open-questions.md` are resolved.
- The Macabacus checks (A1–A11) are captured.

**Exit:** ADR-0001 is accepted, a LICENSE is chosen, and there's a build-vs-fork decision.

### Phase 1 — Spike (throwaway `spike/*` branch)

Each test must pass on Windows desktop (M365 ≥ 2509), Mac (≥ 16.100) and Excel on the web, unless it says otherwise.

| # | Test | Pass criterion |
|---|---|---|
| S1 | Shortcut → format applied latency | Over 30 presses, p95 ≤ 150 ms on Windows desktop and ≤ 300 ms on the web (measured with `performance.now` in the handler plus a screen-recording check) |
| S2 | Undo | 3 cycle presses, then 3× Ctrl+Z, restore the original format exactly |
| S3 | Conflict prompt on Ctrl+Shift+1 | Appears once, and the choice persists across an Excel restart |
| S4 | Cold start | A shortcut works on the first press after opening a workbook, without opening the pane first (`setStartupBehavior(load)`) |
| S5 | **Pane focus** | Test **both a task pane and an Office dialog** (XLerate chose a dialog). After the trace shortcut, Up/Down must reach the trace UI, not the grid, and `range.select()` must leave focus in the trace UI. Pick whichever passes on all platforms; prefer the task pane if both do. |
| S6 | Precedent correctness | Fixture workbook: cross-sheet, names, tables, whole columns, INDIRECT, external refs. Every expected precedent is found, or a documented limitation explains why not. |
| S7 | Settings persistence | Survives an Excel restart on each platform |

**Exit:** a go/no-go note in `docs/spike-results.md`. **If S1 or S5 fail on Windows, reopen ADR-0001** (consider an Excel-DNA companion).

### Phase 2 — Scaffold

**Deliverables:**
- TypeScript in strict mode, ESLint, and Vitest.
- GitHub Actions (lint, typecheck, test, build).
- GitHub Pages hosting.
- XML manifest and `shortcuts.json`.
- `LICENSE`, `CONTRIBUTING.md` (DCO), and a license check that fails the build on GPL or AGPL dependencies.

**Exit:**
- CI passes on an empty feature set.
- The add-in sideloads on all three platforms and shows its ribbon tab.

### Phase 3 — Features 1–3 (cycles)

**Exit:**
- Unit tests cover the cycle engine: wrap-around, the hybrid next-item rule, mixed selections and color normalization. `core/` has ≥ 90% line coverage.
- Each of the 7 v1 cycles passes S1 and S2 on all three platforms.
- The settings pane can edit, reorder, preview, reset, export and import, and an export → import round trip gives identical JSON.
- Hitting the number-format limit produces a clear message rather than a silent failure.

### Phase 4 — Feature 4 (trace precedents v1)

**Exit:**
- On the fixture workbook, every case in §4.4 behaves as specified.
- Expanding a node takes ≤ 2 `context.sync()` calls.
- A formula with ≤ 20 references renders in ≤ 500 ms on Windows desktop.
- Unit tests cover reference extraction and tree-model logic.

### Phase 5 — Release

**Exit:**
- The README has install and sideload instructions for each platform.
- A manual test script has been run and recorded on all three platforms.
- The repo is made public, with a v0.1.0 tag.
- Optionally, a Marketplace submission.

## 6. Risks and mitigations

| Risk | Likelihood / impact | Mitigation |
|---|---|---|
| Office.js shortcut latency feels sluggish next to Macabacus | Medium / High | Spike S1 with hard thresholds. Keep cycle state in memory to avoid reads. Fallback: Excel-DNA companion (ADR-0001). |
| Task pane doesn't take keyboard focus, so arrow keys move the grid instead | Medium / High | Spike S5 first. Fallback: an in-pane "click to focus" and explicit keys in the pane. |
| Users miss Macabacus keys (Ctrl+', Ctrl+Shift+[) | High / Medium | Document the mapping, let users remap in-app, and explain the platform limit in the README. |
| Enterprise users on LTSC or perpetual Office have no add-in undo | Medium / Medium | Detect ExcelApi 1.20, show a one-time notice, and document supported builds. |
| Can't navigate into other workbooks | Certain / Low–Medium | List them with a badge. Revisit with an Excel-DNA companion if users demand it. |
| Sideloading friction before Marketplace listing | High / Medium | Write clear per-platform instructions, and plan a Marketplace submission for v0.1. |
| Clash with Macabacus/FactSet installed alongside | Medium / Low | Office's conflict prompt resolves it per user. Document it. |

## 7. Verification
- **Unit:** Vitest on `core/` (pure logic), in CI.
- **Adapter:** `office-addin-mock` for `excel/` adapters, in CI.
- **End-to-end:** a scripted manual test run on the three platforms using `test/fixtures/*.xlsx`, recorded in `docs/test-runs/`.
- **Performance:** timing hooks in development builds, reporting p50/p95 per command (S1).

## 8. ADR
See [`decisions/0001-platform-architecture.md`](decisions/0001-platform-architecture.md).

## Changelog
- 2026-09-28: first draft from research.
- 2026-09-28: added the build-vs-fork recommendation (research/04). Spike S5 now tests a dialog as well as a task pane. Added the undo/`saveAsync` and selection-size rules.
