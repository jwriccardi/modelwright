# Work plan — Excel Modeling Toolkit

> **Status: Phases 0–3 complete. PRs #1–#4 were merged on 2026-10-06 after the owner tested in Excel, and the automated smoke test `tests/excel-smoke/undo-smoke.ps1` passes. Phase 4 (Trace In) is in progress: 4a core logic, then 4b the trace window.**
>
> - Architecture: [`decisions/0002-excel-dna-windows-first.md`](decisions/0002-excel-dna-windows-first.md). It replaces ADR-0001 (Office.js).
> - Research: [`research/01`](research/01-feature-survey.md) features · [`02`](research/02-architecture-options.md) architectures · [`03`](research/03-licensing.md) license · [`04`](research/04-xlerate-evaluation.md) prior art · [`05`](research/05-keys-and-undo.md) keys and undo.
> - Open items: [`open-questions.md`](open-questions.md).

## 1. Requirements summary

An open-source Excel add-in for financial modelers. It does four things:

1. **Number format cycling**
2. **Font color cycling**
3. **Fill color cycling**
4. **Smart trace precedents**

**Hard requirement (owner, 2026-09-28):** the add-in must use **exactly Macabacus's keyboard shortcuts**. Users must not have to learn new keys.

**Other constraints:**
- Undo must work at least as well as it does in Macabacus.
- Settings are per user and can be exported.
- The project will be open-sourced later.

**Platform:** **Windows desktop in v1.** Exact keys can't be bound in Office.js, which is the only way to run on the web (research/05). Mac and web are deferred.

**Not in v1, but the design leaves room for:**
- AutoColor
- Pro Dependents (Ctrl+Shift+])
- Evaluate functions & groups
- Custom style cycles
- The Mac version

## 2. RALPLAN-DR summary

**Principles**
1. Macabacus muscle memory is the spec: the same keys, and behavior as close to Macabacus as possible.
2. Never make undo worse than Macabacus does.
3. Keep logic out of Excel code: cycle, undo and precedent logic lives in a pure C# library with unit tests.
4. Configuration is data: cycles, palettes and the keymap are JSON that users can export and share.
5. Keep installing easy: one per-user install, no admin rights, and a signed binary.

**Decision drivers**
1. Exact keys
2. Undo parity
3. Precedent navigation parity, including into other workbooks

**Options considered.** Details are in ADR-0002.

| Option | Verdict |
|---|---|
| **Excel-DNA (C#, .NET Framework 4.8)** | **Chosen**: meets all three drivers. |
| Office.js | Rejected: it can't bind `[ ' ; , .`. It comes back only if spike K1 finds that named punctuation keys work. |
| VBA .xlam | Only as a later Mac port. |
| Excel-DNA↔Office.js bridge | A later experiment. |
| VSTO / C++ | Rejected. |

## 3. Build vs fork
XLerate is TypeScript/Office.js, so it can't be forked for this architecture. We use it as a **design reference only**:
- range paging thresholds;
- how it merges duplicate precedents;
- its tree-state and keyboard model;
- its cross-host color-matching lessons.

No code is copied, so no attribution is needed unless we later port a specific algorithm. See research/04.

## 4. Proposed design (no code yet)

### 4.1 Solution layout
```
src/
  Toolkit.Core/      netstandard2.0, no Excel references: cycle engine, color and format
                     matching, settings schema + migrations, undo snapshot model,
                     formula reference extraction (XLParser or ClosedXML.Parser),
                     precedent-tree model
  Toolkit.AddIn/     net48 + Excel-DNA: AutoOpen/AutoClose (key registration),
                     commands, COM adapters, ribbon XML, UndoManager,
                     trace window, settings dialog
  Toolkit.Tests/     xUnit tests for Core (CI runs them on windows-latest)
installer/           per-user installer (no admin rights)
test/fixtures/       fixture workbooks for manual end-to-end runs
```

### 4.2 Keymap (Macabacus defaults; users can remap them)

| Action | Key | `OnKey` string | v1? |
|---|---|---|---|
| General Number cycle | Ctrl+Shift+1 | `^+1` | ✅ |
| Date cycle | Ctrl+Shift+2 | `^+2` | ✅ |
| Local Currency cycle | Ctrl+Shift+4 | `^+4` | ✅ |
| Percent cycle | Ctrl+Shift+5 | `^+5` | ✅ |
| Multiple cycle | Ctrl+Shift+8 | `^+8` | ✅ |
| Font Color cycle | Ctrl+' | `^'` | ✅ |
| Fill Color cycle | Ctrl+Shift+K | `^+k` | ✅ |
| Pro Precedents | Ctrl+Shift+[ | `^+{[}` (also register `^{{}` as a hedge) | ✅ |
| Last Audited Cell | Ctrl+Shift+\ (confirmed, research/06) | `^+\` / `^|` | ✅ |
| Blue-Black toggle | Ctrl+; | `^;` | stretch |
| Increase / Decrease decimals | Ctrl+, / Ctrl+. | `^,` / `^.` | stretch |
| Undo / Redo (formatting stack) | Ctrl+Z / Ctrl+Y | `^z` / `^y` | ✅ (see §4.4) |
| Pro Dependents, Show All arrows | Ctrl+Shift+], Ctrl+Alt+[ ] | … | v2 |

- **Registration.** Keys are registered in `AutoOpen` and cleared in `AutoClose`.
- **"Override" command.** Re-registers every key, as Macabacus's Override button does, for when another add-in has taken them.
- **Coexistence.** If Macabacus is also loaded, the add-in that registered a key last owns it.

### 4.3 Cycles (features 1–3)
- **Definition:** `{ id, name, kind: numberFormat | fontColor | fillColor, items[] }`.
- **Choosing the next item (hybrid rule):**
  1. If the previous command was *this cycle on the same selection address*, move to the next position.
  2. Otherwise, match the active cell's current value against the list. If it equals item *k*, apply item *k+1*; if nothing matches, apply item 1.
  3. The list wraps around from the last item to the first. This must be checked against the installed Macabacus (open-questions A3).
- **Applying a format.** Read only the active cell. Write to the whole selection in one COM call each:
  - `Selection.NumberFormat = code`
  - `Font.Color`
  - `Interior.Color`, or `Interior.Pattern = xlNone` for "No fill"
- **Default contents.** Taken from the owner's installed Macabacus (research/06):
  - The General Number cycle has 4 codes.
  - Font: Blue → Green → Purple → Red → White → Black.
  - Fill: (201,218,248) → (210,242,255) → (244,204,204) → (252,229,205) → Navy (28,69,135) → No fill.
  - Percent, Currency, Multiple and Date lists are still to be captured.
- **Too many number formats.** Excel's "too many number formats" error is caught and shown as a clear message.

### 4.4 Undo: Macabacus parity in v1, native undo through a hybrid in v2

Decided by the owner on 2026-09-28, based on spike K2/K2b/K2c (see `docs/spike-results.md`).

**What the spikes established:**
- **Any COM write** to any cell (even in a hidden workbook, even outside macro context) **erases Excel's undo history**, and the change isn't undoable.
- **Built-in commands** (`ExecuteMso`) dispatched *outside* macro context (from the keyboard hook or a ribbon callback) are recorded exactly like user actions. They give full multi-level undo and keep the earlier history, even with volatile formulas. But they exist only for fixed formats: Bold, Percent Style and similar.
- Built-in commands run *inside* an `OnKey` macro lose the earlier history. They also collapse to one level when the workbook has volatile formulas.

**v1 design: our own `UndoManager` (Macabacus parity).**
- **Snapshot.** Before each change, snapshot the property being changed (number format, font color or fill) of the affected cells.
  - Capture is capped by `undoCellCap` (default 10,000 *reads*; uniform regions cost one read).
  - Uniform ranges are run-length compressed.
- **Ctrl+Z / Ctrl+Y are intercepted.** The thread keyboard hook is preferred: it can *decline* to swallow the key.
- **Ordering rule.**
  - If Excel's native undo is available (`GetEnabledMso("Undo")`), the entries on it must be newer than our last action, because our action wiped the stack. So let the key through to Excel.
  - Otherwise, pop our own stack.
- **Invalidation.** Clear our stack when the workbook is closed, and on structural changes (row or column insert or delete).
- **Where native commands exist, use them (optional).** If a cycle item matches a built-in command, such as Bold, dispatch it from the hook outside macro context. That keeps the user's history intact for that action.

**Refinements from the Phase 3b code review (2026-09-29):**
- **Barrier.** A change we couldn't record (the capture cap was exceeded, a patterned or gradient fill, a failed read) is pushed as a *barrier*. Ctrl+Z stops there with "Can't undo: … could not be recorded", so undo never jumps past it to an older change. Every write, recorded or not, clears our redo stack.
- **Redo staleness.** Our restore wipes Excel's undo *and* redo. So if Excel reports any undo or redo entries while our redo stack is non-empty, the user has acted since, and our redo is stale: clear it and pass the key to Excel. The same applies to Ctrl+Z when Excel's undo is available.
- **Identity and invalidation.**
  - Snapshots are keyed by the workbook's **FullName** (full path).
  - Snapshots are invalidated on `WorkbookBeforeClose`, and on `SheetChange` events whose target is whole rows or columns (row or column inserts and deletes).
  - These events come through a late-bound COM event sink, so no Office PIA dependency is needed.
- **The capture cap counts reads, not cells.** Mixed rectangles are split recursively along their longer side, so a uniform region costs one read and whole-column selections stay cheap.
- **Color fidelity.** A font set to *Automatic* is restored as Automatic, not black. Theme-color links are not preserved; the RGB value is restored.
- **Hook safety.**
  - The hook makes no Excel calls while the mouse is captured, while a menu is open, or while a window is being moved or sized.
  - It guards against re-entry.
  - At restore time it checks Excel's undo state again.
  - Holding Ctrl+Z undoes one step, and Excel's own Undo/Redo buttons are not intercepted.
- **Memory.** A global block budget evicts the oldest snapshots.
- **Owner test finding (2026-10-06).** Read from *macro context* (inside an add-in command), `GetEnabledMso("Undo")` is unreliable right after a native undo: it reported `true` while Excel's undo list was empty. The reading taken in the hook (outside macro context) was correct.
  - The restore therefore no longer re-checks Excel's state. The hook alone decides, including whether our redo is stale.
  - The ribbon Undo/Redo buttons always act on our stack.

**v2 candidate: the hybrid (spike K2d, deferred).**
- The Excel-DNA hook catches the exact keys and forwards the command over a local channel to a small Office.js add-in in the same Excel.
- Office.js applies `numberFormat` / `font.color` / `fill.color`, which are native-undoable in ExcelApi 1.20+.
- **Payoff:**
  - exact keys, exact formats *and* full native undo;
  - a formatting engine that can be reused for Mac and web.

### 4.5 Smart trace precedents (feature 4, "Trace In")

Behavior spec: [research/07](research/07-macabacus-trace-in-spec.md).

- **Invoke.** Ctrl+Shift+[ opens a modeless window owned by Excel, titled "Trace In". It remembers its position.
- **Keyboard model (decided by spike K4):**
  - **Preferred (variant C, likely how Macabacus does it):** the window **doesn't take focus**. Excel keeps focus, and a thread-level keyboard hook sends Up/Down/Left/Right/Enter/Esc/Ctrl+E to the tree. This is the only model that also supports F2 editing in Point mode with the dialog open.
  - **Fallback:** the window takes focus (WPF or WinForms) and is re-activated after each Goto.
- **Layout.**
  - Formula header with color-coded references and a wrap toggle.
  - Tree with columns **Precedents | Argument | Value**.
- **v1 tree ("classic" mode):** the root is the audited cell, with one child per reference in the order written. A reference can be:
  - a cell or range;
  - a name;
  - a table reference;
  - an external workbook reference.

  Values come from `Value2` (the first cell for ranges, plus the count).
- **v1.1 tree ("Evaluate functions & groups", Ctrl+E):** the nodes follow the formula's structure.
  - Parenthesized groups are `(x)` nodes, and functions are `ƒx NAME(...)` nodes.
  - A function's children are its arguments, labelled with Excel's parameter names from a **function signature table** (the top ~100 functions first).
  - Each node's value comes from `Worksheet.Evaluate(subexpression)` in the audited cell's sheet context.
  - A function that returns a reference (INDEX, OFFSET, INDIRECT, CHOOSE) is resolved to its target range, so you can navigate to it.
- **How precedents are found.** Parse `Range.Formula` (invariant A1) with XLParser, which also produces the structure for v1.1. Then resolve each reference:
  - **A1 references** → `Range`.
  - **Names** → `Names(...).RefersToRange`.
  - **Table references** → `ListObjects`. Unqualified `[Col]` is resolved via the table containing the audited cell. Macabacus can't do this, so it's a small improvement.
  - **External `[Book]Sheet!A1`:**
    - If the workbook is open, navigate into it.
    - If it's closed, **try to open it**, as Macabacus does by default.
  - **Check:** compare with `Range.DirectPrecedents` for references on the same sheet.
- **Keys.**

| Key | Action |
|---|---|
| Up / Down | `Application.Goto` the node's range |
| Right | Expand one level (children load only then) |
| Left | Go up a level or collapse |
| Enter / OK | Stay on the current cell |
| Esc / Cancel | Return to the audited cell (confirm, open-questions A7) |
| F2 | Edit in Point mode with the dialog open (depends on variant C) |
| Ctrl+E | Evaluate mode (v1.1) |
| Ctrl+Arrows / Ctrl+Home / Ctrl+End / Shift+Arrows | Move, snap and resize the dialog |

- **History.** Last Audited Cell (`Ctrl+Shift+\`) keeps a stack of up to 20 audits.
- **Edge cases.**
  - Ranges over 50 cells are shown as one node; expanding it shows pages of 100.
  - Circular references are marked ↻.
  - Hidden sheets, rows and columns get a badge.
  - Merged cells are handled: the node shows the merge area.
- **v2:**
  - Trace Out (Ctrl+Shift+], using Excel's own dependency data).
  - Show All Precedents/Dependents (Ctrl+Alt+[ / ]).
  - Clear Arrows (`Ctrl+Alt+\`).
  - AutoTrace (≤ 20 cells).
  - Highlight Navigated Cells, built as an overlay that **doesn't clear undo**, unlike Macabacus's version.

### 4.6 Settings
- **Storage.** `%AppData%\<ProductName>\settings.json`, with a versioned schema.
- **Settings dialog.** Edit, reorder and preview cycles, remap keys, and set the undo cap. Reset, export and import are included.

## 5. Phases and acceptance criteria

### Phase 0 — Decisions (now)
- Resolve D1–D8 in `open-questions.md`.
- Capture the Macabacus observations A1–A11.
- **Exit:** ADR-0002 approved in principle; license chosen; copyright holder decided.

### Phase 1 — Decision spikes (throwaway code, needs owner approval)

**Before starting:** disable Macabacus, or change its keys, on the test machine.

| # | Test | Pass criterion | If it fails |
|---|---|---|---|
| K1 ✅ (the keys work, but ADR-0001 stays superseded: D11) | Office.js named punctuation keys (`Ctrl+Shift+BracketLeft`, `Ctrl+Quote`, `Ctrl+Semicolon`) | The keys register and fire on Windows | Stay with ADR-0002. **If they work → reopen ADR-0001.** |
| K2 | `ExecuteMso` formatting keeps native undo | Type a value → cycle via ExecuteMso → Ctrl+Z twice undoes both | **Done:** partly passes (fixed formats only, outside macro context). v1 uses our own UndoManager (§4.4). |
| K3 ✅ | Excel-DNA binds every key in §4.2 | 100% fire. Over 30 presses on a selection of ≤ 1,000 cells, p95 key → format ≤ 50 ms | Try a thread keyboard hook |
| K4 ✅ (C and B pass; F2 on C passed 2026-10-06) | Trace window keyboard model: **A** WPF with focus, **B** WinForms with focus, **C** a window that doesn't take focus plus a thread keyboard hook (Macabacus-style) | Up/Down/Left/Right/Enter/Esc reach the tree through `Goto` to another sheet and to another workbook. For C, F2 also passes through to Excel. | Choose the best variant that passes; prefer C |

**Exit:** a go/no-go note in `docs/spike-results.md`.

### Phase 2 — Scaffold
- **Contents:**
  - `.sln` with Core / AddIn / Tests projects.
  - GitHub Actions on `windows-latest`: build, test, and a packed `.xll` as an artifact.
  - `LICENSE`, `CONTRIBUTING.md` (DCO sign-off), `THIRD_PARTY_NOTICES.md` (Excel-DNA zlib, XLParser MPL-2.0).
  - A dependency license check that fails the build on GPL/AGPL.
- **Exit:**
  - CI passes.
  - The `.xll` loads in Excel and shows its ribbon tab.
  - One placeholder key fires.

### Phase 3 — Cycles and undo (features 1–3)

Delivered as three PRs:
- **3a:** the cycle engine and the 7 cycles plus Blue-Black on the exact keys, a settings JSON with defaults, ribbon buttons and a timing log.
- **3b:** `UndoManager`, with Ctrl+Z / Ctrl+Y through the thread keyboard hook.
- **3c:** the settings dialog.

Defaults are the **Macabacus factory settings** (v9.9.5 settings export, 2026-10-06; see research/06). Binary (Ctrl+Shift+Y) and Ratio (Alt+Shift+;) cycles were added from the same source.
- **Exit criteria:**
  - xUnit covers the cycle engine: wrap-around, the hybrid rule, mixed selections, color normalization and "No fill". Line coverage of `Toolkit.Core` is ≥ 90%.
  - All 7 v1 cycles fire on their Macabacus keys and meet the K3 latency target.
  - **Undo scenarios pass:**
    - (a) cycle ×3 then Ctrl+Z ×3 restores the original exactly;
    - (b) type → cycle → type → Ctrl+Z ×3 undoes them in reverse order;
    - (c) Ctrl+Y re-applies;
    - (d) inserting a row clears our stack without corrupting anything.
  - The settings dialog can edit, reorder, preview, reset, import and export. Export → import gives identical JSON.

### Phase 4 — Trace precedents (feature 4)
- **Exit criteria:**
  - On the fixture workbook, every §4.5 case behaves as specified: cross-sheet, names, tables, open and closed external workbooks, a range over 50 cells, a circular reference, a hidden sheet, INDIRECT.
  - A formula with ≤ 20 references opens in ≤ 300 ms.
  - Up/Down navigation takes ≤ 100 ms per step.
  - Last Audited Cell returns correctly through 3 levels of history.
- **Notes for 4b (from the 4a review):**
  - When the parser reports a failure (an immediately called `LAMBDA(...)(...)`, or a stack guard tripping on a huge formula), fall back to `Range.DirectPrecedents`.
  - `Book2!Rate` is ambiguous. It can mean a sheet `Book2`, or a workbook-level name in an unsaved `Book2`. If that sheet doesn't exist, try an open workbook with that name.
  - Only INDEX, OFFSET, INDIRECT and CHOOSE are flagged as returning references (per spec). XLOOKUP, IF, IFS, SWITCH and LET can also return references, as in `=SUM(A1:XLOOKUP(...))`. Candidate for v1.1.

### Phase 5 — Release
- **Contents:**
  - A per-user installer and Authenticode signing.
  - Antivirus check: VirusTotal shows 0 detections from major engines.
  - README with install, the Unblock fallback, and Macabacus coexistence notes.
  - A manual test run recorded on Microsoft 365 Current Channel and Office LTSC 2024.
- **Exit:** the repo is made public and v0.1.0 is released.


### v2 roadmap: validated during Phase 1

These are not v1 scope.
- **Office.js component** for Mac and web, and for native undo on Windows (hybrid).
  - K1c proved Office.js can bind the exact v1 keys through undocumented names: `Semicolon`, `Comma`, `Period`, `LeftBracket`, `RightBracket`, `SingleQuote`, `Backslash`. It must use runtime `replaceShortcuts` and a self-test, because one invalid key rejects the whole shortcuts file.
  - Office.js formatting is native-undoable (ExcelApi 1.20+).
  - Its trace can't enter other workbooks, so on Mac/web it would *list* external precedents only.
- **Excel on the web: colors-only edition** (owner idea, 2026-09-29). A lightweight Office.js add-in offering just the font- and fill-color cycles (maybe the number-format cycles too), for people who work in browser Excel.
  - It uses the K1c key names: `Ctrl+SingleQuote` (font cycle); `Ctrl+Shift+K` is standard.
  - Open questions:
    - whether those names also work in browsers (K1c was tested on Windows desktop only);
    - native undo on the web (Microsoft says it's coming "based on customer demand");
    - browser-reserved shortcuts.
  - It shares the cycle and palette JSON with the desktop add-in.
- **Windows hybrid for native undo:** Excel-DNA keeps the hook and cross-workbook trace, and hands formatting to the Office.js component through a local channel (research/05 §3).

## 6. Risks and mitigations

| Risk | Likelihood / impact | Mitigation |
|---|---|---|
| Keys stolen by other add-ins (CapIQ re-binds periodically; Macabacus if installed alongside) | Medium / High | "Override" command, and re-registering keys on `WorkbookActivate`. Optional keyboard hook in v2. |
| Custom undo ordering or corruption bugs | Medium / High | Pure `UndoManager` in Core with thorough unit tests, the §4.4 ordering rule, and clearing the stack on structural changes. |
| AltGr keyboard layouts have no Ctrl+[ | Low for US/UK users / Medium | Keys can be remapped. The keyboard hook (virtual-key codes) is a v2 option. |
| Non-US layouts put `;`, `'` and `[` on different physical keys (for example `;` is Shift+comma on a German layout) | Medium for non-US users / Medium | `OnKey` maps through the active layout, so these chords may be unreachable or land on other keys. Users can remap in settings.json now, and in the dialog in 3c. A v2 option is binding by virtual-key code through the hook. |
| Unloading the add-in while Macabacus is also installed | Medium / Low | `xlcOnKey` without a macro restores *Excel's* default, and the C API can't tell who owns a key. So unloading our add-in also unbinds Macabacus's copies of the shared keys until Macabacus re-registers them (restart Excel or use its Override). Document this in the README coexistence notes (Phase 5). |
| WPF keyboard focus problems inside Excel | Medium / Medium | Spike K4, with a WinForms fallback. |
| Mark-of-the-Web, SmartScreen and antivirus friction | High / Medium | Per-user installer, signing, and VirusTotal checks in the release checklist. |
| Mac and web users left out | Certain / Low for now | JSON settings are portable, so a VBA (Mac) or Office.js (web) port can come later. Recorded as a follow-up. |
| .NET runtime conflicts with other add-ins | Low (net48) / High | Target .NET Framework 4.8, not .NET 6+. |

## 7. Verification
- **Unit tests:** xUnit on `Toolkit.Core`, run in CI.
- **Integration tests:** a manual script against the fixture workbooks, plus an optional Excel-DNA test harness running inside Excel (ExcelDna.Testing).
- **Performance:** timing logs in debug builds (p50/p95 per command).
- **Records:** each test run is saved in `docs/test-runs/`.

## 8. ADR
See [`decisions/0002-excel-dna-windows-first.md`](decisions/0002-excel-dna-windows-first.md).

## Changelog
- 2026-09-28: first draft (Office.js, ADR-0001).
- 2026-09-28: build-vs-fork recommendation added (research/04).
- 2026-09-29: PR #1 (scaffold) merged. Phase 3 approved and split into 3a, 3b and 3c.
- 2026-09-29: Phase 2 (scaffold) approved and started on branch `phase2/scaffold`. The web colors-only edition is logged as a future development.
- 2026-10-06: Phase 3 merged (#2, #3, #4) after owner Excel testing. The undo restore's macro-context re-check was removed (it was unreliable). Automated Excel smoke test added. Phase 4 started.
- 2026-10-06: Phase 4a (Trace In core): formula parser on XLParser 1.7.5 (MPL-2.0) + Irony (MIT), reference extraction, formula structure, precedent tree model and audit history.
- 2026-09-28: **Phase 1 complete.** ADR-0002 accepted. D11: cross-workbook trace essential. K1c: Office.js key names work; recorded as the v2 path.
- 2026-09-28: Undo decision (owner): Macabacus parity in v1, with the Office.js hybrid for native undo as a v2 candidate. §4.4 rewritten from spike K2/K2b/K2c.
- 2026-09-28: Trace In spec from the Macabacus help PDF (research/07): the Argument column, Evaluate mode as v1.1, the focus/hook keyboard model, and the K4 variant C.
- 2026-09-28: Added the owner's Macabacus config (research/06): confirmed default cycles, colors and the full 132-command keymap.
- 2026-09-28: **Pivot.** The owner made exact Macabacus keys a hard requirement. Now Excel-DNA, Windows first (ADR-0002, research/05). Spikes changed from S1–S7 to K1–K4. Undo design changed to Macabacus-parity.
