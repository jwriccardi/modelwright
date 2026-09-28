# Work plan — Excel Modeling Toolkit

> **Status: PENDING APPROVAL.** Planning only; no code has been written yet.
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
| Last Audited Cell | Ctrl+Shift+\ [to confirm] | `^+\` / `^|` | ✅ |
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
- **Default contents.** Macabacus-style codes (research/01 §1), for example `_(#,##0_)_%;(#,##0)_%;_("–"_)_%;_(@_)_%`. Final values will be taken from the installed Macabacus (open-questions A1–A2).
- **Too many number formats.** Excel's "too many number formats" error is caught and shown as a clear message.

### 4.4 Undo (Macabacus parity or better)
- **Preferred approach, if spike K2 passes:** apply formats through `CommandBars.ExecuteMso` built-in commands, so Excel's native undo stack survives.
- **Default approach: our own `UndoManager`.**
  - **Snapshot first.** Before each change, record the number format, font color and fill of the affected cells.
    - Capture is capped at a configurable number of cells, default 10,000. Macabacus has the same kind of setting.
    - Storage is run-length compressed for uniform ranges.
  - **Ctrl+Z / Ctrl+Y are rebound** to the UndoManager.
  - **Ordering rule.** Every action of ours wipes Excel's native undo stack. So if Excel's own undo is available, anything on it must be newer than our last action.
    - If native undo is available (`GetEnabledMso("Undo")`), pass Ctrl+Z through to `Application.Undo`.
    - Otherwise, pop our own stack.
    - This keeps the user's typing and our formatting in the correct order.
  - **Invalidation.** Clear our stack when the workbook is closed. Also clear it when a structural change is detected (rows or columns inserted or deleted, found via the `SheetChange` target shape), because the stored addresses would be wrong.
  - **Ribbon.** A Quick Access Toolbar undo button is optional.

### 4.5 Smart trace precedents (feature 4, v1)
- **Invoke.** Ctrl+Shift+[ opens a **modeless window owned by Excel** (WPF, or WinForms if K4 finds problems with keyboard focus).
- **Layout.**
  - Header: the formula, with references color-coded.
  - Tree with columns **Precedent | Address | Value**.
- **How precedents are found.** Parse `Range.Formula` (invariant A1) with XLParser, then resolve each reference:
  - **A1 references** → `Range`
  - **Names** → `Names(...).RefersToRange`
  - **Table references** → `ListObjects`
  - **External `[Book]Sheet!A1`:**
    - If the workbook is open, resolve and **navigate into it**.
    - If it's closed, list it. An "Open linked workbooks" option opens it (Macabacus parity).
  - **Check:** compare with `Range.DirectPrecedents` for references on the same sheet.
- **Keys.**
  - **Up/Down:** `Application.Goto` the node's range (across sheets and workbooks); focus returns to the window.
  - **Right:** expand the node; its children are loaded only then.
  - **Left:** collapse, or go to the parent.
  - **Enter / OK:** stay on the current cell.
  - **Esc / Cancel:** return to the audited cell. Confirm this matches Macabacus (open-questions A7).
  - **F2:** edit the formula (stretch).
- **History.** Last Audited Cell keeps a stack of up to 20 audits.
- **Edge cases.**
  - Ranges over 50 cells are shown as one node; expanding it shows pages of 100.
  - Circular references are marked ↻.
  - Hidden sheets, rows and columns get a badge. "Unhide rows & columns" is a stretch.
  - INDIRECT/OFFSET are labelled "dynamic"; v2 will evaluate their arguments.

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
| K1 | Office.js named punctuation keys (`Ctrl+Shift+BracketLeft`, `Ctrl+Quote`, `Ctrl+Semicolon`) | The keys register and fire on Windows | Stay with ADR-0002. **If they work → reopen ADR-0001.** |
| K2 | `ExecuteMso` formatting keeps native undo | Type a value → cycle via ExecuteMso → Ctrl+Z twice undoes both | Use our own UndoManager (§4.4) |
| K3 | Excel-DNA binds every key in §4.2 | 100% fire. Over 30 presses on a selection of ≤ 1,000 cells, p95 key → format ≤ 50 ms | Try a thread keyboard hook |
| K4 | Modeless trace window keeps keyboard focus through `Application.Goto` on another sheet and in another workbook | Up/Down stay in the window in both cases | WinForms instead of WPF, or re-activate the window after each Goto |

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

### Phase 5 — Release
- **Contents:**
  - A per-user installer and Authenticode signing.
  - Antivirus check: VirusTotal shows 0 detections from major engines.
  - README with install, the Unblock fallback, and Macabacus coexistence notes.
  - A manual test run recorded on Microsoft 365 Current Channel and Office LTSC 2024.
- **Exit:** the repo is made public and v0.1.0 is released.

## 6. Risks and mitigations

| Risk | Likelihood / impact | Mitigation |
|---|---|---|
| Keys stolen by other add-ins (CapIQ re-binds periodically; Macabacus if installed alongside) | Medium / High | "Override" command, and re-registering keys on `WorkbookActivate`. Optional keyboard hook in v2. |
| Custom undo ordering or corruption bugs | Medium / High | Pure `UndoManager` in Core with thorough unit tests, the §4.4 ordering rule, and clearing the stack on structural changes. |
| AltGr keyboard layouts have no Ctrl+[ | Low for US/UK users / Medium | Keys can be remapped. The keyboard hook (virtual-key codes) is a v2 option. |
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
- 2026-09-28: **Pivot.** The owner made exact Macabacus keys a hard requirement. Now Excel-DNA, Windows first (ADR-0002, research/05). Spikes changed from S1–S7 to K1–K4. Undo design changed to Macabacus-parity.
