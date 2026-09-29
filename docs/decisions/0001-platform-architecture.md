# ADR-0001 — Platform architecture: Office.js + TypeScript

- **Status:** **Superseded by [ADR-0002](0002-excel-dna-windows-first.md)** (2026-09-28). Spike K1c later proved that Office.js *can* bind the exact v1 keys through undocumented names. It remains superseded because Office.js can't trace into other workbooks, which the owner ruled essential. It is the planned basis for the v2 Mac/web and native-undo work.
- **Context:** [research/02-architecture-options.md](../research/02-architecture-options.md)

## Decision

Build the add-in as an **Office Add-in using the Excel JavaScript API (Office.js) in TypeScript**:
- a shared runtime that loads when the document opens;
- keyboard shortcuts declared in the manifest;
- the XML add-in-only manifest, migrating to the unified manifest later;
- hosted on GitHub Pages.

## Drivers

1. **Runs on Windows, Mac and the web.** The owner's first preference.
2. **Undo keeps working.** Since ExcelApi 1.20 (2025), undo works natively for Office.js changes. COM, VBA and XLL code always clear it.
3. **Shortcuts are fast and reliable.** To be confirmed in the spike.

## Alternatives considered

- **Excel-DNA (C#/.xll).**
  - For: any key binding, full COM access, cross-workbook navigation.
  - Against: Windows only, clears undo (a custom undo engine would be needed), and install friction (Mark of the Web, signing).
  - Kept as the fallback.
- **VSTO.** A dead-end platform (.NET Framework 4.8 only), Windows only, clears undo.
- **Native XLL.** Development cost, Windows only, clears undo.
- **VBA .xlam.** Clears undo, blocked by Mark of the Web, binary source, weak UI.
- **Python (PyXLL / xlwings).** Licensing (PyXLL) or COM/undo limits (xlwings).

## Why chosen

It is the only option that meets drivers 1 and 2 together. It is also Microsoft's strategic add-in platform, and the most open to contributors, since it uses TypeScript, npm and CI.

## Consequences

- **Keys.** Macabacus's Ctrl+[, Ctrl+' and Ctrl+; can't be bound. We ship a remappable keymap built from letters and numbers. Overriding a built-in shortcut shows a one-time conflict prompt.
- **External workbooks.** Precedents in other workbooks can be listed but not navigated to.
- **Undo needs a current build.** Microsoft 365 2509 or later (Windows) or 16.100 or later (Mac). LTSC and perpetual Office lose undo, and we show a notice.
- **Round trips.** Every Excel interaction is asynchronous and batched. Latency has to be engineered deliberately: in-memory state, and at most 2 syncs per key press.
- **Hosting.** We need HTTPS hosting and a manifest-based install. Marketplace listing is the long-term way to distribute.

## Follow-ups

- Phase 1 spike S1–S7 (PLAN §5). If S1 (latency) or S5 (pane focus) fails on Windows, revisit this ADR and evaluate an Excel-DNA Windows companion declared through the unified manifest's `alternates`.
- Re-check Office.js keyboard-shortcut key-set limits each year. Microsoft may add punctuation keys.
