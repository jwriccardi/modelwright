# Excel Modeling Toolkit *(working name)*

A source-available Excel add-in for financial modelers:

1. **Number format cycling**: step the selected cells through your own list of formats with one shortcut.
2. **Font color cycling**
3. **Fill color cycling**
4. **Smart trace precedents**: a keyboard-driven tree of a cell's precedents, with values, that works across sheets.

These are the features people pay for in Macabacus, FactSet Spreadsheet Tools, TTS Turbo Macros and similar tools. This project aims to provide them for free, using **exactly the same keyboard shortcuts as Macabacus**. The first target is **Excel for Windows desktop**. Mac and web are deferred, because Office.js can't bind Macabacus's punctuation keys.

## Status: formatting cycles shipped; Trace In in progress

The architecture is decided: an **Excel-DNA (C#) add-in for Windows desktop Excel** ([ADR-0002](docs/decisions/0002-excel-dna-windows-first.md)), validated by decision spikes ([results](docs/spike-results.md)). Working today, on Macabacus's exact keys:
- **Number-format cycles:** Ctrl+Shift+1 / 2 / 4 / 5 / 8, Ctrl+Shift+Y and Alt+Shift+;.
- **Color cycles:** Ctrl+' (font), Ctrl+Shift+K (fill) and Ctrl+; (blue/black toggle).
- **Undo:** Ctrl+Z / Ctrl+Y work alongside Excel's own undo.
- **Settings dialog** (Modeling Toolkit › Settings…).

Trace In (Ctrl+Shift+[) is next ([PLAN](docs/PLAN.md)).

**Excel smoke test:** with Excel open, run `powershell -ExecutionPolicy Bypass -File tests/excel-smoke/undo-smoke.ps1`. It drives Excel with real keystrokes in a scratch workbook.

### Build and test

Requires the .NET SDK 10 on Windows. No Visual Studio needed.

```powershell
dotnet build ExcelModelingToolkit.sln -c Release
dotnet test ExcelModelingToolkit.sln -c Release
powershell -ExecutionPolicy Bypass -File build/check-licenses.ps1
```

**Load in Excel:** go to File › Options › Add-ins, set Manage: Excel Add-ins, click Go… › Browse…, and pick `src/ExcelModelingToolkit.AddIn/bin/Release/net48/publish/ModelingToolkit64-packed.xll`. A **Modeling Toolkit** ribbon tab appears. Turn Macabacus off first, because it uses the same keys.

See [CONTRIBUTING.md](CONTRIBUTING.md). Contributions need a DCO sign-off.

### Docs

| Doc | What it covers |
|---|---|
| [docs/PLAN.md](docs/PLAN.md) | Work plan, design and acceptance criteria |
| [docs/spike-results.md](docs/spike-results.md) | Decision spike results (keys, undo, trace window, Office.js) |
| [docs/open-questions.md](docs/open-questions.md) | Decisions still needed, and points to check against Macabacus |
| [docs/decisions/0002-excel-dna-windows-first.md](docs/decisions/0002-excel-dna-windows-first.md) | ADR: Excel-DNA (C#), Windows first (accepted) |
| [docs/decisions/0001-platform-architecture.md](docs/decisions/0001-platform-architecture.md) | ADR: Office.js (superseded) |
| [docs/research/01-feature-survey.md](docs/research/01-feature-survey.md) | How Macabacus and competitors implement each feature |
| [docs/research/02-architecture-options.md](docs/research/02-architecture-options.md) | Office.js vs Excel-DNA vs VSTO vs VBA vs others |
| [docs/research/03-licensing.md](docs/research/03-licensing.md) | License options (MIT recommended) |
| [docs/research/04-xlerate-evaluation.md](docs/research/04-xlerate-evaluation.md) | Existing open-source prior art: whether to fork or build fresh |
| [docs/research/05-keys-and-undo.md](docs/research/05-keys-and-undo.md) | Which architectures can bind Macabacus's keys, and the options for undo |
| [docs/research/06-macabacus-observed-config.md](docs/research/06-macabacus-observed-config.md) | The owner's Macabacus settings: cycles, colors and the full keymap |
| [docs/research/07-macabacus-trace-in-spec.md](docs/research/07-macabacus-trace-in-spec.md) | How Macabacus's Trace In behaves, and what that means for our design |

## License

[PolyForm Shield 1.0.0](LICENSE), Copyright (c) 2026 Pegasus Technology Group LLC. In short: you may use the add-in for any purpose, including inside your company, and change it privately with no obligation to share your changes; you may not offer a product that competes with it (selling it or a derivative, even free). This is a "source-available" license, not an [OSI open-source](https://opensource.org/osd) one. See [ADR 0003](docs/decisions/0003-polyform-shield-license.md). Third-party components are listed in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
