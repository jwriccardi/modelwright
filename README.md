# Excel Modeling Toolkit *(working name)*

An open-source Excel add-in for financial modelers:

1. **Number format cycling**: step the selected cells through your own list of formats with one shortcut.
2. **Font color cycling**
3. **Fill color cycling**
4. **Smart trace precedents**: a keyboard-driven tree of a cell's precedents, with values, that works across sheets.

These are the features people pay for in Macabacus, FactSet Spreadsheet Tools, TTS Turbo Macros and similar tools. This project aims to provide them for free, using **exactly the same keyboard shortcuts as Macabacus**. The first target is **Excel for Windows desktop**. Mac and web are deferred, because Office.js can't bind Macabacus's punctuation keys.

## Status: Phase 2 (project skeleton)

The architecture is decided: an **Excel-DNA (C#) add-in for Windows desktop Excel** ([ADR-0002](docs/decisions/0002-excel-dna-windows-first.md)), validated by decision spikes ([results](docs/spike-results.md)). The skeleton builds a loadable add-in with a ribbon tab and an About command. The formatting cycles and Trace In come next ([PLAN](docs/PLAN.md)).

### Build and test

Requires the .NET SDK 10 on Windows. No Visual Studio needed.

```powershell
dotnet build ExcelModelingToolkit.sln -c Release
dotnet test ExcelModelingToolkit.sln -c Release
powershell -ExecutionPolicy Bypass -File build/check-licenses.ps1
```

**Load in Excel:** go to File › Options › Add-ins, set Manage: Excel Add-ins, click Go… › Browse…, and pick `src/ExcelModelingToolkit.AddIn/bin/Release/net48/publish/ModelingToolkit64-packed.xll`. A **Modeling Toolkit** ribbon tab appears, and **Ctrl+Alt+Shift+F12** shows the About box.

See [CONTRIBUTING.md](CONTRIBUTING.md). Contributions need a DCO sign-off.

### Docs

| Doc | What it covers |
|---|---|
| [docs/PLAN.md](docs/PLAN.md) | Work plan, design and acceptance criteria |
| [docs/spike-results.md](docs/spike-results.md) | Decision spike results (keys, undo, trace window, Office.js) |
| [docs/open-questions.md](docs/open-questions.md) | Decisions still needed, and points to check against Macabacus |
| [docs/decisions/0002-excel-dna-windows-first.md](docs/decisions/0002-excel-dna-windows-first.md) | ADR: Excel-DNA (C#), Windows first (proposed) |
| [docs/decisions/0001-platform-architecture.md](docs/decisions/0001-platform-architecture.md) | ADR: Office.js (superseded) |
| [docs/research/01-feature-survey.md](docs/research/01-feature-survey.md) | How Macabacus and competitors implement each feature |
| [docs/research/02-architecture-options.md](docs/research/02-architecture-options.md) | Office.js vs Excel-DNA vs VSTO vs VBA vs others |
| [docs/research/03-licensing.md](docs/research/03-licensing.md) | License options (MIT recommended) |
| [docs/research/04-xlerate-evaluation.md](docs/research/04-xlerate-evaluation.md) | Existing open-source prior art: whether to fork or build fresh |
| [docs/research/05-keys-and-undo.md](docs/research/05-keys-and-undo.md) | Which architectures can bind Macabacus's keys, and the options for undo |
| [docs/research/06-macabacus-observed-config.md](docs/research/06-macabacus-observed-config.md) | The owner's Macabacus settings: cycles, colors and the full keymap |
| [docs/research/07-macabacus-trace-in-spec.md](docs/research/07-macabacus-trace-in-spec.md) | How Macabacus's Trace In behaves, and what that means for our design |

## License

[MIT](LICENSE), Copyright (c) 2026 Pegasus Technology Group LLC. Third-party components are listed in [THIRD_PARTY_NOTICES.md](THIRD_PARTY_NOTICES.md).
