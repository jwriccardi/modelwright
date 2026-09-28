# Excel Modeling Toolkit *(working name)*

An open-source Excel add-in for financial modelers:

1. **Number format cycling**: step the selected cells through your own list of formats with one shortcut.
2. **Font color cycling**
3. **Fill color cycling**
4. **Smart trace precedents**: a keyboard-driven tree of a cell's precedents, with values, that works across sheets.

These are the features people pay for in Macabacus, FactSet Spreadsheet Tools, TTS Turbo Macros and similar tools. This project aims to provide them for free on Excel for **Windows, Mac and the web**, with **Ctrl+Z working** after every action.

## Status: planning

No code has been written yet. Start here:

| Doc | What it covers |
|---|---|
| [docs/PLAN.md](docs/PLAN.md) | Work plan, design and acceptance criteria (**pending approval**) |
| [docs/open-questions.md](docs/open-questions.md) | Decisions still needed, and points to check against Macabacus |
| [docs/decisions/0001-platform-architecture.md](docs/decisions/0001-platform-architecture.md) | ADR: Office.js + TypeScript (proposed) |
| [docs/research/01-feature-survey.md](docs/research/01-feature-survey.md) | How Macabacus and competitors implement each feature |
| [docs/research/02-architecture-options.md](docs/research/02-architecture-options.md) | Office.js vs Excel-DNA vs VSTO vs VBA vs others |
| [docs/research/03-licensing.md](docs/research/03-licensing.md) | License options (MIT recommended) |
| [docs/research/04-xlerate-evaluation.md](docs/research/04-xlerate-evaluation.md) | Existing open-source prior art: whether to fork or build fresh |

## License

Not chosen yet (see [research/03](docs/research/03-licensing.md)). Until a license is added, all rights are reserved.
