# ADR-0002 — Excel-DNA (C#), Windows desktop first

- **Status:** **Accepted (2026-09-28)**, after decision spikes K1–K4 and the owner's rulings: exact Macabacus keys are required (D0), Macabacus-parity undo is acceptable for v1 (D10), and cross-workbook trace is **essential** (D11). It replaces [ADR-0001](0001-platform-architecture.md).
- **Context:**
  - [research/02](../research/02-architecture-options.md)
  - [research/05](../research/05-keys-and-undo.md)
  - The owner's requirement: *exact Macabacus keyboard shortcuts are critical*.

## Decision

Build a **Windows desktop Excel add-in in C# using Excel-DNA** (zlib license), packaged as a single `.xll`. It will:

- **Keys:** bind the exact Macabacus keymap through `xlcOnKey` / `Application.OnKey`.
- **Undo:** provide formatting undo at Macabacus parity:
  - preferably by preserving Excel's native stack through `ExecuteMso` (if spike K2 shows this works);
  - otherwise with a multi-level custom stack on Ctrl+Z / Ctrl+Y, as Macabacus does.
- **Trace:** show trace precedents in a modeless, keyboard-driven window.
  - It can navigate across sheets **and into other open workbooks**, and optionally open linked workbooks.
- **Runtime:** target **.NET Framework 4.8**, so users need nothing extra installed.
  - Microsoft's own guidance for Excel-DNA distribution favors this.
  - It avoids the "one .NET Core version per Excel process" conflict with other add-ins.
  - Revisit Native AOT once Excel-DNA 1.10 is stable.

**Mac and web are deferred.** Settings are stored as JSON that describes cycles, palettes and the keymap, so a later Mac/web version can reuse them:

- Mac: VBA with Control-key shortcuts, or Office.js with letter shortcuts.
- Web: Office.js.

Only the settings files carry over to that later version; its code would be separate.

## Drivers

1. **Exact Macabacus keys.** A hard requirement.
2. **Undo that behaves at least as well as Macabacus.**
3. **Precedent navigation that matches Macabacus,** including into other workbooks.

## Alternatives considered

| Alternative | Why not |
|---|---|
| **Office.js (ADR-0001)** | Can't bind `[ ' ; , .` (the schema allows only `[A-Za-z0-9-_+]`), and can't navigate into other workbooks. It would be reinstated only if spike K1 shows that named punctuation keys work. |
| **VBA .xlam (Windows + Mac, one codebase)** | Binds the keys on Windows, and on Mac with Control (punctuation there unverified). Against it: Mark-of-the-Web blocks downloaded add-ins on Windows, the code is binary with export tooling, UserForms are weak for a tree view, there's no unit-test story, and it clears undo just the same. It's a reasonable *Mac* follow-up; Breakdown (MIT) shows it can be done. |
| **Excel-DNA + Office.js bridge** | Might give native undo together with exact keys, but it means two runtimes and a localhost channel. Kept as a later experiment (research/05 §3). |
| **VSTO** | This is what Macabacus uses, but it's stuck on .NET Framework, has awkward shortcut support and no advantage over Excel-DNA. |
| **C++ XLL** | Too costly to build, and few people could contribute. |

## Why chosen

- It is the only option that satisfies all three drivers.
- C#, NuGet, CI builds and unit tests with xUnit are a reasonable environment for open-source contributors.
- Excel-DNA is mature, with the same author active since 2006, and zlib-licensed.

## Consequences

- **Windows desktop only for v1.** This is the owner's third-choice platform scope, accepted because exact keys require it.
- **Undo.** Excel's native undo history is lost after our first formatting action, unless K2 finds a way around it. This is the same as Macabacus today.
- **Coexisting with Macabacus.** If Macabacus is installed too, whichever add-in loaded last owns each key. We document this and ship an "Override" command, as Macabacus does.
  - For development on the owner's machine, Macabacus must be disabled or given different keys.
- **Distribution friction.**
  - A downloaded `.xll` is blocked by Mark-of-the-Web. We need either an installer (per-user, no admin: copy to `%AppData%` and add the `HKCU …\Excel\Options\OPEN` key) or documented Unblock steps.
  - We need Authenticode signing: SignPath Foundation (free for OSS, [UNVERIFIED]) or Azure Trusted Signing.
  - We should test for antivirus false positives.
- **Formula parsing.** Use XLParser (MPL-2.0; fine to consume from an MIT project) or ClosedXML.Parser (MIT). XLerate's TypeScript code becomes a design reference only, not code we can reuse.

## Follow-ups (decision spikes, throwaway code, which the owner must approve)

- **K1 (done).** Office.js *can* bind the exact v1 punctuation keys through undocumented names (`Semicolon`, `Comma`, `Period`, `LeftBracket`, `RightBracket`, `SingleQuote`, `Backslash`) with runtime `replaceShortcuts`. ADR-0001 was reconsidered and **stays superseded**, because Office.js can't navigate into other workbooks, which the owner ruled essential (D11). Office.js is the proven path for v2 (Mac/web, native undo).
- **K2 (done).** Built-in commands dispatched outside macro context keep native undo, but only for fixed formats. Any COM write wipes the history. v1 therefore matches Macabacus with a custom undo stack. A hybrid with Office.js for full native undo is recorded as a v2 candidate (PLAN §4.4).
- **K3 (passed).** Every Macabacus key tested binds through `xlcOnKey`. p95 is 9.8 ms on a 4,770-cell selection.
