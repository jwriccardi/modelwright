# 03 — License choice

*This is engineering analysis, not legal advice.*

## What matters for this project

1. **Adoption by finance professionals.** Their firms' IT and legal teams approve add-ins, and they routinely approve permissive licenses. Copyleft (GPL/AGPL) is often blocked by policy.
2. **Compatibility with the prior art we may borrow from.**
   - Open-source add-ins and samples: XLerate (MIT), Breakdown (MIT), Microsoft's Office-Add-in-samples (MIT).
   - Formula-parsing libraries: excel-formula-tokenizer (MIT), Formualizer (MIT/Apache-2.0).
3. **How it's distributed.** An Office.js add-in is JavaScript served to every user's Excel, so it is "distributed" in license terms. AGPL's network clause would also apply to the hosted site.
4. **Contributions.** We want a low-friction way to accept outside contributions.

## Options

| License | Pros | Cons | Fit |
|---|---|---|---|
| **MIT** | Shortest and best-understood; the same as the closest prior art; nothing for adopters to review | No explicit patent grant; allows closed-source commercial forks | **Recommended** |
| Apache-2.0 | Explicit patent grant and patent-retaliation clause; often preferred by corporate legal | Longer, needs a NOTICE file and statements of changes; not compatible with GPLv2 | Good alternative |
| MPL-2.0 | File-level copyleft: changes to *our* files must be shared, but it can be combined with closed code | Less familiar and needs more explanation; a few more steps for adopters | Only if you want forks' improvements to stay open |
| GPL-3.0 / AGPL-3.0 | Keeps all derivatives open | Deters adoption in banks and funds; clashes with the permissive ecosystem; AGPL reaches the hosted site | Not recommended |

## Recommendation: **MIT**

- **Patent risk.** A formatting and tracing add-in has little patent exposure, so Apache-2.0's patent grant adds little for the extra paperwork.
- **Borrowing code.** MIT matches XLerate and Breakdown, so any code we adopt keeps a single, simple license. We would keep their copyright lines in `LICENSE` or `THIRD_PARTY_NOTICES.md`.
- **Contributions.** Use a **DCO sign-off** (`git commit -s`) rather than a CLA. This keeps contributing easy while recording where each contribution came from.

## Things to settle before publishing

- **Copyright holder: decided 2026-09-28, Pegasus Technology Group LLC.** `LICENSE` will read "Copyright (c) 2026 Pegasus Technology Group LLC".
- **Trademarks and clean room.**
  - Don't use "Macabacus" (or FactSet, Arixcel and the like) in the product name, icon or marketing beyond factual comparisons such as "Macabacus-compatible shortcuts".
  - Don't copy their help text, icons or screenshots.
  - Reimplementing the same *functionality* and similar default keys is normal practice in this category: FactSet and Macabacus already share almost the same keymap.
- **Third-party notices.** Keep `THIRD_PARTY_NOTICES.md` up to date. Under ADR-0002 that means NuGet dependencies: Excel-DNA (zlib) and XLParser (MPL-2.0, which is fine to consume unmodified from an MIT project). CI should fail the build on GPL or AGPL dependencies.
- **Name.** `excel-modeling-toolkit` is a working name. Before going public, check for trademark or name collisions on npm, GitHub and Microsoft Marketplace.

## Status

**Proposed. Waiting for the owner's decision.** No `LICENSE` file has been committed yet, so the private repo is currently "all rights reserved" by default.
