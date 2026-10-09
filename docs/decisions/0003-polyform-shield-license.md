# ADR 0003: License — PolyForm Shield 1.0.0

**Status:** accepted (owner, 2026-10-08). Supersedes the MIT choice recorded in D2 for the scaffold.

## Context

The owner's requirements for the public license (2026-10-08):

1. People working in a private corporate environment can use the product without violating the license.
2. They can modify or extend it without an obligation to contribute changes back.
3. Nobody may download the source and sell the product, or a derivative product, for personal gain.

Requirement 3 rules out every OSI-approved open-source license (MIT, Apache-2.0, GPL, AGPL and the rest): all of them permit selling. Copyleft licenses only require sharing source on distribution, which does not stop a sale.

## Options considered

| License | Corporate use | Private changes, no give-back | Blocks selling | Notes |
|---|---|---|---|---|
| MIT / Apache-2.0 | yes | yes | **no** | Anyone may sell it. |
| AGPL-3.0 | yes | yes (until distributed or offered as a service) | **no** | Selling is allowed with source. |
| PolyForm Noncommercial 1.0.0 | **no** (business use is commercial) | yes | yes | Fails requirement 1. |
| PolyForm Internal Use 1.0.0 | yes | yes | yes | Forbids any redistribution, even free forks on GitHub; too strict for a public project. |
| Apache-2.0 + Commons Clause | yes | yes | yes ("Sell") | Matches the wording, but the Commons Clause is a bolt-on with a contested definition of "sell". |
| **PolyForm Shield 1.0.0** | yes | yes | yes (no competing product) | Any use is permitted except providing a product that competes with the software, free or paid. Free forks and contributions are fine; selling it or a derivative is not. |

## Decision

PolyForm Shield License 1.0.0, with Pegasus Technology Group LLC as the licensor and the required notice line in `LICENSE`. Contributions are accepted under the same license with DCO sign-off.

## Consequences

- The project is **source-available**, not open source in the OSI sense; README and docs say so.
- Shield is stricter than requirement 3 in one way: a *free* competing product is also not allowed. The owner accepted that reading.
- Dependencies stay compatible: XLParser (MPL-2.0, unmodified, source linked), Irony (MIT), Excel-DNA (MIT). MPL-2.0 is file-level copyleft and allows combination with differently licensed code.
- Changing to a permissive license later is easy while Pegasus owns all contributions; the reverse would not be.
