# Model governance

## Frozen Base

- Path: `model/base-v2.4.5.xlsx`
- SHA-256: `5764fa6dc28137bf4cb2348e04400bd73a4663cb22ee14a546ae60e9082f4e15`
- The Base is immutable. Every scenario must derive from the exact approved hash.

## Architecture

The 13 approved tabs, in exact order, are:

1. IC Summary
2. Base Case Bridge
3. Evidence & Market
4. Assumptions
5. Acquisition Schedule
6. Anchor QoE
7. Revenue & Durability
8. Org & Capacity
9. Operating Case
10. Transaction & Liquidity
11. Returns
12. Controls
13. Version Bridge

No tabs may be added and no architecture may be redesigned without explicit approval. Terminal presentation tabs may sit first, but calculation tabs must not depend on them.

## Workbook and formula rules

- No merged cells and no wrapped populated cells.
- Century Gothic, 8 pt.
- Display zeros as `-` and negatives in parentheses.
- Hardcodes are blue; cross-sheet formulas are green; same-sheet formulas are black.
- No formula may exceed 1,000 characters.
- Use explicit standalone formulas only: no shared-formula records or groups.
- No external links.
- Maintain valid calculation-chain behavior and remove stale calculation-chain artifacts when required by the release process.
- Perform deterministic external recalculation and semantic formula regression.
- A release requires native desktop Microsoft Excel to open with no repair and no warning. Deterministic QA does not replace this gate.
