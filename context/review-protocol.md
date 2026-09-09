# Codex review protocol

Codex independently reconstructs the candidate; it does not trust builder assertions. Review must cover arithmetic and formula tracing, double-count prevention, source/evidence classification, debt and liquidity reconstruction, returns reconstruction, scenario ordering, bounded-search logic, IC aggressiveness, and formula/package checks. The native desktop Excel no-repair/no-warning gate remains mandatory.

Every finding uses this schema:

- `finding_id`
- `severity`
- `category`
- `sheet`
- `cell_or_range`
- `issue`
- `economic_effect`
- `blocks_release`
- `requires_jack_decision`
- `required_fix`
- `required_test`
