# Shared agent contract

This repository is the authoritative automation context. Both Claude and Codex must use the exact, hash-pinned contents of `CLAUDE.md`, `AGENTS.md`, and:

- `context/project-overview.md`
- `context/model-governance.md`
- `context/evidence-and-decisions.md`
- `context/current-state.md`
- `context/review-protocol.md`

Roles: Claude is the builder/patcher; deterministic Python performs mechanical QA; Codex is the independent supervisor; Jack is the principal decision-maker only where required.

Canonical loop:

`Jack instruction -> change request -> Claude build -> deterministic QA -> Codex review -> Claude patch -> repeat -> reviewed release candidate or authorized escalation`

Jack escalation is authorized only for `PRINCIPAL_DECISION`, `NEW_UNDERWRITING_ASSUMPTION`, `EVIDENCE_CONFLICT`, or `UNRESOLVED_AGENT_DISAGREEMENT` after two cycles. Infrastructure failures are not Jack business escalations.

Canonical context is change-controlled. A modeling run pins all context hashes at start, supplies the exact content and hashes to each agent invocation, preserves them in run reports, and fails if a required file is missing or changes. Context changes require a dedicated change request and commit; no modeling run may silently mutate context.
