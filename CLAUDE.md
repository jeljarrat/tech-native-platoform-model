# Claude Code operating contract

Treat this repository as the source of truth. Before any modeling work, read `AGENTS.md` and every file in the canonical context set listed there, plus the current change request, relevant specifications, and current-run findings supplied by the orchestrator.

- Never modify `model/base-v2.4.5.xlsx`. Derive scenarios only from its approved SHA-256 recorded in `context/model-governance.md`.
- Implement only an approved change request. Preserve workbook architecture unless that request explicitly authorizes a change.
- Distinguish facts, comparables, external evidence, assumptions, open questions, and Jack principal decisions. Never silently invent or promote an underwriting assumption.
- Follow every Excel release rule in `context/model-governance.md` and every review requirement in `context/review-protocol.md`.
- Send blockers through the orchestrator. Do not ask Jack directly unless the issue fits an authorized escalation category in `AGENTS.md`.
- Do not modify canonical context during a modeling run. Context changes require their own change request and commit.
