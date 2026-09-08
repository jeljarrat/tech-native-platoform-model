# PMC model automation control plane

This repository automates the reviewed model-change loop. Jack supplies one change request; Claude builds or patches the candidate, deterministic QA gates it, Codex independently reviews it, and blocking feedback returns to Claude automatically. The loop stops at a reviewed release candidate or an authorized escalation. It never merges a pull request.

The approved, immutable Base is `model/base-v2.4.5.xlsx` with SHA-256 `5764fa6dc28137bf4cb2348e04400bd73a4663cb22ee14a546ae60e9082f4e15`.

## One-time setup — subscription mode (default)

Use Python 3.12, install dependencies, install Claude Code, and authenticate both CLIs through their paid-product browser login. Do not set API keys for subscription runs.

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
npm install -g @anthropic-ai/claude-code
claude
codex login
gh auth login
```

When Claude opens, choose the Claude App/Pro/Max subscription login, not Anthropic Console API billing. `codex login` opens the ChatGPT browser login; do not use `--with-api-key`.

Verify each login:

```powershell
claude --version
claude auth status --json
codex --version
codex login status
gh auth status
```

If Claude's installed release does not support `claude auth status`, start `claude` and run `/status`. The orchestrator removes `ANTHROPIC_API_KEY`, `OPENAI_API_KEY`, and `CODEX_API_KEY` from subscription subprocesses and warns when those variables are present.

API mode remains available. For that mode only, set `ANTHROPIC_API_KEY` and `OPENAI_API_KEY`; optional model overrides are `ANTHROPIC_MODEL` and `OPENAI_MODEL`.

## Create a change request

Copy `change-requests/template.yaml`, give it a unique `change_request_id`, and state the authorized changes, prohibited changes, evidence/decision IDs, assumptions, and measurable acceptance criteria. Keep the Base version and hash unchanged. A change request must not silently introduce underwriting economics.

## Run one modeling task

```powershell
python scripts/orchestrate_review.py change-requests/CR-YYYY-NNN.yaml --auth subscription
```

Useful options:

- `--dry-run` validates setup and Base identity without API calls.
- `--mock [transcript.json]` rehearses without live API calls.
- `--max-cycles N` changes the cycle cap (default: 2).
- `--no-push` keeps the task branch and commits local and does not open a PR.
- `--auth api` uses the original API adapters and requires both API keys.
- `--allow-api-fallback` permits API fallback after a subscription usage limit; without it, fallback is forbidden.
- `--resume <run-id>` resumes a paused run from its saved builder or supervisor stage.

The command verifies the frozen Base, creates or reuses `change/<change-request-id>`, sends Claude the builder contract/request/specs, applies Claude's explicit repository-local operations, runs deterministic QA, sends QA defects back to Claude, asks Codex for schema-validated JSON findings after QA passes, and sends blockers back to Claude. A clean candidate is committed, pushed, and opened or updated as a PR when GitHub credentials permit.

Every request/response, QA report, finding set, state transition, hash, and final result is retained under `reports/runs/<run-id>/`. `final.json` is the machine-readable handoff; `escalation.json` exists only when Jack must decide.

If a subscription usage limit is reached, the run stops as `SUBSCRIPTION_LIMIT_PAUSED`, saves `checkpoint.json`, and prints the resume command. After the limit resets:

```powershell
python scripts/orchestrate_review.py --resume <run-id> --auth subscription
```

It never silently switches to API billing.

Jack is interrupted only for `PRINCIPAL_DECISION`, `NEW_UNDERWRITING_ASSUMPTION`, `EVIDENCE_CONFLICT`, or `UNRESOLVED_AGENT_DISAGREEMENT` after the configured cycles. Routine build, QA, and supervisor blockers remain inside the automated Claude–Codex loop.

## Desktop Excel gate

Automation stops at `REVIEWED_RELEASE_CANDIDATE`; it does not merge. A changed workbook still requires the desktop Excel full-recalculation/visual verification gate required by `spec/release-rules.yaml`. The recalculated desktop copy is evidence only and must never replace the canonical zero-shared-formula package.

## Development checks

```powershell
python -m compileall scripts tests
python -m unittest discover -s tests -p "test_*.py" -v
python scripts/run_task_c_qa.py
```
