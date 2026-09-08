# PMC underwriting model automation control plane

This repository governs how the PMC underwriting workbook is changed, tested, reviewed, and released. `main` is the source of truth. The workbook itself is deliberately not included yet.

## How the workflow works

1. Jack creates and approves a uniquely identified change request from `change-requests/template.yaml`.
2. A working branch is created for that request. Claude builds or patches the candidate workbook only within the approved scope.
3. Python extracts a machine-readable snapshot and runs deterministic package, formula, economic, scenario, presentation, and version checks.
4. Codex independently reviews the request, candidate, evidence, QA report, and architectural diff.
5. Release blockers return to Claude as structured findings. Claude patches the candidate and the cycle repeats.
6. A candidate can merge only when mandatory tests pass, no release blockers remain, and no unauthorized changes exist.
7. After two unsuccessful cycles, unresolved agent disagreement is escalated to Jack.

## Responsibilities

- **Jack** makes principal business decisions, approves new underwriting assumptions, resolves conflicting evidence, and resolves agent disagreement after two cycles. Jack is not interrupted for routine build or QA failures.
- **Claude** is the workbook builder. It implements an approved change request and fixes structured review findings without inventing assumptions.
- **Codex** is the independent supervisor. It reviews authorization, evidence, architecture, economic behavior, and release readiness; it does not build the workbook.
- **Python** performs repeatable checks and produces machine-readable reports. Code is the deterministic gate, not an economic decision-maker.

## Repository map

- `model/`: frozen approved Base workbook, eventually `base-v2.4.4.xlsx`
- `scenarios/`: scenario workbooks derived from the approved Base
- `spec/`: model contract, evidence, decisions, and release rules
- `tests/`: deterministic QA suites
- `scripts/`: extraction, comparison, regression, and orchestration tools
- `prompts/`: builder and supervisor role contracts
- `change-requests/`: one YAML request per proposed change
- `reports/`: generated snapshots, test reports, and review findings

## Local setup

```bash
python -m venv .venv
python -m pip install -r requirements.txt
python -m compileall scripts tests
```

Run QA after a real candidate manifest exists:

```bash
python scripts/extract_workbook.py path/to/candidate.xlsx --output reports/output-snapshot.json
python scripts/run_regression.py --manifest path/to/manifest.yaml --output reports/test-report.json
python scripts/compare_workbooks.py model/base-v2.4.4.xlsx path/to/candidate.xlsx --output reports/workbook-diff.json
```

Inspect orchestration configuration without invoking agents:

```bash
python scripts/orchestrate_review.py change-requests/CR-2026-001.yaml --manifest scenarios/CR-2026-001/manifest.yaml --dry-run
```

## Add a model change request

Copy `change-requests/template.yaml` to a unique name such as `CR-2026-001.yaml`. Assign the same ID inside the file, reference the approved Base version and SHA-256, describe only authorized changes, list evidence and decision IDs, and define measurable acceptance criteria. Create a branch such as `change/CR-2026-001` before any workbook edit.

Every candidate package must contain a workbook, manifest, machine-readable output snapshot, test report, and Codex review findings. A scenario manifest must identify the approved Base SHA-256 and must not silently alter Base architecture.

## Release approval

GitHub Actions runs on pull requests and manual dispatch. It installs dependencies, validates the infrastructure, runs deterministic QA when given a candidate manifest, and uploads reports. A release is approved only when all mandatory tests pass, findings contain no release blockers, and the supervisor identifies no unauthorized changes.

Agent API calls are not enabled yet. Future adapters will read `ANTHROPIC_API_KEY` and `OPENAI_API_KEY`; the workflow includes explicit TODO placeholders and does not require either secret today.

## Ingest the approved repaired Base v2.4.4

After approval, place the frozen workbook at `model/base-v2.4.4.xlsx`, calculate its SHA-256, replace `TODO_AFTER_BASE_APPROVAL` in `spec/model-spec.yaml`, record the exact tab order and formula color, create the first approved golden-output snapshot, and commit those artifacts through a dedicated Base-ingestion pull request. Do not place scenarios in `model/` or modify the frozen Base in place.
