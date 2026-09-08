#!/usr/bin/env python3
"""Coordinate builder, deterministic QA, and independent supervisor cycles."""

from __future__ import annotations

import argparse
import json
import subprocess
import sys
from pathlib import Path

import jsonschema
import yaml

ROOT = Path(__file__).resolve().parents[1]
ESCALATIONS = {"PRINCIPAL_DECISION", "NEW_UNDERWRITING_ASSUMPTION", "EVIDENCE_CONFLICT", "UNRESOLVED_AGENT_DISAGREEMENT"}


def run_adapter(command: str | None, context: Path) -> None:
    if not command:
        raise RuntimeError("Agent adapter is not configured; see README and TODO API placeholders")
    subprocess.run([command, str(context)], check=True)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("change_request", type=Path)
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--builder-command")
    parser.add_argument("--supervisor-command")
    parser.add_argument("--max-cycles", type=int, default=2)
    parser.add_argument("--dry-run", action="store_true")
    parser.add_argument("--allow-template", action="store_true", help=argparse.SUPPRESS)
    args = parser.parse_args()
    request = yaml.safe_load(args.change_request.read_text(encoding="utf-8"))
    request_id = request.get("change_request_id")
    if not request_id or (request_id == "CR-YYYY-NNN" and not args.allow_template):
        raise ValueError("A unique change_request_id is required")
    if args.dry_run:
        print(json.dumps({"change_request_id": request_id, "status": "ready", "cycles": args.max_cycles, "agents_configured": bool(args.builder_command and args.supervisor_command)}, indent=2))
        return 0
    review_path = ROOT / "reports" / f"{request_id}-review.json"
    test_path = ROOT / "reports" / f"{request_id}-tests.json"
    schema = json.loads((ROOT / "reports" / "review-schema.json").read_text(encoding="utf-8"))
    for cycle in range(1, args.max_cycles + 1):
        run_adapter(args.builder_command, args.change_request)
        qa = subprocess.run([sys.executable, str(ROOT / "scripts" / "run_regression.py"), "--manifest", str(args.manifest), "--output", str(test_path)])
        run_adapter(args.supervisor_command, test_path)
        review = json.loads(review_path.read_text(encoding="utf-8"))
        jsonschema.validate(review, schema)
        blockers = [item for item in review["findings"] if item["blocks_release"]]
        unauthorized = [item for item in review["findings"] if item["category"] == "unauthorized_change"]
        escalations = [item for item in review["findings"] if item.get("category") in ESCALATIONS or item["requires_jack_decision"]]
        if escalations:
            print(json.dumps({"status": "escalate", "cycle": cycle, "findings": escalations}, indent=2))
            return 2
        if qa.returncode == 0 and not blockers and not unauthorized:
            print(json.dumps({"status": "approved", "cycle": cycle}, indent=2))
            return 0
    print(json.dumps({"status": "escalate", "category": "UNRESOLVED_AGENT_DISAGREEMENT", "cycles": args.max_cycles}, indent=2))
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
