"""Shared deterministic QA utilities."""

from __future__ import annotations

import hashlib
import json
from pathlib import Path
from typing import Any

import yaml


def load_yaml(path: Path) -> dict[str, Any]:
    return yaml.safe_load(path.read_text(encoding="utf-8")) or {}


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256(path.read_bytes()).hexdigest()
    return digest


def result(name: str, passed: bool, details: str, blockers: list[dict] | None = None) -> dict:
    return {"test": name, "passed": passed, "details": details, "findings": blockers or []}


def finding(finding_id: str, category: str, issue: str, required_fix: str, required_test: str, *, sheet: str | None = None, cell: str | None = None, severity: str = "high") -> dict:
    return {
        "finding_id": finding_id,
        "severity": severity,
        "category": category,
        "sheet": sheet,
        "cell_or_range": cell,
        "issue": issue,
        "economic_effect": "Requires reviewer assessment",
        "blocks_release": severity in {"critical", "high"},
        "requires_jack_decision": False,
        "required_fix": required_fix,
        "required_test": required_test,
    }


def write_json(path: Path, payload: Any) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(json.dumps(payload, indent=2, sort_keys=True) + "\n", encoding="utf-8")

