#!/usr/bin/env python3
"""Run all deterministic QA suites for one candidate manifest."""

from __future__ import annotations

import argparse
import sys
import traceback
from pathlib import Path
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from scripts.qa_common import load_yaml, write_json
from tests import economic_tests, formula_tests, package_tests, presentation_tests

QA_PROFILES = {"base": None, "task_c": "task_c"}


class ManifestWorkbookPathError(ValueError):
    """The candidate manifest has no usable workbook path."""


def resolve_workbook_path(manifest_path: Path, manifest: dict[str, Any]) -> Path:
    """Resolve canonical and legacy workbook declarations with explicit validation."""
    top_level = manifest.get("workbook_path")
    workbook = manifest.get("workbook")
    if top_level is not None:
        if not isinstance(top_level, str) or not top_level.strip():
            raise ManifestWorkbookPathError("manifest.workbook_path must be a non-empty string")
        value, repo_relative = top_level, False
    elif isinstance(workbook, dict):
        value = workbook.get("path")
        if not isinstance(value, str) or not value.strip():
            raise ManifestWorkbookPathError("manifest.workbook.path must be a non-empty string")
        repo_relative = True
    elif isinstance(workbook, str):
        if not workbook.strip():
            raise ManifestWorkbookPathError("manifest.workbook must be a non-empty string")
        value, repo_relative = workbook, False
    elif workbook is None:
        raise ManifestWorkbookPathError("manifest must define workbook_path or workbook.path")
    else:
        raise ManifestWorkbookPathError("manifest.workbook must be a string or mapping containing path")
    path = Path(value)
    if path.is_absolute():
        return path.resolve()
    return ((ROOT / path) if repo_relative else (manifest_path.parent / path)).resolve()


def run_qa_profile(profile: str, workbook: Path, manifest: dict[str, Any], spec: dict[str, Any]) -> list[dict]:
    if profile not in QA_PROFILES:
        raise ValueError(f"Unsupported qa_profile {profile!r}; expected one of {sorted(QA_PROFILES)}")
    if profile == "task_c":
        from scripts.run_task_c_qa import run_checks
        report = run_checks(ROOT, write_reports=False)
        return [{"test": item["test_id"], "passed": item["passed"], "details": item.get("details", ""),
                 "findings": [] if item["passed"] else [{
                     "finding_id": item["test_id"], "severity": "high", "category": item["suite"],
                     "sheet": None, "cell_or_range": None, "issue": item["test_id"],
                     "economic_effect": item.get("details", "Requires reviewer assessment"),
                     "blocks_release": True, "requires_jack_decision": False,
                     "required_fix": "Return to builder", "required_test": item["test_id"]}]}
                for item in report["tests"]]
    results = []
    for suite in (package_tests, formula_tests, economic_tests, presentation_tests):
        results.extend(suite.run(workbook, manifest, spec))
    return results


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--output", type=Path, default=ROOT / "reports" / "test-report.json")
    args = parser.parse_args()
    try:
        manifest = load_yaml(args.manifest)
        workbook = resolve_workbook_path(args.manifest.resolve(), manifest)
        spec = load_yaml(ROOT / "spec" / "model-spec.yaml")
        results = run_qa_profile(str(manifest.get("qa_profile", "base")), workbook, manifest, spec)
        payload = {"manifest": str(args.manifest), "workbook": str(workbook), "passed": all(item["passed"] for item in results), "results": results}
        write_json(args.output, payload)
        return 0 if payload["passed"] else 1
    except Exception as error:
        payload = {"manifest": str(args.manifest), "passed": False, "results": [],
                   "infrastructure_error": {"type": type(error).__name__, "message": str(error),
                                            "traceback": traceback.format_exc()}}
        write_json(args.output, payload)
        print(payload["infrastructure_error"]["traceback"], file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
