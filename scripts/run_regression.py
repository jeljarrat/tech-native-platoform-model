#!/usr/bin/env python3
"""Run all deterministic QA suites for one candidate manifest."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

from scripts.qa_common import load_yaml, write_json
from tests import economic_tests, formula_tests, package_tests, presentation_tests


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("--manifest", type=Path, required=True)
    parser.add_argument("--output", type=Path, default=ROOT / "reports" / "test-report.json")
    args = parser.parse_args()
    manifest = load_yaml(args.manifest)
    workbook = (args.manifest.parent / manifest["workbook"]).resolve()
    spec = load_yaml(ROOT / "spec" / "model-spec.yaml")
    results = []
    for suite in (package_tests, formula_tests, economic_tests, presentation_tests):
        results.extend(suite.run(workbook, manifest, spec))
    payload = {"manifest": str(args.manifest), "workbook": str(workbook), "passed": all(item["passed"] for item in results), "results": results}
    write_json(args.output, payload)
    return 0 if payload["passed"] else 1


if __name__ == "__main__":
    raise SystemExit(main())

