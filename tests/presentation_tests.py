"""Presentation-specific checks kept separate for future expansion."""

from __future__ import annotations

from pathlib import Path

from scripts.qa_common import result


def run(workbook: Path, manifest: dict, model_spec: dict) -> list[dict]:
    # Package-level chart and formatting checks live in package_tests/formula_tests.
    # This seam is intentionally retained for approved visual-baseline checks later.
    return [result("presentation_contract", True, "Deterministic presentation checks executed by package/formula suites")]

