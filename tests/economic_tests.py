"""Golden outputs, economic invariants, and scenarios declared by the manifest."""

from __future__ import annotations

from pathlib import Path

from scripts.qa_common import finding, result


def _compare(actual, operator: str, expected) -> bool:
    return {"eq": actual == expected, "gte": actual >= expected, "lte": actual <= expected}.get(operator, False)


def run(workbook: Path, manifest: dict, model_spec: dict) -> list[dict]:
    snapshot = manifest.get("output_snapshot", {})
    golden = manifest.get("golden_outputs", {})
    invariants = manifest.get("economic_invariants", [])
    scenarios = manifest.get("scenario_tests", [])
    golden_ok = bool(golden) and all(snapshot.get(key) == value for key, value in golden.items())
    invariant_ok = bool(invariants) and all(_compare(snapshot.get(rule["key"]), rule["operator"], rule["value"]) for rule in invariants)
    scenario_ok = bool(scenarios) and all(item.get("passed") is True for item in scenarios)
    checks = [
        result("golden_output_checks", golden_ok, "Snapshot must match approved golden outputs"),
        result("economic_invariants", invariant_ok, "All declared economic invariants must hold"),
        result("scenario_tests", scenario_ok, "All declared scenario tests must pass"),
    ]
    for index, item in enumerate(checks, 1):
        if not item["passed"]:
            item["findings"] = [finding(f"ECO-{index:03d}", item["test"], item["details"], "Resolve economic mismatch or obtain an approved decision", item["test"])]
    return checks

