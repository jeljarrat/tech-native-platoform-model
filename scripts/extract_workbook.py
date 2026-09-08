#!/usr/bin/env python3
"""Extract a deterministic workbook snapshot without changing the workbook."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
from collections import defaultdict
from pathlib import Path

from openpyxl import load_workbook


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def extract(path: Path) -> dict:
    formulas = load_workbook(path, data_only=False, read_only=False)
    cached = load_workbook(path, data_only=True, read_only=False)
    sheets = []
    for ws in formulas.worksheets:
        cached_ws = cached[ws.title]
        cells = {}
        for row in ws.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                item = {"value": cell.value}
                if isinstance(cell.value, str) and cell.value.startswith("="):
                    item["cached_value"] = cached_ws[cell.coordinate].value
                cells[cell.coordinate] = item
        sheets.append({"name": ws.title, "cells": cells})
    return {"workbook": path.name, "sha256": sha256(path), "tab_order": formulas.sheetnames, "sheets": sheets}


def structural_snapshot(path: Path) -> dict:
    workbook = load_workbook(path, data_only=False, read_only=False)
    return {
        "workbook": path.name,
        "sha256": sha256(path),
        "tab_order": workbook.sheetnames,
        "calculation": {
            "mode": workbook.calculation.calcMode,
            "full_calc_on_load": workbook.calculation.fullCalcOnLoad,
            "force_full_calc": workbook.calculation.forceFullCalc,
            "iterate": workbook.calculation.iterate,
        },
        "sheets": [
            {
                "name": ws.title,
                "max_row": ws.max_row,
                "max_column": ws.max_column,
                "merged_ranges": sorted(str(item) for item in ws.merged_cells.ranges),
                "chart_count": len(ws._charts),
                "table_names": sorted(ws.tables.keys()),
            }
            for ws in workbook.worksheets
        ],
    }


def formula_dependency_snapshot(path: Path) -> dict:
    workbook = load_workbook(path, data_only=False, read_only=False)
    dependencies = defaultdict(set)
    formulas = []
    reference_pattern = re.compile(r"(?:'([^']+)'|([A-Za-z0-9_ &-]+))!\$?[A-Z]{1,3}\$?\d+")
    for ws in workbook.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if not (isinstance(cell.value, str) and cell.value.startswith("=")):
                    continue
                refs = sorted({left or right for left, right in reference_pattern.findall(cell.value)})
                dependencies[ws.title].update(refs)
                formulas.append({"sheet": ws.title, "cell": cell.coordinate, "formula": cell.value, "sheet_dependencies": refs})
    return {
        "workbook": path.name,
        "sha256": sha256(path),
        "formula_count": len(formulas),
        "dependency_graph": {key: sorted(value) for key, value in sorted(dependencies.items())},
        "formulas": formulas,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("workbook", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    parser.add_argument("--structure-output", type=Path)
    parser.add_argument("--formula-output", type=Path)
    args = parser.parse_args()
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(extract(args.workbook), indent=2, sort_keys=True, default=str) + "\n", encoding="utf-8")
    if args.structure_output:
        args.structure_output.parent.mkdir(parents=True, exist_ok=True)
        args.structure_output.write_text(json.dumps(structural_snapshot(args.workbook), indent=2, sort_keys=True, default=str) + "\n", encoding="utf-8")
    if args.formula_output:
        args.formula_output.parent.mkdir(parents=True, exist_ok=True)
        args.formula_output.write_text(json.dumps(formula_dependency_snapshot(args.workbook), indent=2, sort_keys=True, default=str) + "\n", encoding="utf-8")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
