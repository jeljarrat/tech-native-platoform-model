"""Regression tests for the v2.4.5 shared-formula repair engine."""

from __future__ import annotations

import ast
import unittest
import zipfile
from pathlib import Path

from lxml import etree
from openpyxl import load_workbook

from scripts.expand_shared_formulas import MAIN_NS, expand_sheet
from scripts.verify_package_repair import LINEAGE_RANGES, compare_native_excel_controls


ROOT = Path(__file__).resolve().parents[1]


def _imports(path: Path) -> set[str | None]:
    tree = ast.parse(path.read_text(encoding="utf-8"))
    return {
        node.module if isinstance(node, ast.ImportFrom) else alias.name
        for node in ast.walk(tree) if isinstance(node, (ast.Import, ast.ImportFrom))
        for alias in node.names
    }


class SharedFormulaExpansionTests(unittest.TestCase):
    @staticmethod
    def _formula_xml() -> bytes:
        return f'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="{MAIN_NS}"><sheetData><row r="1">
<c r="A1"><f t="shared" ref="A1:B1" si="26">SUM(A2:A3)</f><v>3</v></c>
<c r="B1"><f t="shared" si="26"/><v>7</v></c>
<c r="C1"><f>SUM(C2:C3)</f><v>11</v></c>
</row></sheetData></worksheet>'''.encode()

    def test_self_closing_shared_member_and_following_formula(self) -> None:
        output, report = expand_sheet(self._formula_xml())
        root = etree.fromstring(output)
        formulas = [(node.text, dict(node.attrib)) for node in root.iter(f"{{{MAIN_NS}}}f")]
        self.assertEqual(formulas, [("SUM(A2:A3)", {}), ("SUM(B2:B3)", {}), ("SUM(C2:C3)", {})])
        self.assertEqual(report["shared_formula_records_before"], 2)
        self.assertEqual(report["shared_formula_groups_before"], 1)

    def test_expander_and_validator_do_not_share_parser_or_resolver(self) -> None:
        expansion_imports = _imports(ROOT / "scripts" / "expand_shared_formulas.py")
        validation_imports = _imports(ROOT / "scripts" / "verify_package_repair.py")
        self.assertIn("lxml", expansion_imports)
        self.assertNotIn("lxml", validation_imports)
        self.assertNotIn("expand_shared_formulas", validation_imports)
        self.assertNotIn("openpyxl.formula.translate", expansion_imports)

    def test_explicit_lineage_matches_resolved_v244(self) -> None:
        source = load_workbook(ROOT / "model" / "base-v2.4.4.xlsx", data_only=False)
        candidate = load_workbook(ROOT / "model" / "base-v2.4.5.xlsx", data_only=False)
        for sheet, ranges in LINEAGE_RANGES.items():
            for address in ranges:
                for source_row, candidate_row in zip(source[sheet][address], candidate[sheet][address]):
                    for source_cell, candidate_cell in zip(source_row, candidate_row):
                        self.assertEqual(source_cell.value, candidate_cell.value, f"{sheet}!{source_cell.coordinate}")

    def test_raw_candidate_has_no_shared_formula_records(self) -> None:
        with zipfile.ZipFile(ROOT / "model" / "base-v2.4.5.xlsx") as package:
            worksheet_xml = [package.read(name) for name in package.namelist() if name.startswith("xl/worksheets/") and name.endswith(".xml")]
        self.assertFalse(any(b't="shared"' in xml for xml in worksheet_xml))

    def test_cached_values_match_source_before_native_recalc(self) -> None:
        source = load_workbook(ROOT / "model" / "base-v2.4.4.xlsx", data_only=True)
        candidate = load_workbook(ROOT / "model" / "base-v2.4.5.xlsx", data_only=True)
        for source_ws, candidate_ws in zip(source.worksheets, candidate.worksheets):
            for source_row, candidate_row in zip(source_ws.iter_rows(), candidate_ws.iter_rows()):
                for source_cell, candidate_cell in zip(source_row, candidate_row):
                    self.assertEqual(source_cell.value, candidate_cell.value, f"{source_ws.title}!{source_cell.coordinate}")

    def test_native_excel_control_gate_has_66_controls(self) -> None:
        candidate = ROOT / "model" / "base-v2.4.5.xlsx"
        result = compare_native_excel_controls(candidate, candidate)
        self.assertEqual(result["control_count"], 66)
        self.assertTrue(result["passed"])


if __name__ == "__main__":
    unittest.main()
