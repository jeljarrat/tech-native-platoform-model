#!/usr/bin/env python3
"""Expand XLSX shared formulas with XML parsing and explicit A1 translation."""

from __future__ import annotations

import argparse
import hashlib
import json
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from lxml import etree
from openpyxl.formula.tokenizer import Tokenizer
from openpyxl.utils.cell import get_column_letter, column_index_from_string


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
CELL_REF_RE = re.compile(r"^(?P<col>\$?[A-Z]{1,3})(?P<row>\$?\d+)$")
COL_RANGE_RE = re.compile(r"^(?P<left>\$?[A-Z]{1,3}):(?P<right>\$?[A-Z]{1,3})$")
ROW_RANGE_RE = re.compile(r"^(?P<top>\$?\d+):(?P<bottom>\$?\d+)$")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _translate_column(value: str, delta: int) -> str:
    if value.startswith("$"):
        return value
    translated = column_index_from_string(value) + delta
    if translated < 1:
        raise ValueError(f"Column translation moved {value} outside the worksheet")
    return get_column_letter(translated)


def _translate_row(value: str, delta: int) -> str:
    if value.startswith("$"):
        return value
    translated = int(value) + delta
    if translated < 1:
        raise ValueError(f"Row translation moved {value} outside the worksheet")
    return str(translated)


def _translate_reference(value: str, row_delta: int, col_delta: int) -> str:
    match = CELL_REF_RE.match(value)
    if match:
        return _translate_column(match.group("col"), col_delta) + _translate_row(match.group("row"), row_delta)
    match = COL_RANGE_RE.match(value)
    if match:
        return f"{_translate_column(match.group('left'), col_delta)}:{_translate_column(match.group('right'), col_delta)}"
    match = ROW_RANGE_RE.match(value)
    if match:
        return f"{_translate_row(match.group('top'), row_delta)}:{_translate_row(match.group('bottom'), row_delta)}"
    if ":" in value:
        left, right = value.split(":", 1)
        translated_left = _translate_reference(left, row_delta, col_delta)
        translated_right = _translate_reference(right, row_delta, col_delta)
        if translated_left != left or translated_right != right:
            return f"{translated_left}:{translated_right}"
    return value


def _split_sheet_prefix(value: str) -> tuple[str, str]:
    in_quote = False
    bracket_depth = 0
    last_bang = -1
    for index, char in enumerate(value):
        if char == "'":
            in_quote = not in_quote
        elif not in_quote and char == "[":
            bracket_depth += 1
        elif not in_quote and char == "]":
            bracket_depth = max(0, bracket_depth - 1)
        elif not in_quote and bracket_depth == 0 and char == "!":
            last_bang = index
    return (value[: last_bang + 1], value[last_bang + 1 :]) if last_bang >= 0 else ("", value)


def translate_formula(formula: str, origin: str, destination: str) -> str:
    """Translate relative A1 references without using openpyxl's Translator helper."""
    origin_match = CELL_REF_RE.match(origin)
    destination_match = CELL_REF_RE.match(destination)
    if not origin_match or not destination_match:
        raise ValueError(f"Invalid cell coordinate: {origin} -> {destination}")
    row_delta = int(destination_match.group("row")) - int(origin_match.group("row"))
    col_delta = column_index_from_string(destination_match.group("col").replace("$", "")) - column_index_from_string(origin_match.group("col").replace("$", ""))
    tokens = Tokenizer("=" + formula).items
    translated: list[str] = []
    for token in tokens:
        value = token.value
        if token.type == "OPERAND" and token.subtype == "RANGE":
            prefix, reference = _split_sheet_prefix(value)
            value = prefix + _translate_reference(reference, row_delta, col_delta)
        translated.append(value)
    return "".join(translated)


def expand_sheet(xml_bytes: bytes) -> tuple[bytes, dict]:
    parser = etree.XMLParser(remove_blank_text=False, resolve_entities=False)
    root = etree.fromstring(xml_bytes, parser)
    formula_tag = f"{{{MAIN_NS}}}f"
    cell_tag = f"{{{MAIN_NS}}}c"
    anchors: dict[str, tuple[str, str]] = {}
    shared_cells: list[tuple[etree._Element, etree._Element, str, str]] = []
    groups: set[str] = set()
    formula_count = 0
    for cell in root.iter(cell_tag):
        formula = cell.find(formula_tag)
        if formula is None:
            continue
        formula_count += 1
        if formula.get("t") != "shared":
            continue
        shared_index = formula.get("si")
        coordinate = cell.get("r")
        if not shared_index or not coordinate:
            raise ValueError("Shared formula is missing si or cell coordinate")
        groups.add(shared_index)
        text = formula.text or ""
        if text:
            if shared_index in anchors:
                raise ValueError(f"Shared formula group {shared_index} has multiple anchors")
            anchors[shared_index] = (coordinate, text)
        shared_cells.append((cell, formula, coordinate, shared_index))
    for _, _, coordinate, shared_index in shared_cells:
        if shared_index not in anchors:
            raise ValueError(f"Shared formula group {shared_index} at {coordinate} has no anchor")
    for _, formula, coordinate, shared_index in shared_cells:
        anchor_coordinate, anchor_formula = anchors[shared_index]
        resolved = anchor_formula if coordinate == anchor_coordinate else translate_formula(anchor_formula, anchor_coordinate, coordinate)
        formula.attrib.clear()
        formula.text = resolved
    output = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=None)
    return output, {
        "shared_formula_records_before": len(shared_cells),
        "shared_formula_records_after": 0,
        "shared_formula_groups_before": len(groups),
        "shared_formula_groups_after": 0,
        "formula_count_before": formula_count,
        "formula_count_after": formula_count,
    }


def copy_zipinfo(info: zipfile.ZipInfo) -> zipfile.ZipInfo:
    clone = zipfile.ZipInfo(info.filename, date_time=info.date_time)
    for name in ("compress_type", "comment", "extra", "create_system", "create_version", "extract_version", "reserved", "flag_bits", "volume", "internal_attr", "external_attr"):
        setattr(clone, name, getattr(info, name))
    return clone


def expand_workbook(source: Path, destination: Path) -> dict:
    if destination.exists():
        raise FileExistsError(destination)
    sheet_reports: dict[str, dict] = {}
    changed_members: list[str] = []
    destination.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.NamedTemporaryFile(dir=destination.parent, suffix=".xlsx", delete=False) as temp_handle:
        temp_path = Path(temp_handle.name)
    try:
        with zipfile.ZipFile(source, "r") as source_zip, zipfile.ZipFile(temp_path, "w") as target_zip:
            for info in source_zip.infolist():
                payload = source_zip.read(info.filename)
                if info.filename.startswith("xl/worksheets/") and info.filename.endswith(".xml"):
                    expanded, report = expand_sheet(payload)
                    sheet_reports[info.filename] = report
                    if report["shared_formula_records_before"]:
                        changed_members.append(info.filename)
                        payload = expanded
                target_zip.writestr(copy_zipinfo(info), payload)
        totals = {
            key: sum(report[key] for report in sheet_reports.values())
            for key in ("shared_formula_records_before", "shared_formula_records_after", "shared_formula_groups_before", "shared_formula_groups_after", "formula_count_before", "formula_count_after")
        }
        if totals["formula_count_before"] != totals["formula_count_after"] or totals["shared_formula_records_after"] or totals["shared_formula_groups_after"]:
            raise RuntimeError(json.dumps(totals, indent=2))
        shutil.move(temp_path, destination)
        return {"source_sha256": sha256(source), "output_sha256": sha256(destination), **totals, "changed_members": changed_members, "worksheets": sheet_reports}
    finally:
        temp_path.unlink(missing_ok=True)


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("source", type=Path)
    parser.add_argument("destination", type=Path)
    parser.add_argument("--report", type=Path, required=True)
    args = parser.parse_args()
    report = expand_workbook(args.source, args.destination)
    args.report.parent.mkdir(parents=True, exist_ok=True)
    args.report.write_text(json.dumps(report, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    print(json.dumps({key: value for key, value in report.items() if key != "worksheets"}, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
