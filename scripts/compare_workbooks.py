#!/usr/bin/env python3
"""Compare deterministic snapshots of Base and candidate workbooks."""

from __future__ import annotations

import argparse
import json
from pathlib import Path

from extract_workbook import extract


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("base", type=Path)
    parser.add_argument("candidate", type=Path)
    parser.add_argument("--output", type=Path, required=True)
    args = parser.parse_args()
    base, candidate = extract(args.base), extract(args.candidate)
    diff = {
        "base_sha256": base["sha256"],
        "candidate_sha256": candidate["sha256"],
        "architecture_changed": base["tab_order"] != candidate["tab_order"],
        "base_tab_order": base["tab_order"],
        "candidate_tab_order": candidate["tab_order"],
        "changed_sheets": sorted({s["name"] for s in base["sheets"]} | {s["name"] for s in candidate["sheets"]}),
    }
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(diff, indent=2, sort_keys=True) + "\n", encoding="utf-8")
    return 1 if diff["architecture_changed"] else 0


if __name__ == "__main__":
    raise SystemExit(main())

