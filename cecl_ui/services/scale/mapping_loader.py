"""Load SCALE mapping CSVs.

A mapping CSV has columns: ``field_code, sheet, cell``. Each row tells
``runner.fill_template`` which Solr field to write into which cell on
which sheet.
"""
from __future__ import annotations

import csv
from pathlib import Path
from typing import Iterable


REQUIRED_HEADERS = {"field_code", "sheet", "cell"}


def load_rows(path: str | Path) -> list[dict]:
    p = Path(path)
    rows: list[dict] = []
    with p.open("r", encoding="utf-8-sig", newline="") as f:
        reader = csv.DictReader(f)
        headers = {h.lower() for h in (reader.fieldnames or [])}
        if not REQUIRED_HEADERS.issubset(headers):
            raise ValueError(
                "Mapping CSV must have headers: field_code, sheet, cell"
            )
        for row in reader:
            rows.append(
                {
                    "field_code": (row.get("field_code") or "").strip(),
                    "sheet": (row.get("sheet") or "").strip(),
                    "cell": (row.get("cell") or "").strip(),
                }
            )
    return rows


def summarize(rows: Iterable[dict]) -> dict:
    rows = list(rows)
    sheets: dict[str, int] = {}
    for r in rows:
        sheets[r["sheet"]] = sheets.get(r["sheet"], 0) + 1
    return {
        "total": len(rows),
        "sheets": sheets,
        "field_codes": sorted({r["field_code"] for r in rows if r["field_code"]}),
    }
