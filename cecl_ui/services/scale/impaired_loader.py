"""SCALE Impaired Loans loader.

Reads the standardized ``CECL-SCALE Impaired Loans`` workbook a credit
union uploads and copies the data-entry rows (cols A:J) into the SCALE
output workbook's `` Impaired Loans ASC 310-10`` sheet, preserving the
template's calc columns K:Q (formulas) and summary K6:Q26.

Pools and impairment types are canonical across all credit unions for
the SCALE methodology, so we do NOT remap pool names or look up credit
grades — we just stream the rows through.

Source layout (per template ``CECL-SCALE Impaired Loans.xlsx``):
    Sheet:       " Impaired Loans ASC 310-10" (note leading space)
    Header row:  29
    Data rows:   30..414 (stop at first blank Impairment Type / col A)
    Columns:     A=Impairment Type, B=Member #, C=Loan Suffix,
                 D=Loan Pool, E=Current Balance, F=Days Delinquent,
                 G=Balance at Other Lender, H=Collateral Value,
                 I=Allowance Provided, J=Notes
"""
from __future__ import annotations

import warnings
from pathlib import Path
from typing import Any

import openpyxl


SHEET_CANDIDATES = (
    " Impaired Loans ASC 310-10",
    "Impaired Loans ASC 310-10",
)

# 1-based column letters mirrored as 0-based indexes for clarity.
_COLS = ("A", "B", "C", "D", "E", "F", "G", "H", "I", "J")
_FIELDS = (
    "impairment_type", "member", "suffix", "loan_pool",
    "current_balance", "days_delinquent", "other_lender_balance",
    "collateral_value", "allowance_provided", "notes",
)
_DATA_START_ROW = 30
_DATA_END_ROW = 414  # matches template formulas' $A$30:$A$414 range


def _find_sheet(wb) -> str | None:
    names = {s.strip().lower(): s for s in wb.sheetnames}
    for cand in SHEET_CANDIDATES:
        hit = names.get(cand.strip().lower())
        if hit:
            return hit
    # Fallback: any sheet containing "impaired" + "310" in its name.
    for s in wb.sheetnames:
        low = s.lower()
        if "impaired" in low and "310" in low:
            return s
    return None


def _coerce_num(value: Any) -> float:
    if value in (None, ""):
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    try:
        return float(str(value).strip().replace(",", "").replace("$", ""))
    except (TypeError, ValueError):
        return 0.0


def _str(value: Any) -> str:
    if value is None:
        return ""
    return str(value).strip()


def parse_file(path: str | Path) -> dict[str, Any]:
    """Parse a CU's SCALE Impaired Loans workbook.

    Returns ``{ok, error, sheet_used, cu_name, period, rows, row_count,
    total_balance}``. ``rows`` is a list of dicts using ``_FIELDS``
    above plus a derived ``member_suffix`` key.
    """
    out: dict[str, Any] = {
        "ok": False,
        "error": "",
        "sheet_used": "",
        "cu_name": "",
        "period": "",
        "rows": [],
        "row_count": 0,
        "total_balance": 0.0,
    }
    p = Path(path)
    if not p.exists():
        out["error"] = f"File not found: {p}"
        return out
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(p, data_only=True, read_only=False)
    except Exception as exc:  # noqa: BLE001
        out["error"] = f"Failed to open workbook: {exc}"
        return out

    sheet = _find_sheet(wb)
    if not sheet:
        out["error"] = (
            f"Could not find ' Impaired Loans ASC 310-10' sheet. "
            f"Sheets present: {wb.sheetnames}"
        )
        return out
    ws = wb[sheet]
    out["sheet_used"] = sheet
    # Header metadata (best-effort)
    try:
        out["cu_name"] = _str(ws["A2"].value)
    except Exception:  # noqa: BLE001
        pass
    try:
        out["period"] = _str(ws["B5"].value)
    except Exception:  # noqa: BLE001
        pass

    rows: list[dict[str, Any]] = []
    total = 0.0
    for r in range(_DATA_START_ROW, _DATA_END_ROW + 1):
        a = ws.cell(row=r, column=1).value
        if a in (None, ""):
            # Stop on first blank Impairment Type — matches template UX.
            break
        record = {}
        for idx, field in enumerate(_FIELDS):
            v = ws.cell(row=r, column=idx + 1).value
            if field in (
                "current_balance", "days_delinquent",
                "other_lender_balance", "collateral_value",
                "allowance_provided",
            ):
                record[field] = _coerce_num(v)
            else:
                record[field] = _str(v)
        record["member_suffix"] = (
            f"{record['member']}-{record['suffix']}"
            if record["member"] or record["suffix"]
            else ""
        )
        rows.append(record)
        total += record["current_balance"]

    out["ok"] = True
    out["rows"] = rows
    out["row_count"] = len(rows)
    out["total_balance"] = round(total, 2)
    return out


def apply_impaired_rows(
    workbook_path: str | Path, rows: list[dict[str, Any]],
) -> dict[str, Any]:
    """Write ``rows`` into the output workbook's impaired sheet.

    Clears A30:J<end> first so a re-run doesn't leave stale rows.
    Calc columns K:Q are left untouched — the template formulas pick up
    the new data automatically.
    """
    result = {
        "ok": False,
        "applied": 0,
        "cleared": 0,
        "sheet_used": "",
        "error": "",
    }
    if not rows:
        # Still clear any prior writes so re-runs with an emptied list
        # remove old data.
        rows = []
    p = Path(workbook_path)
    if not p.exists():
        result["error"] = f"Output workbook not found: {p}"
        return result
    try:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            wb = openpyxl.load_workbook(p)
    except Exception as exc:  # noqa: BLE001
        result["error"] = f"Failed to open output workbook: {exc}"
        return result
    sheet = _find_sheet(wb)
    if not sheet:
        result["error"] = (
            "Output template is missing ' Impaired Loans ASC 310-10' sheet."
        )
        return result
    ws = wb[sheet]
    result["sheet_used"] = sheet
    # Clear A:J for the full data range.
    cleared = 0
    for r in range(_DATA_START_ROW, _DATA_END_ROW + 1):
        for c in range(1, 11):  # cols A..J
            cell = ws.cell(row=r, column=c)
            if cell.value not in (None, ""):
                cell.value = None
                cleared += 1
    # Write rows (cap at template range).
    applied = 0
    max_rows = _DATA_END_ROW - _DATA_START_ROW + 1
    for i, row in enumerate(rows[:max_rows]):
        r = _DATA_START_ROW + i
        ws.cell(row=r, column=1).value = row.get("impairment_type") or None
        ws.cell(row=r, column=2).value = row.get("member") or None
        ws.cell(row=r, column=3).value = row.get("suffix") or None
        ws.cell(row=r, column=4).value = row.get("loan_pool") or None
        ws.cell(row=r, column=5).value = row.get("current_balance") or None
        ws.cell(row=r, column=6).value = row.get("days_delinquent") or None
        ws.cell(row=r, column=7).value = row.get("other_lender_balance") or None
        ws.cell(row=r, column=8).value = row.get("collateral_value") or None
        ws.cell(row=r, column=9).value = row.get("allowance_provided") or None
        ws.cell(row=r, column=10).value = row.get("notes") or None
        applied += 1
    wb.save(p)
    result["ok"] = True
    result["applied"] = applied
    result["cleared"] = cleared
    if len(rows) > max_rows:
        result["error"] = (
            f"Truncated: source has {len(rows)} rows but template only "
            f"supports {max_rows} (rows 30..{_DATA_END_ROW})."
        )
    return result
