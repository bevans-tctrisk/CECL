"""Parser for the recurring 'Monthly Balance by Pool/Type' workbook a CU
sends each quarter.

Public API
==========
- ``analyse_file(path)`` -> dict with detected layout + parsed labels/dates.
- ``normalize_to_month_end(value)`` -> ``datetime.date | None``. Snaps any
  date-ish value to the last day of its calendar month.
"""
from __future__ import annotations

import calendar
import re
from datetime import date, datetime
from pathlib import Path
from typing import Any

from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


# ---------------------------------------------------------------------------
# Date normalization
# ---------------------------------------------------------------------------

_MONTHS = {
    "jan": 1, "january": 1,
    "feb": 2, "february": 2,
    "mar": 3, "march": 3,
    "apr": 4, "april": 4,
    "may": 5,
    "jun": 6, "june": 6,
    "jul": 7, "july": 7,
    "aug": 8, "august": 8,
    "sep": 9, "sept": 9, "september": 9,
    "oct": 10, "october": 10,
    "nov": 11, "november": 11,
    "dec": 12, "december": 12,
}


def _last_day(year: int, month: int) -> date:
    return date(year, month, calendar.monthrange(year, month)[1])


def _parse_string_date(s: str) -> tuple[int, int, int] | None:
    """Return (year, month, day) or None. Day defaults to 1 if absent."""
    s = s.strip()
    if not s:
        return None

    # ISO YYYY-MM-DD or YYYY-MM
    m = re.match(r"^(\d{4})[-/](\d{1,2})(?:[-/](\d{1,2}))?$", s)
    if m:
        y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3) or 1)
        if 1 <= mo <= 12:
            return y, mo, d

    # M/D/YYYY or M-D-YYYY (US-style)
    m = re.match(r"^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})$", s)
    if m:
        mo, d, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
        if y < 100:
            y += 2000 if y < 70 else 1900
        if 1 <= mo <= 12:
            return y, mo, d

    # MM-YYYY or MM/YYYY (no day)
    m = re.match(r"^(\d{1,2})[/-](\d{4})$", s)
    if m:
        mo, y = int(m.group(1)), int(m.group(2))
        if 1 <= mo <= 12:
            return y, mo, 1

    # "Mon YYYY" / "Month YYYY" / "Mon-YY"
    m = re.match(r"^([A-Za-z]+)[\s\-/]+(\d{2,4})$", s)
    if m:
        mo = _MONTHS.get(m.group(1).lower())
        y = int(m.group(2))
        if y < 100:
            y += 2000 if y < 70 else 1900
        if mo:
            return y, mo, 1

    # "YYYY Mon" reverse
    m = re.match(r"^(\d{4})[\s\-/]+([A-Za-z]+)$", s)
    if m:
        y = int(m.group(1))
        mo = _MONTHS.get(m.group(2).lower())
        if mo:
            return y, mo, 1

    return None


def normalize_to_month_end(value: Any) -> date | None:
    """Snap any date-ish value to the last day of its calendar month."""
    if value is None or value == "":
        return None
    if isinstance(value, datetime):
        return _last_day(value.year, value.month)
    if isinstance(value, date):
        return _last_day(value.year, value.month)
    if isinstance(value, (int, float)):
        # Excel serial date (1900-based, ignoring the 1900 leap-year bug).
        try:
            from openpyxl.utils.datetime import from_excel
            d = from_excel(float(value))
            if isinstance(d, datetime):
                return _last_day(d.year, d.month)
            if isinstance(d, date):
                return _last_day(d.year, d.month)
        except Exception:  # noqa: BLE001
            return None
    if isinstance(value, str):
        parts = _parse_string_date(value)
        if parts:
            y, mo, _d = parts
            return _last_day(y, mo)
    return None


# ---------------------------------------------------------------------------
# File analysis
# ---------------------------------------------------------------------------

_LIKELY_LABEL_HEADERS = {
    "loan type", "loan pool", "pool", "type", "category", "description",
    "loan category",
}

_SKIP_LABELS = {
    "total", "total loans", "subtotal", "sub-total", "sub total",
    "grand total",
}

# Phrases that, when present in a label, identify the row holding the
# Allowance for Credit Loss balance per month.
_ACL_LABEL_PATTERNS = [
    "alll balance",
    "acl balance",
    "allowance for credit loss",
    "allowance for loan loss",
    "allowance for loan and lease loss",
    "allowance balance",
]


def _is_acl_label(value: Any) -> bool:
    if not isinstance(value, str):
        return False
    s = value.strip().lower()
    if not s:
        return False
    return any(p in s for p in _ACL_LABEL_PATTERNS)


def _is_label_row(value: Any) -> bool:
    if not isinstance(value, str):
        return False
    s = value.strip()
    if not s:
        return False
    if s.lower() in _SKIP_LABELS:
        return False
    return True


def _scan_sheet(ws) -> dict[str, Any]:
    """Find header row + first date column on a single worksheet."""
    best: dict[str, Any] | None = None

    # Look at the first ~30 rows for a row containing 2+ date-like cells.
    for row_idx, row in enumerate(
        ws.iter_rows(min_row=1, max_row=30, values_only=False), start=1
    ):
        date_cols: list[tuple[int, date, Any]] = []
        for cell in row:
            if cell.value is None:
                continue
            norm = normalize_to_month_end(cell.value)
            if norm is not None:
                date_cols.append((cell.column, norm, cell.value))
        if len(date_cols) >= 2:
            score = len(date_cols)
            if best is None or score > best["score"]:
                best = {
                    "score": score,
                    "header_row": row_idx,
                    "first_date_col_idx": date_cols[0][0],
                    "dates": [
                        {
                            "col": get_column_letter(c),
                            "raw": str(raw),
                            "normalized": d.isoformat(),
                        }
                        for (c, d, raw) in date_cols
                    ],
                }

    if not best:
        return {"ok": False, "error": "No date-like header row found in first 30 rows."}

    # Pool/type label column: scan col A first 60 rows for any string row
    # below the header.
    label_col_idx = 1
    labels: list[str] = []
    seen: set[str] = set()
    acl_row: int | None = None
    acl_label: str = ""
    for row_idx in range(best["header_row"] + 1, best["header_row"] + 60):
        cell = ws.cell(row=row_idx, column=label_col_idx)
        if cell.value is None:
            continue
        # ACL row detection (independent of pool-label collection).
        if acl_row is None and _is_acl_label(cell.value):
            acl_row = row_idx
            acl_label = str(cell.value).strip()
        if not _is_label_row(cell.value):
            continue
        s = cell.value.strip()
        key = s.lower()
        if key in seen:
            continue
        seen.add(key)
        labels.append(s)

    # Extract per-date ACL history if a row was identified.
    acl_history: dict[str, float] = {}
    if acl_row is not None:
        for entry in best["dates"]:
            col_letter = entry["col"]
            col_idx = _col_letter_to_idx(col_letter)
            if not col_idx:
                continue
            v = _coerce_number(ws.cell(row=acl_row, column=col_idx).value)
            if v is not None:
                acl_history[entry["normalized"]] = abs(v)

    return {
        "ok": True,
        "sheet": ws.title,
        "header_row": best["header_row"],
        "pool_name_col": get_column_letter(label_col_idx),
        "first_date_col": get_column_letter(best["first_date_col_idx"]),
        "dates": best["dates"],
        "parsed_pool_labels": labels,
        "acl_row": acl_row,
        "acl_label": acl_label,
        "acl_history": acl_history,
    }


def analyse_file(path: str | Path) -> dict[str, Any]:
    """Open the workbook, pick the most data-rich sheet, and return layout."""
    p = Path(path)
    if not p.exists():
        return {"ok": False, "error": f"File not found: {path}"}

    try:
        wb = load_workbook(p, read_only=True, data_only=True)
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "error": f"Could not open workbook: {exc}"}

    best: dict[str, Any] | None = None
    errors: list[str] = []
    try:
        for ws in wb.worksheets:
            try:
                result = _scan_sheet(ws)
            except Exception as exc:  # noqa: BLE001
                errors.append(f"{ws.title}: {exc}")
                continue
            if not result.get("ok"):
                continue
            score = len(result.get("dates", [])) + len(result.get("parsed_pool_labels", []))
            if best is None or score > best["_score"]:
                result["_score"] = score
                best = result
    finally:
        wb.close()

    if not best:
        msg = "Could not detect a header row with month-end dates."
        if errors:
            msg += " (" + "; ".join(errors[:3]) + ")"
        return {"ok": False, "error": msg}

    best.pop("_score", None)
    return best


def extract_row_history(
    saved_path: str | Path,
    sheet: str,
    header_row: int,
    target_row: int,
) -> dict[str, Any]:
    """Re-read a single row's per-date values from the monthly balance file.

    Used when the user overrides the auto-detected ACL row number.

    Returns ``{"ok": bool, "error": str|None, "label": str,
    "history": {"YYYY-MM-DD": float, ...}}``.
    """
    p = Path(saved_path)
    if not p.exists():
        return {"ok": False, "error": f"File not found: {saved_path}",
                "label": "", "history": {}}
    if not sheet or not header_row or not target_row:
        return {"ok": False, "error": "Missing sheet / header_row / target_row",
                "label": "", "history": {}}
    try:
        wb = load_workbook(p, read_only=True, data_only=True)
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "error": f"Could not open workbook: {exc}",
                "label": "", "history": {}}
    try:
        if sheet not in wb.sheetnames:
            return {"ok": False, "error": f"Sheet '{sheet}' not found",
                    "label": "", "history": {}}
        ws = wb[sheet]
        # Walk the header row, collect (col_idx, normalized_date) pairs.
        date_cols: list[tuple[int, date]] = []
        for cell in ws[header_row]:
            d = normalize_to_month_end(cell.value)
            if d is not None:
                date_cols.append((cell.column, d))
        history: dict[str, float] = {}
        for col_idx, d in date_cols:
            v = _coerce_number(ws.cell(row=target_row, column=col_idx).value)
            if v is not None:
                history[d.isoformat()] = abs(v)
        label_cell = ws.cell(row=target_row, column=1)
        label = str(label_cell.value).strip() if label_cell.value else ""
        return {"ok": True, "error": None, "label": label, "history": history}
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Pool-map auto-seeding
# ---------------------------------------------------------------------------

def seed_pool_map(
    parsed_labels: list[str],
    balance_title_map: dict[str, str] | None,
) -> tuple[dict[str, str], dict[str, str]]:
    """Return (pool_map, status_per_label).

    ``status_per_label`` is one of: ``matched``, ``ignored`` (mapped to
    empty in WARM), or ``new`` (label not present in WARM map).
    """
    btm = {k.strip().lower(): v for k, v in (balance_title_map or {}).items()}
    pool_map: dict[str, str] = {}
    status: dict[str, str] = {}
    for label in parsed_labels:
        key = label.strip().lower()
        if key in btm:
            value = (btm[key] or "").strip()
            if value:
                pool_map[label] = value
                status[label] = "matched"
            else:
                pool_map[label] = ""
                status[label] = "ignored"
        else:
            pool_map[label] = ""
            status[label] = "new"
    return pool_map, status


# ---------------------------------------------------------------------------
# Per-pool balance extraction (latest period)
# ---------------------------------------------------------------------------

def _coerce_number(v: Any) -> float | None:
    """Best-effort numeric coercion for a balance cell."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        try:
            return float(v)
        except (TypeError, ValueError):
            return None
    s = str(v).strip()
    if not s:
        return None
    neg = False
    if s.startswith("(") and s.endswith(")"):
        neg = True
        s = s[1:-1]
    s = s.replace("$", "").replace(",", "").replace(" ", "")
    if s.startswith("-"):
        neg = True
        s = s[1:]
    try:
        n = float(s)
    except ValueError:
        return None
    return -n if neg else n


def _col_letter_to_idx(letter: str) -> int:
    """1-based column index from a letter (A->1)."""
    s = str(letter or "").strip().upper()
    if not s.isalpha():
        return 0
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - ord("A") + 1)
    return n


def pool_balances_for_latest_period(
    saved_path: str | Path,
    sheet: str,
    header_row: int,
    pool_name_col: str,
    label_to_pool: dict[str, str] | None = None,
    period: str | None = None,
) -> dict[str, Any]:
    """Read the monthly balance file and return per-pool balances for the
    latest detected period (or for ``period`` if given as ISO ``YYYY-MM-DD``).

    Returns::

        {
          "ok": bool,
          "error": str | None,
          "period": "YYYY-MM-DD",
          "by_pool": {pool_name: balance_float},
          "raw_rows": [{label, balance, mapped_pool}, ...],
        }

    ``label_to_pool`` is the wizard's ``monthly_bal.pool_map`` (raw label ->
    canonical pool name). Labels that map to an empty/None pool are
    treated as "ignored" and dropped from ``by_pool`` totals (but still
    appear in ``raw_rows``).
    """
    p = Path(saved_path)
    if not p.exists():
        return {"ok": False, "error": f"File not found: {saved_path}",
                "period": "", "by_pool": {}, "raw_rows": []}
    if not sheet or not header_row or not pool_name_col:
        return {"ok": False, "error": "Missing sheet / header_row / pool_name_col",
                "period": "", "by_pool": {}, "raw_rows": []}

    try:
        wb = load_workbook(p, read_only=True, data_only=True)
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "error": f"Could not open workbook: {exc}",
                "period": "", "by_pool": {}, "raw_rows": []}

    try:
        if sheet not in wb.sheetnames:
            return {"ok": False, "error": f"Sheet '{sheet}' not found",
                    "period": "", "by_pool": {}, "raw_rows": []}
        ws = wb[sheet]

        label_col_idx = _col_letter_to_idx(pool_name_col) or 1

        # Walk the header row, collect (col_idx, normalized_date) pairs.
        date_cols: list[tuple[int, date]] = []
        for cell in ws[header_row]:
            d = normalize_to_month_end(cell.value)
            if d is not None:
                date_cols.append((cell.column, d))
        if not date_cols:
            return {"ok": False,
                    "error": f"No date columns found on row {header_row}",
                    "period": "", "by_pool": {}, "raw_rows": []}

        # Pick the target period: explicit, else latest.
        target: tuple[int, date] | None = None
        if period:
            try:
                want = date.fromisoformat(period)
                for c, d in date_cols:
                    if d == want:
                        target = (c, d)
                        break
            except ValueError:
                target = None
        if target is None:
            target = max(date_cols, key=lambda t: t[1])
        target_col, target_date = target

        ltp = {
            (k or "").strip().lower(): (v or "").strip()
            for k, v in (label_to_pool or {}).items()
        }

        by_pool: dict[str, float] = {}
        raw_rows: list[dict[str, Any]] = []
        seen_labels: set[str] = set()

        # Walk data rows below the header; stop after a stretch of blanks.
        max_row = ws.max_row or (header_row + 200)
        blanks = 0
        for r in range(header_row + 1, max_row + 1):
            label_cell = ws.cell(row=r, column=label_col_idx).value
            if not _is_label_row(label_cell):
                blanks += 1
                if blanks >= 15:
                    break
                continue
            blanks = 0
            label = str(label_cell).strip()
            key = label.lower()
            if key in seen_labels:
                continue
            seen_labels.add(key)

            bal = _coerce_number(
                ws.cell(row=r, column=target_col).value
            )
            mapped = ltp.get(key, "")
            raw_rows.append({
                "label": label,
                "balance": bal,
                "mapped_pool": mapped,
            })
            if mapped and bal is not None:
                by_pool[mapped] = by_pool.get(mapped, 0.0) + bal
    finally:
        wb.close()

    return {
        "ok": True,
        "error": None,
        "period": target_date.isoformat(),
        "by_pool": by_pool,
        "raw_rows": raw_rows,
    }
