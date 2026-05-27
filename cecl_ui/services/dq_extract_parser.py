"""Derive historical delinquency balances from month-end loan-data extracts.

Mirrors ``extract_hist_processor.rollup_dataframe`` but produces a
per-loan_code DQ rollup instead of a balance rollup. For each file::

    For each row in the extract:
      pool_code  = row[mapping['loan_pool_code']]
      balance    = clean(row[mapping['current_balance']])
      days_dq    = clean(row[mapping['days_delinquent']])

    rollup[pool_code] += balance           -> total_balance
    if days_dq >= dq_threshold:
        rollup[pool_code]._dq += balance   -> dq_amount

Returns one record per distinct loan_pool_code value seen in the file.
The caller upserts via ``delinquency_hist_processor.upsert_month``.
"""
from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any

from cecl_ui.services import (
    delinquency_hist_processor,
    extract_hist_processor,
)


DEFAULT_DQ_THRESHOLD = 60  # days; matches WARM convention (>=60 days)


def _coerce_days(val: Any) -> int | None:
    """Parse a 'days delinquent' cell to an int. Empty/non-numeric -> None."""
    if val is None:
        return None
    s = str(val).strip()
    if not s:
        return None
    # Strip non-digit/minus characters (handles "30 days", "30+", etc.)
    cleaned = re.sub(r"[^0-9\-]", "", s)
    if not cleaned or cleaned == "-":
        return None
    try:
        return int(cleaned)
    except ValueError:
        try:
            return int(float(s))
        except ValueError:
            return None


_DATE_RX_LIST = [
    (re.compile(r"(\d{4})[-_](\d{1,2})[-_](\d{1,2})"), "ymd"),
    (re.compile(r"(\d{1,2})[-_](\d{1,2})[-_](\d{4})"), "mdy"),
    (re.compile(r"(\d{4})[-_](\d{1,2})(?!\d)"),         "ym"),
    (re.compile(r"(\d{4})(\d{2})(\d{2})"),               "ymd"),
    (re.compile(r"(\d{4})(\d{2})(?!\d)"),                "ym"),
]


def guess_as_of_date(filename: str) -> str | None:
    """Best-effort ``YYYY-MM-DD`` extraction from a filename.

    Picks the last calendar day of the month when only year+month are
    present. Returns ``None`` if no date is recognisable.
    """
    name = Path(filename).stem
    for rx, kind in _DATE_RX_LIST:
        m = rx.search(name)
        if not m:
            continue
        try:
            if kind == "ymd":
                y, mo, d = int(m.group(1)), int(m.group(2)), int(m.group(3))
            elif kind == "mdy":
                mo, d, y = int(m.group(1)), int(m.group(2)), int(m.group(3))
            elif kind == "ym":
                y, mo = int(m.group(1)), int(m.group(2))
                d = _last_day(y, mo)
            else:
                continue
            return date(y, mo, d).isoformat()
        except (ValueError, TypeError):
            continue
    return None


def _last_day(year: int, month: int) -> int:
    import calendar
    return calendar.monthrange(year, month)[1]


def rollup_dataframe(
    df,
    mapping: dict[str, str],
    *,
    dq_threshold: int = DEFAULT_DQ_THRESHOLD,
) -> dict[str, Any]:
    """Aggregate ``df`` to one record per loan_pool_code.

    Returns::

        {
            "rows": [{loan_code, total_balance, dq_amount, dq_pct}, ...],
            "loan_count": int,
            "loan_count_dq": int,
            "total_balance": float,
            "total_dq": float,
            "warnings": [str, ...],
        }

    ``mapping`` must include ``loan_pool_code``, ``current_balance``,
    and ``days_delinquent``. Missing days_delinquent column is treated
    as a fatal error (no DQ data can be derived without it).
    """
    code_col = mapping.get("loan_pool_code")
    bal_col = mapping.get("current_balance")
    dq_col = mapping.get("days_delinquent")
    warnings: list[str] = []
    if not code_col or not bal_col:
        raise ValueError(
            "Column mappings missing 'loan_pool_code' and/or "
            "'current_balance'."
        )
    if not dq_col:
        raise ValueError(
            "Column mappings missing 'days_delinquent' — required to "
            "derive DQ history from loan-data extracts. Set it on the "
            "Column Mappings step."
        )
    for col in (code_col, bal_col, dq_col):
        if col not in df.columns:
            raise ValueError(f"Column '{col}' not found in file.")

    by_code: dict[str, dict[str, Any]] = {}
    loan_count = 0
    loan_count_dq = 0
    missing_dq_rows = 0
    for raw_code, raw_bal, raw_dq in zip(
        df[code_col], df[bal_col], df[dq_col], strict=False
    ):
        code = str(raw_code).strip()
        if not code:
            continue
        bal = extract_hist_processor._clean_balance(raw_bal)
        days = _coerce_days(raw_dq)
        rec = by_code.setdefault(
            code,
            {
                "loan_code": code,
                "total_balance": 0.0,
                "dq_amount": 0.0,
                "loan_count": 0,
                "loan_count_dq": 0,
            },
        )
        rec["total_balance"] += bal
        rec["loan_count"] += 1
        loan_count += 1
        if days is None:
            missing_dq_rows += 1
            continue
        if days >= dq_threshold:
            rec["dq_amount"] += bal
            rec["loan_count_dq"] += 1
            loan_count_dq += 1
    if missing_dq_rows and missing_dq_rows == loan_count:
        raise ValueError(
            f"Could not parse a numeric 'days_delinquent' value on any "
            f"of the {loan_count} loan rows. Check the column mapping."
        )
    if missing_dq_rows:
        warnings.append(
            f"{missing_dq_rows} loan rows had a blank/non-numeric "
            f"days_delinquent value and were treated as current."
        )

    rows: list[dict[str, Any]] = []
    total_balance = 0.0
    total_dq = 0.0
    for code, rec in by_code.items():
        tb = round(rec["total_balance"], 2)
        dq = round(rec["dq_amount"], 2)
        pct = (dq / tb) if tb else None
        rows.append({
            "loan_code": code,
            "total_balance": tb,
            "dq_amount": dq,
            "dq_pct": round(pct, 8) if pct is not None else None,
            "loan_count": rec["loan_count"],
            "loan_count_dq": rec["loan_count_dq"],
        })
        total_balance += tb
        total_dq += dq

    rows.sort(key=lambda r: r["loan_code"])
    return {
        "rows": rows,
        "loan_count": loan_count,
        "loan_count_dq": loan_count_dq,
        "total_balance": round(total_balance, 2),
        "total_dq": round(total_dq, 2),
        "warnings": warnings,
    }


def process_files(
    cu: str,
    files: list[dict[str, Any]],
    mapping: dict[str, str],
    *,
    dq_threshold: int = DEFAULT_DQ_THRESHOLD,
    source_tag: str = "loan_extract",
) -> dict[str, Any]:
    """Parse each file in ``files`` and upsert its rollup.

    ``files`` is a list of dicts with at least ``path``, and optionally
    ``as_of_date`` (otherwise inferred from filename) and ``name``.

    Returns a summary dict suitable for the wizard UI.
    """
    out: dict[str, Any] = {
        "ok": False,
        "error": None,
        "files": [],
        "rows_written": 0,
        "dates_written": [],
    }
    if not cu:
        out["error"] = "Credit union name not set on Identity step."
        return out
    if not files:
        out["error"] = "No loan-data extract files provided."
        return out

    try:
        delinquency_hist_processor.ensure_table()
    except Exception as exc:  # noqa: BLE001
        out["error"] = (
            f"Could not ensure loan_code_delinquency_history table: {exc}"
        )
        return out

    dates: list[str] = []
    for entry in files:
        path_str = entry.get("path") or ""
        name = entry.get("name") or Path(path_str).name
        as_of = entry.get("as_of_date") or guess_as_of_date(name)
        info: dict[str, Any] = {
            "name": name,
            "as_of_date": as_of,
            "ok": False,
            "error": None,
            "rows": 0,
            "total_balance": 0.0,
            "total_dq": 0.0,
            "dq_pct": None,
            "warnings": [],
        }
        out["files"].append(info)
        if not as_of:
            info["error"] = (
                "Could not detect an as_of_date from the filename. "
                "Rename to include YYYY-MM-DD or YYYY-MM."
            )
            continue
        try:
            df = extract_hist_processor._read_file(Path(path_str))
        except Exception as exc:  # noqa: BLE001
            info["error"] = f"Read error: {exc}"
            continue
        try:
            roll = rollup_dataframe(df, mapping, dq_threshold=dq_threshold)
        except Exception as exc:  # noqa: BLE001
            info["error"] = str(exc)
            continue
        info["rows"] = len(roll["rows"])
        info["total_balance"] = roll["total_balance"]
        info["total_dq"] = roll["total_dq"]
        info["dq_pct"] = (
            roll["total_dq"] / roll["total_balance"]
            if roll["total_balance"]
            else None
        )
        info["warnings"] = roll["warnings"]
        try:
            written = delinquency_hist_processor.upsert_month(
                cu, as_of, roll["rows"],
                source=f"{source_tag}:{name}",
            )
        except Exception as exc:  # noqa: BLE001
            info["error"] = f"DB upsert failed: {exc}"
            continue
        info["ok"] = True
        out["rows_written"] += written
        dates.append(as_of)
    out["dates_written"] = sorted(set(dates))
    out["ok"] = any(f["ok"] for f in out["files"])
    if not out["ok"] and not out["error"]:
        out["error"] = "No files were processed successfully."
    return out
