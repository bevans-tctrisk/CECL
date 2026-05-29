"""Aggregate per-month Charge-Off / Recovery files into the historical DB.

The wizard's Historical step lets a user upload many monthly CO / Recovery
spreadsheets at once (one file per month).  Each file typically has a header
row + a few transaction rows like::

    Account #   Loan Type Code   Recovery Amount   Charge off Amount   Date
    28883L102   43               197.15            (blank)             ...
    49198L102   42                                 100                 ...

Some credit unions ship combined CO+Recovery files (both amount columns in
one workbook), others ship separate files.  This module reads each file,
sums per ``loan_code`` for the requested kind (``"co"`` or ``"recov"``), and
writes the result via ``chargeoff_hist_processor`` /
``recovery_hist_processor`` ``upsert_month``.

Files dated to the same ``as_of_date`` are merged before upsert so the
processor's "replace prior rows" behaviour doesn't drop earlier files of
the same month.
"""
from __future__ import annotations

import re
from collections import defaultdict
from pathlib import Path
from typing import Any

from cecl_ui.services import (
    chargeoff_hist_processor,
    extract_hist_service,
    recovery_hist_processor,
)


# Header-name keyword sets used to auto-locate columns when the file has a
# header row.  Lowered + stripped before matching.
_LOAN_CODE_HEADERS = (
    "loan type code", "loan_type_code", "loan code", "loan_code",
    "type code", "type_code", "pool code", "pool_code",
    "security code", "security_code",
    "loan type", "loan_type", "loan class", "loan_class",
    "product code", "product_code", "product type", "product_type",
)
_MEMBER_HEADERS = (
    "member number", "member_number", "member#", "member #", "member id",
    "member", "acct member", "member no",
)
_ACCOUNT_HEADERS = (
    "account number", "account_number", "account#", "account #",
    "loan number", "loan_number", "loan#", "loan #",
    "suffix", "account suffix", "loan suffix", "share suffix",
    "id",  # last-resort fallback
)
_CO_AMOUNT_HEADERS = (
    "charge off amount", "charge-off amount", "chargeoff amount",
    "charge off", "chargeoff", "chg off amount", "chg-off amount",
    "co amount", "co_amount",
)
_RECOV_AMOUNT_HEADERS = (
    "recovery amount", "recovery_amount", "recoveries amount",
    "recovery", "recoveries", "recov amount", "recov_amount",
)
_DATE_HEADERS = (
    "charge off date", "chargeoff date", "charge-off date",
    "recovery date", "recoveries date",
    "effective date", "transaction date", "posting date", "date",
)


def _read_rows(path: Path) -> list[list[Any]]:
    """Return all sheet rows as a list of lists. CSV or XLSX/XLSM/XLS.

    For multi-sheet workbooks, picks the worksheet with the most
    non-blank cells — prevents an empty leading sheet from masking
    the real data tab.
    """
    ext = path.suffix.lower()
    if ext == ".csv":
        import csv

        out: list[list[Any]] = []
        with open(path, newline="", encoding="utf-8-sig") as fh:
            for row in csv.reader(fh):
                out.append(list(row))
        return out

    from openpyxl import load_workbook

    wb = load_workbook(path, read_only=True, data_only=True)
    best_rows: list[list[Any]] = []
    best_score = -1
    for ws in wb.worksheets:
        sheet_rows: list[list[Any]] = []
        for r in ws.iter_rows(values_only=True):
            sheet_rows.append(list(r) if r else [])
        score = sum(
            1 for row in sheet_rows
            for c in row if c not in (None, "")
        )
        if score > best_score:
            best_score = score
            best_rows = sheet_rows
    return best_rows


def _norm(v: Any) -> str:
    if v is None:
        return ""
    return re.sub(r"\s+", " ", str(v).strip()).lower()


def _find_header_row(rows: list[list[Any]]) -> int | None:
    """Find the first row whose cells look like headers (contain a loan-code
    column AND at least one amount column)."""
    for i, row in enumerate(rows[:25]):  # don't scan forever
        cells = [_norm(c) for c in row]
        has_code = any(
            any(k in c for k in _LOAN_CODE_HEADERS) for c in cells
        )
        has_amount = any(
            any(k in c for k in (*_CO_AMOUNT_HEADERS, *_RECOV_AMOUNT_HEADERS))
            for c in cells
        )
        if has_code and has_amount:
            return i
    return None


def _pick_column(
    headers: list[str], keywords: tuple[str, ...]
) -> int | None:
    """Return the index of the first header that exactly equals or contains
    one of ``keywords``.  Prefers exact equality over substring match."""
    norm = [_norm(h) for h in headers]
    for kw in keywords:
        if kw in norm:
            return norm.index(kw)
    for kw in keywords:
        for i, h in enumerate(norm):
            if kw in h:
                return i
    return None


def _to_float(v: Any) -> float | None:
    """Coerce a cell to a float.  Returns None for blank/non-numeric."""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        f = float(v)
        if f != f:  # NaN
            return None
        return f
    s = str(v).strip()
    if not s:
        return None
    neg = s.startswith("(") and s.endswith(")")
    s = s.strip("()").replace("$", "").replace(",", "").strip()
    if not s:
        return None
    try:
        f = float(s)
    except ValueError:
        return None
    return -f if neg else f


def _loan_code_str(v: Any) -> str | None:
    """Normalise a loan-code cell to a short uppercase string.

    Strips trailing ``"/anything"`` suffixes (e.g. ``"VA/01"`` -> ``"VA"``)
    and converts numeric codes (43, 42) to ``"43"`` / ``"42"``.
    """
    if v is None or v == "":
        return None
    if isinstance(v, float) and v != v:  # NaN
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    if not s:
        return None
    if "/" in s:
        s = s.split("/", 1)[0].strip()
    return s.upper()


def _pick_column_by_name(headers: list[str], name: str) -> int | None:
    if not name:
        return None
    norm = [_norm(h) for h in headers]
    target = _norm(name)
    if target in norm:
        return norm.index(target)
    for i, h in enumerate(norm):
        if target and target in h:
            return i
    return None


def inspect_file(path: Path) -> dict[str, Any]:
    """Peek at a CO/Recov file and return its header row + suggested column
    matches for the column-mapping UI.

    Returns::

        {
          "ok": bool, "error": str | None,
          "filename": str,
          "headers": [str, ...],
          "suggested": {"code": str, "co_amount": str,
                        "recov_amount": str, "date": str},
          "preview_rows": [[...], ...],  # up to 5 rows after the header
        }
    """
    out: dict[str, Any] = {
        "ok": False, "error": None,
        "filename": path.name,
        "headers": [], "suggested": {}, "preview_rows": [],
    }
    try:
        raw = _read_rows(path)
    except Exception as exc:  # noqa: BLE001
        out["error"] = f"Could not read file: {exc}"
        return out
    if not raw:
        out["error"] = "File is empty."
        return out

    hdr_idx = _find_header_row(raw)
    if hdr_idx is None:
        # Fallback: the keyword heuristic missed (file uses non-standard
        # header text). Pick the row with the most non-blank string
        # cells inside the first 25 rows so the manual dropdowns at
        # least get useful options.
        best_i, best_score = -1, 0
        for i, row in enumerate(raw[:25]):
            score = sum(
                1 for c in row
                if isinstance(c, str) and c.strip()
            )
            if score > best_score:
                best_score, best_i = score, i
        if best_score >= 2:
            hdr_idx = best_i
            out["error"] = (
                "Couldn't auto-detect the header row from the usual "
                "keywords \u2014 used the row with the most text cells "
                f"(row {best_i + 1}) as a best guess. Confirm and pick "
                "the columns manually below."
            )
        else:
            out["error"] = (
                "Could not find a header row. Pick the columns "
                "manually below."
            )
            # Still expose the first row as a best-effort header list.
            out["headers"] = [
                str(c) if c is not None else ""
                for c in (raw[0] if raw else [])
            ]
            out["preview_rows"] = [list(r) for r in raw[1:6]]
            return out

    headers = [str(c) if c is not None else "" for c in raw[hdr_idx]]

    def _name_at(idx: int | None) -> str:
        if idx is None or idx >= len(headers):
            return ""
        return headers[idx]

    out["headers"] = headers
    out["preview_rows"] = [
        list(r) for r in raw[hdr_idx + 1: hdr_idx + 6]
    ]
    out["suggested"] = {
        "code": _name_at(_pick_column(headers, _LOAN_CODE_HEADERS)),
        "co_amount": _name_at(_pick_column(headers, _CO_AMOUNT_HEADERS)),
        "recov_amount": _name_at(
            _pick_column(headers, _RECOV_AMOUNT_HEADERS)
        ),
        "date": _name_at(_pick_column(headers, _DATE_HEADERS)),
        "member": _name_at(_pick_column(headers, _MEMBER_HEADERS)),
        "account": _name_at(_pick_column(headers, _ACCOUNT_HEADERS)),
    }
    out["ok"] = True
    return out


def parse_file(
    path: Path,
    kind: str,
    *,
    code_header: str | None = None,
    amount_header: str | None = None,
    date_header: str | None = None,
) -> dict[str, Any]:
    """Parse one monthly CO/Recov file.

    ``kind`` is ``"co"`` or ``"recov"``.  Optional ``code_header`` /
    ``amount_header`` force a specific header-name match (case/whitespace
    insensitive); when not supplied we fall back to keyword auto-detection.

    Returns::

        {
          "ok": bool, "error": str | None,
          "filename": str,
          "as_of_date": "YYYY-MM-DD" | "",
          "date_source": "filename" | "mtime" | "header" | "",
          "date_confidence": "high"|"medium"|"low"|"none",
          "rows": [{"loan_code": str, "amount": float}, ...],  # aggregated
          "total_rows": int,        # raw rows used
          "total_amount": float,    # sum across rows
          "empty_ok": bool,         # True when file parsed cleanly but had
                                    # zero non-blank amount rows
        }
    """
    out: dict[str, Any] = {
        "ok": False, "error": None,
        "filename": path.name,
        "as_of_date": "", "date_source": "", "date_confidence": "none",
        "rows": [], "total_rows": 0, "total_amount": 0.0,
        "empty_ok": False,
    }
    try:
        raw = _read_rows(path)
    except Exception as exc:  # noqa: BLE001
        out["error"] = f"Could not read file: {exc}"
        return out
    if not raw:
        out["error"] = "File is empty."
        return out

    hdr_idx = _find_header_row(raw)
    if hdr_idx is None:
        # Fallback to densest-text-row heuristic so files with
        # non-standard header text still parse when the user has
        # supplied an explicit column mapping.
        best_i, best_score = -1, 0
        for i, row in enumerate(raw[:25]):
            score = sum(
                1 for c in row
                if isinstance(c, str) and c.strip()
            )
            if score > best_score:
                best_score, best_i = score, i
        if best_score >= 2 and (code_header or amount_header):
            hdr_idx = best_i
        else:
            out["error"] = (
                "Could not find a header row with a Loan Code column and an "
                "amount column."
            )
            return out

    headers = [str(c) if c is not None else "" for c in raw[hdr_idx]]
    if code_header:
        code_col = _pick_column_by_name(headers, code_header)
    else:
        code_col = _pick_column(headers, _LOAN_CODE_HEADERS)
    if amount_header:
        amount_col = _pick_column_by_name(headers, amount_header)
    else:
        amount_keywords = (
            _CO_AMOUNT_HEADERS if kind == "co" else _RECOV_AMOUNT_HEADERS
        )
        amount_col = _pick_column(headers, amount_keywords)
    if date_header:
        date_col = _pick_column_by_name(headers, date_header)
    else:
        date_col = _pick_column(headers, _DATE_HEADERS)

    if code_col is None:
        if code_header:
            out["error"] = (
                f"Configured Loan Code column '{code_header}' was not found "
                f"in the header row."
            )
        else:
            out["error"] = (
                "Could not find a Loan Code column in the header row."
            )
        return out
    if amount_col is None:
        label = "Charge-Off" if kind == "co" else "Recovery"
        if amount_header:
            out["error"] = (
                f"Configured {label} Amount column '{amount_header}' was not "
                f"found in the header row."
            )
        else:
            out["error"] = (
                f"Could not find a {label} Amount column in the header row."
            )
        return out

    # Aggregate by loan code.
    agg: dict[str, float] = defaultdict(float)
    used = 0
    header_dates: list[Any] = []
    for row in raw[hdr_idx + 1:]:
        if not row:
            continue
        code = _loan_code_str(
            row[code_col] if code_col < len(row) else None
        )
        amt = _to_float(
            row[amount_col] if amount_col < len(row) else None
        )
        if code is None or amt is None or amt == 0:
            continue
        agg[code] += amt
        used += 1
        if date_col is not None and date_col < len(row):
            dv = row[date_col]
            if dv not in (None, ""):
                header_dates.append(dv)

    # Date detection.
    det = extract_hist_service.detect_as_of_date(path.name, path)
    out["as_of_date"] = det.get("date") or ""
    out["date_source"] = det.get("source") or ""
    out["date_confidence"] = det.get("confidence") or "none"

    # If filename detection was low/none, try the latest in-row date.
    if (out["date_confidence"] in ("low", "none")) and header_dates:
        latest = _latest_in_row_date(header_dates)
        if latest:
            out["as_of_date"] = latest
            out["date_source"] = "header"
            out["date_confidence"] = "medium"

    out["rows"] = [
        {"loan_code": k, "amount": round(v, 2)}
        for k, v in sorted(agg.items())
    ]
    out["total_rows"] = used
    out["total_amount"] = round(sum(agg.values()), 2)
    out["empty_ok"] = used == 0
    out["ok"] = True
    return out


def _latest_in_row_date(values: list[Any]) -> str | None:
    """Best-effort: scan a list of date-ish values and return the latest as
    YYYY-MM-DD."""
    from datetime import date, datetime

    latest: date | None = None
    for v in values:
        d: date | None = None
        if isinstance(v, datetime):
            d = v.date()
        elif isinstance(v, date):
            d = v
        else:
            s = str(v).strip()
            for fmt in ("%Y-%m-%d", "%m/%d/%Y", "%m-%d-%Y", "%Y/%m/%d"):
                try:
                    d = datetime.strptime(s, fmt).date()
                    break
                except ValueError:
                    continue
        if d is None:
            continue
        if latest is None or d > latest:
            latest = d
    return latest.isoformat() if latest else None


def aggregate_all(state: dict[str, Any], kind: str) -> dict[str, Any]:
    """Parse every monthly file in ``state.hist_scan`` for ``kind`` and
    upsert into the historical CO/Recov DB.

    ``kind`` is ``"co"`` (uses ``monthly_co_files``) or ``"recov"`` (uses
    ``monthly_recov_files``).
    """
    out: dict[str, Any] = {
        "ok": False, "error": None,
        "kind": kind,
        "files": [],          # per-file summary list
        "months_written": [], # distinct as_of_date values written
        "total_files": 0,
        "total_parsed": 0,
        "total_rows_written": 0,
        "total_amount": 0.0,
    }

    cu = (state.get("credit_union") or "").strip()
    if not cu:
        out["error"] = "Set the credit union on the Identity step first."
        return out

    hs = state.get("hist_scan") or {}
    key = "monthly_co_files" if kind == "co" else "monthly_recov_files"
    files = hs.get(key) or []
    out["total_files"] = len(files)
    if not files:
        out["error"] = "No monthly files have been uploaded yet."
        return out

    # Column overrides (from the wizard's Column Mapping card).
    col_key = "co_columns" if kind == "co" else "recov_columns"
    overrides = state.get(col_key) or {}
    code_header = (overrides.get("loan_code") or "").strip() or None
    amount_header = (overrides.get("amount") or "").strip() or None
    date_header = (overrides.get("date") or "").strip() or None

    # First pass: parse each file. ``parsed_per_date`` is a normal dict so
    # an empty-but-validated month registers a key with an empty inner dict,
    # which downstream upserts as a zero-row baseline for that month.
    parsed_per_date: dict[str, dict[str, float]] = {}
    per_date_sources: dict[str, list[str]] = defaultdict(list)
    for entry in files:
        name = entry.get("name") or ""
        path_str = entry.get("path") or ""
        item: dict[str, Any] = {
            "name": name, "ok": False, "error": None,
            "as_of_date": "", "rows": 0, "amount": 0.0,
            "empty": False,
        }
        if not path_str or not Path(path_str).exists():
            item["error"] = "Saved file is missing on disk."
            out["files"].append(item)
            continue
        res = parse_file(
            Path(path_str), kind,
            code_header=code_header, amount_header=amount_header,
            date_header=date_header,
        )
        item["as_of_date"] = res.get("as_of_date") or ""
        if not res.get("ok"):
            item["error"] = res.get("error") or "Parse failed."
            out["files"].append(item)
            continue
        if not item["as_of_date"]:
            item["error"] = (
                "Could not determine an as-of date for this file. Rename "
                "the file to include a date (e.g. MMDDYYYY)."
            )
            out["files"].append(item)
            continue
        item["ok"] = True
        item["rows"] = res["total_rows"]
        item["amount"] = res["total_amount"]
        item["empty"] = bool(res.get("empty_ok"))
        out["total_parsed"] += 1
        out["files"].append(item)
        per_date_sources[item["as_of_date"]].append(name)
        # Register the date even if no rows so an empty file still serves
        # as a baseline marker for the month.
        bucket = parsed_per_date.setdefault(item["as_of_date"], {})
        for r in res["rows"]:
            bucket[r["loan_code"]] = bucket.get(r["loan_code"], 0.0) + r["amount"]

    if not parsed_per_date:
        out["error"] = "Nothing to aggregate \u2014 see per-file errors."
        return out

    # Second pass: upsert each (cu, as_of_date) bundle.
    field = "chargeoff_amount" if kind == "co" else "recovery_amount"
    processor = (
        chargeoff_hist_processor if kind == "co"
        else recovery_hist_processor
    )
    write_errors: list[str] = []
    try:
        processor.ensure_table()
    except Exception as exc:  # noqa: BLE001
        out["error"] = f"Could not prepare DB table: {exc}"
        return out

    for as_of, code_to_amt in sorted(parsed_per_date.items()):
        rows = [
            {"loan_code": code, field: round(amt, 2)}
            for code, amt in sorted(code_to_amt.items())
        ]
        src = "monthly:" + ",".join(per_date_sources[as_of])
        try:
            n = processor.upsert_month(cu, as_of, rows, src)
        except Exception as exc:  # noqa: BLE001
            write_errors.append(f"{as_of}: {exc}")
            continue
        out["months_written"].append(as_of)
        out["total_rows_written"] += n
        out["total_amount"] += sum(code_to_amt.values())

    if write_errors and not out["months_written"]:
        out["error"] = "Database write failed: " + "; ".join(write_errors)
        return out
    if write_errors:
        out["error"] = "Some months failed: " + "; ".join(write_errors)
    out["ok"] = bool(out["months_written"])
    out["total_amount"] = round(out["total_amount"], 2)
    return out
