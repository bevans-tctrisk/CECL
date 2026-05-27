"""NCUA 5300 backfill for ``loan_code_history``.

Fills *quarter-end* months that are missing from a CU's history table by
pulling 5300 call-report data from the TCT NCUA Solr core and writing
one ``loan_code_history`` row per canonical 5300 field code.

The canonical 5300 → loan_code map lives in a static CSV at
``cecl_ui/data/ncua_5300_loan_codes.csv``. Each row defines one
``loan_code`` (a human-readable bucket like "New Vehicles") and the
5300 ``field_code`` whose value populates it. Per-CU pool assignment
happens later in the wizard's Loan Code Mapping step (``step4_pools``);
this service just seeds the rows.

This is one rung up from the "manual fill" fallback: it gets users
quarter-end totals automatically, but it can't reconstruct intra-quarter
monthly snapshots or FICO histograms (those are stored empty).
"""
from __future__ import annotations

import calendar
import csv
import re
from datetime import date
from pathlib import Path
from typing import Any

try:
    import requests
except ImportError:  # pragma: no cover
    requests = None  # type: ignore

from cecl_ui.services import extract_hist_processor


_CANONICAL_CSV = (
    Path(__file__).resolve().parent.parent / "data" / "ncua_5300_loan_codes.csv"
)
_QUARTER_END_DAY = {3: 31, 6: 30, 9: 30, 12: 31}


def canonical_map_path() -> Path:
    return _CANONICAL_CSV


def load_canonical_map(csv_path: Path | str | None = None) -> list[dict[str, str]]:
    """Read the canonical 5300 → loan_code map from disk.

    Returns a list of ``{field_code, loan_code, description}`` dicts in
    file order. Missing file returns an empty list.
    """
    path = Path(csv_path) if csv_path else _CANONICAL_CSV
    if not path.exists():
        return []
    out: list[dict[str, str]] = []
    with path.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for row in reader:
            fc = (row.get("field_code") or "").strip()
            lc = (row.get("loan_code") or "").strip()
            if not fc or not lc:
                continue
            out.append({
                "field_code": fc,
                "loan_code": lc,
                "description": (row.get("description") or "").strip(),
            })
    return out


def _parse_yyyy_mm_dd(s: str) -> date | None:
    if not s:
        return None
    try:
        y, m, d = s.split("-")
        return date(int(y), int(m), int(d))
    except Exception:  # noqa: BLE001
        return None


def _quarter_end_date(year: int, month: int) -> date:
    return date(year, month, _QUARTER_END_DAY[month])


def _month_end(year: int, month: int) -> date:
    return date(year, month, calendar.monthrange(year, month)[1])


def quarter_fill_dates(qe_iso: str) -> list[str]:
    """Return the 3 month-end ISO dates a 5300 quarter should populate.

    e.g. ``"2026-03-31"`` → ``["2026-01-31", "2026-02-28", "2026-03-31"]``.
    Since 5300 reports are quarterly, the user has asked us to copy the
    quarter-end balance to the two prior month-ends (loan balances only —
    charge-offs and recoveries will be handled differently later).
    """
    d = _parse_yyyy_mm_dd(qe_iso)
    if not d or d.month not in _QUARTER_END_DAY:
        return [qe_iso]
    m1_year, m1_month = d.year, d.month - 2
    m2_year, m2_month = d.year, d.month - 1
    if m1_month <= 0:
        m1_month += 12
        m1_year -= 1
    if m2_month <= 0:
        m2_month += 12
        m2_year -= 1
    return [
        _month_end(m1_year, m1_month).isoformat(),
        _month_end(m2_year, m2_month).isoformat(),
        d.isoformat(),
    ]


def expected_quarter_ends(
    target_period: str, history_months: int
) -> list[str]:
    """Return quarter-end ISO dates inside the lookback window.

    ``target_period`` is the report month (any day in it works, but the
    wizard saves a YYYY-MM-DD).  ``history_months`` is the lookback in
    months (e.g. 84 = 7 years).  Result is newest-first.
    """
    tp = _parse_yyyy_mm_dd(target_period)
    if not tp:
        return []
    # Walk back month-by-month and collect quarter ends.
    out: list[str] = []
    y, m = tp.year, tp.month
    # Step back history_months from target_period.
    for _ in range(max(1, int(history_months)) + 1):
        if m in (3, 6, 9, 12):
            qe = _quarter_end_date(y, m)
            if qe <= tp:
                out.append(qe.isoformat())
        # decrement
        m -= 1
        if m == 0:
            m = 12
            y -= 1
    # Dedup + sort newest-first
    return sorted(set(out), reverse=True)


def quarter_date_string_for_charterdate(charter: int, period_iso: str) -> str:
    """Solr ``charterdate`` composite key for a given quarter end."""
    d = _parse_yyyy_mm_dd(period_iso)
    if not d:
        raise ValueError(f"Bad period: {period_iso!r}")
    return f"{charter}-{d.month:02d}/{d.day:02d}/{d.year}"


def fetch_solr_doc(
    solr_url: str,
    core: str,
    charter: int,
    period_iso: str,
    *,
    charter_field: str = "charter",
    charterdate_field: str = "charterdate",
    username: str | None = None,
    password: str | None = None,
    timeout: int = 20,
) -> dict[str, Any] | None:
    """Return the matching Solr doc or ``None`` if no doc found.

    Raises on transport / HTTP errors.
    """
    if requests is None:
        raise RuntimeError("'requests' library not installed in venv.")
    target = quarter_date_string_for_charterdate(charter, period_iso)
    base = solr_url.rstrip("/")
    url = f"{base}/{core}/select"
    q = f'{charter_field}:{charter} AND {charterdate_field}:"{target}"'
    params = {"q": q, "rows": 1, "wt": "json"}
    auth = (username, password) if username and password else None
    r = requests.get(url, params=params, auth=auth, timeout=timeout)
    r.raise_for_status()
    docs = (r.json().get("response") or {}).get("docs") or []
    return docs[0] if docs else None


def _coerce_number(value: Any) -> float:
    if value is None:
        return 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, list):
        # Solr multi-valued fields — take first numeric.
        for v in value:
            n = _coerce_number(v)
            if n:
                return n
        return 0.0
    s = str(value).strip()
    if not s:
        return 0.0
    neg = s.startswith("(") and s.endswith(")")
    s = s.replace("(", "").replace(")", "")
    s = s.replace("$", "").replace(",", "").replace(" ", "")
    if s.endswith("%"):
        s = s[:-1]
    if s in ("", "-"):
        return 0.0
    try:
        n = float(s)
    except ValueError:
        return 0.0
    return -n if neg else n


_FIELD_SPLIT = re.compile(r"[,\s;|]+")


def parse_field_list(raw: str) -> list[str]:
    """Split a user-entered field list (commas / spaces / etc.)."""
    if not raw:
        return []
    return [tok for tok in _FIELD_SPLIT.split(raw.strip()) if tok]


def test_connection(
    solr_url: str,
    core: str,
    charter: int,
    period_iso: str,
    *,
    username: str | None = None,
    password: str | None = None,
) -> dict[str, Any]:
    """Quick sanity check for the UI.

    Returns a dict the template can render directly.  Never raises.
    """
    out: dict[str, Any] = {
        "ok": False,
        "error": None,
        "period": period_iso,
        "charter": charter,
        "field_count": 0,
        "doc_keys_sample": [],
    }
    try:
        doc = fetch_solr_doc(
            solr_url, core, int(charter), period_iso,
            username=username, password=password,
        )
    except Exception as exc:  # noqa: BLE001
        out["error"] = f"{type(exc).__name__}: {exc}"
        return out
    if doc is None:
        out["error"] = (
            f"No Solr doc found for charter {charter} at {period_iso}."
        )
        return out
    keys = [k for k in doc.keys() if not k.startswith("_")]
    # Prefer keys that look like 5300 account codes (start with a letter).
    code_keys = sorted([k for k in keys if re.match(r"^[A-Za-z]\w*$", k)])
    out["ok"] = True
    out["field_count"] = len(keys)
    out["doc_keys_sample"] = code_keys[:40]
    return out


def backfill_missing_quarters(
    cu: str,
    charter: int,
    solr_url: str,
    core: str,
    target_period: str,
    history_months: int,
    *,
    canonical_map: list[dict[str, str]] | None = None,
    existing_dates: set[str] | None = None,
    overwrite: bool = False,
    username: str | None = None,
    password: str | None = None,
) -> dict[str, Any]:
    """Walk expected quarter-ends and upsert any that are missing.

    Uses the canonical 5300 → loan_code map (loaded from the static CSV
    if not supplied). Each 5300 field code becomes one row in
    ``loan_code_history`` keyed by its canonical ``loan_code`` label.
    """
    out: dict[str, Any] = {
        "ok": False,
        "error": None,
        "expected": [],
        "months_filled": [],
        "months_skipped": [],
        "months_no_data": [],
        "rows_written": 0,
        "stale_rows_removed": 0,
        "new_loan_codes": [],
    }
    if not cu:
        out["error"] = "Credit union name not set on Identity step."
        return out
    if not charter:
        out["error"] = "Charter number not set on Identity step."
        return out
    cmap = canonical_map if canonical_map is not None else load_canonical_map()
    if not cmap:
        out["error"] = (
            f"Canonical 5300 map is empty or missing: {_CANONICAL_CSV}"
        )
        return out
    try:
        extract_hist_processor.ensure_table()
    except Exception as exc:  # noqa: BLE001
        out["error"] = f"Could not ensure loan_code_history table: {exc}"
        return out

    # Cleanup: drop rows previously written by this backfill whose
    # loan_code is no longer in the canonical map. Scoped to source
    # LIKE '5300:%' so user-uploaded historical extracts are untouched.
    try:
        from sqlalchemy import text as _sql_text
        eng = extract_hist_processor._engine_lazy()
        canonical_codes = sorted({e["loan_code"] for e in cmap})
        with eng.begin() as conn:
            res = conn.execute(
                _sql_text(
                    "DELETE FROM loan_code_history "
                    "WHERE cu = :cu "
                    "AND source LIKE '5300:%' "
                    "AND loan_code <> ALL(:codes)"
                ),
                {"cu": cu, "codes": canonical_codes},
            )
            out["stale_rows_removed"] = int(res.rowcount or 0)
    except Exception as exc:  # noqa: BLE001
        # Non-fatal — log to skipped list and continue.
        out["months_skipped"].append({
            "period": "(cleanup)",
            "reason": f"stale-row cleanup failed: {type(exc).__name__}: {exc}",
        })

    quarter_ends = expected_quarter_ends(target_period, history_months)
    out["expected"] = quarter_ends
    existing = set(existing_dates or [])
    seen_codes: set[str] = set()

    for qe in quarter_ends:
        fill_dates = quarter_fill_dates(qe)
        # Decide which of the (up to 3) month-ends still need filling.
        if overwrite:
            targets = list(fill_dates)
        else:
            targets = [d for d in fill_dates if d not in existing]
        if not targets:
            out["months_skipped"].append({
                "period": qe,
                "reason": "quarter already covered (all 3 months present)",
            })
            continue
        try:
            doc = fetch_solr_doc(
                solr_url, core, int(charter), qe,
                username=username, password=password,
            )
        except Exception as exc:  # noqa: BLE001
            out["months_skipped"].append({
                "period": qe,
                "reason": f"fetch error: {type(exc).__name__}: {exc}",
            })
            continue
        if doc is None:
            out["months_no_data"].append(qe)
            continue
        # Build rollup rows from canonical map. One row per loan_code
        # bucket, summing the value(s) of its 5300 field code(s).
        # If multiple canonical rows share the same loan_code, their
        # field values get summed into a single bucket.
        bucket_totals: dict[str, float] = {}
        for entry in cmap:
            code = entry["loan_code"]
            fc = entry["field_code"]
            val = _coerce_number(doc.get(fc)) if fc in doc else 0.0
            bucket_totals[code] = bucket_totals.get(code, 0.0) + val
            seen_codes.add(code)
        rollup_rows = [
            {
                "loan_code": code,
                "total_balance": round(total, 2),
                "loan_count": 0,
                "fico_histogram": {},
                "fico_balance": {},
            }
            for code, total in bucket_totals.items()
        ]
        # Write the same balances to each missing month-end inside the
        # quarter (loan balances only — the source label notes which
        # quarter the data was copied from).
        for as_of in targets:
            label = "quarter-end" if as_of == qe else "copied from quarter-end"
            try:
                written = extract_hist_processor.upsert_month(
                    cu, as_of, rollup_rows, source=f"5300:{qe}",
                )
            except Exception as exc:  # noqa: BLE001
                out["months_skipped"].append({
                    "period": as_of,
                    "reason": f"upsert error: {exc}",
                })
                continue
            out["months_filled"].append({
                "period": as_of,
                "quarter_end": qe,
                "rows_written": written,
                "codes": len(rollup_rows),
                "note": label,
            })
            out["rows_written"] += written
            existing.add(as_of)

    out["new_loan_codes"] = sorted(seen_codes)
    out["ok"] = True
    return out
