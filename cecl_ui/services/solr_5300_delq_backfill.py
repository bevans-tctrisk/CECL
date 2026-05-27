"""NCUA 5300 backfill for ``loan_code_delinquency_history``.

Parallel to ``solr_5300_co_backfill`` but for DELINQUENCY balances
instead of charge-offs. The key difference: DQ is reported as a
point-in-time snapshot, so we do NOT YTD-difference. Each quarter-end's
value is written as-is.

The canonical 5300-DQ → loan_code map lives at
``cecl_ui/data/ncua_5300_delinquency_codes.csv``. Rows with a blank
``field_code`` are silently skipped — populate them before relying on
this backfill (see ``ncua_5300_delinquency_codes.README.md``).
"""
from __future__ import annotations

import csv
from pathlib import Path
from typing import Any

from sqlalchemy import text as _sql_text

from cecl_ui.services import delinquency_hist_processor
from cecl_ui.services.solr_5300_backfill import (
    _coerce_number,
    expected_quarter_ends,
    fetch_solr_doc,
)


_CANONICAL_CSV = (
    Path(__file__).resolve().parent.parent
    / "data"
    / "ncua_5300_delinquency_codes.csv"
)


def canonical_map_path() -> Path:
    return _CANONICAL_CSV


def load_canonical_map(
    csv_path: Path | str | None = None,
) -> list[dict[str, str]]:
    """Read the canonical 5300-DQ → loan_code map from disk."""
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


def map_status() -> dict[str, Any]:
    """Report how many rows in the canonical CSV have field_codes set."""
    path = _CANONICAL_CSV
    total = 0
    mapped = 0
    if path.exists():
        with path.open("r", encoding="utf-8-sig", newline="") as fh:
            for row in csv.DictReader(fh):
                lc = (row.get("loan_code") or "").strip()
                fc = (row.get("field_code") or "").strip()
                if not lc:
                    continue
                total += 1
                if fc:
                    mapped += 1
    return {"total": total, "mapped": mapped, "path": str(path)}


def backfill_missing_delinquency_quarters(
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
    """Walk expected quarter-ends and upsert any missing DQ rows.

    Unlike charge-offs, DQ values are point-in-time snapshots — written
    as-is for each quarter-end with no YTD differencing.
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
        status = map_status()
        if status["total"] and not status["mapped"]:
            out["error"] = (
                f"Canonical 5300 delinquency map has {status['total']} "
                f"loan_code rows but no field_codes are populated. Edit "
                f"{status['path']} and fill in the field_code column "
                f"(see the sidecar README)."
            )
        else:
            out["error"] = (
                f"Canonical 5300 delinquency map is empty or missing: "
                f"{_CANONICAL_CSV}"
            )
        return out
    try:
        delinquency_hist_processor.ensure_table()
    except Exception as exc:  # noqa: BLE001
        out["error"] = (
            f"Could not ensure loan_code_delinquency_history table: {exc}"
        )
        return out

    # Cleanup: drop ALL rows previously written by this backfill for
    # this cu, regardless of date or loan_code. Scoped strictly to
    # ``source LIKE '5300DQ:%'`` so user-uploaded extracts (which use
    # a different source tag) are left untouched.
    try:
        eng = delinquency_hist_processor._engine_lazy()
        with eng.begin() as conn:
            res = conn.execute(
                _sql_text(
                    "DELETE FROM loan_code_delinquency_history "
                    "WHERE cu = :cu "
                    "AND source LIKE '5300DQ:%'"
                ),
                {"cu": cu},
            )
            out["stale_rows_removed"] = int(res.rowcount or 0)
    except Exception as exc:  # noqa: BLE001
        out["months_skipped"].append({
            "period": "(cleanup)",
            "reason": f"stale-row cleanup failed: {type(exc).__name__}: {exc}",
        })

    quarter_ends = expected_quarter_ends(target_period, history_months)
    out["expected"] = quarter_ends
    # Recompute existing AFTER the purge so dates we just cleared are
    # not skipped. Any survivors came from a different source.
    try:
        hv = delinquency_hist_processor.history_matrix(cu)
        existing = set((hv or {}).get("months") or [])
    except Exception:  # noqa: BLE001
        existing = set(existing_dates or [])
    seen_codes: set[str] = set()

    for qe in sorted(quarter_ends):
        if not overwrite and qe in existing:
            out["months_skipped"].append({
                "period": qe,
                "reason": "quarter-end already present",
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

        # Sum per loan_code (one loan_code may span multiple DQ buckets,
        # e.g. 60-179d + ≥180d — both get added).
        by_code: dict[str, float] = {}
        for entry in cmap:
            code = entry["loan_code"]
            fc = entry["field_code"]
            val = _coerce_number(doc.get(fc)) if fc in doc else 0.0
            by_code[code] = by_code.get(code, 0.0) + val
            seen_codes.add(code)

        rollup_rows = [
            {
                "loan_code": code,
                "dq_amount": round(amt, 2),
                "total_balance": None,
                "dq_pct": None,
            }
            for code, amt in by_code.items()
        ]
        try:
            written = delinquency_hist_processor.upsert_month(
                cu, qe, rollup_rows, source=f"5300DQ:{qe}",
            )
        except Exception as exc:  # noqa: BLE001
            out["months_skipped"].append({
                "period": qe,
                "reason": f"upsert error: {exc}",
            })
            continue
        out["months_filled"].append({
            "period": qe,
            "quarter_end": qe,
            "rows_written": written,
            "codes": len(rollup_rows),
            "note": "point-in-time",
        })
        out["rows_written"] += written
        existing.add(qe)

    out["new_loan_codes"] = sorted(seen_codes)
    out["ok"] = True
    return out
