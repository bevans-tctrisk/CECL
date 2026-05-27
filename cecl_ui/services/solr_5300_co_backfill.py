"""NCUA 5300 backfill for ``loan_code_chargeoff_history``.

Parallel to ``solr_5300_backfill`` (which fills loan *balances*). Pulls
year-to-date net-charge-off totals for each quarter-end inside the
lookback window and writes one ``loan_code_chargeoff_history`` row per
canonical 5300 charge-off field code.

The canonical 5300-CO → loan_code map lives at
``cecl_ui/data/ncua_5300_chargeoff_codes.csv``.

YTD DIFFERENCING:
    NCUA 5300 reports charge-offs year-to-date (cumulative since Jan 1).
    To get per-quarter (and per-month) amounts we difference each
    quarter-end's YTD value against the running YTD for prior quarters
    in the same calendar year::

        March  (Q1) = Q1_YTD
        June   (Q2) = Q2_YTD - Q1_YTD
        Sept   (Q3) = Q3_YTD - Q2_YTD
        Dec    (Q4) = Q4_YTD - Q3_YTD

    The differenced amount is written ONLY at the quarter-end date.
    Non-quarter month-ends (Jan/Feb/Apr/May/Jul/Aug/Oct/Nov) are left
    untouched — absence of a row is treated as zero charge-offs by
    downstream consumers (which sum over date ranges).
"""
from __future__ import annotations

import csv
from pathlib import Path
from typing import Any

from cecl_ui.services import chargeoff_hist_processor
from cecl_ui.services.solr_5300_backfill import (
    _coerce_number,
    expected_quarter_ends,
    fetch_solr_doc,
)


_CANONICAL_CSV = (
    Path(__file__).resolve().parent.parent
    / "data"
    / "ncua_5300_chargeoff_codes.csv"
)


def canonical_map_path() -> Path:
    return _CANONICAL_CSV


def load_canonical_map(csv_path: Path | str | None = None) -> list[dict[str, str]]:
    """Read the canonical 5300-CO → loan_code map from disk."""
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


def backfill_missing_chargeoff_quarters(
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
    """Walk expected quarter-ends and upsert any missing CO rows."""
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
            f"Canonical 5300 charge-off map is empty or missing: "
            f"{_CANONICAL_CSV}"
        )
        return out
    try:
        chargeoff_hist_processor.ensure_table()
    except Exception as exc:  # noqa: BLE001
        out["error"] = (
            f"Could not ensure loan_code_chargeoff_history table: {exc}"
        )
        return out

    # Cleanup: drop ALL rows previously written by this backfill for
    # this cu, regardless of date or loan_code. Scoped strictly to
    # ``source LIKE '5300CO:%'`` so user-uploaded CO extracts (which
    # use a different source tag) are left untouched. This lets us
    # safely re-run the backfill and pick up logic changes (e.g. the
    # YTD-differencing switch that no longer writes to non-quarter-end
    # month-ends), and ensures stale orphan rows from prior runs are
    # purged.
    try:
        from sqlalchemy import text as _sql_text
        eng = chargeoff_hist_processor._engine_lazy()
        with eng.begin() as conn:
            res = conn.execute(
                _sql_text(
                    "DELETE FROM loan_code_chargeoff_history "
                    "WHERE cu = :cu "
                    "AND source LIKE '5300CO:%'"
                ),
                {"cu": cu},
            )
            out["stale_rows_removed"] = int(res.rowcount or 0)
    except Exception as exc:  # noqa: BLE001
        out["months_skipped"].append({
            "period": "(cleanup)",
            "reason": f"stale-row cleanup failed: {type(exc).__name__}: {exc}",
        })

    # Walk quarter-ends OLDEST -> NEWEST so we can maintain a running
    # YTD total per year/loan_code and difference each quarter against
    # the prior YTD (Q2_only = Q2_YTD - Q1_YTD, etc.).
    quarter_ends = expected_quarter_ends(target_period, history_months)
    out["expected"] = quarter_ends
    quarter_ends_chrono = sorted(quarter_ends)
    # Recompute ``existing`` from the DB AFTER the purge above so the
    # caller's pre-purge snapshot doesn't make us incorrectly skip
    # quarter-ends we just cleared. Any dates still listed here came
    # from a non-5300CO source (e.g. user-uploaded extract) and should
    # be preserved.
    try:
        hv = chargeoff_hist_processor.history_matrix(cu)
        existing = set((hv or {}).get("months") or [])
    except Exception:  # noqa: BLE001
        existing = set(existing_dates or [])
    seen_codes: set[str] = set()
    # {year: {loan_code: running_ytd_total_through_prior_quarter}}
    prior_ytd: dict[int, dict[str, float]] = {}

    for qe in quarter_ends_chrono:
        try:
            year = int(qe[:4])
        except (TypeError, ValueError):
            out["months_skipped"].append({
                "period": qe,
                "reason": "malformed quarter-end date",
            })
            continue
        year_prior = prior_ytd.setdefault(year, {})

        # Always fetch — even when we won't write — so the YTD chain
        # stays accurate for later quarters in the same year.
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

        bucket_ytd: dict[str, float] = {}
        for entry in cmap:
            code = entry["loan_code"]
            fc = entry["field_code"]
            val = _coerce_number(doc.get(fc)) if fc in doc else 0.0
            bucket_ytd[code] = bucket_ytd.get(code, 0.0) + val
            seen_codes.add(code)

        # Difference against prior-quarter YTD (within this calendar
        # year) to get the quarter-only net charge-off amount.
        rollup_rows: list[dict[str, Any]] = []
        for code, ytd in bucket_ytd.items():
            prior = year_prior.get(code, 0.0)
            quarter_only = round(ytd - prior, 2)
            rollup_rows.append({
                "loan_code": code,
                "chargeoff_amount": quarter_only,
            })
            # Update running YTD regardless of whether we write below,
            # so subsequent quarters difference correctly.
            year_prior[code] = ytd

        if not overwrite and qe in existing:
            out["months_skipped"].append({
                "period": qe,
                "reason": "quarter-end already present (YTD chain updated)",
            })
            continue

        try:
            written = chargeoff_hist_processor.upsert_month(
                cu, qe, rollup_rows, source=f"5300CO:{qe}",
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
            "note": "quarter-only (YTD differenced)",
        })
        out["rows_written"] += written
        existing.add(qe)

    out["new_loan_codes"] = sorted(seen_codes)
    out["ok"] = True
    return out
