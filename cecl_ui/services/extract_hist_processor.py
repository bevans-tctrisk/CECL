"""Roll up monthly loan-data extracts into the ``loan_code_history`` table.

Slice 3 of the no-WARM Historical Data flow.  Reads each in-window file
using its column profile's mapping, aggregates per raw ``loan_pool_code``
(here treated as the loan-code grain), bins original FICO scores into
10-point buckets so credit-grade bands can be re-derived later without
having to rescan, and upserts a row per (cu, as_of_date, loan_code).
"""
from __future__ import annotations

import json
from datetime import datetime
from pathlib import Path
from typing import Any

import pandas as pd
from sqlalchemy import text

from cecl_credentials import get_database_url


# ---------------------------------------------------------------------------
# DB helpers
# ---------------------------------------------------------------------------

_DDL = """
CREATE TABLE IF NOT EXISTS loan_code_history (
    cu              TEXT        NOT NULL,
    as_of_date      DATE        NOT NULL,
    loan_code       TEXT        NOT NULL,
    total_balance   NUMERIC(20, 2) NOT NULL DEFAULT 0,
    loan_count      INTEGER     NOT NULL DEFAULT 0,
    fico_histogram  JSONB,
    fico_balance    JSONB,
    source          TEXT,
    updated_at      TIMESTAMP   NOT NULL DEFAULT NOW(),
    PRIMARY KEY (cu, as_of_date, loan_code)
);
"""

_engine = None


def _engine_lazy():
    global _engine
    if _engine is None:
        from sqlalchemy import create_engine
        _engine = create_engine(get_database_url())
    return _engine


def ensure_table() -> None:
    eng = _engine_lazy()
    with eng.begin() as conn:
        conn.execute(text(_DDL))


# ---------------------------------------------------------------------------
# File reading + roll-up
# ---------------------------------------------------------------------------

def _read_file(path: Path) -> pd.DataFrame:
    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path, dtype=str, keep_default_na=False)
    # Excel — first visible sheet by default; keep everything as string so
    # we don't lose leading zeros on loan codes.
    return pd.read_excel(path, sheet_name=0, dtype=str, keep_default_na=False)


def _clean_balance(val: Any) -> float:
    if val is None:
        return 0.0
    s = str(val).strip()
    if not s:
        return 0.0
    neg = s.startswith("(") and s.endswith(")")
    s = s.replace("(", "").replace(")", "").replace("$", "").replace(",", "")
    s = s.replace(" ", "")
    if not s or s == "-":
        return 0.0
    try:
        v = float(s)
    except ValueError:
        return 0.0
    return -v if neg else v


def _clean_fico(val: Any) -> int | None:
    if val is None:
        return None
    s = str(val).strip()
    if not s:
        return None
    try:
        n = int(float(s))
    except ValueError:
        return None
    if n <= 0 or n < 300 or n > 900:
        return None
    return n


def _bucket(score: int | None) -> str:
    if score is None:
        return "no_score"
    return str((score // 10) * 10)


def rollup_dataframe(df: pd.DataFrame, mapping: dict[str, str]) -> list[dict[str, Any]]:
    """Aggregate ``df`` to one row per loan_code.

    ``mapping`` must include ``loan_pool_code``, ``current_balance``, and
    (optionally) ``original_fico_score``.  Other fields are ignored at this
    stage of the pipeline.
    """
    code_col = mapping.get("loan_pool_code")
    bal_col = mapping.get("current_balance")
    fico_col = mapping.get("original_fico_score")
    if not code_col or not bal_col:
        raise ValueError(
            "Profile is missing required column mappings "
            "(need at minimum loan_pool_code and current_balance)."
        )
    if code_col not in df.columns:
        raise ValueError(f"Column '{code_col}' (loan_pool_code) not in file.")
    if bal_col not in df.columns:
        raise ValueError(f"Column '{bal_col}' (current_balance) not in file.")
    has_fico = bool(fico_col) and fico_col in df.columns

    by_code: dict[str, dict[str, Any]] = {}
    for raw_code, raw_bal, raw_fico in zip(
        df[code_col],
        df[bal_col],
        df[fico_col] if has_fico else [None] * len(df),
        strict=False,
    ):
        code = str(raw_code).strip()
        if not code:
            continue
        bal = _clean_balance(raw_bal)
        score = _clean_fico(raw_fico) if has_fico else None
        bkt = _bucket(score)
        rec = by_code.setdefault(
            code,
            {
                "loan_code": code,
                "total_balance": 0.0,
                "loan_count": 0,
                "fico_histogram": {},
                "fico_balance": {},
            },
        )
        rec["total_balance"] += bal
        rec["loan_count"] += 1
        rec["fico_histogram"][bkt] = rec["fico_histogram"].get(bkt, 0) + 1
        rec["fico_balance"][bkt] = (
            rec["fico_balance"].get(bkt, 0.0) + bal
        )
    # Round balances for storage.
    for rec in by_code.values():
        rec["total_balance"] = round(rec["total_balance"], 2)
        rec["fico_balance"] = {
            k: round(v, 2) for k, v in rec["fico_balance"].items()
        }
    return list(by_code.values())


# ---------------------------------------------------------------------------
# Upsert
# ---------------------------------------------------------------------------

_UPSERT = text("""
    INSERT INTO loan_code_history
        (cu, as_of_date, loan_code, total_balance, loan_count,
         fico_histogram, fico_balance, source, updated_at)
    VALUES
        (:cu, :as_of_date, :loan_code, :total_balance, :loan_count,
         CAST(:fico_histogram AS JSONB), CAST(:fico_balance AS JSONB),
         :source, NOW())
    ON CONFLICT (cu, as_of_date, loan_code) DO UPDATE SET
        total_balance  = EXCLUDED.total_balance,
        loan_count     = EXCLUDED.loan_count,
        fico_histogram = EXCLUDED.fico_histogram,
        fico_balance   = EXCLUDED.fico_balance,
        source         = EXCLUDED.source,
        updated_at     = NOW();
""")

_DELETE_MONTH = text("""
    DELETE FROM loan_code_history
    WHERE cu = :cu AND as_of_date = :as_of_date;
""")


def upsert_month(
    cu: str, as_of_date: str, rollup_rows: list[dict[str, Any]], source: str
) -> int:
    """Replace prior rows for (cu, as_of_date) with ``rollup_rows``."""
    eng = _engine_lazy()
    with eng.begin() as conn:
        conn.execute(_DELETE_MONTH, {"cu": cu, "as_of_date": as_of_date})
        if not rollup_rows:
            return 0
        params = [
            {
                "cu": cu,
                "as_of_date": as_of_date,
                "loan_code": r["loan_code"],
                "total_balance": r["total_balance"],
                "loan_count": r["loan_count"],
                "fico_histogram": json.dumps(r["fico_histogram"]),
                "fico_balance": json.dumps(r["fico_balance"]),
                "source": source,
            }
            for r in rollup_rows
        ]
        conn.execute(_UPSERT, params)
        return len(params)


# ---------------------------------------------------------------------------
# Top-level driver
# ---------------------------------------------------------------------------

def process_scan(
    cu: str,
    scan: dict[str, Any],
    profiles: list[dict[str, Any]],
    anchor_files: list[dict[str, Any]] | None = None,
    ignored_signatures: list[str] | None = None,
) -> dict[str, Any]:
    """Process every in-window file from a scan result.

    Picks the first file per month when multiple are present.  Returns a
    summary suitable for stashing onto ``state.hist_extracts.scan_results``.
    """
    out: dict[str, Any] = {
        "started_at": datetime.utcnow().isoformat(timespec="seconds") + "Z",
        "ok": False,
        "error": None,
        "months_processed": [],
        "months_skipped": [],
        "loan_codes_total": 0,
        "rows_written": 0,
    }
    if not cu:
        out["error"] = "Credit union name not set on Identity step."
        return out

    profiles_by_id = {p.get("id"): p for p in (profiles or [])}
    anchors_by_name = {
        e.get("name"): e for e in (anchor_files or [])
    }
    ignored_set = set(ignored_signatures or [])

    try:
        ensure_table()
    except Exception as exc:  # noqa: BLE001
        out["error"] = f"Could not ensure loan_code_history table: {exc}"
        return out

    files = scan.get("files") or []
    months = scan.get("months") or {}
    expected = scan.get("expected_months") or []
    in_window_by_period: dict[str, list[dict[str, Any]]] = {}
    for f in files:
        if not f.get("in_window"):
            continue
        # Skip files whose layout the user explicitly ignored.
        if (f.get("signature") or "") in ignored_set:
            continue
        in_window_by_period.setdefault(f.get("period"), []).append(f)

    all_codes: set[str] = set()
    for ym in expected:
        bucket = in_window_by_period.get(ym) or []
        if not bucket:
            out["months_skipped"].append({"period": ym, "reason": "missing"})
            continue
        # Process EVERY mapped file in the bucket and merge their rollups,
        # so a month covered by multiple files (e.g. AIRES + a separate
        # mortgage extract sharing the same period) doesn't lose codes
        # from whichever file isn't picked first.
        mapped = [f for f in bucket if f.get("profile_id")]
        if not mapped:
            # Fall back to first file (will be flagged below if no profile
            # resolves via anchor lookup either).
            mapped = [bucket[0]]
        merged: dict[str, dict[str, Any]] = {}
        chosen_names: list[str] = []
        as_of = ""
        per_file_errors: list[str] = []
        for chosen in mapped:
            prof_id = chosen.get("profile_id")
            if not prof_id:
                anchor = anchors_by_name.get(chosen.get("name"))
                if anchor:
                    prof_id = anchor.get("profile_id")
            profile = profiles_by_id.get(prof_id) if prof_id else None
            if not profile:
                per_file_errors.append(
                    f"{chosen.get('name')}: no column profile"
                )
                continue
            mapping = profile.get("column_mappings") or {}
            if (not mapping.get("loan_pool_code")
                    or not mapping.get("current_balance")):
                per_file_errors.append(
                    f"{chosen.get('name')}: profile "
                    f"{profile.get('label') or prof_id} missing column mapping"
                )
                continue
            path = Path(chosen.get("path") or "")
            file_as_of = chosen.get("detected_date") or ""
            if file_as_of and not as_of:
                as_of = file_as_of
            try:
                df = _read_file(path)
                rollup = rollup_dataframe(df, mapping)
            except Exception as exc:  # noqa: BLE001
                per_file_errors.append(f"{chosen.get('name')}: {exc}")
                continue
            chosen_names.append(path.name)
            # Merge rollup rows into ``merged`` by loan_code: sum balances /
            # counts, additively merge fico_histogram + fico_balance.
            for r in rollup:
                code = r["loan_code"]
                acc = merged.setdefault(code, {
                    "loan_code": code,
                    "total_balance": 0.0,
                    "loan_count": 0,
                    "fico_histogram": {},
                    "fico_balance": {},
                })
                acc["total_balance"] += r["total_balance"]
                acc["loan_count"] += r["loan_count"]
                for k, v in (r.get("fico_histogram") or {}).items():
                    acc["fico_histogram"][k] = (
                        acc["fico_histogram"].get(k, 0) + v
                    )
                for k, v in (r.get("fico_balance") or {}).items():
                    acc["fico_balance"][k] = (
                        acc["fico_balance"].get(k, 0.0) + v
                    )
        if not merged:
            reason = (
                "; ".join(per_file_errors)
                if per_file_errors
                else "no column profile (upload as anchor + map cols)"
            )
            out["months_skipped"].append(
                {"period": ym, "name": ", ".join(
                    f.get("name") or "" for f in mapped
                ), "reason": reason}
            )
            continue
        # Round merged balances for storage.
        for rec in merged.values():
            rec["total_balance"] = round(rec["total_balance"], 2)
            rec["fico_balance"] = {
                k: round(v, 2) for k, v in rec["fico_balance"].items()
            }
        rollup_rows = list(merged.values())
        source_label = " + ".join(chosen_names) or "(unknown)"
        try:
            written = upsert_month(cu, as_of, rollup_rows, source=source_label)
        except Exception as exc:  # noqa: BLE001
            out["months_skipped"].append(
                {"period": ym, "name": source_label,
                 "reason": f"upsert error: {exc}"}
            )
            continue
        codes = {r["loan_code"] for r in rollup_rows}
        all_codes |= codes
        out["months_processed"].append({
            "period": ym,
            "as_of_date": as_of,
            "name": source_label,
            "loan_codes": len(codes),
            "rows_written": written,
            "extras": len(bucket) - len(chosen_names),
        })
        out["rows_written"] += written

    out["loan_codes_total"] = len(all_codes)
    out["ok"] = True
    return out


# ---------------------------------------------------------------------------
# Read-back helpers (used by the UI to surface code-appearance changes)
# ---------------------------------------------------------------------------

def codes_by_month(cu: str) -> dict[str, list[str]]:
    """Return ``{"YYYY-MM-DD": [loan_code, ...]}`` for this CU."""
    eng = _engine_lazy()
    out: dict[str, list[str]] = {}
    sql = text("""
        SELECT as_of_date, loan_code
        FROM loan_code_history
        WHERE cu = :cu
        ORDER BY as_of_date DESC, loan_code
    """)
    with eng.connect() as conn:
        for row in conn.execute(sql, {"cu": cu}):
            key = row[0].isoformat() if hasattr(row[0], "isoformat") else str(row[0])
            out.setdefault(key, []).append(row[1])
    return out


def history_matrix(cu: str) -> dict[str, Any]:
    """Pivot ``loan_code_history`` into a month-by-loan-code matrix.

    Returns a dict shaped for direct template rendering::

        {
          "months":  ["2026-03-31", "2026-02-28", ...],   # newest first
          "codes":   ["AUT", "MTG", ...],                 # union, sorted
          "cells":   {month: {code: {"balance", "count"}}},
          "totals_by_month": {month: {"balance", "count", "codes"}},
          "totals_by_code":  {code:  {"balance", "count", "months"}},
          "new_in_month":      {month: [code, ...]},
          "dropped_after_month": {month: [code, ...]},
          "row_count": int,
        }
    """
    eng = _engine_lazy()
    out: dict[str, Any] = {
        "months": [],
        "codes": [],
        "cells": {},
        "totals_by_month": {},
        "totals_by_code": {},
        "new_in_month": {},
        "dropped_after_month": {},
        "row_count": 0,
    }
    try:
        ensure_table()
    except Exception:  # noqa: BLE001
        return out
    if not cu:
        return out

    sql = text("""
        SELECT as_of_date, loan_code, total_balance, loan_count
        FROM loan_code_history
        WHERE cu = :cu
        ORDER BY as_of_date DESC, loan_code
    """)
    months_set: set[str] = set()
    codes_set: set[str] = set()
    with eng.connect() as conn:
        for row in conn.execute(sql, {"cu": cu}):
            d = row[0]
            month = d.isoformat() if hasattr(d, "isoformat") else str(d)
            code = row[1]
            bal = float(row[2] or 0)
            cnt = int(row[3] or 0)
            months_set.add(month)
            codes_set.add(code)
            out["cells"].setdefault(month, {})[code] = {
                "balance": bal,
                "count": cnt,
            }
            tm = out["totals_by_month"].setdefault(
                month, {"balance": 0.0, "count": 0, "codes": 0}
            )
            tm["balance"] += bal
            tm["count"] += cnt
            tm["codes"] += 1
            tc = out["totals_by_code"].setdefault(
                code, {"balance": 0.0, "count": 0, "months": 0}
            )
            tc["balance"] += bal
            tc["count"] += cnt
            tc["months"] += 1
            out["row_count"] += 1

    months_sorted = sorted(months_set, reverse=True)   # newest first
    codes_sorted = sorted(codes_set)
    out["months"] = months_sorted
    out["codes"] = codes_sorted

    # New / dropped flags.  Walk the months OLD -> NEW so we can spot
    # codes that show up for the first time, and NEW -> OLD to spot codes
    # that disappear afterwards.
    seen: set[str] = set()
    for month in reversed(months_sorted):
        present = set(out["cells"].get(month, {}).keys())
        new_here = sorted(present - seen)
        if new_here:
            out["new_in_month"][month] = new_here
        seen |= present
    for i, month in enumerate(months_sorted):
        present = set(out["cells"].get(month, {}).keys())
        # "Dropped after this month" = present here but missing in the
        # *next-newer* month (i-1 in the newest-first list).
        if i == 0:
            continue
        newer = set(out["cells"].get(months_sorted[i - 1], {}).keys())
        gone = sorted(present - newer)
        if gone:
            out["dropped_after_month"][month] = gone

    return out
