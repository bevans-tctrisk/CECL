"""Roll up historical delinquency snapshots.

Parallel to ``chargeoff_hist_processor`` (which handles charge-offs).
Each row stores the delinquent balance for one loan_code at one
``as_of_date`` (typically a quarter-end). Unlike charge-offs (a flow),
delinquency is a point-in-time snapshot, so the value at each date
stands on its own (no YTD differencing).

Optional columns ``total_balance`` and ``dq_pct`` let the caller store
either:
  * raw balances (dq_amount + optional total_balance), with the
    downstream consumer computing pct as dq_amount / total_balance, OR
  * an explicit percentage (dq_pct), used as-is when present.

This dual support is needed because the three wizard sources produce
different things:
  * Loan-extract derivation: dq_amount + total_balance (both real $).
  * 5300 backfill:           dq_amount (sum of >=60-day buckets).
  * Manual entry:            dq_pct (user types a percentage).
"""
from __future__ import annotations

from typing import Any

from sqlalchemy import text

from cecl_credentials import get_database_url


_DDL = """
CREATE TABLE IF NOT EXISTS loan_code_delinquency_history (
    cu              TEXT        NOT NULL,
    as_of_date      DATE        NOT NULL,
    loan_code       TEXT        NOT NULL,
    dq_amount       NUMERIC(20, 2) NOT NULL DEFAULT 0,
    total_balance   NUMERIC(20, 2),
    dq_pct          NUMERIC(12, 8),
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
        # Best-effort migration: add the optional columns if an older
        # version of the table exists without them.
        for col_ddl in (
            "ALTER TABLE loan_code_delinquency_history "
            "ADD COLUMN IF NOT EXISTS total_balance NUMERIC(20, 2)",
            "ALTER TABLE loan_code_delinquency_history "
            "ADD COLUMN IF NOT EXISTS dq_pct NUMERIC(12, 8)",
        ):
            try:
                conn.execute(text(col_ddl))
            except Exception:  # noqa: BLE001
                pass


_UPSERT = text("""
    INSERT INTO loan_code_delinquency_history
        (cu, as_of_date, loan_code, dq_amount, total_balance, dq_pct,
         source, updated_at)
    VALUES
        (:cu, :as_of_date, :loan_code, :dq_amount, :total_balance, :dq_pct,
         :source, NOW())
    ON CONFLICT (cu, as_of_date, loan_code) DO UPDATE SET
        dq_amount     = EXCLUDED.dq_amount,
        total_balance = EXCLUDED.total_balance,
        dq_pct        = EXCLUDED.dq_pct,
        source        = EXCLUDED.source,
        updated_at    = NOW();
""")

_DELETE_MONTH = text("""
    DELETE FROM loan_code_delinquency_history
    WHERE cu = :cu AND as_of_date = :as_of_date;
""")


def upsert_month(
    cu: str, as_of_date: str, rows: list[dict[str, Any]], source: str
) -> int:
    """Replace prior DQ rows for (cu, as_of_date) with ``rows``.

    Each row dict must have ``loan_code``. Optional keys:
    ``dq_amount`` (default 0), ``total_balance`` (default None),
    ``dq_pct`` (default None).
    """
    eng = _engine_lazy()
    with eng.begin() as conn:
        conn.execute(_DELETE_MONTH, {"cu": cu, "as_of_date": as_of_date})
        if not rows:
            return 0
        params = [
            {
                "cu": cu,
                "as_of_date": as_of_date,
                "loan_code": r["loan_code"],
                "dq_amount": float(r.get("dq_amount") or 0.0),
                "total_balance": (
                    float(r["total_balance"])
                    if r.get("total_balance") not in (None, "")
                    else None
                ),
                "dq_pct": (
                    float(r["dq_pct"])
                    if r.get("dq_pct") not in (None, "")
                    else None
                ),
                "source": source,
            }
            for r in rows
        ]
        conn.execute(_UPSERT, params)
        return len(params)


def delete_rows_by_source_prefix(cu: str, source_prefix: str) -> int:
    """Drop every row for ``cu`` whose ``source`` starts with ``source_prefix``.

    Used by the 5300 backfill to purge stale prior-run rows before
    rewriting. Returns the number of rows deleted.
    """
    eng = _engine_lazy()
    with eng.begin() as conn:
        res = conn.execute(
            text(
                "DELETE FROM loan_code_delinquency_history "
                "WHERE cu = :cu AND source LIKE :pfx"
            ),
            {"cu": cu, "pfx": f"{source_prefix}%"},
        )
        return int(res.rowcount or 0)


def history_matrix(cu: str) -> dict[str, Any]:
    """Return a month-by-loan-code matrix of DQ amounts / percentages.

    Shape mirrors :func:`chargeoff_hist_processor.history_matrix` with
    extra ``total_balance`` and ``dq_pct`` keys inside each cell so the
    UI can show the underlying numbers and either source mode.
    """
    out: dict[str, Any] = {
        "months": [],
        "codes": [],
        "cells": {},
        "totals_by_month": {},
        "totals_by_code": {},
        "row_count": 0,
    }
    if not cu:
        return out
    try:
        ensure_table()
    except Exception:  # noqa: BLE001
        return out
    eng = _engine_lazy()
    with eng.begin() as conn:
        rows = conn.execute(
            text(
                "SELECT as_of_date, loan_code, dq_amount, "
                "       total_balance, dq_pct, source "
                "FROM loan_code_delinquency_history "
                "WHERE cu = :cu "
                "ORDER BY as_of_date DESC, loan_code"
            ),
            {"cu": cu},
        ).fetchall()
    months_set: set[str] = set()
    codes_set: set[str] = set()
    for r in rows:
        m = r[0].isoformat() if hasattr(r[0], "isoformat") else str(r[0])
        c = r[1]
        amt = float(r[2] or 0.0)
        tot = float(r[3]) if r[3] is not None else None
        pct = float(r[4]) if r[4] is not None else None
        src = r[5] or ""
        months_set.add(m)
        codes_set.add(c)
        out["cells"].setdefault(m, {})[c] = {
            "amount": amt,
            "total_balance": tot,
            "dq_pct": pct,
            "source": src,
        }
        tm = out["totals_by_month"].setdefault(
            m, {"amount": 0.0, "total_balance": 0.0, "codes": 0}
        )
        tm["amount"] += amt
        if tot is not None:
            tm["total_balance"] += tot
        tm["codes"] += 1
        tc = out["totals_by_code"].setdefault(
            c, {"amount": 0.0, "months": 0}
        )
        tc["amount"] += amt
        tc["months"] += 1
        out["row_count"] += 1

    out["months"] = sorted(months_set, reverse=True)
    out["codes"] = sorted(codes_set)
    return out


def existing_dates(cu: str) -> set[str]:
    """Return the set of ``YYYY-MM-DD`` dates that already have DQ rows."""
    if not cu:
        return set()
    try:
        ensure_table()
    except Exception:  # noqa: BLE001
        return set()
    eng = _engine_lazy()
    with eng.begin() as conn:
        rows = conn.execute(
            text(
                "SELECT DISTINCT as_of_date "
                "FROM loan_code_delinquency_history WHERE cu = :cu"
            ),
            {"cu": cu},
        ).fetchall()
    return {
        (r[0].isoformat() if hasattr(r[0], "isoformat") else str(r[0]))
        for r in rows
    }
