"""Roll up historical NCUA 5300 recovery rows.

Parallel to ``chargeoff_hist_processor`` — recoveries live in their own
table for symmetry.
"""
from __future__ import annotations

from typing import Any

from sqlalchemy import text

from cecl_credentials import get_database_url


_DDL = """
CREATE TABLE IF NOT EXISTS loan_code_recovery_history (
    cu              TEXT        NOT NULL,
    as_of_date      DATE        NOT NULL,
    loan_code       TEXT        NOT NULL,
    recovery_amount NUMERIC(20, 2) NOT NULL DEFAULT 0,
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


_UPSERT = text("""
    INSERT INTO loan_code_recovery_history
        (cu, as_of_date, loan_code, recovery_amount, source, updated_at)
    VALUES
        (:cu, :as_of_date, :loan_code, :recovery_amount, :source, NOW())
    ON CONFLICT (cu, as_of_date, loan_code) DO UPDATE SET
        recovery_amount = EXCLUDED.recovery_amount,
        source          = EXCLUDED.source,
        updated_at      = NOW();
""")

_DELETE_MONTH = text("""
    DELETE FROM loan_code_recovery_history
    WHERE cu = :cu AND as_of_date = :as_of_date;
""")


def upsert_month(
    cu: str, as_of_date: str, rows: list[dict[str, Any]], source: str
) -> int:
    """Replace prior recovery rows for (cu, as_of_date) with ``rows``."""
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
                "recovery_amount": r["recovery_amount"],
                "source": source,
            }
            for r in rows
        ]
        conn.execute(_UPSERT, params)
        return len(params)


def history_matrix(cu: str) -> dict[str, Any]:
    """Return a month-by-loan-code matrix of recovery amounts.

    Shape mirrors :func:`extract_hist_processor.history_matrix`.
    """
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
                "SELECT as_of_date, loan_code, recovery_amount "
                "FROM loan_code_recovery_history "
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
        months_set.add(m)
        codes_set.add(c)
        out["cells"].setdefault(m, {})[c] = {"amount": amt}
        tm = out["totals_by_month"].setdefault(
            m, {"amount": 0.0, "codes": 0}
        )
        tm["amount"] += amt
        tm["codes"] += 1
        tc = out["totals_by_code"].setdefault(
            c, {"amount": 0.0, "months": 0}
        )
        tc["amount"] += amt
        tc["months"] += 1
        out["row_count"] += 1

    months_sorted = sorted(months_set, reverse=True)
    codes_sorted = sorted(codes_set)
    out["months"] = months_sorted
    out["codes"] = codes_sorted

    seen: set[str] = set()
    for month in reversed(months_sorted):
        present = set(out["cells"].get(month, {}).keys())
        new_here = sorted(present - seen)
        if new_here:
            out["new_in_month"][month] = new_here
        seen |= present
    for i, month in enumerate(months_sorted):
        if i == 0:
            continue
        present = set(out["cells"].get(month, {}).keys())
        newer = set(out["cells"].get(months_sorted[i - 1], {}).keys())
        gone = sorted(present - newer)
        if gone:
            out["dropped_after_month"][month] = gone
    return out
