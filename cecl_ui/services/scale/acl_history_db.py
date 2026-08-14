"""PostgreSQL-backed persistence for SCALE prior-quarter ACL values.

This table replaces the previous "read formula cells from the prior
quarter's workbook via openpyxl" approach which was fragile because
openpyxl cannot evaluate Excel formulas and therefore relies on a
cached-on-disk value -- only present after Excel itself has opened and
saved the file. Without pywin32 to drive Excel headlessly, that read
silently returned None and the runner fell back to writing literal 0
into ``Historical Data!AY2..AY5`` of the new workbook (which made the
"prior" column on Executive Summary-Vizo equal the "current" column,
or zero).

Schema (one row per CU + quarter-end)::

    CREATE TABLE IF NOT EXISTS scale_acl_history (
        cu_short                TEXT           NOT NULL,
        period_end              DATE           NOT NULL,
        total_expected_losses   NUMERIC(20, 2),
        acl_balance             NUMERIC(20, 2),
        adjustment              NUMERIC(20, 2),
        acl_ratio               NUMERIC(12, 8),
        source                  TEXT,
        updated_at              TIMESTAMP      NOT NULL DEFAULT NOW(),
        PRIMARY KEY (cu_short, period_end)
    );

Population paths (in priority order):

* Runtime capture: after a SCALE report is generated and the runner's
  best-effort post-run recalc completes, ``capture_from_workbook``
  reads the four cached values out of the just-written workbook and
  upserts them. Works when pywin32 successfully ran COM recalc.
* Manual backfill: ``scripts/scale_acl_backfill.py`` walks every
  ``Generated_Reports/<short>/<period>/*.xlsx`` and runs the same
  capture function. Useful after the user opens existing workbooks in
  Excel and saves them so cached values populate.
* Future: a manual-entry UI on the run dashboard so analysts can type
  values directly when neither auto-capture path works. (Not yet
  implemented; deferred until needed.)

Read path:

* ``runner._inject_prior_acl`` calls ``lookup`` for the prior quarter
  end. On hit, it skips the prior-workbook read entirely and writes
  the four values as literals into the new workbook's
  ``Historical Data!AY2..AY5``. On miss, it falls back to the existing
  Excel-read logic so behaviour is backwards-compatible.
"""
from __future__ import annotations

from datetime import date, datetime
from pathlib import Path
from typing import Any

from sqlalchemy import text

from cecl_credentials import get_database_url


_DDL = """
CREATE TABLE IF NOT EXISTS scale_acl_history (
    cu_short                TEXT           NOT NULL,
    period_end              DATE           NOT NULL,
    total_expected_losses   NUMERIC(20, 2),
    acl_balance             NUMERIC(20, 2),
    adjustment              NUMERIC(20, 2),
    acl_ratio               NUMERIC(12, 8),
    source                  TEXT,
    updated_at              TIMESTAMP      NOT NULL DEFAULT NOW(),
    PRIMARY KEY (cu_short, period_end)
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
    """Create the table if missing. Idempotent."""
    eng = _engine_lazy()
    with eng.begin() as conn:
        conn.execute(text(_DDL))


def _period_to_date(period: str | date | datetime) -> date | None:
    """Coerce ``"YYYY-MM"`` / ``"YYYY-MM-DD"`` / date / datetime to a
    quarter-end date. ``YYYY-MM`` resolves to the last day of that month.
    """
    if isinstance(period, datetime):
        return period.date()
    if isinstance(period, date):
        return period
    s = str(period or "").strip()
    if not s:
        return None
    parts = s.split("-")
    if len(parts) == 3:
        try:
            return date(int(parts[0]), int(parts[1]), int(parts[2]))
        except (TypeError, ValueError):
            return None
    if len(parts) == 2:
        try:
            import calendar
            y, m = int(parts[0]), int(parts[1])
            last = calendar.monthrange(y, m)[1]
            return date(y, m, last)
        except (TypeError, ValueError):
            return None
    return None


_UPSERT = text("""
    INSERT INTO scale_acl_history
        (cu_short, period_end, total_expected_losses, acl_balance,
         adjustment, acl_ratio, source, updated_at)
    VALUES
        (:cu_short, :period_end, :total_expected_losses, :acl_balance,
         :adjustment, :acl_ratio, :source, NOW())
    ON CONFLICT (cu_short, period_end) DO UPDATE SET
        total_expected_losses = COALESCE(
            EXCLUDED.total_expected_losses,
            scale_acl_history.total_expected_losses
        ),
        acl_balance = COALESCE(
            EXCLUDED.acl_balance, scale_acl_history.acl_balance
        ),
        adjustment = COALESCE(
            EXCLUDED.adjustment, scale_acl_history.adjustment
        ),
        acl_ratio = COALESCE(
            EXCLUDED.acl_ratio, scale_acl_history.acl_ratio
        ),
        source = EXCLUDED.source,
        updated_at = NOW();
""")


def upsert(
    cu_short: str,
    period: str | date | datetime,
    values: dict[str, Any],
    source: str = "",
) -> dict[str, Any]:
    """Insert or update one row.

    ``values`` may use either the ``AY2..AY5`` keys (matching
    ``runs_service.read_prior_acl_values`` output) OR the explicit
    column names. Empty/non-numeric entries are coerced to None.

    Returns ``{"ok": bool, "wrote": int, "error": str}``.
    """
    result: dict[str, Any] = {"ok": False, "wrote": 0, "error": ""}
    if not cu_short:
        result["error"] = "cu_short is required"
        return result
    pe = _period_to_date(period)
    if pe is None:
        result["error"] = f"could not parse period {period!r}"
        return result

    # Accept both shapes.
    tel = (values.get("AY2") if isinstance(values, dict) else None)
    if tel is None and isinstance(values, dict):
        tel = values.get("total_expected_losses")
    bal = (values.get("AY3") if isinstance(values, dict) else None)
    if bal is None and isinstance(values, dict):
        bal = values.get("acl_balance")
    adj = (values.get("AY4") if isinstance(values, dict) else None)
    if adj is None and isinstance(values, dict):
        adj = values.get("adjustment")
    ratio = (values.get("AY5") if isinstance(values, dict) else None)
    if ratio is None and isinstance(values, dict):
        ratio = values.get("acl_ratio")

    def _num(v):
        if isinstance(v, (int, float)):
            return float(v)
        try:
            return float(v) if v not in (None, "") else None
        except (TypeError, ValueError):
            return None

    payload = {
        "cu_short": cu_short,
        "period_end": pe,
        "total_expected_losses": _num(tel),
        "acl_balance": _num(bal),
        "adjustment": _num(adj),
        "acl_ratio": _num(ratio),
        "source": source or "",
    }

    # Don't waste a row when every numeric field is None.
    if all(payload[k] is None for k in (
        "total_expected_losses", "acl_balance", "adjustment", "acl_ratio"
    )):
        result["error"] = "all values are None; nothing to persist"
        return result

    try:
        ensure_table()
        eng = _engine_lazy()
        with eng.begin() as conn:
            conn.execute(_UPSERT, payload)
        result["ok"] = True
        result["wrote"] = 1
    except Exception as exc:  # noqa: BLE001
        result["error"] = f"upsert failed: {exc}"
    return result


_SELECT_ONE = text("""
    SELECT total_expected_losses, acl_balance, adjustment, acl_ratio,
           source, updated_at
    FROM scale_acl_history
    WHERE cu_short = :cu_short AND period_end = :period_end
""")


def lookup(
    cu_short: str, period: str | date | datetime
) -> dict[str, Any] | None:
    """Return ``{"values": {AY2..AY5}, "source": str, "updated_at":
    datetime, "period_end": date}`` for the exact (cu, period), or
    None when missing. Never raises; logs error to stdout on failure.
    """
    if not cu_short:
        return None
    pe = _period_to_date(period)
    if pe is None:
        return None
    try:
        ensure_table()
        eng = _engine_lazy()
        with eng.connect() as conn:
            row = conn.execute(
                _SELECT_ONE, {"cu_short": cu_short, "period_end": pe},
            ).mappings().first()
    except Exception as exc:  # noqa: BLE001
        print(f"scale_acl_history.lookup({cu_short}, {pe}) failed: {exc}")
        return None
    if row is None:
        return None
    values: dict[str, float] = {}
    for key, coord in (
        ("total_expected_losses", "AY2"),
        ("acl_balance", "AY3"),
        ("adjustment", "AY4"),
        ("acl_ratio", "AY5"),
    ):
        v = row.get(key)
        if v is not None:
            try:
                values[coord] = float(v)
            except (TypeError, ValueError):
                pass
    if not values:
        return None
    return {
        "values": values,
        "source": row.get("source") or "",
        "updated_at": row.get("updated_at"),
        "period_end": pe,
    }


_SELECT_LATEST_BEFORE = text("""
    SELECT period_end, total_expected_losses, acl_balance, adjustment,
           acl_ratio, source, updated_at
    FROM scale_acl_history
    WHERE cu_short = :cu_short AND period_end < :before
    ORDER BY period_end DESC
    LIMIT 1
""")


def lookup_latest_before(
    cu_short: str, before_period: str | date | datetime
) -> dict[str, Any] | None:
    """Return the newest row strictly older than ``before_period``.

    Used as a fallback when the exact-prior quarter has no row -- for
    example when an analyst skips a quarter (Sep -> Mar) or when a
    previous run failed before reaching the DB capture step.
    """
    if not cu_short:
        return None
    before = _period_to_date(before_period)
    if before is None:
        return None
    try:
        ensure_table()
        eng = _engine_lazy()
        with eng.connect() as conn:
            row = conn.execute(
                _SELECT_LATEST_BEFORE,
                {"cu_short": cu_short, "before": before},
            ).mappings().first()
    except Exception as exc:  # noqa: BLE001
        print(
            f"scale_acl_history.lookup_latest_before({cu_short}, {before}) "
            f"failed: {exc}"
        )
        return None
    if row is None:
        return None
    values: dict[str, float] = {}
    for key, coord in (
        ("total_expected_losses", "AY2"),
        ("acl_balance", "AY3"),
        ("adjustment", "AY4"),
        ("acl_ratio", "AY5"),
    ):
        v = row.get(key)
        if v is not None:
            try:
                values[coord] = float(v)
            except (TypeError, ValueError):
                pass
    if not values:
        return None
    return {
        "values": values,
        "source": row.get("source") or "",
        "updated_at": row.get("updated_at"),
        "period_end": row.get("period_end"),
    }


def capture_from_workbook(
    cu_short: str,
    period: str | date | datetime,
    workbook_path: str | Path,
    source: str = "captured-from-workbook",
) -> dict[str, Any]:
    """Read U29..U33 cached values from the given SCALE workbook and
    upsert them into the DB.

    This is the "passive" capture path -- works only when the
    workbook has cached formula values (set by Excel on save, or by
    pywin32 COM recalc). Returns ``{"ok": bool, "wrote": int,
    "values": dict, "skipped": bool, "error": str}``.

    ``skipped=True`` indicates the workbook had no usable cached
    numeric values -- the caller should treat this as a soft miss,
    not a hard failure (it just means the file hasn't been opened in
    Excel yet).
    """
    # Local import to avoid a circular dep: runs_service is in the same
    # package and imports nothing from acl_history_db.
    from cecl_ui.services.scale import runs_service

    result: dict[str, Any] = {
        "ok": False, "wrote": 0, "values": {}, "skipped": False, "error": "",
    }
    read = runs_service.read_prior_acl_values(workbook_path)
    if not read.get("ok"):
        # ``read_prior_acl_values`` already attempted cached + raw passes.
        # No numeric values means we cannot capture anything.
        result["skipped"] = True
        result["error"] = read.get("error", "no cached values available")
        return result
    values = read.get("values") or {}
    if not values:
        result["skipped"] = True
        result["error"] = "no numeric values returned"
        return result
    up = upsert(cu_short, period, values, source=source)
    result["ok"] = bool(up.get("ok"))
    result["wrote"] = int(up.get("wrote") or 0)
    result["values"] = values
    if not result["ok"]:
        result["error"] = up.get("error") or ""
    return result
