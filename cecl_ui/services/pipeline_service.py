"""Wrappers around the existing CECL pipeline scripts."""
from __future__ import annotations

import contextlib
import importlib
import io
import os
from pathlib import Path
from typing import Any


def _client_configs_dir() -> Path:
    """Return the client_configs directory.

    Respects CECL_WORKSPACE_ROOT so the data root can live on a shared drive
    while the code lives in a local clone.  Falls back to the historical
    layout (sibling of cecl_ui/) when the env var is unset.
    """
    ws = os.environ.get("CECL_WORKSPACE_ROOT", "").strip()
    root = Path(ws) if ws else Path(__file__).resolve().parents[2]
    return root / "client_configs"


def run_import(client_short_name: str, specific_file: str | None = None) -> int:
    """Invoke import_data.process_client; returns row count."""
    import_data = importlib.import_module("import_data")
    return import_data.process_client(client_short_name, specific_file=specific_file)


def run_reports(
    client_short_name: str,
    snapshot_date: str | None = None,
    reports: list[str] | None = None,
) -> tuple[list[str], str]:
    """Invoke generate_report.generate_report.

    Returns (output_paths, captured_stdout). The stdout is useful when no
    reports were generated — it carries the underlying ERROR line(s).
    """
    gr = importlib.import_module("generate_report")
    buf = io.StringIO()
    with contextlib.redirect_stdout(buf):
        out = gr.generate_report(
            client_short_name,
            snapshot_date=snapshot_date,
            reports=reports,
        )
    return [str(p) for p in (out or [])], buf.getvalue()


def run_impdet_report(client_short_name: str, snapshot_date: str | None = None) -> str | None:
    gid = importlib.import_module("generate_impdet_report")
    out = gid.generate_report(client_short_name, snap=snapshot_date)
    return str(out) if out else None


def fetch_economic_data(state: str, county: str | None = None) -> dict[str, Any]:
    fed = importlib.import_module("fetch_econ_data")
    try:
        return fed.fetch_economic_data(state, county) or {}
    except Exception as exc:  # noqa: BLE001
        return {"error": str(exc)}


def list_snapshots_for_cu(client_short_name: str) -> list[str]:
    """Return distinct snapshot_date values currently in the DB for the CU."""
    try:
        cfg_mod = importlib.import_module("cecl_credentials")
        url = cfg_mod.get_database_url()
    except Exception:
        return []
    try:
        from sqlalchemy import create_engine, text
        cfg = importlib.import_module("yaml").safe_load(
            (_client_configs_dir()
             / f"{client_short_name}.yaml").read_text(encoding="utf-8")
        )
        cu_name = cfg.get("credit_union", client_short_name)
        eng = create_engine(url)
        with eng.connect() as conn:
            rows = conn.execute(
                text(
                    "SELECT DISTINCT snapshot_date FROM monthly_loan_data "
                    "WHERE credit_union = :cu ORDER BY snapshot_date DESC"
                ),
                {"cu": cu_name},
            ).fetchall()
        return [str(r[0]) for r in rows]
    except Exception:
        return []


def list_pools_for_cu(client_short_name: str) -> list[str]:
    """Return distinct loan_pool values currently in the DB for the CU."""
    try:
        cfg_mod = importlib.import_module("cecl_credentials")
        url = cfg_mod.get_database_url()
        from sqlalchemy import create_engine, text
        cfg = importlib.import_module("yaml").safe_load(
            (_client_configs_dir()
             / f"{client_short_name}.yaml").read_text(encoding="utf-8")
        )
        cu_name = cfg.get("credit_union", client_short_name)
        eng = create_engine(url)
        with eng.connect() as conn:
            rows = conn.execute(
                text(
                    "SELECT DISTINCT loan_pool FROM monthly_loan_data "
                    "WHERE credit_union = :cu AND loan_pool IS NOT NULL "
                    "ORDER BY loan_pool"
                ),
                {"cu": cu_name},
            ).fetchall()
        return [str(r[0]) for r in rows if r[0]]
    except Exception:
        return []


def pool_distribution_for_cu(
    client_short_name: str, snapshot_date: str | None = None
) -> list[dict]:
    """Return ``[{pool, count, balance}, ...]`` for the CU.

    If ``snapshot_date`` is None, uses the latest snapshot. Returns an
    empty list on any error (missing DB, no rows, etc).
    """
    try:
        cfg_mod = importlib.import_module("cecl_credentials")
        url = cfg_mod.get_database_url()
        from sqlalchemy import create_engine, text
        cfg = importlib.import_module("yaml").safe_load(
            (_client_configs_dir()
             / f"{client_short_name}.yaml").read_text(encoding="utf-8")
        )
        cu_name = cfg.get("credit_union", client_short_name)
        eng = create_engine(url)
        with eng.connect() as conn:
            if not snapshot_date:
                snapshot_date = conn.execute(
                    text(
                        "SELECT MAX(snapshot_date) FROM monthly_loan_data "
                        "WHERE credit_union = :cu"
                    ),
                    {"cu": cu_name},
                ).scalar()
                if not snapshot_date:
                    return []
            rows = conn.execute(
                text(
                    "SELECT loan_pool, COUNT(*), "
                    "       COALESCE(SUM(current_balance), 0) "
                    "FROM monthly_loan_data "
                    "WHERE credit_union = :cu AND snapshot_date = :s "
                    "GROUP BY loan_pool "
                    "ORDER BY COALESCE(SUM(current_balance), 0) DESC"
                ),
                {"cu": cu_name, "s": snapshot_date},
            ).fetchall()
        return [
            {"pool": (r[0] or "(null)"),
             "count": int(r[1] or 0),
             "balance": float(r[2] or 0)}
            for r in rows
        ]
    except Exception:
        return []


def reclassify_loan_pool(
    client_short_name: str,
    from_pool: str,
    to_pool: str,
    snapshot_date: str | None = None,
) -> int:
    """Bulk-update ``monthly_loan_data.loan_pool`` rows for this CU from
    ``from_pool`` to ``to_pool``. Optionally limit to a single snapshot.

    Returns the number of rows updated. Raises on DB errors so the
    caller can surface them.
    """
    cfg_mod = importlib.import_module("cecl_credentials")
    url = cfg_mod.get_database_url()
    from sqlalchemy import create_engine, text
    cfg = importlib.import_module("yaml").safe_load(
        (_client_configs_dir()
         / f"{client_short_name}.yaml").read_text(encoding="utf-8")
    )
    cu_name = cfg.get("credit_union", client_short_name)
    eng = create_engine(url)
    params = {"cu": cu_name, "frm": from_pool, "to": to_pool}
    sql = (
        "UPDATE monthly_loan_data SET loan_pool = :to "
        "WHERE credit_union = :cu AND loan_pool = :frm"
    )
    if snapshot_date:
        sql += " AND snapshot_date = :s"
        params["s"] = snapshot_date
    with eng.begin() as conn:
        result = conn.execute(text(sql), params)
        return int(result.rowcount or 0)


def import_warm_as_baseline(
    cu_name: str,
    warm_source_path: str | Path,
    as_of_date: str,
    overwrite: bool = False,
) -> dict[str, Any]:
    """Copy a Manual WARM workbook into a baseline folder under Reports/
    so the existing report engine can use it as a "prior TCT report."

    The engine at generate_report._find_prior_tct_report() and
    load_prior_tct_hist_bal() walk Reports/ (recursively) looking for files
    matching::

        <YYYY-MM-DD>_CECL_Migration_<safe_cu>_TCT_Model.xlsx

    where ``safe_cu`` is the CU name with spaces replaced by underscores and
    slashes by dashes.  We store the WARM workbook under the canonical name
    in the subdirectory ``Reports/_warm_baselines/`` so subsequent report
    runs (which write the generated TCT to the top-level ``Reports/``
    folder) never overwrite the WARM source.

    Returns a dict with::

        {"ok": bool, "dest_path": str, "filename": str,
         "skipped": bool,           # True if file already existed and overwrite=False
         "error": str | None}
    """
    import shutil
    from pathlib import Path as _Path

    workspace_root = _Path(__file__).resolve().parents[2]
    reports_dir = workspace_root / "Reports" / "_warm_baselines"
    reports_dir.mkdir(parents=True, exist_ok=True)

    safe_cu = (cu_name or "").strip().replace(" ", "_").replace("/", "-")
    if not safe_cu:
        return {"ok": False, "dest_path": "", "filename": "", "skipped": False,
                "error": "Empty CU name"}
    if not as_of_date or len(as_of_date) < 10:
        return {"ok": False, "dest_path": "", "filename": "", "skipped": False,
                "error": f"Invalid as-of date: {as_of_date!r}"}

    fname = f"{as_of_date}_CECL_Migration_{safe_cu}_TCT_Model.xlsx"
    dest = reports_dir / fname

    if dest.exists() and not overwrite:
        return {"ok": True, "dest_path": str(dest), "filename": fname,
                "skipped": True, "error": None}

    try:
        shutil.copy2(str(warm_source_path), str(dest))
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "dest_path": str(dest), "filename": fname,
                "skipped": False, "error": str(exc)}

    return {"ok": True, "dest_path": str(dest), "filename": fname,
            "skipped": False, "error": None}
