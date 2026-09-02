"""Wrappers around the existing CECL pipeline scripts."""
from __future__ import annotations

import contextlib
import importlib
import io
import os
import shutil
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


def _workspace_root() -> Path:
    """Return the workspace root used by import_data.BASE_FOLDER."""
    ws = os.environ.get("CECL_WORKSPACE_ROOT", "").strip()
    if ws:
        return Path(ws)
    # Fall back to the historical layout (sibling of cecl_ui/).
    return Path(__file__).resolve().parents[2]


def restore_archived_files(client_short_name: str) -> dict[str, list[str]]:
    """Copy files from ``Archive/<short>/`` back into ``Raw_Uploads/<short>/``.

    Use case: ``archive_imported_files: true`` (the default for most YAMLs)
    sweeps successfully-imported loan extracts out of ``Raw_Uploads`` and into
    ``Archive`` so a subsequent ``run_import`` call is a no-op.  When the user
    drops a fresh credit-pull file and wants the prior quarters re-imported
    against the new scores, ``Raw_Uploads`` is effectively empty — restoring
    from ``Archive`` brings the files back so the importer has something to
    chew on.

    Uses ``shutil.copy2`` (NOT ``shutil.move``) so the archive copy survives
    as a safety net.  Files already present in the destination are skipped to
    avoid overwriting any in-flight uploads the user just made.

    Returns ``{restored: [names], skipped_existing: [names], errors: [strs]}``.
    Silently no-ops when the Archive folder doesn't exist (first-time setup).
    """
    out: dict[str, list[str]] = {
        "restored": [],
        "skipped_existing": [],
        "errors": [],
    }
    ws = _workspace_root()
    archive_dir = ws / "Archive" / client_short_name
    upload_dir = ws / "Raw_Uploads" / client_short_name
    if not archive_dir.exists() or not archive_dir.is_dir():
        return out
    try:
        upload_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:  # pragma: no cover - filesystem race
        out["errors"].append(f"cannot create {upload_dir}: {exc}")
        return out
    # Walk recursively because import_data preserves relative paths when
    # archiving (e.g. nested folders for per-extract subdirectories).
    for src in archive_dir.rglob("*"):
        if not src.is_file():
            continue
        try:
            rel = src.relative_to(archive_dir)
        except ValueError:
            continue
        dst = upload_dir / rel
        if dst.exists():
            out["skipped_existing"].append(str(rel))
            continue
        try:
            dst.parent.mkdir(parents=True, exist_ok=True)
            shutil.copy2(src, dst)
            out["restored"].append(str(rel))
        except OSError as exc:
            out["errors"].append(f"{rel}: {exc}")
    return out


def run_import(
    client_short_name: str,
    specific_file: str | None = None,
    *,
    restore_archive: bool = True,
) -> int:
    """Invoke import_data.process_client; returns row count.

    When ``restore_archive`` is ``True`` (the default), any files previously
    swept into ``Archive/<short>/`` are first copied back into
    ``Raw_Uploads/<short>/`` so a fresh credit-pull replacement automatically
    re-imports the historical quarters.  See ``restore_archived_files`` for
    details — the call is a no-op when the Archive folder doesn't exist.
    """
    if restore_archive and not specific_file:
        try:
            res = restore_archived_files(client_short_name)
            if res["restored"]:
                names = ", ".join(res["restored"][:4])
                more = (
                    f" (+{len(res['restored']) - 4} more)"
                    if len(res["restored"]) > 4 else ""
                )
                print(
                    f"Restored {len(res['restored'])} archived file(s) into "
                    f"Raw_Uploads/{client_short_name}/: {names}{more}"
                )
        except Exception as exc:  # noqa: BLE001 - never block import on this
            print(f"Archive restore skipped (non-fatal): {exc}")
    import_data = importlib.import_module("import_data")
    return import_data.process_client(client_short_name, specific_file=specific_file)


def reimport_period(client_short_name: str, period: str) -> dict[str, Any]:
    """Re-import ONE historical period on demand (outlier re-runs).

    Pulls exactly that period's loan files from the authoritative source
    folder — ``loan_source_folder`` config, else ``loan_file_folder``, else
    the CU's ``Archive/<short>/``, else ``Raw_Uploads/<short>/`` — copies just
    those files into a temp folder, imports ONLY that snapshot (idempotent
    per-snapshot replace), then removes the temp folder. Leaves the standing
    Raw_Uploads folder (the current period) untouched.

    ``period`` accepts ``YYYY-MM`` or ``YYYY-MM-DD``. Returns
    ``{ok, period, cu, source_folder, files, rows, error}``.
    """
    import os
    import re
    import shutil
    import calendar
    import tempfile

    import_data = importlib.import_module("import_data")
    try:
        config = import_data.load_client_config(client_short_name)
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "error": f"Could not load config: {exc}"}
    cu_name = config.get("credit_union")

    # Resolve the target snapshot as a month-end ISO date.
    p = (period or "").strip()
    if re.fullmatch(r"\d{4}-\d{2}-\d{2}", p):
        target = p
    elif re.fullmatch(r"\d{4}-\d{2}", p):
        y, m = int(p[:4]), int(p[5:7])
        if not (1 <= m <= 12):
            return {"ok": False, "error": f"Invalid month in period: {p!r}."}
        target = f"{y:04d}-{m:02d}-{calendar.monthrange(y, m)[1]:02d}"
    else:
        return {"ok": False,
                "error": f"Invalid period {p!r} — use YYYY-MM or YYYY-MM-DD."}

    # Resolve the source folder to pull the period's files from.
    source = config.get("loan_source_folder") or config.get("loan_file_folder")
    if source:
        source = import_data.resolve_path(source)
    else:
        candidates = [
            os.path.join(import_data.ARCHIVE_FOLDER, client_short_name),
            os.path.join(import_data.UPLOAD_FOLDER, client_short_name),
            import_data.UPLOAD_FOLDER,
        ]
        source = next((c for c in candidates if os.path.isdir(c)), None)
    if not source or not os.path.isdir(source):
        return {"ok": False, "period": target,
                "error": ("No source folder found to pull the period from. "
                          "Set 'loan_source_folder' in the config.")}

    # Compile the loan file patterns (top-level + per-extract).
    pats = list(import_data._compile_file_patterns(
        config.get("file_pattern"), label="top"))
    for e in (config.get("loan_data_extracts") or []):
        pats += import_data._compile_file_patterns(
            (e or {}).get("file_pattern"), label=(e or {}).get("label", "?"))

    # Find files whose snapshot resolves to the target AND match a loan
    # pattern. Search the primary source folder first; if it yields nothing
    # fall back to the CU's Archive and Raw_Uploads folders. This matters
    # because ``process_client`` moves each imported file into
    # ``Archive/<short>/`` after a successful import, so the current
    # quarter's loan file is no longer under ``loan_source_folder`` when an
    # on-demand re-import runs (e.g. after a Balance Adjustment pool remap).
    search_dirs: list[str] = []
    for d in (
        source,
        os.path.join(import_data.ARCHIVE_FOLDER, client_short_name),
        os.path.join(import_data.UPLOAD_FOLDER, client_short_name),
        import_data.UPLOAD_FOLDER,
    ):
        if d and os.path.isdir(d) and d not in search_dirs:
            search_dirs.append(d)

    matches: list[tuple[str, str]] = []  # (full_path, base_dir it was found under)
    used_source = source
    for base in search_dirs:
        found: list[tuple[str, str]] = []
        seen_names: set[str] = set()
        for root, _dirs, files in os.walk(base):
            for fn in files:
                if fn.startswith("~$"):
                    continue
                full = os.path.join(root, fn)
                rel = os.path.relpath(full, base)
                if pats and not import_data._patterns_match_any(pats, fn, rel):
                    continue
                # Honor date_source='path': some cores deliver a loan file
                # whose name carries no date (e.g. Symitar 'AIRESLOANS.xlsx')
                # but the dated folder does. Resolve the snapshot from the
                # relative path in that case.
                date_src = rel if str(config.get("date_source", "filename")).lower() == "path" else fn
                try:
                    snap = import_data.extract_snapshot_date(date_src, config)
                except Exception:  # noqa: BLE001
                    snap = None
                if snap == target:
                    if fn in seen_names:
                        continue  # avoid basename collisions when staging
                    seen_names.add(fn)
                    found.append((full, base))
        # First folder that yields matches wins — keeps a single-source
        # semantics and avoids importing the same period from two places.
        if found:
            matches = found
            used_source = base
            break
    if not matches:
        return {"ok": False, "period": target, "source_folder": source,
                "error": (f"No loan files for {target} found under {source} "
                          "(or the CU's Archive / Raw_Uploads folders).")}

    tmp = tempfile.mkdtemp(prefix=f"cecl_reimport_{client_short_name}_")
    try:
        # When the snapshot comes from the folder path (date_source='path'),
        # preserve the dated subfolder so the import step can resolve the date
        # too; otherwise stage flat (filename carries the date).
        stage_by_path = str(config.get("date_source", "filename")).lower() == "path"
        for src, base in matches:
            if stage_by_path:
                dst = os.path.join(tmp, os.path.relpath(src, base))
                os.makedirs(os.path.dirname(dst), exist_ok=True)
            else:
                dst = os.path.join(tmp, os.path.basename(src))
            shutil.copy2(src, dst)
        files_processed = import_data.process_client(
            client_short_name, scan_folder_override=tmp)
    finally:
        shutil.rmtree(tmp, ignore_errors=True)

    # Report the resulting DB row count for the period (process_client returns
    # a file count, not a loan-row count).
    row_count = 0
    try:
        from sqlalchemy import create_engine, text as _text
        from cecl_credentials import get_database_url
        _eng = create_engine(get_database_url())
        with _eng.connect() as _c:
            row_count = int(_c.execute(
                _text("SELECT COUNT(*) FROM monthly_loan_data "
                      "WHERE credit_union = :cu AND snapshot_date = :s"),
                {"cu": cu_name, "s": target}).scalar() or 0)
    except Exception:  # noqa: BLE001
        row_count = 0

    return {"ok": True, "period": target, "cu": cu_name,
            "source_folder": used_source,
            "files": [os.path.basename(m) for m, _base in matches],
            "files_processed": int(files_processed or 0),
            "rows": row_count}



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


def _hybrid_flag(cfg: dict[str, Any]) -> dict[str, Any]:
    """Return the ``hybrid_router`` config block (empty when unset)."""
    block = cfg.get("hybrid_router")
    return block if isinstance(block, dict) else {}


def run_reports_hybrid(
    client_short_name: str,
    snapshot_date: str | None = None,
    reports: list[str] | None = None,
) -> tuple[list[str], str, list[str]]:
    """Run reports through the hybrid validation router, then the engine.

    Behaves exactly like :func:`run_reports` but first passes each selected
    report type through the local-first / AI-fallback schema router
    (``cecl_ui.services.hybrid``). The router is **opt-in per CU** via the
    client YAML::

        hybrid_router:
          enabled: true          # default: false (Option A — off)
          proactive: true        # escalate when a CU has no schema yet
          model: claude-opus-4.8

    Returns ``(output_paths, captured_stdout, hybrid_notes)``. When the flag is
    off — or any router dependency is missing — it is a transparent pass-through
    to :func:`run_reports` and ``hybrid_notes`` is empty, so existing behavior
    is unchanged.
    """
    notes: list[str] = []
    try:
        cfg = importlib.import_module("yaml").safe_load(
            (_client_configs_dir()
             / f"{client_short_name}.yaml").read_text(encoding="utf-8")
        ) or {}
    except Exception as exc:  # noqa: BLE001 - never block on config read
        cfg = {}
        notes.append(f"hybrid: config unreadable ({exc}); router skipped")

    flag = _hybrid_flag(cfg)
    if not flag.get("enabled", False):
        paths, log = run_reports(client_short_name, snapshot_date, reports)
        return paths, log, notes

    # Feature is on — run the router (best-effort, never blocks the report).
    try:
        from cecl_ui.services import hybrid

        if not hybrid.is_available():
            notes.append(
                "hybrid enabled but Pydantic is not installed "
                "(pip install pydantic); proceeding without validation."
            )
        else:
            cred = importlib.import_module("cecl_credentials")
            api_key = None
            try:
                api_key = cred.get_anthropic_api_key()
            except Exception:  # noqa: BLE001
                api_key = None
            from sqlalchemy import create_engine
            engine = create_engine(cred.get_database_url())
            router = hybrid.HybridReportRouter(
                engine,
                api_key=api_key,
                model=flag.get("model"),
                proactive=bool(flag.get("proactive", True)),
            )
            selected = reports or [
                r for r, on in (cfg.get("reports") or {}).items() if on
            ]
            for rt in selected:
                try:
                    decision = router.resolve(cfg, rt)
                    notes.append(decision.summary())
                except Exception as exc:  # noqa: BLE001 - per-type isolation
                    notes.append(f"{rt}: hybrid error ({exc}); proceeding")
    except Exception as exc:  # noqa: BLE001 - router import/setup failure
        notes.append(f"hybrid router unavailable ({exc}); proceeding")

    paths, log = run_reports(client_short_name, snapshot_date, reports)
    return paths, log, notes


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


def pool_score_coverage_for_cu(
    cu_name: str, snapshot_date: str | None = None
) -> dict[str, dict]:
    """Return per-pool credit-score coverage for a credit union.

    Takes the credit-union *name* (``monthly_loan_data.credit_union``) rather
    than a config short-name, so it works during the setup wizard before the
    YAML config is saved. For the latest (or given) snapshot, returns::

        {loan_pool: {"total": float, "scored": float, "frac": float}}

    where ``scored`` is the balance of loans carrying a real credit score
    (``original_fico_score > 0``) and ``frac`` is ``scored / total``. Returns
    an empty dict when the CU has no imported data or the DB is unavailable.
    """
    cu_name = (cu_name or "").strip()
    if not cu_name:
        return {}
    try:
        cfg_mod = importlib.import_module("cecl_credentials")
        url = cfg_mod.get_database_url()
        from sqlalchemy import create_engine, text
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
                    return {}
            rows = conn.execute(
                text(
                    "SELECT loan_pool, "
                    "       COALESCE(SUM(current_balance), 0) AS total, "
                    "       COALESCE(SUM(CASE WHEN original_fico_score > 0 "
                    "                        THEN current_balance ELSE 0 END), 0) "
                    "         AS scored "
                    "FROM monthly_loan_data "
                    "WHERE credit_union = :cu AND snapshot_date = :s "
                    "GROUP BY loan_pool"
                ),
                {"cu": cu_name, "s": snapshot_date},
            ).fetchall()
        out: dict[str, dict] = {}
        for pool, total, scored in rows:
            if pool is None:
                continue
            total = float(total or 0)
            scored = float(scored or 0)
            out[str(pool)] = {
                "total": total,
                "scored": scored,
                "frac": (scored / total) if total else 0.0,
            }
        return out
    except Exception:
        return {}


def balance_coverage_for_cu(cu_name: str) -> dict:
    """Return the snapshot coverage of a CU's imported loan data.

    Used to describe how far the historical monthly balances (derived from the
    imported loan extracts) actually span, so a validator can see the range::

        {"start": "YYYY-MM-DD", "end": "YYYY-MM-DD",
         "months": int, "pools": int}

    Returns ``{}`` when the CU has no data or the DB is unavailable.
    """
    cu_name = (cu_name or "").strip()
    if not cu_name:
        return {}
    try:
        cfg_mod = importlib.import_module("cecl_credentials")
        url = cfg_mod.get_database_url()
        from sqlalchemy import create_engine, text
        eng = create_engine(url)
        with eng.connect() as conn:
            row = conn.execute(
                text(
                    "SELECT MIN(snapshot_date), MAX(snapshot_date), "
                    "       COUNT(DISTINCT snapshot_date), "
                    "       COUNT(DISTINCT loan_pool) "
                    "FROM monthly_loan_data WHERE credit_union = :cu"
                ),
                {"cu": cu_name},
            ).fetchone()
        if not row or row[0] is None:
            return {}
        return {
            "start": str(row[0]),
            "end": str(row[1]),
            "months": int(row[2] or 0),
            "pools": int(row[3] or 0),
        }
    except Exception:
        return {}


def _co_recov_source_group(source: str) -> str:
    """Map a raw history ``source`` tag to a human-readable origin group."""
    s = (source or "").strip().lower()
    if s.startswith("5300"):
        return "NCUA 5300 Call Report backfill"
    if s.startswith("monthly"):
        return "Monthly charge-off / recovery summary files"
    if s.startswith("workbook"):
        return "Uploaded historical workbook"
    if not s:
        return "Unlabeled source"
    return source.split(":", 1)[0]


def co_recov_coverage_for_cu(cu_name: str) -> dict:
    """Return DB-stored charge-off / recovery coverage grouped by origin.

    Describes what actually landed in ``loan_code_chargeoff_history`` /
    ``loan_code_recovery_history`` for the CU, grouped by source (NCUA 5300
    backfill, monthly summaries, uploaded workbook). Note: for CUs whose
    charge-offs/recoveries are parsed from files at report time
    (``historical_file_formats``), that portion is NOT in these tables — the
    provenance view combines this DB picture with the configured CU files.

    Shape::

        {"chargeoff": {"groups": [{"group","years","total","rows"}, ...],
                       "year_min": int, "year_max": int, "total": float},
         "recovery":  {...}}

    Returns ``{}`` on any error / no data.
    """
    cu_name = (cu_name or "").strip()
    if not cu_name:
        return {}
    try:
        cfg_mod = importlib.import_module("cecl_credentials")
        url = cfg_mod.get_database_url()
        from sqlalchemy import create_engine, text
        eng = create_engine(url)
        out: dict[str, Any] = {}
        with eng.connect() as conn:
            for key, tbl, amt in (
                ("chargeoff", "loan_code_chargeoff_history", "chargeoff_amount"),
                ("recovery", "loan_code_recovery_history", "recovery_amount"),
            ):
                rows = conn.execute(
                    text(
                        "SELECT source, "
                        "       MIN(EXTRACT(YEAR FROM as_of_date))::int, "
                        "       MAX(EXTRACT(YEAR FROM as_of_date))::int, "
                        "       COUNT(*), COALESCE(SUM(" + amt + "), 0) "
                        "FROM " + tbl + " WHERE cu = :cu GROUP BY source"
                    ),
                    {"cu": cu_name},
                ).fetchall()
                if not rows:
                    out[key] = {"groups": [], "year_min": None,
                                "year_max": None, "total": 0.0}
                    continue
                agg: dict[str, dict] = {}
                y_min = None
                y_max = None
                total = 0.0
                for source, ymn, ymx, cnt, amt_sum in rows:
                    grp = _co_recov_source_group(source)
                    g = agg.setdefault(
                        grp, {"group": grp, "year_min": ymn, "year_max": ymx,
                              "total": 0.0, "rows": 0}
                    )
                    if ymn is not None:
                        g["year_min"] = ymn if g["year_min"] is None else min(g["year_min"], ymn)
                        y_min = ymn if y_min is None else min(y_min, ymn)
                    if ymx is not None:
                        g["year_max"] = ymx if g["year_max"] is None else max(g["year_max"], ymx)
                        y_max = ymx if y_max is None else max(y_max, ymx)
                    g["total"] += float(amt_sum or 0)
                    g["rows"] += int(cnt or 0)
                    total += float(amt_sum or 0)
                groups = []
                for g in agg.values():
                    yr = ""
                    if g["year_min"] is not None:
                        yr = (f"{g['year_min']}" if g["year_min"] == g["year_max"]
                              else f"{g['year_min']}\u2013{g['year_max']}")
                    groups.append({"group": g["group"], "years": yr,
                                   "total": round(g["total"], 2), "rows": g["rows"]})
                groups.sort(key=lambda x: x["group"])
                out[key] = {"groups": groups, "year_min": y_min,
                            "year_max": y_max, "total": round(total, 2)}
        return out
    except Exception:
        return {}


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

    workspace_root = _workspace_root()
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
