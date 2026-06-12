"""Run quarterly reports for an already-configured CU."""
from __future__ import annotations

import calendar
import re
import shutil
from datetime import date
from pathlib import Path

from flask import (
    Blueprint, current_app, flash, redirect, render_template, request,
    send_file, url_for,
)
from werkzeug.utils import secure_filename

from cecl_ui.services import config_service, pipeline_service


run_bp = Blueprint("run", __name__)


def _normalize_folder_path(folder_str: str) -> str:
    """Trim whitespace + surrounding quotes a user may have pasted in."""
    s = (folder_str or "").strip()
    if len(s) >= 2 and s[0] in ('"', "'") and s[-1] == s[0]:
        s = s[1:-1].strip()
    return s


def _period_to_snapshot(period: str) -> str | None:
    """Convert a 'YYYY-MM' period string to ISO month-end date."""
    s = (period or "").strip()
    if len(s) != 7 or s[4] != "-":
        return None
    try:
        y, m = int(s[:4]), int(s[5:7])
        last = calendar.monthrange(y, m)[1]
        return date(y, m, last).isoformat()
    except (ValueError, calendar.IllegalMonthError):
        return None


@run_bp.route("/", methods=["GET"])
def select_cu():
    clients = config_service.list_existing_clients(current_app.config["WORKSPACE_ROOT"])
    return render_template("run/select_cu.html", clients=clients)


@run_bp.route("/<short_name>", methods=["GET"])
def client_dashboard(short_name: str):
    ws = current_app.config["WORKSPACE_ROOT"]
    cfg = config_service.load_client_config(ws, short_name)
    snapshots = pipeline_service.list_snapshots_for_cu(short_name)
    upload_dir = config_service.raw_uploads_dir(ws) / short_name
    upload_dir.mkdir(parents=True, exist_ok=True)
    pending = [p.name for p in upload_dir.glob("*") if p.is_file()]
    return render_template(
        "run/client_dashboard.html",
        short_name=short_name,
        cfg=cfg,
        snapshots=snapshots,
        pending_files=pending,
    )


@run_bp.route("/<short_name>/upload", methods=["POST"])
def upload(short_name: str):
    ws = current_app.config["WORKSPACE_ROOT"]
    upload_dir = config_service.raw_uploads_dir(ws) / short_name
    upload_dir.mkdir(parents=True, exist_ok=True)
    files = request.files.getlist("files")
    saved = 0
    for f in files:
        if not f or not f.filename:
            continue
        fn = secure_filename(f.filename)
        f.save(upload_dir / fn)
        saved += 1
    flash(f"Uploaded {saved} file(s) to Raw_Uploads/{short_name}/.", "success")
    return redirect(url_for("run.client_dashboard", short_name=short_name))


@run_bp.route("/<short_name>/upload/delete", methods=["POST"])
def delete_upload(short_name: str):
    """Delete a single file from ``Raw_Uploads/<short_name>/``.

    Filename is supplied as a form field. Path traversal is rejected by
    resolving the candidate against the upload directory.
    """
    ws = current_app.config["WORKSPACE_ROOT"]
    upload_dir = (config_service.raw_uploads_dir(ws) / short_name).resolve()
    target_name = (request.form.get("filename") or "").strip()
    if not target_name:
        flash("Nothing to delete.", "error")
    else:
        try:
            candidate = (upload_dir / target_name).resolve()
            if upload_dir != candidate and upload_dir not in candidate.parents:
                flash(f"Refusing to delete outside upload folder: {target_name}", "error")
            elif not candidate.exists() or not candidate.is_file():
                flash(f"File not found: {target_name}", "error")
            else:
                candidate.unlink()
                flash(f"Deleted {candidate.name}.", "success")
        except OSError as exc:
            flash(f"Could not delete {target_name}: {exc}", "error")
    return redirect(url_for("run.client_dashboard", short_name=short_name))


@run_bp.route("/<short_name>/reclassify_pool", methods=["POST"])
def reclassify_pool(short_name: str):
    """Bulk-move ``monthly_loan_data`` rows from one pool name to another
    for this CU without re-importing. Useful when a code-mapping change
    is made *after* import and the original loan-data files are no
    longer available to re-import.
    """
    from_pool = (request.form.get("from_pool") or "").strip()
    to_sel = (request.form.get("to_pool_select") or "").strip()
    to_new = (request.form.get("to_pool_new") or "").strip()
    to_pool = to_new if to_sel == "__new__" else to_sel
    snapshot = (request.form.get("snapshot_date") or "").strip() or None
    if not from_pool or not to_pool:
        flash("Reclassify needs both source and destination pool.", "error")
        return redirect(url_for("run.edit_pool_map", short_name=short_name))
    if to_pool == "Ignore":
        to_pool = "Exclude"  # match the runtime exclusion sentinel
    try:
        n = pipeline_service.reclassify_loan_pool(
            short_name, from_pool, to_pool, snapshot
        )
        scope = f"snapshot {snapshot}" if snapshot else "all snapshots"
        flash(
            f"Reclassified {n} loan row(s): {from_pool!r} → {to_pool!r} "
            f"({scope}).",
            "success",
        )
    except Exception as exc:  # noqa: BLE001
        flash(f"Reclassify failed: {exc}", "error")
    return redirect(url_for("run.edit_pool_map", short_name=short_name))


@run_bp.route("/<short_name>/import", methods=["POST"])
def import_data(short_name: str):
    try:
        count = pipeline_service.run_import(short_name)
        if count:
            flash(f"Imported {count} loan file(s) into the database.", "success")
        else:
            flash(
                "No loan files were imported. Check that uploaded filenames "
                "match the file pattern AND that a snapshot date can be parsed "
                "from the filename (e.g. '2026-03' or 'Mar_2026').",
                "warning",
            )
    except Exception as exc:  # noqa: BLE001
        flash(f"Import failed: {exc}", "error")
    return redirect(url_for("run.client_dashboard", short_name=short_name))


@run_bp.route("/<short_name>/reports", methods=["POST"])
def reports(short_name: str):
    snap = request.form.get("snapshot_date") or None
    selected: list[str] = []
    for r in ("tct", "vizo", "vizo_supp"):
        if request.form.get(r) == "on":
            selected.append(r)
    impdet = request.form.get("impdet") == "on"

    outputs: list[str] = []
    errors: list[str] = []

    if selected:
        try:
            paths, log = pipeline_service.run_reports(short_name, snap, selected)
            outputs.extend(paths)
            if not paths:
                # Surface the underlying ERROR line(s) so the user isn't
                # left staring at an empty results page.
                lines = [ln for ln in (log or "").splitlines() if ln.strip()]
                err_lines = [
                    ln for ln in lines
                    if "ERROR" in ln or "Error" in ln or "Warning" in ln
                    or "No data" in ln or "no data" in ln
                    or "not found" in ln or "Not found" in ln
                ]
                if err_lines:
                    msg = "; ".join(err_lines)
                elif lines:
                    # Show the tail of the log so the user can see what happened.
                    msg = "No reports were generated. Last log lines: " + " | ".join(lines[-8:])
                else:
                    msg = "No reports were generated (no log output captured)."
                errors.append(f"Standard reports: {msg}")
        except Exception as exc:  # noqa: BLE001
            errors.append(f"Standard reports: {exc}")

    if impdet:
        try:
            p = pipeline_service.run_impdet_report(short_name, snap)
            if p:
                outputs.append(p)
        except Exception as exc:  # noqa: BLE001
            errors.append(f"Impaired/Deteriorated: {exc}")

    if not selected and not impdet:
        flash("Pick at least one report to generate.", "error")
        return redirect(url_for("run.client_dashboard", short_name=short_name))

    # If reports succeeded and this CU has no Completed-setup entry yet,
    # auto-adopt its YAML config so it shows up alongside wizard-built
    # CUs on the home page. Lazy-imported to avoid a circular import
    # between home.py and run.py at module load.
    if outputs and not errors:
        try:
            from cecl_ui.routes.home import adopt_config_to_completed
            adopt_config_to_completed(
                current_app.config["WORKSPACE_ROOT"], short_name,
            )
        except Exception:  # noqa: BLE001
            # Auto-adoption is a UX-nicety; never block the results page.
            pass

    return render_template(
        "run/results.html",
        short_name=short_name,
        outputs=outputs,
        errors=errors,
    )


@run_bp.route("/download")
def download():
    """Stream a generated report file back to the browser."""
    target = request.args.get("path", "")
    if not target:
        return ("Missing path", 400)
    p = Path(target).resolve()
    ws = Path(current_app.config["WORKSPACE_ROOT"]).resolve()
    # Containment check — only allow files inside the workspace root.
    try:
        p.relative_to(ws)
    except ValueError:
        return ("Forbidden", 403)
    if not p.exists():
        return ("Not found", 404)
    return send_file(p, as_attachment=True, download_name=p.name)


@run_bp.route("/<short_name>/new-quarter", methods=["GET", "POST"])
def new_quarter(short_name: str):
    """Streamlined per-quarter ingest + run.

    GET: shows a single-page form (period picker + multi-file upload +
    optional folder path + report checkboxes pre-checked from the CU's
    saved ``cfg.reports``).

    POST: copies the supplied files into ``Raw_Uploads/<short>/``,
    invokes ``run_import``, then runs the selected reports for the
    requested period and renders the existing results template.
    """
    ws = current_app.config["WORKSPACE_ROOT"]
    cfg = config_service.load_client_config(ws, short_name)
    upload_dir = config_service.raw_uploads_dir(ws) / short_name
    upload_dir.mkdir(parents=True, exist_ok=True)

    if request.method == "GET":
        return render_template(
            "run/new_quarter.html",
            short_name=short_name,
            cfg=cfg,
        )

    # POST
    period = (request.form.get("period") or "").strip()
    snapshot = _period_to_snapshot(period)
    if not snapshot:
        flash("Pick a valid reporting period (YYYY-MM).", "error")
        return redirect(url_for("run.new_quarter", short_name=short_name))

    # ``stage`` distinguishes the initial submit ("") from a resubmit after
    # the user has filled in the new-loan-code mapping form
    # (``stage="mapped"``). On the resubmit we skip file staging (files are
    # already on disk from phase 1) and instead apply the user-supplied
    # mappings to ``cfg["pool_map"]`` before continuing with import +
    # reports.
    stage = (request.form.get("stage") or "").strip()

    if stage == "mapped":
        posted_codes = request.form.getlist("map_code")
        posted_pools = request.form.getlist("map_pool")
        posted_new_pools = request.form.getlist("map_pool_new")
        new_map = dict(cfg.get("pool_map") or {})
        added = 0
        for i, code in enumerate(posted_codes):
            code = (code or "").strip()
            if not code:
                continue
            sel = (posted_pools[i] if i < len(posted_pools) else "").strip()
            np_name = (posted_new_pools[i] if i < len(posted_new_pools) else "").strip()
            pool = np_name if sel == "__new__" else sel
            if not pool:
                continue
            new_map[code] = pool
            added += 1
        if added:
            cfg["pool_map"] = new_map
            try:
                config_service.save_client_config(ws, short_name, cfg, overwrite=True)
                flash(f"Mapped {added} new loan code(s) and saved YAML.", "success")
            except (OSError, ValueError) as exc:
                flash(f"Could not save pool mappings: {exc}", "error")
                return redirect(url_for("run.new_quarter", short_name=short_name))
        # Skip phase-1 staging block; jump to the import step below.
        saved = 0
        copied = 0
        folder_skipped: list[str] = []
        # Fall through to the import + report block.

    folder_raw = _normalize_folder_path(request.form.get("folder_path") or "") if stage != "mapped" else ""

    # 1) Save uploaded files
    saved = 0
    files = request.files.getlist("files") if stage != "mapped" else []
    for f in files:
        if not f or not f.filename:
            continue
        fn = secure_filename(f.filename)
        f.save(upload_dir / fn)
        saved += 1

    # 2) Copy from folder path (recursive, common spreadsheet formats).
    # Subfolder layouts like 2026-Q1/{Jan,Feb,Mar} 2026/ are common, so
    # walk the tree and flatten every spreadsheet into Raw_Uploads/.
    copied = 0
    folder_skipped: list[str] = []
    if folder_raw:
        src = Path(folder_raw)
        if not src.is_dir():
            flash(f"Folder not found: {folder_raw}", "error")
            return redirect(url_for("run.new_quarter", short_name=short_name))
        allowed = {".xlsx", ".xlsm", ".xls", ".csv"}
        for entry in src.rglob("*"):
            if not entry.is_file():
                continue
            if entry.name.startswith("~$") or entry.name.startswith("."):
                continue
            if entry.suffix.lower() not in allowed:
                folder_skipped.append(entry.name)
                continue
            try:
                shutil.copy2(entry, upload_dir / entry.name)
                copied += 1
            except OSError as exc:
                flash(f"Could not copy {entry.name}: {exc}", "error")

    if saved == 0 and copied == 0 and stage != "mapped":
        flash(
            "No files were uploaded or copied. Pick at least one file or "
            "supply a folder path.",
            "error",
        )
        return redirect(url_for("run.new_quarter", short_name=short_name))

    if stage != "mapped":
        flash(
            f"Staged {saved} uploaded + {copied} folder file(s) into Raw_Uploads/{short_name}/.",
            "success",
        )
    if folder_skipped:
        flash(
            f"Skipped non-spreadsheet file(s): {', '.join(folder_skipped[:6])}"
            + ("…" if len(folder_skipped) > 6 else ""),
            "info",
        )

    # 2b) Update YAML's ``report_period`` to the user-selected period so
    # the import step targets the correct snapshot. Filenames without a
    # parseable date (e.g. ``February Loan File Upload.xlsx``) fall back
    # to ``report_period``'s year inside ``import_data.process_client``;
    # if we left a stale period in the YAML those files would be
    # mis-attributed to a previous year.
    if period and (cfg.get("report_period") or "") != period:
        try:
            cfg["report_period"] = period
            config_service.save_client_config(ws, short_name, cfg, overwrite=True)
        except (OSError, ValueError) as exc:
            flash(f"Could not persist report_period={period}: {exc}", "warning")

    # 2c) On the initial submit, scan the staged files for any new
    # ``loan_pool_code`` values that aren't yet in ``cfg["pool_map"]``.
    # If any are found, re-render the page with a mapping section so the
    # user can assign each new code to an existing pool BEFORE we run
    # import/reports — otherwise those codes would silently fall through
    # to ``default_pool`` (often "Ignore") and the report would show a
    # spurious "Ignore" pool. The user fills the form, submits with
    # ``stage="mapped"``, and we land at the top of this handler again
    # to apply the mappings + continue.
    if stage != "mapped":
        try:
            unmapped = _scan_unmapped_loan_codes(cfg, short_name)
        except Exception:  # noqa: BLE001
            unmapped = []
        if unmapped:
            report_selection = {
                r: (request.form.get(r) == "on")
                for r in ("tct", "vizo", "vizo_supp", "impdet")
            }
            return render_template(
                "run/new_quarter.html",
                short_name=short_name,
                cfg=cfg,
                stage="awaiting_mappings",
                unmapped_codes=unmapped,
                all_pools=_all_known_pools(cfg, short_name),
                staged_period=period,
                staged_reports=report_selection,
                staged_count=saved + copied,
            )

    # 3) Import
    try:
        n_imported = pipeline_service.run_import(short_name)
        if n_imported:
            flash(f"Imported {n_imported} loan file(s).", "success")
        else:
            # Surface a hint listing which staged spreadsheets did NOT
            # match the configured file_pattern so the user can spot
            # version-suffix / spelling drift quickly.
            unmatched: list[str] = []
            try:
                pat = (cfg or {}).get("file_pattern") or ""
                if pat:
                    rx = re.compile(pat)
                    for entry in upload_dir.iterdir():
                        if not entry.is_file():
                            continue
                        if entry.suffix.lower() not in {".xlsx", ".xlsm", ".xls", ".csv"}:
                            continue
                        if not rx.match(entry.name):
                            unmatched.append(entry.name)
            except (re.error, OSError):
                unmatched = []
            base = (
                "No loan files were imported. Confirm filenames match the "
                "configured file_pattern and that a snapshot date can be "
                "parsed."
            )
            if unmatched:
                preview = ", ".join(unmatched[:5])
                more = f" (+{len(unmatched) - 5} more)" if len(unmatched) > 5 else ""
                base += (
                    f" Files staged but not matched by file_pattern: {preview}{more}."
                    " Re-run the wizard's Sample step (or hand-edit the YAML)"
                    " to relax the pattern."
                )
            flash(base, "warning")
    except Exception as exc:  # noqa: BLE001
        flash(f"Import failed: {exc}", "error")
        return redirect(url_for("run.new_quarter", short_name=short_name))

    # 4) Run reports — selection comes from form (defaults seeded from cfg)
    selected: list[str] = []
    for r in ("tct", "vizo", "vizo_supp"):
        if request.form.get(r) == "on":
            selected.append(r)
    impdet = request.form.get("impdet") == "on"

    outputs: list[str] = []
    errors: list[str] = []
    if selected:
        try:
            paths, log = pipeline_service.run_reports(short_name, snapshot, selected)
            outputs.extend(paths)
            if not paths:
                lines = [ln for ln in (log or "").splitlines() if ln.strip()]
                err_lines = [
                    ln for ln in lines
                    if "ERROR" in ln or "Error" in ln or "Warning" in ln
                    or "No data" in ln or "no data" in ln
                    or "not found" in ln or "Not found" in ln
                ]
                if err_lines:
                    msg = "; ".join(err_lines)
                elif lines:
                    msg = "No reports were generated. Last log lines: " + " | ".join(lines[-8:])
                else:
                    msg = "No reports were generated (no log output captured)."
                errors.append(f"Standard reports: {msg}")
        except Exception as exc:  # noqa: BLE001
            errors.append(f"Standard reports: {exc}")

    if impdet:
        try:
            p = pipeline_service.run_impdet_report(short_name, snapshot)
            if p:
                outputs.append(p)
        except Exception as exc:  # noqa: BLE001
            errors.append(f"Impaired/Deteriorated: {exc}")

    if not selected and not impdet:
        flash("Pick at least one report to generate.", "warning")
        return redirect(url_for("run.client_dashboard", short_name=short_name))

    return render_template(
        "run/results.html",
        short_name=short_name,
        outputs=outputs,
        errors=errors,
    )


@run_bp.route("/<short_name>/settings", methods=["GET", "POST"])
def edit_settings(short_name: str):
    """Edit pool-level settings (risk_rated, ACL months, mgmt-adj defaults)
    and global mgmt_adj for a previously-configured CU.

    Reads/writes through ``config_service.get_pools`` / ``set_pools`` so the
    canonical ``pools`` block is the source of truth; legacy
    ``pool_order`` / ``not_risk_rated`` / ``excluded_pools`` /
    ``acl_months_by_pool`` fields are kept in sync automatically.
    """
    ws = current_app.config["WORKSPACE_ROOT"]
    cfg = config_service.load_client_config(ws, short_name)

    pools_list = config_service.get_pools(cfg)
    known_names = {p["name"] for p in pools_list}

    # Pull in any pool names the DB knows about that aren't yet declared
    # (legacy CUs whose pools were inferred from pool_map only). They land
    # at the end as risk-rated by default; the user can flip them.
    db_pools = [n for n in pipeline_service.list_pools_for_cu(short_name)
                if n and n not in known_names and n != "Exclude"]
    for n in db_pools:
        pools_list.append({
            "name": n, "risk_rated": True, "acl_months": None,
            "excluded": False, "use_default_mgmt_adj": False,
        })

    if request.method == "POST":
        posted_names = request.form.getlist("pool_name")
        seen: set[str] = set()
        new_pools: list[dict] = []
        for i, n in enumerate(posted_names):
            n = (n or "").strip()
            if not n or n in seen:
                continue
            seen.add(n)
            rr_val = (request.form.get(f"rr_{i}") or "yes").strip().lower()
            acl_raw = (request.form.get(f"acl_{i}") or "").strip()
            try:
                acl_int = int(acl_raw) if acl_raw else None
                if acl_int is not None and acl_int <= 0:
                    acl_int = None
            except ValueError:
                acl_int = None
            new_pools.append({
                "name": n,
                "risk_rated": rr_val == "yes",
                "acl_months": acl_int,
                "excluded": rr_val == "excluded",
                "use_default_mgmt_adj": False,
            })

        config_service.set_pools(cfg, new_pools)

        # Global mgmt_adj knobs
        ma = dict(cfg.get("mgmt_adj") or {})
        for key in ("ltv_baseline", "probability_factor"):
            v = (request.form.get(f"ma_{key}") or "").strip()
            if v:
                try:
                    ma[key] = float(v)
                except ValueError:
                    pass
        cfg["mgmt_adj"] = ma

        config_service.save_client_config(ws, short_name, cfg, overwrite=True)
        flash("Pool settings saved.", "success")
        return redirect(url_for("run.edit_settings", short_name=short_name))

    rows = [
        {
            "name": p["name"],
            "risk_rated": p["risk_rated"] and not p["excluded"],
            "excluded": p["excluded"],
            "acl_months": p["acl_months"] or "",
        }
        for p in pools_list
    ]
    return render_template(
        "run/edit_settings.html",
        short_name=short_name,
        cfg=cfg,
        rows=rows,
        mgmt_adj=cfg.get("mgmt_adj") or {},
    )


def _resolves_to_pool(code: str, pool_map: dict, split: str) -> bool:
    """Mirror generate_report's pool-code resolution: case-insensitive
    direct hit, then split on ``pool_code_split`` and try each part."""
    if not code:
        return False
    pm_lower = {str(k).strip().lower(): v for k, v in (pool_map or {}).items()}
    if str(code).strip().lower() in pm_lower:
        return True
    if split:
        import re as _re
        for part in _re.split(_re.escape(split), str(code)):
            if part.strip().lower() in pm_lower:
                return True
    return False


def _scan_unmapped_loan_codes(cfg: dict, short_name: str) -> list[dict]:
    """Scan available loan-data files and return raw loan_pool_code values
    that don't resolve to any pool via ``cfg['pool_map']``.

    Looks under ``Raw_Uploads/<short_name>/`` and (as a fallback) the
    ``credit_pull.source_folder`` for files matching any
    ``loan_data_extracts[].file_pattern`` (or the top-level ``file_pattern``).
    Reads just the configured ``loan_pool_code`` column from each match
    and aggregates distinct unmapped codes with row counts.
    """
    import os
    import re
    import pandas as pd

    ws = current_app.config["WORKSPACE_ROOT"]
    pool_map = cfg.get("pool_map") or {}
    split = cfg.get("pool_code_split") or ""

    # Build (file_pattern_re, column_mappings) tuples. Try per-extract
    # patterns first (most specific), then fall back to the top-level
    # ``file_pattern`` + top-level ``column_mappings`` so files that no
    # extract matches still get scanned.
    def _compile_list(value):
        """str | list[str] -> list[Pattern]; invalid regexes silently skipped."""
        if not value:
            return []
        items = [value] if isinstance(value, str) else [
            v for v in value if isinstance(v, str) and v
        ]
        out: list = []
        for s in items:
            try:
                out.append(re.compile(s, re.IGNORECASE))
            except re.error:
                continue
        return out

    extracts = cfg.get("loan_data_extracts") or []
    patterns: list[tuple] = []
    for e in extracts:
        pres = _compile_list((e or {}).get("file_pattern"))
        if not pres:
            continue
        cm = (e or {}).get("column_mappings") or {}
        # One tuple per regex so the existing single-Pattern consumer
        # below works unchanged for multi-pattern extracts.
        for pre in pres:
            patterns.append((pre, cm))
    top_pres = _compile_list(cfg.get("file_pattern"))
    if top_pres:
        top_cm = cfg.get("column_mappings") or {}
        for pre in top_pres:
            patterns.append((pre, top_cm))
    if not patterns:
        return []

    # Folders to scan.
    folders: list[Path] = []
    raw_dir = config_service.raw_uploads_dir(ws) / short_name
    if raw_dir.is_dir():
        folders.append(raw_dir)
    cp_folder = ((cfg.get("credit_pull") or {}).get("source_folder") or "").strip()
    if cp_folder and os.path.isdir(cp_folder):
        folders.append(Path(cp_folder))

    counts: dict[str, int] = {}
    samples: dict[str, set] = {}

    def _read_column(path: Path, col_mappings: dict) -> pd.Series | None:
        # Static code (e.g. credit-card extract). Treat as a single-value
        # column so it gets counted properly.
        static = (col_mappings.get("loan_pool_code_static") or "").strip()
        col = col_mappings.get("loan_pool_code")
        try:
            suffix = path.suffix.lower()
            if static:
                # Read just enough to know how many rows there are.
                if suffix == ".csv":
                    df = pd.read_csv(path, usecols=[0])
                else:
                    df = pd.read_excel(path, usecols=[0])
                return pd.Series([static] * len(df))
            if not col:
                return None
            if suffix == ".csv":
                df = pd.read_csv(path, usecols=[col], dtype=str)
            elif suffix in (".xlsx", ".xls"):
                df = pd.read_excel(path, usecols=[col], dtype=str)
            else:
                return None
            return df[col]
        except Exception:
            return None

    seen_files: set[str] = set()
    for folder in folders:
        for p in folder.iterdir():
            if not p.is_file() or p.name.startswith("~$"):
                continue
            key = p.name.lower()
            if key in seen_files:
                continue
            for pre, cm in patterns:
                if not pre.search(p.name):
                    continue
                series = _read_column(p, cm)
                if series is None:
                    continue
                seen_files.add(key)
                for raw in series.dropna().astype(str):
                    code = raw.strip()
                    if not code or _resolves_to_pool(code, pool_map, split):
                        continue
                    counts[code] = counts.get(code, 0) + 1
                    samples.setdefault(code, set())
                break

    # Also surface any loan_code values present in the historical-data DB
    # tables (charge-offs, recoveries, delinquency) that don't resolve to
    # any pool. These come from per-month aggregation, 5300 backfills, and
    # the wizard's historical-data steps, and are what populate the Vizo
    # Display CO-Recov-DQ tab. Unmapped codes there silently fall back to
    # the default pool, so the user needs to see them.
    db_counts = _scan_unmapped_codes_from_db(cfg, pool_map, split)
    for code, n in db_counts.items():
        counts[code] = counts.get(code, 0) + n

    return [
        {"code": c, "count": counts[c]}
        for c in sorted(counts, key=lambda k: (-counts[k], k.lower()))
    ]


def _scan_unmapped_codes_from_db(cfg: dict, pool_map: dict,
                                 split: str) -> dict[str, int]:
    """Pull distinct loan_code values from the historical-data DB tables
    (chargeoffs, recoveries, delinquency) that don't resolve to any pool
    via ``cfg['pool_map']``. Returns ``{code: row_count}``.

    Returns an empty dict on any error (missing tables, no DB access).
    """
    cu = (cfg.get("credit_union") or "").strip()
    if not cu:
        return {}
    try:
        from cecl_credentials import get_database_url
        from sqlalchemy import create_engine, text as _sql_text
    except Exception:  # noqa: BLE001
        return {}
    try:
        eng = create_engine(get_database_url())
    except Exception:  # noqa: BLE001
        return {}

    counts: dict[str, int] = {}
    for table in ("loan_code_history",
                  "loan_code_chargeoff_history",
                  "loan_code_recovery_history",
                  "loan_code_delinquency_history"):
        try:
            with eng.begin() as conn:
                rows = conn.execute(
                    _sql_text(
                        f"SELECT loan_code, COUNT(*) AS n "
                        f"FROM {table} WHERE cu = :cu "
                        f"GROUP BY loan_code"
                    ),
                    {"cu": cu},
                ).fetchall()
        except Exception:  # noqa: BLE001
            # Table may not exist yet; skip.
            continue
        for r in rows:
            code = (r[0] or "").strip()
            if not code or _resolves_to_pool(code, pool_map, split):
                continue
            counts[code] = counts.get(code, 0) + int(r[1] or 0)
    return counts


def _all_known_pools(cfg: dict, short_name: str) -> list[str]:
    """Return a sorted list of every pool name we know about for this CU.

    Used to populate the pool drop-down on the loan-code-mapping editor so
    the user can re-assign codes to any existing pool. Driven off the
    canonical ``pools`` registry (with a fallback to legacy fields and the
    DB so older configs still surface every pool).
    """
    pool_set: set[str] = {p["name"] for p in config_service.get_pools(cfg)}
    pool_set.update((cfg.get("pool_map") or {}).values())
    pool_set.update(pipeline_service.list_pools_for_cu(short_name))
    pool_set.discard("")
    pool_set.discard(None)
    pool_set.discard("Exclude")
    pool_set.discard("Ignore")
    return sorted(p for p in pool_set if p)


@run_bp.route("/<short_name>/pool_map", methods=["GET", "POST"])
def edit_pool_map(short_name: str):
    """Edit the loan-code -> pool mapping for a previously-configured CU.

    Lets the user add new loan codes, change which pool a code maps to,
    add a brand-new pool name on the fly, delete codes, and update the
    ``default_pool`` (where unmapped codes land).
    """
    ws = current_app.config["WORKSPACE_ROOT"]
    cfg = config_service.load_client_config(ws, short_name)

    if request.method == "POST":
        codes = request.form.getlist("code")
        pool_select = request.form.getlist("pool_select")
        pool_new = request.form.getlist("pool_new")
        delete_flags = request.form.getlist("delete")
        delete_set = set(delete_flags)

        new_map: dict[str, str] = {}
        seen: set[str] = set()
        for i, code in enumerate(codes):
            code = (code or "").strip()
            if not code or str(i) in delete_set:
                continue
            sel = (pool_select[i] if i < len(pool_select) else "").strip()
            new_pool = (pool_new[i] if i < len(pool_new) else "").strip()
            # "__new__" sentinel means "use the typed-in name to its right"
            pool = new_pool if sel == "__new__" else sel
            if not pool:
                continue
            if code in seen:
                continue
            seen.add(code)
            new_map[code] = pool

        cfg["pool_map"] = new_map

        default_sel = (request.form.get("default_pool_select") or "").strip()
        default_new = (request.form.get("default_pool_new") or "").strip()
        default_pool = default_new if default_sel == "__new__" else default_sel
        if default_pool:
            cfg["default_pool"] = default_pool

        config_service.save_client_config(ws, short_name, cfg, overwrite=True)
        flash(f"Loan code mapping saved ({len(new_map)} codes).", "success")
        flash(
            "Existing snapshots in the database still use the pool labels "
            "from when they were imported. Re-run Import Data on the client "
            "dashboard so mapping changes (including any new Ignore "
            "assignments) take effect in reports.",
            "warning",
        )
        return redirect(url_for("run.edit_pool_map", short_name=short_name))

    pool_map = cfg.get("pool_map") or {}
    # Sort by code for predictable display.
    rows = [{"code": k, "pool": v} for k, v in sorted(pool_map.items(), key=lambda kv: str(kv[0]).lower())]
    all_pools = _all_known_pools(cfg, short_name)
    unmapped_codes = _scan_unmapped_loan_codes(cfg, short_name)
    pool_distribution = pipeline_service.pool_distribution_for_cu(short_name)
    snapshots = pipeline_service.list_snapshots_for_cu(short_name)
    return render_template(
        "run/edit_pool_map.html",
        short_name=short_name,
        cfg=cfg,
        rows=rows,
        all_pools=all_pools,
        default_pool=cfg.get("default_pool", ""),
        unmapped_codes=unmapped_codes,
        pool_distribution=pool_distribution,
        snapshots=snapshots,
    )
