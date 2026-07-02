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

from cecl_ui.services import balance_check as balance_check_service
from cecl_ui.services import config_service, pipeline_service
from cecl_ui.services import impaired_check_service


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


def _refresh_staged_monthly_balance(
    cfg: dict, folder_src: str, upload_dir: Path,
) -> tuple[str, str] | None:
    """Look among the just-staged files for an updated copy of the CU's
    monthly balance workbook and refresh the wizard-staged saved_path.

    The wizard captures ``monthly_balance.saved_path`` as a copy under
    ``%TEMP%\\cecl_ui_monthly_bal\\`` at setup time. Subsequent quarters'
    updated Historical Balance Sheets files (with new month columns
    appended) are dropped into the source folder but never touched by
    Run New Quarter's default file staging (which only routes files
    through Raw_Uploads/ for the loan importer). This helper closes
    that gap: when we detect a filename-matching file among the newly
    staged uploads OR in the source folder, and it's newer than the
    existing staged copy, we copy it over ``saved_path`` so the
    Balance Adjustment Review + report engine both pick up May/June/
    etc. columns without a manual wizard re-visit.

    Returns ``(flash_msg, category)`` or ``None`` when no refresh
    happened. Best-effort: never raises.
    """
    mb = (cfg or {}).get("monthly_balance") or {}
    saved_path = str(mb.get("saved_path") or "").strip()
    if not saved_path:
        return None
    canonical = str(mb.get("filename") or "").strip() or Path(saved_path).name

    def _norm(name: str) -> str:
        return re.sub(r"[\s_\-]+", "_", str(name).strip().lower())

    target = _norm(canonical)
    if not target:
        return None

    candidates: list[Path] = []
    try:
        for entry in upload_dir.iterdir():
            if entry.is_file() and _norm(entry.name) == target:
                candidates.append(entry)
    except OSError:
        pass
    if folder_src:
        try:
            for entry in Path(folder_src).rglob("*"):
                if entry.is_file() and _norm(entry.name) == target:
                    candidates.append(entry)
        except OSError:
            pass
    if not candidates:
        return None
    try:
        candidates.sort(key=lambda p: p.stat().st_mtime, reverse=True)
    except OSError:
        return None
    src = candidates[0]
    dest = Path(saved_path)
    # Skip when the staged copy is already at least as fresh (avoids
    # clobbering a user's hand-edited local copy on repeat submits).
    try:
        if dest.exists() and dest.stat().st_mtime >= src.stat().st_mtime:
            return None
    except OSError:
        pass
    try:
        dest.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy2(src, dest)
    except OSError as exc:
        return (
            f"Found updated monthly balance file '{src.name}' but could "
            f"not refresh the staged copy: {exc}",
            "warning",
        )
    return (
        f"Refreshed staged monthly balance file from '{src.name}' — "
        "Balance Adjustment Review will use the latest data.",
        "success",
    )


def _ingest_monthly_co_recov(cfg: dict, upload_dir: Path) -> str | None:
    """Fold any staged *monthly-summary* Charge-Off / Recovery files into
    the historical CO/recovery DB.

    Some CUs (e.g. Central Susquehanna) ship a small monthly workbook that
    summarizes charge-offs and recoveries *by loan type* — no per-loan
    account numbers, and no date column: the reporting month is encoded in
    the filename (``…Totals 1225.xlsx`` → Dec 2025). The wizard's
    Historical step already knows how to aggregate these, but a quarterly
    Run New Quarter pass never touched them, so a freshly-dropped month
    never reached the loss-rate calc.

    This detects such files among the staged uploads and runs the exact
    same aggregation the wizard uses
    (:func:`monthly_co_recov_aggregator.aggregate_all`), which upserts one
    month per file (DELETE+INSERT, so re-runs are idempotent).

    A file qualifies only when it has BOTH a loan-code/type column and a
    charge-off or recovery *amount* column, AND **no** date column (that
    last check keeps transaction-level CO/recovery workbooks — which need
    per-row date binning — out of this filename-dated path). Files that
    match the CU's loan ``file_pattern`` are skipped so the main loan
    extract is never mis-ingested.

    Returns a human-readable summary string, or ``None`` when nothing
    qualified. Best-effort — never raises.
    """
    try:
        from cecl_ui.services import monthly_co_recov_aggregator as agg
        from cecl_ui.services import extract_hist_service
    except Exception:  # noqa: BLE001
        return None

    cu = (cfg or {}).get("credit_union") or ""
    if not cu:
        return None

    # Compile the loan file_pattern so we can exclude the loan extract.
    loan_rx = None
    pat = (cfg or {}).get("file_pattern") or ""
    if pat:
        try:
            loan_rx = re.compile(pat, re.IGNORECASE)
        except re.error:
            loan_rx = None

    allowed = {".xlsx", ".xlsm", ".xls", ".csv"}
    co_files: list[dict] = []
    recov_files: list[dict] = []
    for entry in sorted(upload_dir.iterdir()):
        if not entry.is_file() or entry.suffix.lower() not in allowed:
            continue
        if entry.name.startswith("~$"):
            continue
        if loan_rx is not None and loan_rx.search(entry.name):
            continue  # this is the loan extract, not a CO/recovery summary
        try:
            insp = agg.inspect_file(entry)
        except Exception:  # noqa: BLE001
            continue
        if not insp.get("ok"):
            continue
        sug = insp.get("suggested") or {}
        has_code = bool((sug.get("code") or "").strip())
        has_co = bool((sug.get("co_amount") or "").strip())
        has_recov = bool((sug.get("recov_amount") or "").strip())
        has_date = bool((sug.get("date") or "").strip())
        if not has_code or has_date or not (has_co or has_recov):
            continue
        # The month must be resolvable from the filename, else we can't
        # place it — skip rather than let it fall back to file mtime.
        det = extract_hist_service.detect_as_of_date(entry.name, entry)
        if not det.get("date") or det.get("source") != "filename":
            continue
        item = {"name": entry.name, "path": str(entry)}
        if has_co:
            co_files.append(item)
        if has_recov:
            recov_files.append(item)

    if not co_files and not recov_files:
        return None

    state = {
        "credit_union": cu,
        "pool_map": (cfg or {}).get("pool_map") or {},
        "hist_scan": {
            "monthly_co_files": co_files,
            "monthly_recov_files": recov_files,
        },
    }
    parts: list[str] = []
    for kind, files in (("co", co_files), ("recov", recov_files)):
        if not files:
            continue
        try:
            res = agg.aggregate_all(state, kind)
        except Exception:  # noqa: BLE001
            continue
        if res.get("ok") and res.get("months_written"):
            label = "charge-off" if kind == "co" else "recovery"
            months = sorted(res["months_written"])
            span = (f"{months[0][:7]}…{months[-1][:7]}"
                    if len(months) > 1 else months[0][:7])
            parts.append(
                f"{len(months)} {label} month(s) [{span}]"
            )
    if not parts:
        return None
    return "Ingested monthly CO/recovery summary file(s): " + ", ".join(parts) + "."


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
    acl_cfg = cfg.get("acl") or {}
    acl_history = acl_cfg.get("history") or {}
    return render_template(
        "run/client_dashboard.html",
        short_name=short_name,
        cfg=cfg,
        snapshots=snapshots,
        pending_files=pending,
        acl_history=acl_history,
        report_period_end=_period_to_snapshot(cfg.get("report_period") or ""),
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
    for r in ("tct", "vizo", "vizo_supp", "mgmt_adj_napkin"):
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


def _month_end_iso(value: str) -> str | None:
    """Normalize a user-supplied date to an ISO month-end string.

    Accepts ``YYYY-MM`` or a full ``YYYY-MM-DD`` and always returns the
    last calendar day of that month (e.g. ``2026-06`` → ``2026-06-30``),
    matching how ACL history / snapshot dates are keyed elsewhere.
    """
    s = (value or "").strip()
    if not s:
        return None
    if len(s) == 7 and s[4] == "-":
        return _period_to_snapshot(s)
    try:
        d = date.fromisoformat(s)
    except ValueError:
        return None
    last = calendar.monthrange(d.year, d.month)[1]
    return date(d.year, d.month, last).isoformat()


@run_bp.route("/<short_name>/set-acl", methods=["POST"])
def set_acl(short_name: str):
    """Override the ACL (allowance) balance for a single month-end without
    re-running the wizard.

    Writes the value into ``cfg['acl']['history'][<month-end>]`` — the
    authoritative source the report engine merges with precedence over
    the monthly-balance file (see ``_merge_acl_history`` in
    ``generate_report.py``). When the CU's ACL source is ``manual`` the
    ``acl.manual`` map is kept in sync too, and the top-level
    ``acl_balance`` is refreshed when the edited month is the current
    ``report_period``.
    """
    ws = current_app.config["WORKSPACE_ROOT"]
    cfg = config_service.load_client_config(ws, short_name)

    month_end = _month_end_iso(request.form.get("acl_month") or "")
    raw_amount = (request.form.get("acl_balance") or "").strip().replace(",", "")
    if not month_end:
        flash("Pick a valid month-end for the ACL balance.", "error")
        return redirect(url_for("run.client_dashboard", short_name=short_name))
    try:
        amount = float(raw_amount)
    except ValueError:
        flash(f"Invalid ACL balance: {raw_amount!r}", "error")
        return redirect(url_for("run.client_dashboard", short_name=short_name))
    if amount < 0:
        flash("ACL balance cannot be negative.", "error")
        return redirect(url_for("run.client_dashboard", short_name=short_name))

    acl = dict(cfg.get("acl") or {})
    history = dict(acl.get("history") or {})
    history[month_end] = amount
    acl["history"] = history
    # Keep the manual map aligned when the CU sources ACL manually.
    if (acl.get("source") or "").strip().lower() == "manual":
        manual = dict(acl.get("manual") or {})
        manual[month_end] = amount
        acl["manual"] = manual
    cfg["acl"] = acl

    # Refresh the scalar top-level balance when editing the current period.
    period_end = _period_to_snapshot(cfg.get("report_period") or "")
    if period_end and period_end == month_end:
        cfg["acl_balance"] = amount

    config_service.save_client_config(ws, short_name, cfg, overwrite=True)
    flash(
        f"ACL balance for {month_end} set to ${amount:,.2f}. "
        f"Re-generate the report to apply it.",
        "success",
    )
    return redirect(url_for("run.client_dashboard", short_name=short_name))


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

    if stage == "balance_checked":
        # User reviewed the per-pool balance comparison and clicked
        # "Continue → Run Reports". Files were already staged and
        # imported on the prior submit; jump straight to the report
        # block. Mirror the ``mapped`` skip pattern.
        saved = 0
        copied = 0
        folder_skipped = []
        # Fall through to the report block (import is skipped below).

    if stage == "balance_remapped":
        # User adjusted one or more loan-code → pool mappings on the
        # Balance Adjustment Review page. Apply the new mappings to
        # cfg["pool_map"], persist to YAML, and re-run import so the
        # ``monthly_loan_data`` snapshot re-stamps each row with the
        # new pool name. Then fall through to the balance-check
        # intercept again so the user sees updated totals and can
        # either keep editing or click Continue → Run Reports.
        posted_codes = request.form.getlist("map_code")
        posted_pools = request.form.getlist("map_pool")
        new_map = dict(cfg.get("pool_map") or {})
        changed = 0
        for i, code in enumerate(posted_codes):
            code = (code or "").strip()
            if not code:
                continue
            new_pool = (posted_pools[i] if i < len(posted_pools) else "").strip()
            old_pool = new_map.get(code, "")
            if new_pool == old_pool:
                continue
            new_map[code] = new_pool
            changed += 1
        if changed:
            cfg["pool_map"] = new_map
            try:
                config_service.save_client_config(ws, short_name, cfg, overwrite=True)
            except (OSError, ValueError) as exc:
                flash(f"Could not save pool mappings: {exc}", "error")
                return redirect(url_for("run.new_quarter", short_name=short_name))
            # Re-import so the DB snapshot reflects the new pool_map.
            try:
                pipeline_service.run_import(short_name)
                flash(
                    f"Reassigned {changed} loan code(s) and re-imported. "
                    "Review the updated balances below.",
                    "success",
                )
            except Exception as exc:  # noqa: BLE001
                flash(f"Re-import after remap failed: {exc}", "error")
                return redirect(url_for("run.new_quarter", short_name=short_name))
        else:
            flash("No mapping changes detected.", "info")
        saved = 0
        copied = 0
        folder_skipped = []
        # Fall through; import block is skipped via _skip_stage below.

    if stage == "impaired_checked":
        # User reviewed the impaired-loans verification page and clicked
        # "Continue → Run Reports". Files were already staged and
        # imported on the prior submit; jump straight to the report
        # block. Mirror the ``balance_checked`` skip pattern.
        saved = 0
        copied = 0
        folder_skipped = []
        # Fall through to the report block (import is skipped below).

    if stage == "impaired_remapped":
        # User adjusted one or more loan-code → pool mappings on the
        # Impaired Loans Verification page. Apply the new mappings to
        # cfg["pool_map"], persist to YAML, and re-run import so the
        # loan-data extract lookup sees the new pools. Then fall
        # through to the impaired-check intercept so the user sees
        # updated rows.
        posted_codes = request.form.getlist("map_code")
        posted_pools = request.form.getlist("map_pool")
        new_map = dict(cfg.get("pool_map") or {})
        changed = 0
        for i, code in enumerate(posted_codes):
            code = (code or "").strip()
            if not code:
                continue
            new_pool = (posted_pools[i] if i < len(posted_pools) else "").strip()
            old_pool = new_map.get(code, "")
            if new_pool == old_pool:
                continue
            new_map[code] = new_pool
            changed += 1
        if changed:
            cfg["pool_map"] = new_map
            try:
                config_service.save_client_config(ws, short_name, cfg, overwrite=True)
            except (OSError, ValueError) as exc:
                flash(f"Could not save pool mappings: {exc}", "error")
                return redirect(url_for("run.new_quarter", short_name=short_name))
            try:
                pipeline_service.run_import(short_name)
                flash(
                    f"Reassigned {changed} loan code(s) and re-imported. "
                    "Review the updated impaired loans below.",
                    "success",
                )
            except Exception as exc:  # noqa: BLE001
                flash(f"Re-import after remap failed: {exc}", "error")
                return redirect(url_for("run.new_quarter", short_name=short_name))
        else:
            flash("No mapping changes detected.", "info")
        saved = 0
        copied = 0
        folder_skipped = []
        # Fall through; import block is skipped via _skip_stage below.

    _skip_stage = stage in (
        "mapped", "balance_checked", "balance_remapped",
        "impaired_checked", "impaired_remapped",
    )
    folder_raw = _normalize_folder_path(request.form.get("folder_path") or "") if not _skip_stage else ""

    # 1) Save uploaded files
    saved = 0
    files = request.files.getlist("files") if not _skip_stage else []
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

    if saved == 0 and copied == 0 and not _skip_stage:
        flash(
            "No files were uploaded or copied. Pick at least one file or "
            "supply a folder path.",
            "error",
        )
        return redirect(url_for("run.new_quarter", short_name=short_name))

    if not _skip_stage:
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

    # 2a) Refresh the wizard-staged monthly balance workbook if the
    # source folder / uploads contain a newer copy (e.g. Vizo dropped
    # an updated Historical Balance Sheets file with May/June columns
    # into 2026-Q2/June 2026/). Best-effort — never blocks the run.
    if not _skip_stage:
        try:
            _mb_msg = _refresh_staged_monthly_balance(cfg, folder_raw, upload_dir)
            if _mb_msg:
                flash(_mb_msg[0], _mb_msg[1])
        except Exception as _mb_exc:  # noqa: BLE001
            flash(
                f"Monthly-balance auto-refresh skipped: {_mb_exc}",
                "warning",
            )

    # 2a-2) Monthly Charge-Off / Recovery ingestion. When a CU ships a
    # by-loan-type monthly summary workbook (date encoded in the filename,
    # no account numbers), fold any freshly-staged month(s) into the
    # historical CO/recovery DB so the loss-rate calc reflects them. Mirrors
    # the wizard's Historical step; idempotent per month. Best-effort.
    if not _skip_stage:
        try:
            _cr_msg = _ingest_monthly_co_recov(cfg, upload_dir)
            if _cr_msg:
                flash(_cr_msg, "success")
        except Exception as _cr_exc:  # noqa: BLE001
            flash(
                f"Monthly CO/recovery ingestion skipped: {_cr_exc}",
                "warning",
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

    # 2b-2) ACL Balance override — Phase 9.36. The user can override the
    # default ACL source for this run from the new-quarter form:
    #
    #   * ``acl_mode="manual"`` → record ``acl_manual_value`` as the
    #     ACL Balance for this quarter, written to
    #     ``cfg['acl']['history'][snapshot]`` so ``_merge_acl_history``
    #     picks it up at report time. The YAML history entry persists
    #     so subsequent quarters can carry it forward; if the user
    #     wants to revert they re-pick "default" on a future run with
    #     a non-matching snapshot (the old history entry stays for
    #     that quarter, which is what an audit trail should do).
    #   * ``acl_mode="solr"`` → enable the 5300 fallback with the
    #     user's preferred field code (default ``AAS0048``). Written to
    #     ``cfg['acl']['use_5300_fallback']=True`` and
    #     ``cfg['acl']['solr_fields']=[<field>]`` so the engine probes
    #     only the chosen field. Legacy fields (A007/A718A3/A718A5/A719)
    #     are NOT auto-appended — per user direction AAS0048 is the
    #     canonical source. Users who need legacy backstops can list
    #     multiple fields by submitting them comma-separated.
    #
    # Empty / "default" mode leaves cfg['acl'] untouched, preserving
    # whatever the wizard configured originally.
    if not _skip_stage:
        acl_mode = (request.form.get("acl_mode") or "default").strip()
        if acl_mode in ("manual", "solr"):
            acl_block = dict(cfg.get("acl") or {})
            persist = False
            flash_msg = ""
            if acl_mode == "manual":
                raw = (request.form.get("acl_manual_value") or "").strip()
                try:
                    val = float(raw.replace(",", "")) if raw else 0.0
                except ValueError:
                    val = 0.0
                if val > 0:
                    hist_map = dict(acl_block.get("history") or {})
                    hist_map[snapshot] = round(val, 2)
                    acl_block["history"] = hist_map
                    cfg["acl"] = acl_block
                    cfg["acl_balance"] = round(val, 2)
                    persist = True
                    flash_msg = (
                        f"Saved manual ACL Balance ${val:,.2f} for "
                        f"{snapshot} (acl.history[{snapshot}])."
                    )
                else:
                    flash(
                        "Manual ACL mode selected but no positive value "
                        "provided; using YAML defaults instead.",
                        "warning",
                    )
            else:  # acl_mode == "solr"
                raw_field = (request.form.get("acl_solr_field") or "").strip()
                # Allow comma-separated lists so power users can add
                # legacy backstops (e.g. "AAS0048,A007,A718A3"). Single
                # field is the common case.
                if raw_field:
                    field_order = [
                        f.strip().upper()
                        for f in raw_field.split(",")
                        if f.strip()
                    ]
                else:
                    field_order = []
                if field_order:
                    acl_block["use_5300_fallback"] = True
                    acl_block["solr_fields"] = field_order
                    cfg["acl"] = acl_block
                    persist = True
                    if len(field_order) == 1:
                        flash_msg = (
                            f"Enabled NCUA 5300 ACL fallback with "
                            f"field '{field_order[0]}'."
                        )
                    else:
                        flash_msg = (
                            f"Enabled NCUA 5300 ACL fallback (probe order: "
                            f"{', '.join(field_order)})."
                        )
                else:
                    flash(
                        "5300 ACL mode selected but no field code provided; "
                        "using YAML defaults instead.",
                        "warning",
                    )
            if persist:
                try:
                    config_service.save_client_config(
                        ws, short_name, cfg, overwrite=True)
                    if flash_msg:
                        flash(flash_msg, "success")
                except (OSError, ValueError) as exc:
                    flash(
                        f"Could not persist ACL override: {exc}",
                        "warning",
                    )

    # 2c) On the initial submit, scan the staged files for any new
    # ``loan_pool_code`` values that aren't yet in ``cfg["pool_map"]``.
    # If any are found, re-render the page with a mapping section so the
    # user can assign each new code to an existing pool BEFORE we run
    # import/reports — otherwise those codes would silently fall through
    # to ``default_pool`` (often "Ignore") and the report would show a
    # spurious "Ignore" pool. The user fills the form, submits with
    # ``stage="mapped"``, and we land at the top of this handler again
    # to apply the mappings + continue.
    if not _skip_stage:
        try:
            unmapped = _scan_unmapped_loan_codes(cfg, short_name)
        except Exception:  # noqa: BLE001
            unmapped = []
        if unmapped:
            report_selection = {
                r: (request.form.get(r) == "on")
                for r in ("tct", "vizo", "vizo_supp", "impdet", "mgmt_adj_napkin")
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

    # 3) Import (skip on balance_checked / balance_remapped /
    # impaired_checked / impaired_remapped resubmits — import already
    # ran on the prior submit or inside the remap branch).
    if stage in ("balance_checked", "balance_remapped",
                 "impaired_checked", "impaired_remapped"):
        n_imported = 0  # already imported on prior submit
    else:
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

    # 3b) Balance Adjustment intercept — show per-pool monthly-vs-loan
    # balance comparison so the user can review discrepancies BEFORE
    # the reports are produced. Mirrors wizard Step 14. On the resubmit
    # the user clicks "Continue → Run Reports" and we fall through.
    # Also skip once the user has advanced PAST balance to the impaired
    # intercept — otherwise clicking Continue on the impaired page would
    # bounce them back to the balance page (loop).
    if stage not in ("balance_checked", "impaired_checked",
                     "impaired_remapped"):
        try:
            comparison = balance_check_service.compare_run(
                cfg, snapshot, short_name=short_name)
        except Exception:  # noqa: BLE001
            comparison = None
        # Only intercept when we actually have a populated comparison.
        # If the loan side failed (empty DB) or both sides are empty,
        # skip the review and let the report block surface its own
        # error — we don't want to dead-end the user on a useless page.
        if comparison and comparison.get("ok") and comparison.get("rows"):
            report_selection = {
                r: (request.form.get(r) == "on")
                for r in ("tct", "vizo", "vizo_supp", "impdet", "mgmt_adj_napkin")
            }
            # Build the pool-name dropdown for the inline remap UI.
            try:
                pool_choices = _all_known_pools(cfg, short_name)
            except Exception:  # noqa: BLE001
                pool_choices = []
            return render_template(
                "run/new_quarter.html",
                short_name=short_name,
                cfg=cfg,
                stage="awaiting_balance_check",
                comparison=comparison,
                staged_period=period,
                staged_reports=report_selection,
                pool_choices=pool_choices,
            )

    # 3c) Impaired Loans Verification intercept — show the parsed
    # impaired workbook (if one is on disk for this period) so the user
    # can confirm each impaired loan resolves to a real pool BEFORE the
    # reports are produced. Mirrors wizard Step 16. Skipped when no
    # impaired workbook is found — most CUs don't ship one and there's
    # nothing to review.
    if stage != "impaired_checked":
        try:
            impaired_result = impaired_check_service.verify_for_run(
                cfg, snapshot, short_name=short_name)
        except Exception:  # noqa: BLE001
            impaired_result = None
        if (impaired_result
                and impaired_result.get("ok")
                and impaired_result.get("found")
                and (impaired_result.get("data_rows")
                     or impaired_result.get("validation", {}).get(
                         "unmapped_codes"))):
            report_selection = {
                r: (request.form.get(r) == "on")
                for r in ("tct", "vizo", "vizo_supp", "impdet", "mgmt_adj_napkin")
            }
            return render_template(
                "run/new_quarter.html",
                short_name=short_name,
                cfg=cfg,
                stage="awaiting_impaired_check",
                impaired=impaired_result,
                staged_period=period,
                staged_reports=report_selection,
            )
        # Surface non-fatal parse errors so the user isn't blindsided if
        # the workbook exists but couldn't be parsed. Report generation
        # will still proceed — the report engine has its own tolerance
        # for a missing/malformed impaired file.
        if impaired_result and impaired_result.get("error"):
            flash(
                "Could not verify the impaired-loans workbook for "
                f"{snapshot}: {impaired_result['error']}. Reports will "
                "continue with best-effort defaults.",
                "warning",
            )

    # 4) Run reports — selection comes from form (defaults seeded from cfg)
    selected: list[str] = []
    for r in ("tct", "vizo", "vizo_supp", "mgmt_adj_napkin"):
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


def _impaired_data_dir(cfg: dict, ws: str | Path, short_name: str) -> Path:
    """Resolve the folder the report engine searches for the standalone
    Impaired Loans workbook.

    ``generate_report.find_standalone_impaired_file`` walks
    ``cfg['data_directory']`` (relative paths are resolved against the
    workspace root). When the CU has no ``data_directory`` configured we
    fall back to ``Raw_Uploads/<short>/`` — which is where the loan
    importer already stages files, and is inside the workspace so the
    download containment check keeps working.
    """
    data_dir = str((cfg or {}).get("data_directory") or "").strip()
    if data_dir:
        p = Path(data_dir)
        if not p.is_absolute():
            p = Path(ws) / data_dir
        return p
    return config_service.raw_uploads_dir(ws) / short_name


def _canonical_impaired_name(cfg: dict, snapshot: str, original: str) -> str:
    """Build a filename the strict impaired-file matcher will accept.

    ``find_standalone_impaired_file`` matches
    ``^YYYY\\s*[-_]?\\s*MM.*Impaired[\\s_-]+Loans.*\\.xlsx$`` (plus a few
    looser fallbacks). Saving under ``YYYY-MM Impaired Loans - <CU>.xlsx``
    guarantees a hit regardless of what the CU named the file they sent —
    so the very next report run discovers it with no config change.
    Preserves the uploaded file's extension.
    """
    ext = Path(original).suffix.lower() or ".xlsx"
    prefix = (snapshot or "")[:7]  # YYYY-MM
    cu = str((cfg or {}).get("credit_union") or "").strip()
    cu = re.sub(r'[\\/:*?"<>|]', "", cu).strip()  # strip path-hostile chars
    base = f"{prefix} Impaired Loans"
    if cu:
        base += f" - {cu}"
    return secure_filename(f"{base}{ext}") or f"{prefix}_Impaired_Loans{ext}"


def _supersede_existing_impaired(
    target_dir: Path, snapshot: str, keep_name: str,
) -> list[str]:
    """Archive any impaired workbooks for the same period already on disk.

    ``find_standalone_impaired_file`` walks the folder and returns the
    *first* ``.xlsx`` that matches the period — filesystem order is not
    deterministic, so leaving an older workbook for the same month next
    to the freshly-uploaded one risks the report engine silently reading
    the stale file. To make "upload a new report" mean "use this one",
    we rename every other same-period impaired workbook by appending a
    ``.superseded-<UTC timestamp>`` suffix. The suffix pushes the name
    past the ``\\.xlsx$`` anchor so it no longer matches discovery, while
    nothing is deleted (the analyst can recover it if needed).

    ``keep_name`` (the canonical target we're about to write) is never
    archived. Returns the list of archived basenames for flash display.
    Best-effort — logs nothing, raises nothing.
    """
    prefix = (snapshot or "")[:7]  # YYYY-MM
    if not prefix or len(prefix) != 7:
        return []
    year, month = prefix[:4], prefix[5:7]
    # Same loose date matching the finder uses (YYYY-MM, MM-YYYY, month name).
    month_names = {
        "01": r"jan(?:uary)?", "02": r"feb(?:ruary)?", "03": r"mar(?:ch)?",
        "04": r"apr(?:il)?", "05": r"may", "06": r"jun(?:e)?",
        "07": r"jul(?:y)?", "08": r"aug(?:ust)?",
        "09": r"sep(?:t(?:ember)?)?", "10": r"oct(?:ober)?",
        "11": r"nov(?:ember)?", "12": r"dec(?:ember)?",
    }
    mnum = str(int(month))
    alts = [
        rf"{year}[-_.\s]*0?{mnum}(?!\d)",
        rf"(?<!\d)0?{mnum}[-_.\s]*{year}",
    ]
    mname = month_names.get(month, "")
    if mname:
        alts.append(rf"{mname}[-_.\s]*{year}")
        alts.append(rf"{year}[-_.\s]*{mname}")
    try:
        date_rx = re.compile("|".join(alts), re.IGNORECASE)
        imp_rx = re.compile(r"impaired[\s_\-]+loans", re.IGNORECASE)
    except re.error:
        return []

    from datetime import datetime, timezone
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    archived: list[str] = []
    try:
        entries = list(target_dir.iterdir())
    except OSError:
        return []
    for entry in entries:
        if not entry.is_file():
            continue
        name = entry.name
        if name == keep_name:
            continue
        if name.startswith("~$") or name.upper().startswith("DNU"):
            continue
        if "WARM" in name.upper():
            continue
        if not name.lower().endswith(".xlsx"):
            continue
        if not (imp_rx.search(name) and date_rx.search(name)):
            continue
        archive_to = entry.with_name(f"{name}.superseded-{stamp}")
        try:
            entry.rename(archive_to)
            archived.append(name)
        except OSError:
            continue
    return archived


@run_bp.route("/<short_name>/impaired", methods=["GET"])
def upload_impaired_form(short_name: str):
    """Standalone Impaired Loans upload page.

    Lets a user drop a *new* Impaired Loans workbook for an existing CU
    without re-running the wizard or the full new-quarter pass. The file
    is saved into the CU's ``data_directory`` under a canonical name so
    the report engine's ``find_standalone_impaired_file`` picks it up on
    the next report run. Shows the currently-discovered impaired file (if
    any) for the selected period.
    """
    ws = current_app.config["WORKSPACE_ROOT"]
    cfg = config_service.load_client_config(ws, short_name)
    snapshots = pipeline_service.list_snapshots_for_cu(short_name)
    # Default the period to the CU's configured report_period, else the
    # latest snapshot in the DB.
    default_snap = ""
    rp = _period_to_snapshot((cfg or {}).get("report_period") or "")
    if rp:
        default_snap = rp
    elif snapshots:
        default_snap = snapshots[0]
    return render_template(
        "run/impaired_upload.html",
        short_name=short_name,
        cfg=cfg,
        snapshots=snapshots,
        default_snap=default_snap,
        verification=None,
    )


@run_bp.route("/<short_name>/impaired", methods=["POST"])
def upload_impaired(short_name: str):
    """Save an uploaded Impaired Loans workbook and verify it.

    The file is written into the CU's impaired-search folder using a
    canonical, matcher-friendly name for the requested period, then
    parsed + enriched via ``impaired_check_service.verify_for_run`` so
    the user immediately sees the parsed rows and loan-data match counts.
    No import or report generation is triggered — the file simply becomes
    the authoritative impaired list for that period's next report run.
    """
    ws = current_app.config["WORKSPACE_ROOT"]
    cfg = config_service.load_client_config(ws, short_name)

    snapshot = (request.form.get("snapshot_date") or "").strip()
    if not snapshot:
        # Allow a YYYY-MM period picker too.
        snapshot = _period_to_snapshot(request.form.get("period") or "") or ""
    if not snapshot:
        flash("Pick the reporting period this impaired file is for.", "error")
        return redirect(url_for("run.upload_impaired_form", short_name=short_name))

    f = request.files.get("impaired_file")
    if not f or not f.filename:
        flash("Choose an Impaired Loans file to upload.", "error")
        return redirect(url_for("run.upload_impaired_form", short_name=short_name))

    ext = Path(f.filename).suffix.lower()
    if ext not in (".xlsx", ".xlsm"):
        flash(
            "Impaired Loans workbook must be an .xlsx or .xlsm file "
            f"(got {ext or 'no extension'}).",
            "error",
        )
        return redirect(url_for("run.upload_impaired_form", short_name=short_name))

    target_dir = _impaired_data_dir(cfg, ws, short_name)
    try:
        target_dir.mkdir(parents=True, exist_ok=True)
    except OSError as exc:
        flash(f"Could not access the CU's data folder ({target_dir}): {exc}",
              "error")
        return redirect(url_for("run.upload_impaired_form", short_name=short_name))

    dest_name = _canonical_impaired_name(cfg, snapshot, f.filename)
    dest = target_dir / dest_name

    # Archive any other impaired workbook already present for the same
    # period so discovery deterministically resolves to this upload.
    try:
        archived = _supersede_existing_impaired(target_dir, snapshot, dest_name)
    except Exception:  # noqa: BLE001
        archived = []
    if archived:
        preview = ", ".join(archived[:4]) + ("…" if len(archived) > 4 else "")
        flash(
            f"Archived {len(archived)} older impaired file(s) for this "
            f"period (renamed with a .superseded suffix): {preview}",
            "info",
        )

    try:
        f.save(dest)
    except OSError as exc:
        flash(f"Could not save the impaired file: {exc}", "error")
        return redirect(url_for("run.upload_impaired_form", short_name=short_name))

    flash(
        f"Saved impaired workbook as '{dest_name}' in "
        f"{target_dir}. It will be used by the next report run for "
        f"{snapshot}.",
        "success",
    )

    # Verify + preview so the user can confirm the file parses and its
    # loans resolve to real pools before running reports.
    verification = None
    try:
        verification = impaired_check_service.verify_for_run(
            cfg, snapshot, short_name=short_name)
    except Exception as exc:  # noqa: BLE001
        flash(f"Uploaded, but verification failed: {exc}", "warning")
    if verification and verification.get("error"):
        flash(
            f"Uploaded, but the workbook did not fully verify: "
            f"{verification['error']}",
            "warning",
        )

    snapshots = pipeline_service.list_snapshots_for_cu(short_name)
    return render_template(
        "run/impaired_upload.html",
        short_name=short_name,
        cfg=cfg,
        snapshots=snapshots,
        default_snap=snapshot,
        verification=verification,
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
        # Guard: credit_pull.source_folder can point at a SHARED wizard temp
        # staging area (e.g. ``%TEMP%\\cecl_ui_samples``) that holds sample
        # loan files uploaded for OTHER credit unions during setup. Scanning
        # it here pulls in foreign LOAN TYPE codes (another CU's numeric
        # coding scheme), which then surface as bogus "new loan codes" for
        # THIS CU — with inflated counts summed across every CU's samples.
        # The wizard's temp folders all live under a ``cecl_ui_*`` directory;
        # skip any source folder that resolves inside one. Real credit-pull
        # folders live on the client's network share and never match.
        if "cecl_ui_" not in cp_folder.replace("\\", "/").lower():
            folders.append(Path(cp_folder))

    counts: dict[str, int] = {}
    sources: dict[str, set] = {}
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
                    sources.setdefault(code, set()).add(f"file:{p.name}")
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
        sources.setdefault(code, set()).add("db:historical")

    def _src_label(code: str) -> str:
        srcs = sources.get(code) or set()
        has_file = any(s.startswith("file:") for s in srcs)
        has_db = "db:historical" in srcs
        if has_file and has_db:
            return "loan files + historical data"
        if has_file:
            return "loan files"
        if has_db:
            return "historical data (DB)"
        return ""

    return [
        {"code": c, "count": counts[c], "source": _src_label(c)}
        for c in sorted(counts, key=lambda k: (-counts[k], k.lower()))
    ]


def _scan_unmapped_codes_from_db(cfg: dict, pool_map: dict,
                                 split: str) -> dict[str, int]:
    """Pull distinct loan_code values from the historical-data DB tables
    (chargeoffs, recoveries, delinquency) that don't resolve to any pool
    via ``cfg['pool_map']`` AND don't resolve via the NCUA canonical
    fallback (``_resolve_pool_with_ncua``). Returns ``{code: row_count}``.

    The NCUA fallback check is critical: rows in ``loan_code_*_history``
    populated by 5300 backfills (``source LIKE '5300%'``) are keyed by
    NCUA canonical labels like 'First Liens', 'New Vehicles',
    'Unsecured Credit Card'. Those resolve automatically at report time
    via :func:`generate_report._resolve_pool_with_ncua` — surfacing them
    here as "needs mapping" misleads the user (they have nothing to map;
    the resolution is automatic). Only codes that resolve to NEITHER
    pool_map NOR NCUA fallback are genuinely unmapped.

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

    # Build NCUA canonical -> pool fallback (best-effort; if generate_report
    # isn't importable for any reason we degrade to pool_map-only matching).
    ncua_lookup: dict[str, str] = {}
    try:
        import generate_report as _gr
        ncua_lookup = _gr._build_ncua_canonical_pool_lookup(cfg) or {}
    except Exception:  # noqa: BLE001
        ncua_lookup = {}
    default_pool = (cfg.get("default_pool") or "").strip()

    # Build set of known pool names (lower-cased) from the canonical pools
    # registry. DB historical rows from 5300 backfills (5300-distributed:*,
    # 5300CO:*, 5300REC:*, 5300DQ:*) often carry the ALREADY-RESOLVED pool
    # name as ``loan_code`` (because the backfill writers persist the
    # destination pool, not a raw NCUA label, when one resolves). Surfacing
    # those as "needs mapping" misleads the user — the code IS the pool.
    known_pool_names: set[str] = set()
    try:
        for _p in config_service.get_pools(cfg) or []:
            nm = (_p.get("name") or "").strip().lower()
            if nm:
                known_pool_names.add(nm)
    except Exception:  # noqa: BLE001
        known_pool_names = set()

    def _resolves(code: str) -> bool:
        if _resolves_to_pool(code, pool_map, split):
            return True
        if ncua_lookup:
            key = str(code).strip().lower()
            pool = ncua_lookup.get(key) or ncua_lookup.get(" ".join(key.split()))
            if pool and pool.strip().lower() not in ("ignore", "exclude"):
                return True
        # Code IS already a configured pool name (DB historical rows often
        # carry the resolved pool name as loan_code — nothing to map).
        if known_pool_names and str(code).strip().lower() in known_pool_names:
            return True
        # default_pool resolution: any non-empty default_pool is a deliberate
        # user choice for how unmapped codes should flow. For DB-historical
        # NCUA-canonical aggregations (e.g. 'Agricultural',
        # 'Commercial and Industrial' for a non-commercial CU) the user has
        # already opted into Ignore via default_pool — flagging them every
        # quarter forces the same answer the user already gave. File-side
        # scanning (in the caller) still surfaces genuinely-new codes via
        # the loan_data_extracts path; this DB-only path defers to the
        # configured default.
        if default_pool:
            return True
        return False

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
            if not code or _resolves(code):
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
