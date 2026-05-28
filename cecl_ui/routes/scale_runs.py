"""Routes for running a new quarter against an already-configured SCALE
credit union.

Two modes:

* **Carry historical from prior report** -- copies the most recent
  prior SCALE workbook on disk and overlays only the target quarter's
  data. Fast; what we run when older 5300 filings haven't been
  restated. Calls ``scale_runner.run_quarter_carry_history``.
* **Re-pull all from 5300** -- fresh template, full multi-quarter run.
  Use when the credit union has amended older 5300 filings. Calls
  ``scale_runner.run_multi_quarter``.

Both modes additionally hard-code the Prior ACL block on Executive
Summary-Vizo from the previous quarter's report (see
``runner._inject_prior_acl``).
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

from flask import (
    Blueprint, current_app, flash, redirect, render_template, request,
    session, url_for,
)
from werkzeug.utils import secure_filename

from cecl_ui.routes.setup import STATE_KEY
from cecl_ui.routes.scale_setup import _default_scale_block
from cecl_ui.services.scale import (
    impaired_loader,
    runner as scale_runner,
    runs_service,
    solr_fetcher,
    template_loader,
)


scale_runs_bp = Blueprint(
    "scale_runs", __name__, template_folder="../templates"
)


def _workspace_root() -> str:
    return current_app.config["WORKSPACE_ROOT"]


def _save_run_impaired_upload(
    workspace_root: str, short_name: str, period: str, file_storage,
) -> Path:
    """Save an uploaded impaired-loans workbook alongside the quarter's
    generated reports.

    Lives at ``Generated_Reports/<short>/<period>/uploads/<filename>``
    so it's archived with the report it was used to generate.
    """
    sub = (
        Path(workspace_root) / "Generated_Reports" / short_name
        / period / "uploads"
    )
    sub.mkdir(parents=True, exist_ok=True)
    fn = secure_filename(file_storage.filename or "impaired.xlsx")
    target = sub / fn
    file_storage.save(target)
    return target


@scale_runs_bp.route("/", methods=["GET"])
def index():
    cus = runs_service.list_scale_cus(_workspace_root())
    return render_template("scale_runs/index.html", cus=cus)


def _period_choices() -> list[str]:
    seen: set[str] = set()
    seen.update(template_loader.list_available_map_periods())
    seen.update(template_loader.list_available_template_periods())
    return sorted(seen, reverse=True)


def _default_next_period(latest: str, choices: list[str]) -> str:
    """Pick the most reasonable target period for the New Run form."""
    try:
        nxt_map = {"03": "06", "06": "09", "09": "12", "12": "03"}
        y, m = latest.split("-")
        ny = int(y) + 1 if m == "12" else int(y)
        nm = nxt_map[m]
        candidate = f"{ny}-{nm}"
    except Exception:  # noqa: BLE001
        candidate = latest
    if candidate in choices:
        return candidate
    if latest in choices:
        return latest
    return choices[0] if choices else ""


def _solr_available_periods(state: dict | None) -> dict:
    """Best-effort: ask Solr which quarter-end periods have a 5300 doc
    for the CU's charter. Returns ``{ok, periods:set[str], error}``.
    Caller falls back to the unfiltered template/map period list when
    ok is False (Solr down, no creds, missing charter, etc.).
    """
    if not state:
        return {"ok": False, "periods": set(), "error": "no draft"}
    charter = state.get("charter_number") or state.get("charter")
    sc = state.get("scale") or {}
    solr_url = sc.get("solr_url")
    solr_core = sc.get("solr_core")
    if not (charter and solr_url and solr_core):
        return {
            "ok": False, "periods": set(),
            "error": "missing charter or Solr config in wizard draft",
        }
    try:
        charter_int = int(str(charter).strip())
    except (TypeError, ValueError):
        return {
            "ok": False, "periods": set(),
            "error": f"charter_number is not numeric: {charter!r}",
        }
    return solr_fetcher.list_charter_periods(
        solr_url, solr_core, charter_int,
        charter_field=sc.get("charter_field") or "charter",
        charterdate_field=sc.get("charterdate_field") or "charterdate",
        username=sc.get("solr_user") or None,
        password=sc.get("solr_pass") or None,
    )


@scale_runs_bp.route("/<short_name>", methods=["GET"])
def cu_dashboard(short_name: str):
    workspace_root = _workspace_root()
    all_cus = runs_service.list_scale_cus(workspace_root)
    cu = next((c for c in all_cus if c["short_name"] == short_name), None)
    if cu is None:
        flash(f"No SCALE reports found for {short_name}.", "warning")
        return redirect(url_for("scale_runs.index"))

    state = runs_service.load_state_for_cu(workspace_root, short_name)
    draft_present = state is not None
    all_choices = _period_choices()

    # Filter the dropdown to quarters Solr actually has a 5300 doc for
    # this charter. NCUA publishes 5300 data ~6-10 weeks after each
    # quarter-end, so the template/map period list typically gets ahead
    # of what's actually queryable. When Solr is unreachable we fall
    # back to the full list and surface the error.
    solr_probe = _solr_available_periods(state)
    if solr_probe["ok"]:
        period_choices = [p for p in all_choices if p in solr_probe["periods"]]
        if not period_choices:
            # Defensive: don't strand the user with an empty dropdown if
            # the intersection is empty (e.g. brand-new charter).
            period_choices = all_choices
            solr_probe["error"] = (
                "Solr has no 5300 docs for this charter that match the "
                "available SCALE template/map periods. Showing the full list."
            )
    else:
        period_choices = all_choices
    default_target = _default_next_period(cu["latest_period"], period_choices)

    # Surface the impaired file the wizard currently has saved (used
    # as fallback when no per-run file is uploaded).
    saved_impaired: dict[str, Any] = {}
    if state:
        imp = (state.get("scale") or {}).get("impaired_file") or {}
        if imp.get("saved_path"):
            parsed = imp.get("parsed") or {}
            saved_impaired = {
                "filename": imp.get("uploaded_filename") or "",
                "row_count": parsed.get("row_count", 0),
                "total_balance": parsed.get("total_balance", 0.0),
                "period": parsed.get("period", ""),
            }

    # List on-disk runs (one row per period) for the history panel.
    runs_root = Path(workspace_root) / "Generated_Reports" / short_name
    runs: list[dict[str, Any]] = []
    if runs_root.exists():
        for period_dir in sorted(runs_root.iterdir(), reverse=True):
            if not period_dir.is_dir():
                continue
            files = sorted(period_dir.glob("*CECL_SCALE_*.xlsx"))
            if not files:
                continue
            runs.append({
                "period": period_dir.name,
                "files": [
                    {"name": f.name, "path": str(f)} for f in files
                ],
            })

    return render_template(
        "scale_runs/cu_dashboard.html",
        cu=cu,
        runs=runs,
        period_choices=period_choices,
        default_target=default_target,
        draft_present=draft_present,
        saved_impaired=saved_impaired,
        solr_probe=solr_probe,
    )


@scale_runs_bp.route("/<short_name>/run", methods=["POST"])
def run(short_name: str):
    workspace_root = _workspace_root()
    state = runs_service.load_state_for_cu(workspace_root, short_name)
    if state is None:
        flash(
            f"No saved wizard draft for {short_name}. Re-open the wizard "
            "for this CU and save it once before running new quarters.",
            "error",
        )
        return redirect(url_for("scale_runs.cu_dashboard",
                                short_name=short_name))

    period = (request.form.get("period") or "").strip()
    mode = (request.form.get("mode") or "carry").strip()
    variant = (request.form.get("report_variant") or "").strip()
    try:
        quarters = int(request.form.get("quarters") or 32)
    except (TypeError, ValueError):
        quarters = 32
    quarters = max(1, min(quarters, 80))

    if not period:
        flash("Pick a target period.", "error")
        return redirect(url_for("scale_runs.cu_dashboard",
                                short_name=short_name))

    sc = state.setdefault("scale", {})
    # Seed any missing SCALE keys (solr_url, solr_core, ...) from
    # admin defaults. CUs configured via the regular TCT/Vizo wizard
    # have no ``scale`` block; carry-mode reruns still need Solr
    # connectivity for the target quarter's 5300 data.
    defaults = _default_scale_block()
    for k, v in defaults.items():
        if not sc.get(k):
            sc[k] = v
    sc["period"] = period
    if variant:
        sc["report_variant"] = variant

    # Optional per-run impaired-loans upload. When the user attaches a
    # file on the New Run form, save it under the quarter's output
    # folder and override sc["impaired_file"] for this run only — we
    # do NOT persist the override into the wizard draft (the draft
    # keeps whatever the SCALE wizard last saved).
    impaired_override: dict[str, Any] | None = None
    f = request.files.get("impaired_file")
    if f and f.filename:
        try:
            target = _save_run_impaired_upload(
                workspace_root, short_name, period, f,
            )
        except Exception as exc:  # noqa: BLE001
            flash(f"Impaired file save failed: {exc}", "error")
            return redirect(url_for("scale_runs.cu_dashboard",
                                    short_name=short_name))
        parsed = impaired_loader.parse_file(target)
        if not parsed.get("ok"):
            flash(
                f"Impaired file parse failed: {parsed.get('error') or 'unknown error'}. "
                "Run aborted.",
                "error",
            )
            return redirect(url_for("scale_runs.cu_dashboard",
                                    short_name=short_name))
        impaired_override = {
            "saved_path": str(target),
            "uploaded_filename": f.filename,
            "parsed": parsed,
        }
        sc["impaired_file"] = impaired_override
        flash(
            f"Using uploaded impaired file ({parsed['row_count']} row(s), "
            f"${parsed['total_balance']:,.2f}) for this run.",
            "info",
        )

    if mode == "refetch_all":
        result = scale_runner.run_multi_quarter(state, workspace_root,
                                                quarters=quarters)
    else:
        result = scale_runner.run_quarter_carry_history(state, workspace_root)

    # Stash result on the live session under a dedicated key so we can
    # render it without polluting the wizard's setup_state.
    sess_state = session.get(STATE_KEY) or {}
    sess_state["scale_runs_last"] = {
        "short_name": short_name,
        "mode": mode,
        "result": result,
        "impaired_override": impaired_override,
    }
    session[STATE_KEY] = sess_state

    if result.get("ok"):
        flash(
            f"Generated {period} report for {short_name}. "
            f"({'carry' if mode != 'refetch_all' else 'refetch_all'} mode)",
            "success",
        )
    else:
        for err in result.get("errors") or []:
            flash(err, "error")
    return redirect(url_for("scale_runs.cu_dashboard",
                            short_name=short_name))
