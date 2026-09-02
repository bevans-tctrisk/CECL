"""CECL — Simple (SCALE) wizard routes.

This is a sibling to the Migration wizard (``setup.py``). It reuses the
existing Identity step (``setup.step1_identity``) for CU name / charter
/ state / county, then routes through a much shorter set of SCALE-only
steps:

  1. Identity        -- shared with Migration (``setup.step1_identity``)
  2. Solr & Period   -- ``scale_setup.step_solr``
  3. Template & Map  -- ``scale_setup.step_template_map``
  4. Review          -- ``scale_setup.step_review``
  5. Run             -- ``scale_setup.step_run``

State lives in the same ``session['setup_state']`` dict the Migration
wizard uses so the Identity step works unchanged. SCALE-specific
fields hang off ``state['scale']`` and ``state['model'] == 'scale'``
tags the draft.
"""
from __future__ import annotations

import tempfile
from pathlib import Path
from typing import Any

from flask import (
    Blueprint, current_app, flash, redirect, render_template, request,
    session, url_for,
)
from werkzeug.utils import secure_filename

from cecl_ui.routes.setup import (
    STATE_KEY, _state, _save_state, _wizard_ctx,
)
from cecl_ui.services import admin_defaults, wizard_drafts
from cecl_ui.services.scale import (
    impaired_loader, lol_writer, mapping_loader, mgmt_adj_writer,
    qfactor_loader, runner as scale_runner, solr_fetcher,
    template_loader,
)


scale_setup_bp = Blueprint("scale_setup", __name__)

_SCALE_UPLOAD_DIR = Path(tempfile.gettempdir()) / "cecl_ui_scale"


# Mgmt Adjustment pools that should have "Use Default" checked the
# first time a credit union is set up. Names must match the SCALE
# template's ``Scale Calculation!C9:C21`` labels exactly.
_MGMT_ADJ_DEFAULT_ON_POOLS = frozenset({
    "Loans/Lines of Credit Secured by a First Lien on a single 1- to 4-Family Residential Property",
    "Loans/Lines of Credit Secured by a Junior Lien on a single 1- to 4-Family Residential Property",
    "All Other (Non-Commercial) Real Estate Loans/Lines of Credit",
    "Commercial Loans/Lines of Credit Real Estate Secured",
    "Commercial Loans/Lines of Credit Not Real Estate Secured",
})


# -------------------------------------------------------------------
# Step list + endpoints (consumed by setup._wizard_ctx via the
# ``model == 'scale'`` branch).
# -------------------------------------------------------------------
WIZARD_STEPS_SCALE = [
    ("identity",       "1. CU Identity"),
    ("scale_solr",     "2. Solr Connection"),
    ("scale_template", "3. Template & Map"),
    ("scale_qfactors", "4. Q-Factors"),
    ("scale_lol",      "5. Life of Loan Months"),
    ("scale_mgmt_adj", "6. Mgmt Adjustments"),
    ("scale_impaired", "7. Impaired Loans"),
    ("scale_review",   "8. Review"),
    ("scale_run",      "9. Run"),
]

SCALE_STEP_ENDPOINTS: dict[str, str] = {
    "scale_solr":     "scale_setup.step_solr",
    "scale_template": "scale_setup.step_template_map",
    "scale_qfactors": "scale_setup.step_qfactors",
    "scale_lol":      "scale_setup.step_lol",
    "scale_mgmt_adj": "scale_setup.step_mgmt_adj",
    "scale_impaired": "scale_setup.step_impaired",
    "scale_review":   "scale_setup.step_review",
    "scale_run":      "scale_setup.step_run",
}


# -------------------------------------------------------------------
# State helpers
# -------------------------------------------------------------------

def _default_scale_block() -> dict[str, Any]:
    admin = {}
    try:
        admin = admin_defaults.load() or {}
    except Exception:  # noqa: BLE001
        admin = {}
    scale_admin = admin.get("scale") or {}
    try:
        admin_default_mgmt_adj = float(admin.get("default_mgmt_adj") or 0.0011)
    except (TypeError, ValueError):
        admin_default_mgmt_adj = 0.0011
    return {
        "solr_url": scale_admin.get("solr_url") or "http://searchserver1.tctrisk.com:8983/solr",
        "solr_core": scale_admin.get("solr_core") or "ncua",
        "solr_user": "",
        "solr_pass": "",
        "period": "",
        # The N most recent quarter-ends (oldest -> newest) the wizard will
        # auto-generate reports for, resolved from Solr 5300 availability.
        "auto_periods": [],
        "template_override_path": "",
        "template_override_name": "",
        "map_override_path": "",
        "map_override_name": "",
        "qfactor_overrides": {},   # {"Sheet|Cell": bps_float}
        # Management Adjustment inputs written to the 'Management
        # Adjustment' tab. ``pool_rows`` is keyed by pool *name* so the
        # values survive a reorder of the template's pool list; the
        # writer re-aligns to the template's row order at run-time.
        "mgmt_adj": {
            # Decimal default (0.0011 == 0.11%). Seeded from
            # admin_defaults below.
            "default_pct": admin_default_mgmt_adj,
            # {pool_name: {"hard_code_pct": float, "use_default": bool}}
            "pool_rows": {},
            "portfolio": {"hard_code_pct": 0.0, "use_default": False},
        },
        # Per-pool Life-of-Loan month overrides written to the
        # SCALE template's ``Industry Data`` tab. Keyed by pool name
        # (matches ``Scale Calculation``!C9:C21) so the override
        # survives a template-pool reorder. Values are positive ints;
        # absence/0 means "keep template default".
        "life_of_loan_overrides": {},   # {pool_name: months_int}
        "impaired_file": {},       # {saved_path, uploaded_filename, parsed:{...}}
        "report_variant": "both",  # 'tct' | 'vizo' | 'both'
        "last_test": {},   # {ok, status, message, ran_at}
        "last_run": {},
    }


def _scale(state: dict[str, Any]) -> dict[str, Any]:
    if not isinstance(state.get("scale"), dict):
        state["scale"] = _default_scale_block()
    # Ensure new keys exist on older drafts.
    defaults = _default_scale_block()
    for k, v in defaults.items():
        state["scale"].setdefault(k, v)
    return state["scale"]


def _persist_scale_draft(state: dict[str, Any], active_step: str) -> None:
    """Write the wizard draft to disk after a successful SCALE run so
    the runs dashboard can rehydrate it for future quarters. Failures
    here are surfaced as a warning but do not break the success flow
    (the workbook is already on disk).
    """
    if not (state.get("short_name") or state.get("credit_union")):
        return
    try:
        wizard_drafts.save_draft(
            current_app.config["WORKSPACE_ROOT"], state,
            active_step=active_step, model="scale",
        )
    except Exception as exc:  # noqa: BLE001
        flash(
            f"Workbook written, but the wizard draft could not be saved "
            f"to disk ({exc}). Click Save Progress to enable future-quarter "
            f"runs from the SCALE Runs dashboard.",
            "warning",
        )


def _ensure_scale_mode(state: dict[str, Any]) -> None:
    state["model"] = "scale"
    _scale(state)


def _flash_recalc_warnings(result: dict[str, Any]) -> None:
    """Surface any ``_recalc_outputs`` skips/failures as user-visible
    flash warnings. The runner's post-write Excel recalc is
    best-effort and silent on failure; without this, a missing
    ``pywin32`` install (or any other recalc failure) leaves every
    output workbook with stale template-cached formula values and the
    analyst sees identical numbers across periods until they open
    each file in Excel and trigger a manual recalc (F9 / Save).
    """
    recalc = result.get("recalc") or []
    if not isinstance(recalc, list):
        return
    skipped: list[str] = []
    failed: list[str] = []
    skip_reason = ""
    for entry in recalc:
        if not isinstance(entry, dict):
            continue
        path = str(entry.get("path") or "")
        name = Path(path).name if path else "(unknown)"
        if entry.get("skipped"):
            skipped.append(name)
            if not skip_reason:
                skip_reason = str(entry.get("error") or "").strip()
        elif not entry.get("ok"):
            err = str(entry.get("error") or "").strip()
            failed.append(f"{name}: {err}" if err else name)
    if skipped:
        n = len(skipped)
        sample = ", ".join(skipped[:3])
        more = f" (+{n - 3} more)" if n > 3 else ""
        detail = f" — {skip_reason}" if skip_reason else ""
        flash(
            f"Workbook formulas were NOT auto-recalculated for {n} "
            f"file(s){detail}. Open each file in Excel and press F9 "
            f"(or Save) so cached values refresh; otherwise viewers "
            f"that don't recompute (SharePoint preview, Outlook "
            f"preview, Excel set to manual calc) will show stale "
            f"template values that may look identical across "
            f"periods. Files: {sample}{more}.",
            "warning",
        )
    if failed:
        n = len(failed)
        sample = "; ".join(failed[:3])
        more = f" (+{n - 3} more)" if n > 3 else ""
        flash(
            f"Excel recalc failed for {n} file(s): {sample}{more}. "
            "Open the file(s) in Excel manually so cached values "
            "refresh before reviewing the ACL numbers.",
            "warning",
        )


def _flash_prior_acl_warnings(result: dict[str, Any]) -> None:
    """Surface ``_inject_prior_acl`` fallbacks as user-visible flash
    warnings, and announce DB hits / captures as success info.

    Resolution order in the runner:

    * Path 1 (preferred) -- read from ``scale_acl_history`` PostgreSQL
      table. No Excel dependency, no pywin32.
    * Path 2 -- read cached values from the prior workbook (best
      effort; requires Excel to have saved cached values).
    * Path 3 -- write literal 0 to AY2..AY5 to break the formula leak.

    This helper surfaces:

    * The DB-hit case as an info-level confirmation.
    * The Path-2 capture-into-DB success so the user knows the next
      quarter will be automatic.
    * The Path-3 fallback as a warning with actionable remediation.
    * Any ``db_captures`` from the just-finished run so the user can
      see whether the current period's values are now persisted.
    """
    prior = result.get("prior_acl") or {}
    if not isinstance(prior, dict):
        prior = {}
    source = str(prior.get("source") or "")
    fallback_used = bool(prior.get("fallback_used"))
    prev_period = str(prior.get("prior_period") or "")

    # Case A: DB hit -- no Excel dependency.
    if source.startswith("db:") and not fallback_used:
        flash(
            f"Prior ACL values were loaded from the database "
            f"(scale_acl_history) for {prev_period}; no Excel "
            f"dependency required.",
            "info",
        )

    # Case B: read from prior workbook and persisted to DB.
    db_persist = prior.get("db_persist") or {}
    if source == "prior-workbook" and db_persist.get("ok"):
        flash(
            f"Prior ACL values were read from the prior-quarter "
            f"workbook for {prev_period} and persisted to the "
            f"database. Next quarter's run will use the DB copy "
            f"and won't require Excel access.",
            "info",
        )

    # Case C: fallback to zero -- analyst needs to act.
    if fallback_used:
        reason = str(prior.get("error") or "").strip()
        prior_path = str(prior.get("prior_path") or "")
        prior_name = Path(prior_path).name if prior_path else "(unknown)"
        applied = prior.get("applied") or []
        n_written = sum(
            1 for a in applied if isinstance(a, dict) and a.get("ok")
        )
        detail = f" — {reason}" if reason else ""
        flash(
            f"Prior ACL values for {prev_period} could NOT be loaded "
            f"from the database OR read from the prior-quarter "
            f"workbook ({prior_name}); wrote 0 to "
            f"Historical Data!AY2..AY5 on {n_written} output file(s) "
            f"to prevent the 'prior column duplicates current column' "
            f"leak{detail}. To populate the database for this period: "
            f"(1) open the prior workbook in Excel and Save so cached "
            f"values populate, then run "
            f"`scripts/scale_acl_backfill.py` to harvest them; or "
            f"(2) install pywin32 in the wizard Python environment so "
            f"the runner can headlessly recalc prior workbooks. After "
            f"either, re-run this report and the DB will be populated "
            f"automatically.",
            "warning",
        )

    # Always surface DB captures for the CURRENT period.
    db_captures = result.get("db_captures") or []
    if isinstance(db_captures, list) and db_captures:
        ok_count = sum(1 for c in db_captures if isinstance(c, dict) and c.get("ok"))
        skipped_count = sum(
            1 for c in db_captures
            if isinstance(c, dict) and c.get("skipped")
        )
        if ok_count:
            flash(
                f"Persisted ACL values for the current period to the "
                f"database ({ok_count} file(s)); next quarter's run "
                f"will read them automatically.",
                "info",
            )
        elif skipped_count:
            flash(
                f"Could not persist ACL values for the current period "
                f"to the database: {skipped_count} output file(s) have "
                f"no cached formula values yet (Excel hasn't evaluated "
                f"them). Open the workbook in Excel + Save, then run "
                f"`scripts/scale_acl_backfill.py` to populate the DB. "
                f"Or install pywin32 so the runner can do this "
                f"automatically.",
                "warning",
            )


# -------------------------------------------------------------------
# Entry — model_select POSTs model=scale here
# -------------------------------------------------------------------

@scale_setup_bp.route("/start", methods=["GET", "POST"])
def start():
    # Start a brand-new SCALE setup from a clean slate so a prior CU's
    # data (name, charter, Solr, pools, ...) can't bleed into this one.
    session.pop(STATE_KEY, None)
    session.modified = True
    state = _state()
    _ensure_scale_mode(state)
    _save_state(state)
    return redirect(url_for("setup.step1_identity"))


# -------------------------------------------------------------------
# Step 2 — Solr & Period
# -------------------------------------------------------------------

VALID_Q_MONTHS = ("03", "06", "09", "12")


def _period_choices() -> list[str]:
    """Union of available canonical map + template periods."""
    seen: set[str] = set()
    seen.update(template_loader.list_available_map_periods())
    seen.update(template_loader.list_available_template_periods())
    return sorted(seen, reverse=True)


# Number of most-recent quarters the SCALE wizard auto-generates.
AUTO_PERIOD_COUNT = 4


def _resolve_auto_periods(state: dict[str, Any], n: int = AUTO_PERIOD_COUNT) -> dict[str, Any]:
    """Resolve the ``n`` most recent quarter-ends to auto-run for this CU.

    Uses Solr 5300 availability for the CU's charter (intersected with the
    periods we have a mapping CSV for), falling back to the newest
    available mapping period when Solr can't be reached. Returns
    ``{ok, periods (oldest->newest), latest, source, error}``.
    """
    sc = state.get("scale") or {}
    solr_url = (sc.get("solr_url") or "").strip()
    solr_core = (sc.get("solr_core") or "").strip()
    map_periods = set(template_loader.list_available_map_periods())
    latest = ""
    source = ""
    err = ""

    try:
        charter_int = int(str(state.get("charter_number") or "").strip())
    except (TypeError, ValueError):
        charter_int = None

    if charter_int is not None and solr_url and solr_core:
        probe = solr_fetcher.list_charter_periods(
            solr_url, solr_core, charter_int,
            username=sc.get("solr_user") or None,
            password=sc.get("solr_pass") or None,
        )
        if probe.get("ok"):
            usable = sorted(
                (set(probe.get("periods") or set()) & map_periods),
                reverse=True,
            )
            if usable:
                latest, source = usable[0], "solr"
            else:
                err = ("Solr has no 5300 filings for this charter that match "
                       "an available SCALE mapping period.")
        else:
            err = probe.get("error") or "Solr availability check failed."

    if not latest:
        avail = sorted(map_periods, reverse=True)
        if avail:
            latest = avail[0]
            source = source or "map_fallback"
        else:
            return {"ok": False, "periods": [], "latest": "",
                    "source": "", "error": err or "No mapping periods available."}

    # N consecutive quarters ending at ``latest``, keeping only those with a
    # mapping CSV (maps must match the period exactly; templates fall back).
    seq = [p for p in scale_runner.make_periods(latest, n) if p in map_periods]
    periods = list(reversed(seq))  # oldest -> newest
    return {"ok": bool(periods), "periods": periods, "latest": latest,
            "source": source, "error": err}


def _apply_auto_periods(state: dict[str, Any]) -> dict[str, Any]:
    """Resolve and store the auto-run periods on the draft. Sets
    ``scale.period`` to the latest (so the template/map/LoL/mgmt steps
    resolve against the current quarter) and ``scale.auto_periods`` to the
    full oldest->newest list."""
    sc = _scale(state)
    res = _resolve_auto_periods(state)
    if res["ok"]:
        sc["auto_periods"] = res["periods"]
        sc["period"] = res["latest"]
    return res


def _run_auto_periods(state: dict[str, Any], workspace_root: str) -> list[dict[str, Any]]:
    """Generate a report for each auto-resolved quarter (oldest -> newest).

    The oldest quarter is a full multi-quarter build (establishes the
    Historical Data history); each newer quarter carries forward from the
    prior report. The uploaded Impaired Loans file is applied only to the
    most recent quarter (it's point-in-time), so the older reports trend
    cleanly. Returns ``[{period, result}, ...]``.
    """
    sc = _scale(state)
    periods = [p for p in (sc.get("auto_periods") or []) if p]
    if not periods and sc.get("period"):
        periods = [sc["period"]]
    saved_impaired = sc.get("impaired_file")
    results: list[dict[str, Any]] = []
    for i, p in enumerate(periods):
        sc["period"] = p
        is_latest = (i == len(periods) - 1)
        sc["impaired_file"] = saved_impaired if is_latest else {}
        if i == 0:
            r = scale_runner.run_multi_quarter(state, workspace_root, quarters=32)
        else:
            r = scale_runner.run_quarter_carry_history(state, workspace_root)
        results.append({"period": p, "result": r})
    # Restore the uploaded impaired file and pin the period to the latest
    # quarter so the dashboard / later steps default to the current report.
    sc["impaired_file"] = saved_impaired
    if periods:
        sc["period"] = periods[-1]
    return results


@scale_setup_bp.route("/step/solr", methods=["GET", "POST"])
def step_solr():
    state = _state()
    _ensure_scale_mode(state)
    sc = _scale(state)
    if request.method == "POST":
        action = request.form.get("action", "save")
        sc["solr_url"] = request.form.get("solr_url", "").strip()
        sc["solr_core"] = request.form.get("solr_core", "").strip()
        sc["solr_user"] = request.form.get("solr_user", "").strip()
        # Password not echoed in form when blank — only overwrite when
        # the user types a fresh value.
        new_pass = request.form.get("solr_pass", "")
        if new_pass:
            sc["solr_pass"] = new_pass

        if action == "test":
            res = solr_fetcher.test_connection(sc["solr_url"], sc["solr_core"])
            sc["last_test"] = res
            _save_state(state)
            flash(
                ("Solr ping OK." if res["ok"]
                 else f"Solr ping failed: {res.get('message') or res.get('status')}"),
                "success" if res["ok"] else "error",
            )
            return redirect(url_for("scale_setup.step_solr"))

        if not sc["solr_url"] or not sc["solr_core"]:
            flash("Solr URL and core are required.", "error")
            _save_state(state)
            return redirect(url_for("scale_setup.step_solr"))

        # Auto-resolve the quarters to run (no manual period selection).
        auto = _apply_auto_periods(state)
        _save_state(state)
        if not auto["ok"]:
            flash(
                f"Could not determine the quarters to run automatically: "
                f"{auto.get('error') or 'no available periods'}. Check the "
                f"charter number and Solr connection.",
                "error",
            )
            return redirect(url_for("scale_setup.step_solr"))
        if auto.get("source") == "map_fallback" and auto.get("error"):
            flash(
                f"Solr availability check note: {auto['error']} Using the "
                f"latest available template period instead.",
                "warning",
            )
        if action == "next":
            return redirect(url_for("scale_setup.step_template_map"))
        flash(
            "Detected the latest 5300 quarter and the reports to generate.",
            "success",
        )
        return redirect(url_for("scale_setup.step_solr"))

    # GET: auto-resolve so the page can show what will run.
    auto = _apply_auto_periods(state)
    _save_state(state)
    return render_template(
        "setup/scale/step_solr.html",
        auto=auto,
        **_wizard_ctx("scale_solr"),
    )


# -------------------------------------------------------------------
# Step 3 — Template & Map (canonical preview + optional overrides)
# -------------------------------------------------------------------

def _save_scale_upload(file_storage, kind: str) -> Path:
    """Persist a per-CU override upload. ``kind`` is ``template`` or ``map``."""
    sub = _SCALE_UPLOAD_DIR / kind
    sub.mkdir(parents=True, exist_ok=True)
    fn = secure_filename(file_storage.filename or f"{kind}.bin")
    target = sub / fn
    file_storage.save(target)
    return target


@scale_setup_bp.route("/step/template-map", methods=["GET", "POST"])
def step_template_map():
    state = _state()
    _ensure_scale_mode(state)
    sc = _scale(state)
    if not sc.get("period"):
        flash("Pick a target period on the Solr step first.", "error")
        return redirect(url_for("scale_setup.step_solr"))

    if request.method == "POST":
        action = request.form.get("action", "save")

        if action == "upload_template":
            f = request.files.get("template_file")
            if f and f.filename:
                saved = _save_scale_upload(f, "template")
                sc["template_override_path"] = str(saved)
                sc["template_override_name"] = f.filename
                _save_state(state)
                flash(f"Template override saved: {f.filename}", "success")
            else:
                flash("Pick a .xlsx file first.", "error")
            return redirect(url_for("scale_setup.step_template_map"))

        if action == "clear_template":
            sc["template_override_path"] = ""
            sc["template_override_name"] = ""
            _save_state(state)
            flash("Template override cleared.", "info")
            return redirect(url_for("scale_setup.step_template_map"))

        if action == "upload_map":
            f = request.files.get("map_file")
            if f and f.filename:
                saved = _save_scale_upload(f, "map")
                # Validate it parses before accepting.
                try:
                    rows = mapping_loader.load_rows(saved)
                except Exception as exc:  # noqa: BLE001
                    flash(f"Mapping CSV invalid: {exc}", "error")
                    return redirect(url_for("scale_setup.step_template_map"))
                sc["map_override_path"] = str(saved)
                sc["map_override_name"] = f.filename
                _save_state(state)
                flash(
                    f"Mapping override saved: {f.filename} ({len(rows)} rows)",
                    "success",
                )
            else:
                flash("Pick a .csv file first.", "error")
            return redirect(url_for("scale_setup.step_template_map"))

        if action == "clear_map":
            sc["map_override_path"] = ""
            sc["map_override_name"] = ""
            _save_state(state)
            flash("Mapping override cleared.", "info")
            return redirect(url_for("scale_setup.step_template_map"))

        if action == "next":
            return redirect(url_for("scale_setup.step_qfactors"))

    tmpl = template_loader.resolve_template(
        sc["period"], sc.get("template_override_path") or None
    )
    mp = template_loader.resolve_map(
        sc["period"], sc.get("map_override_path") or None
    )
    map_summary = None
    if mp.get("ok"):
        try:
            map_summary = mapping_loader.summarize(
                mapping_loader.load_rows(mp["path"])
            )
        except Exception as exc:  # noqa: BLE001
            mp["message"] = f"{mp.get('message') or ''} {exc}".strip()
    return render_template(
        "setup/scale/step_template_map.html",
        tmpl=tmpl,
        mp=mp,
        map_summary=map_summary,
        **_wizard_ctx("scale_template"),
    )


# -------------------------------------------------------------------
# Step 4 — Q-Factors (qualitative overlays applied AFTER 5300 fill)
# -------------------------------------------------------------------

def _coerce_bps_form(value: str) -> float:
    if value is None:
        return 0.0
    try:
        return float(str(value).strip() or 0)
    except (TypeError, ValueError):
        return 0.0


@scale_setup_bp.route("/step/qfactors", methods=["GET", "POST"])
def step_qfactors():
    state = _state()
    _ensure_scale_mode(state)
    sc = _scale(state)
    defaults = qfactor_loader.load_defaults()
    if request.method == "POST":
        action = request.form.get("action", "save")
        # Collect overrides from form (fields named qf_<i>_bps with
        # hidden qf_<i>_key) — only keep values that differ from the
        # default.
        overrides: dict[str, float] = {}
        for i, row in enumerate(defaults):
            key = request.form.get(f"qf_{i}_key", "")
            if key != row["key"]:
                continue
            bps = _coerce_bps_form(request.form.get(f"qf_{i}_bps", ""))
            if abs(bps - row["default_bps"]) > 1e-9:
                overrides[row["key"]] = bps
        if action == "reset":
            sc["qfactor_overrides"] = {}
            _save_state(state)
            flash("Q-factor overrides cleared.", "info")
            return redirect(url_for("scale_setup.step_qfactors"))
        sc["qfactor_overrides"] = overrides
        _save_state(state)
        if action == "next":
            return redirect(url_for("scale_setup.step_lol"))
        flash("Q-factor overrides saved.", "success")
        return redirect(url_for("scale_setup.step_qfactors"))
    entries = qfactor_loader.merge_with_overrides(
        defaults, sc.get("qfactor_overrides") or {}
    )
    return render_template(
        "setup/scale/step_qfactors.html",
        entries=entries,
        defaults_path=str(qfactor_loader.defaults_path()),
        **_wizard_ctx("scale_qfactors"),
    )


# -------------------------------------------------------------------
# Step 5 — Life of Loan Months (per-pool ACL Lifetime Months / WAM)
# -------------------------------------------------------------------

def _coerce_months_form(value: str) -> int:
    """Form value -> non-negative int; blank or invalid returns 0."""
    if value is None:
        return 0
    s = str(value).strip()
    if not s:
        return 0
    try:
        n = int(round(float(s)))
    except (TypeError, ValueError):
        return 0
    return max(n, 0)


@scale_setup_bp.route("/step/lol", methods=["GET", "POST"])
def step_lol():
    state = _state()
    _ensure_scale_mode(state)
    sc = _scale(state)
    overrides = sc.setdefault("life_of_loan_overrides", {})

    tmpl_path = _resolved_template_path(sc)
    template_rows: list[dict[str, Any]] = []
    if tmpl_path:
        try:
            template_rows = lol_writer.read_lol_months(tmpl_path)
        except Exception as exc:  # noqa: BLE001
            current_app.logger.warning(
                "lol: could not read months from %s: %s", tmpl_path, exc,
            )

    if request.method == "POST":
        action = request.form.get("action", "save")

        if action == "reset":
            sc["life_of_loan_overrides"] = {}
            _save_state(state)
            flash(
                "Life of Loan overrides cleared. Template defaults will "
                "apply on the next run.", "info",
            )
            return redirect(url_for("scale_setup.step_lol"))

        new_overrides: dict[str, int] = {}
        for i, tr in enumerate(template_rows):
            name = tr["name"]
            raw = request.form.get(f"lol_months__{i}", "")
            months = _coerce_months_form(raw)
            template_default = tr.get("months") or 0
            # Only persist values that differ from the template default.
            # 0 / blank means "use template default" -> drop the entry.
            if months > 0 and months != template_default:
                new_overrides[name] = months
        sc["life_of_loan_overrides"] = new_overrides
        _save_state(state)

        if action == "next":
            return redirect(url_for("scale_setup.step_mgmt_adj"))
        flash("Life of Loan overrides saved.", "success")
        return redirect(url_for("scale_setup.step_lol"))

    # Build display rows merging template defaults with any saved overrides.
    saved = overrides if isinstance(overrides, dict) else {}
    pool_display: list[dict[str, Any]] = []
    for i, tr in enumerate(template_rows):
        name = tr["name"]
        default_months = tr.get("months")
        override = saved.get(name)
        try:
            override_int = int(override) if override is not None else 0
        except (TypeError, ValueError):
            override_int = 0
        pool_display.append({
            "idx": i,
            "name": name,
            "default_months": default_months,
            "override_months": override_int,
            "effective_months": (
                override_int if override_int > 0 else (default_months or 0)
            ),
            "is_overridden": override_int > 0 and override_int != (default_months or 0),
        })

    return render_template(
        "setup/scale/step_lol.html",
        pool_rows=pool_display,
        template_path=tmpl_path,
        template_missing=(not tmpl_path),
        override_count=sum(1 for p in pool_display if p["is_overridden"]),
        **_wizard_ctx("scale_lol"),
    )


# -------------------------------------------------------------------
# Step 6 — Management Adjustments
# -------------------------------------------------------------------

def _resolved_template_path(sc: dict) -> str:
    """Return the .xlsx path the SCALE run will use, or ''."""
    period = (sc.get("period") or "").strip()
    if not period:
        return ""
    res = template_loader.resolve_template(
        period, sc.get("template_override_path") or None
    )
    return res.get("path") or "" if res.get("ok") else ""


def _coerce_pct_form(value: str) -> float:
    """Form inputs are entered as a percentage (e.g. 0.11 = 0.11%)."""
    if value is None:
        return 0.0
    try:
        return float(str(value).strip() or 0) / 100.0
    except (TypeError, ValueError):
        return 0.0


def _coerce_bool_form(value: Any) -> bool:
    return str(value).strip().lower() in ("1", "true", "on", "yes")


@scale_setup_bp.route("/step/mgmt-adj", methods=["GET", "POST"])
def step_mgmt_adj():
    state = _state()
    _ensure_scale_mode(state)
    sc = _scale(state)
    ma = sc.setdefault("mgmt_adj", {
        "default_pct": 0.0011, "pool_rows": {}, "portfolio": {},
    })

    # Re-read the admin default each GET so an updated admin value
    # surfaces. The user can still override it on the form.
    try:
        admin = admin_defaults.load() or {}
        admin_default_pct = float(admin.get("default_mgmt_adj") or 0.0011)
    except Exception:  # noqa: BLE001
        admin_default_pct = 0.0011

    tmpl_path = _resolved_template_path(sc)
    pool_names: list[str] = []
    if tmpl_path:
        try:
            pool_names = mgmt_adj_writer.list_pool_names(tmpl_path)
        except Exception as exc:  # noqa: BLE001
            current_app.logger.warning(
                "mgmt_adj: could not read pool names from %s: %s",
                tmpl_path, exc,
            )

    if request.method == "POST":
        action = request.form.get("action", "save")

        if action == "reset":
            ma["default_pct"] = admin_default_pct
            ma["pool_rows"] = {}
            ma["portfolio"] = {"hard_code_pct": 0.0, "use_default": False}
            _save_state(state)
            flash("Management Adjustment overrides reset to defaults.", "info")
            return redirect(url_for("scale_setup.step_mgmt_adj"))

        # Default rate (entered as percent).
        ma["default_pct"] = _coerce_pct_form(
            request.form.get("default_pct", "")
        )

        # Per-pool rows.
        new_rows: dict[str, dict] = {}
        for i, name in enumerate(pool_names):
            new_rows[name] = {
                "hard_code_pct": _coerce_pct_form(
                    request.form.get(f"pool_pct__{i}", "")
                ),
                "use_default": _coerce_bool_form(
                    request.form.get(f"pool_use_default__{i}", "")
                ),
            }
        ma["pool_rows"] = new_rows

        # Whole-portfolio overlay row.
        ma["portfolio"] = {
            "hard_code_pct": _coerce_pct_form(
                request.form.get("portfolio_pct", "")
            ),
            "use_default": _coerce_bool_form(
                request.form.get("portfolio_use_default", "")
            ),
        }
        _save_state(state)

        if action == "next":
            return redirect(url_for("scale_setup.step_impaired"))
        flash("Management Adjustment overrides saved.", "success")
        return redirect(url_for("scale_setup.step_mgmt_adj"))

    # Build display rows merging template pool list with saved values.
    # For pools that have no saved entry yet, fall back to the
    # app-wide "default on" set so new credit unions start with the
    # expected pools pre-checked.
    saved_rows = ma.get("pool_rows") or {}
    pool_display: list[dict] = []
    for i, name in enumerate(pool_names):
        if name in saved_rows:
            sv = saved_rows[name] or {}
            use_default = bool(sv.get("use_default"))
        else:
            sv = {}
            use_default = name in _MGMT_ADJ_DEFAULT_ON_POOLS
        pool_display.append({
            "idx": i,
            "name": name,
            "hard_code_pct": float(sv.get("hard_code_pct") or 0.0),
            "use_default": use_default,
        })

    port = ma.get("portfolio") or {}
    portfolio_display = {
        "hard_code_pct": float(port.get("hard_code_pct") or 0.0),
        "use_default": bool(port.get("use_default")),
    }

    return render_template(
        "setup/scale/step_mgmt_adj.html",
        pool_rows=pool_display,
        portfolio=portfolio_display,
        default_pct=float(ma.get("default_pct") or 0.0),
        admin_default_pct=admin_default_pct,
        template_path=tmpl_path,
        template_missing=(not tmpl_path),
        **_wizard_ctx("scale_mgmt_adj"),
    )


# -------------------------------------------------------------------
# Step 6 — Impaired Loans (ASC 310-10)
# -------------------------------------------------------------------

_IMPAIRED_UPLOAD_SUBDIR = "impaired"


def _save_impaired_upload(file_storage) -> Path:
    sub = _SCALE_UPLOAD_DIR / _IMPAIRED_UPLOAD_SUBDIR
    sub.mkdir(parents=True, exist_ok=True)
    fn = secure_filename(file_storage.filename or "impaired.xlsx")
    target = sub / fn
    file_storage.save(target)
    return target


@scale_setup_bp.route("/step/impaired", methods=["GET", "POST"])
def step_impaired():
    state = _state()
    _ensure_scale_mode(state)
    sc = _scale(state)
    imp = sc.setdefault("impaired_file", {})
    if request.method == "POST":
        action = request.form.get("action", "save")
        if action == "upload":
            f = request.files.get("impaired_file")
            if not f or not f.filename:
                flash("Pick an Impaired Loans .xlsx file to upload.", "error")
                return redirect(url_for("scale_setup.step_impaired"))
            try:
                target = _save_impaired_upload(f)
            except Exception as exc:  # noqa: BLE001
                flash(f"Save failed: {exc}", "error")
                return redirect(url_for("scale_setup.step_impaired"))
            parsed = impaired_loader.parse_file(target)
            sc["impaired_file"] = {
                "saved_path": str(target),
                "uploaded_filename": f.filename,
                "parsed": parsed,
            }
            _save_state(state)
            if parsed.get("ok"):
                flash(
                    f"Parsed {parsed['row_count']} impaired loan(s) totalling "
                    f"${parsed['total_balance']:,.2f}.",
                    "success",
                )
            else:
                flash(f"Parse failed: {parsed.get('error')}", "error")
            return redirect(url_for("scale_setup.step_impaired"))
        if action == "clear":
            sc["impaired_file"] = {}
            _save_state(state)
            flash("Impaired loans file cleared.", "info")
            return redirect(url_for("scale_setup.step_impaired"))
        if action == "next":
            return redirect(url_for("scale_setup.step_review"))
        return redirect(url_for("scale_setup.step_impaired"))
    parsed = (imp or {}).get("parsed") or {}
    preview = (parsed.get("rows") or [])[:15]
    return render_template(
        "setup/scale/step_impaired.html",
        imp=imp,
        parsed=parsed,
        preview=preview,
        **_wizard_ctx("scale_impaired"),
    )


# -------------------------------------------------------------------
# Step 6 — Review
# -------------------------------------------------------------------

@scale_setup_bp.route("/step/review", methods=["GET", "POST"])
def step_review():
    state = _state()
    _ensure_scale_mode(state)
    sc = _scale(state)
    if request.method == "POST":
        variant = (request.form.get("report_variant") or "both").strip().lower()
        if variant not in ("tct", "vizo", "both"):
            variant = "both"
        sc["report_variant"] = variant
        _save_state(state)
        return redirect(url_for("scale_setup.step_run"))
    tmpl = template_loader.resolve_template(
        sc.get("period", ""), sc.get("template_override_path") or None
    )
    mp = template_loader.resolve_map(
        sc.get("period", ""), sc.get("map_override_path") or None
    )
    map_summary = None
    if mp.get("ok"):
        try:
            map_summary = mapping_loader.summarize(
                mapping_loader.load_rows(mp["path"])
            )
        except Exception:  # noqa: BLE001
            pass
    qf_entries = qfactor_loader.merge_with_overrides(
        qfactor_loader.load_defaults(),
        sc.get("qfactor_overrides") or {},
    )
    qf_active = sum(
        1 for r in qf_entries
        if abs(r["effective_bps"] - r["default_bps"]) > 1e-9
    )
    lol_overrides = sc.get("life_of_loan_overrides") or {}
    lol_count = sum(
        1 for v in lol_overrides.values()
        if isinstance(v, (int, float)) and int(v) > 0
    )
    imp = sc.get("impaired_file") or {}
    imp_parsed = imp.get("parsed") or {}
    return render_template(
        "setup/scale/step_review.html",
        tmpl=tmpl,
        mp=mp,
        map_summary=map_summary,
        qf_total=len(qf_entries),
        qf_active=qf_active,
        lol_count=lol_count,
        imp=imp,
        imp_parsed=imp_parsed,
        auto_periods=sc.get("auto_periods") or [],
        **_wizard_ctx("scale_review"),
    )


# -------------------------------------------------------------------
# Step 6 — Run
# -------------------------------------------------------------------

@scale_setup_bp.route("/step/run", methods=["GET", "POST"])
def step_run():
    state = _state()
    _ensure_scale_mode(state)
    sc = _scale(state)
    sn = (state.get("short_name") or "").strip()

    def _dashboard_or_self() -> str:
        if sn:
            return url_for("scale_runs.cu_dashboard", short_name=sn)
        return url_for("scale_setup.step_run")

    def _mark_setup_completed() -> None:
        if not (sn or state.get("credit_union")):
            return
        try:
            wizard_drafts.mark_completed(
                current_app.config["WORKSPACE_ROOT"],
                wizard_drafts.draft_key_for_state(state),
                model="scale",
            )
        except Exception as exc:  # noqa: BLE001
            flash(f"Could not mark setup completed: {exc}", "warning")

    def _apply_economic_form(econ: dict[str, Any]) -> dict[str, Any]:
        for _k in ("state", "county"):
            _v = (request.form.get(f"econ_{_k}") or "").strip()
            if _v:
                econ[_k] = _v
        _ur = (request.form.get("econ_unemployment_rate") or "").strip().rstrip("%")
        if _ur:
            try:
                _urf = float(_ur)
                if _urf > 1.0:
                    _urf = _urf / 100.0
                if 0.0 <= _urf <= 1.0:
                    econ["unemployment_rate"] = round(_urf, 6)
            except ValueError:
                pass
        for _k in ("population", "bankruptcies", "foreclosures"):
            _raw = (request.form.get(f"econ_{_k}") or "").strip().replace(",", "")
            if _raw:
                try:
                    _n = int(float(_raw))
                    if _n >= 0:
                        econ[_k] = _n
                except ValueError:
                    pass
        return econ

    if request.method == "POST":
        action = request.form.get("action", "")
        if action in ("save_economic", "fetch_economic"):
            econ = dict(state.get("economic_data") or {})
            econ = _apply_economic_form(econ)
            if action == "fetch_economic":
                st = (econ.get("state") or "").strip()
                cty = (econ.get("county") or "").strip()
                if not st:
                    flash("Set the State first, then fetch.", "error")
                    return redirect(url_for("scale_setup.step_run"))
                try:
                    import fetch_econ_data

                    fetched = fetch_econ_data.fetch_economic_data(st, cty) or {}
                except Exception as exc:  # noqa: BLE001
                    flash(f"Fetch failed: {exc}", "error")
                    return redirect(url_for("scale_setup.step_run"))

                got: list[str] = []
                for key, label in (
                    ("unemployment_rate", "unemployment"),
                    ("population", "population"),
                    ("bankruptcies", "bankruptcies"),
                ):
                    val = fetched.get(key)
                    if val not in (None, ""):
                        econ[key] = val
                        got.append(label)
                state["economic_data"] = econ
                _save_state(state)
                _persist_scale_draft(state, "scale_run")
                if got:
                    flash(
                        "Fetched " + ", ".join(got)
                        + " from federal sources. Foreclosures has no federal source "
                        "- enter it manually if needed.",
                        "success",
                    )
                else:
                    flash(
                        "No values could be fetched for "
                        f"{st}{(' / ' + cty) if cty else ''}. Check the state name, "
                        "or enter factors manually.",
                        "warning",
                    )
            else:
                state["economic_data"] = econ
                _save_state(state)
                _persist_scale_draft(state, "scale_run")
                flash("Environmental factors saved.", "success")
            return redirect(url_for("scale_setup.step_run"))

        if action == "run":
            results = _run_auto_periods(
                state, current_app.config["WORKSPACE_ROOT"],
            )
            # Keep the latest quarter's detailed result for the Run page.
            latest_res = results[-1]["result"] if results else {}
            sc["last_run"] = latest_res
            sc["last_auto_run"] = [
                {"period": r["period"], "ok": bool(r["result"].get("ok")),
                 "errors": r["result"].get("errors") or []}
                for r in results
            ]
            _save_state(state)
            ok_periods = [r["period"] for r in results if r["result"].get("ok")]
            fail = [r for r in results if not r["result"].get("ok")]
            if ok_periods:
                _persist_scale_draft(state, "scale_run")
                _mark_setup_completed()
                flash(
                    f"Generated {len(ok_periods)} SCALE report(s): "
                    f"{', '.join(ok_periods)}.",
                    "success",
                )
                _flash_recalc_warnings(latest_res)
                _flash_prior_acl_warnings(latest_res)
            for r in fail:
                errs = "; ".join(r["result"].get("errors") or ["unknown error"])
                flash(f"{r['period']}: {errs}", "error")
            return redirect(_dashboard_or_self())
        if action == "run_single":
            result = scale_runner.run_single_quarter(
                state, current_app.config["WORKSPACE_ROOT"]
            )
            sc["last_run"] = result
            _save_state(state)
            if result.get("ok"):
                _persist_scale_draft(state, "scale_run")
                _mark_setup_completed()
                flash(
                    f"SCALE workbook written: applied {result['applied']} "
                    f"of {result['total_rows']} cells.",
                    "success",
                )
                _flash_recalc_warnings(result)
                _flash_prior_acl_warnings(result)
            else:
                for err in result.get("errors", []):
                    flash(err, "error")
            return redirect(_dashboard_or_self())
    return render_template(
        "setup/scale/step_run.html",
        last_run=sc.get("last_run") or {},
        last_auto_run=sc.get("last_auto_run") or [],
        auto_periods=sc.get("auto_periods") or [],
        economic=state.get("economic_data") or {},
        short_name=sn,
        **_wizard_ctx("scale_run"),
    )


@scale_setup_bp.route("/step/run/complete", methods=["POST"])
def complete_and_home():
    """Mark the SCALE wizard as completed for this CU and jump to
    that CU's SCALE Runs dashboard.

    Triggered by the "Continue to SCALE Runs" button on the Run step.
    Stamps the draft's ``_draft_meta.completed_at`` so the home
    dashboard moves it from "Resume in-progress setup" to "Completed
    setup". The draft itself is kept so the user can still Edit setup
    or Delete it later. Redirect target is the per-CU SCALE Runs page
    (``scale_runs.cu_dashboard``) so the user lands on the page that
    matches their next step of the workflow -- not the generic home
    page, which was the pre-fix behaviour and led to confusion when
    Migration and SCALE dashboards look similar.
    """
    state = _state()
    _ensure_scale_mode(state)
    sn = state.get("short_name") or ""
    # Make sure the on-disk draft exists (Run-step saves are usually
    # already persisted, but call save_draft as a safety net before
    # stamping completion).
    if sn or state.get("credit_union"):
        try:
            wizard_drafts.save_draft(
                current_app.config["WORKSPACE_ROOT"], state,
                active_step="scale_run", model="scale",
            )
            wizard_drafts.mark_completed(
                current_app.config["WORKSPACE_ROOT"],
                wizard_drafts.draft_key_for_state(state),
                model="scale",
            )
            flash(
                f"Marked {state.get('credit_union') or sn} "
                f"setup as completed.",
                "success",
            )
        except Exception as exc:  # noqa: BLE001
            flash(f"Could not mark setup completed: {exc}", "warning")
    if sn:
        return redirect(url_for("scale_runs.cu_dashboard", short_name=sn))
    # Missing short_name (edge case: user hit Complete before Identity
    # step saved). Fall back to the SCALE Runs index.
    return redirect(url_for("scale_runs.index"))
