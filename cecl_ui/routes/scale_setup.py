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
    impaired_loader, mapping_loader, mgmt_adj_writer,
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
    ("scale_solr",     "2. Solr & Period"),
    ("scale_template", "3. Template & Map"),
    ("scale_qfactors", "4. Q-Factors"),
    ("scale_impaired", "5. Impaired Loans"),
    ("scale_review",   "6. Review"),
    ("scale_run",      "7. Run"),
]

SCALE_STEP_ENDPOINTS: dict[str, str] = {
    "scale_solr":     "scale_setup.step_solr",
    "scale_template": "scale_setup.step_template_map",
    "scale_qfactors": "scale_setup.step_qfactors",
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


# -------------------------------------------------------------------
# Entry — model_select POSTs model=scale here
# -------------------------------------------------------------------

@scale_setup_bp.route("/start", methods=["GET", "POST"])
def start():
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
        period = request.form.get("period", "").strip()
        if period:
            sc["period"] = period
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
        # Validate period before saving
        if not sc["period"]:
            flash("Pick a target period (YYYY-MM).", "error")
        else:
            try:
                _y, m = sc["period"].split("-")
                if m not in VALID_Q_MONTHS:
                    raise ValueError
            except Exception:  # noqa: BLE001
                flash("Period must be YYYY-MM with month 03/06/09/12.", "error")
                _save_state(state)
                return redirect(url_for("scale_setup.step_solr"))
        if not sc["solr_url"] or not sc["solr_core"]:
            flash("Solr URL and core are required.", "error")
            _save_state(state)
            return redirect(url_for("scale_setup.step_solr"))
        _save_state(state)
        if action == "next":
            return redirect(url_for("scale_setup.step_template_map"))
        return redirect(url_for("scale_setup.step_solr"))
    return render_template(
        "setup/scale/step_solr.html",
        periods=_period_choices(),
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
            return redirect(url_for("scale_setup.step_mgmt_adj"))
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
# Step 5 — Management Adjustments
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
    imp = sc.get("impaired_file") or {}
    imp_parsed = imp.get("parsed") or {}
    return render_template(
        "setup/scale/step_review.html",
        tmpl=tmpl,
        mp=mp,
        map_summary=map_summary,
        qf_total=len(qf_entries),
        qf_active=qf_active,
        imp=imp,
        imp_parsed=imp_parsed,
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
    if request.method == "POST":
        action = request.form.get("action", "")
        if action == "run":
            result = scale_runner.run_multi_quarter(
                state, current_app.config["WORKSPACE_ROOT"], quarters=32,
            )
            sc["last_run"] = result
            _save_state(state)
            if result.get("ok"):
                _persist_scale_draft(state, "scale_run")
                flash(
                    f"SCALE workbook written: "
                    f"{result.get('quarters_written', 0)} of "
                    f"{result.get('quarters_requested', 32)} quarter(s) filled.",
                    "success",
                )
            else:
                for err in result.get("errors", []):
                    flash(err, "error")
            return redirect(url_for("scale_setup.step_run"))
        if action == "run_single":
            result = scale_runner.run_single_quarter(
                state, current_app.config["WORKSPACE_ROOT"]
            )
            sc["last_run"] = result
            _save_state(state)
            if result.get("ok"):
                _persist_scale_draft(state, "scale_run")
                flash(
                    f"SCALE workbook written: applied {result['applied']} "
                    f"of {result['total_rows']} cells.",
                    "success",
                )
            else:
                for err in result.get("errors", []):
                    flash(err, "error")
            return redirect(url_for("scale_setup.step_run"))
    return render_template(
        "setup/scale/step_run.html",
        last_run=sc.get("last_run") or {},
        **_wizard_ctx("scale_run"),
    )
