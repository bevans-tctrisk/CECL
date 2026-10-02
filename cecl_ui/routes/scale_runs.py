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
from cecl_ui.routes.scale_setup import (
    _MGMT_ADJ_DEFAULT_ON_POOLS,
    _coerce_bool_form,
    _coerce_months_form,
    _coerce_pct_form,
    _default_scale_block,
    _resolved_template_path,
    _save_impaired_upload,
)
from cecl_ui.services import wizard_drafts
from cecl_ui.services.scale import (
    impaired_loader,
    lol_writer,
    mgmt_adj_writer,
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
    state = runs_service.load_state_for_cu(workspace_root, short_name)
    if cu is None:
        # No SCALE workbook on disk yet, but if a wizard draft exists we
        # still show the CU dashboard so the user can run the first
        # quarter. This is the common "just finished the SCALE wizard"
        # landing target from ``scale_setup.complete_and_home``.
        if state is None:
            flash(
                f"No SCALE reports or wizard draft found for {short_name}.",
                "warning",
            )
            return redirect(url_for("scale_runs.index"))
        cu = {
            "short_name": short_name,
            "credit_union": state.get("credit_union") or short_name,
            "latest_period": "",
            "latest_path": "",
            "periods": [],
            "draft_present": True,
        }
        flash(
            f"No SCALE workbook on disk yet for {cu['credit_union']}. "
            "Run the first quarter below.",
            "info",
        )
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
        economic=(state.get("economic_data") if state else {}) or {},
        cu_name_override=((state.get("scale") or {}).get("cu_name_override", "") if state else ""),
    )


def _apply_economic_form(econ: dict, form) -> dict:
    """Merge submitted environmental-factor fields into ``econ`` in place
    and return it. Blank fields are left unchanged."""
    for _k in ("state", "county"):
        _v = (form.get(f"econ_{_k}") or "").strip()
        if _v:
            econ[_k] = _v
    _ur = (form.get("econ_unemployment_rate") or "").strip().rstrip("%")
    if _ur:
        try:
            _urf = float(_ur)
            if _urf > 1.0:  # entered as a percent (e.g. 3.8)
                _urf = _urf / 100.0
            if 0.0 <= _urf <= 1.0:
                econ["unemployment_rate"] = round(_urf, 6)
        except ValueError:
            pass
    for _k in ("population", "bankruptcies", "foreclosures"):
        _raw = (form.get(f"econ_{_k}") or "").strip().replace(",", "")
        if _raw:
            try:
                _n = int(float(_raw))
                if _n >= 0:
                    econ[_k] = _n
            except ValueError:
                pass
    return econ


@scale_runs_bp.route("/<short_name>/economic", methods=["POST"])
def save_economic(short_name: str):
    """Save the CU's environmental factors (``economic_data``) from the
    dashboard's Environmental Factors card into the SCALE wizard draft."""
    workspace_root = _workspace_root()
    state = runs_service.load_state_for_cu(workspace_root, short_name)
    if state is None:
        flash(f"No saved wizard draft for {short_name}.", "error")
        return redirect(url_for("scale_runs.cu_dashboard", short_name=short_name))
    econ = dict(state.get("economic_data") or {})
    _apply_economic_form(econ, request.form)
    state["economic_data"] = econ
    wizard_drafts.save_draft(
        workspace_root, state, active_step="scale_run", model="scale",
    )
    flash("Environmental factors saved.", "success")
    return redirect(url_for("scale_runs.cu_dashboard", short_name=short_name))


@scale_runs_bp.route("/<short_name>/economic/fetch", methods=["POST"])
def fetch_economic(short_name: str):
    """Fetch environmental factors from federal sources (BLS unemployment,
    Census population, U.S. Courts bankruptcies) for the CU's state and
    save them into the draft. Foreclosures has no federal source and is
    left unchanged."""
    workspace_root = _workspace_root()
    state = runs_service.load_state_for_cu(workspace_root, short_name)
    if state is None:
        flash(f"No saved wizard draft for {short_name}.", "error")
        return redirect(url_for("scale_runs.cu_dashboard", short_name=short_name))

    econ = dict(state.get("economic_data") or {})
    # Use any just-typed state/county from the form so the user can fetch
    # without a separate save first.
    st = (request.form.get("econ_state") or econ.get("state") or "").strip()
    cty = (request.form.get("econ_county") or econ.get("county") or "").strip()
    if not st:
        flash("Set the State first, then fetch.", "error")
        return redirect(url_for("scale_runs.cu_dashboard", short_name=short_name))

    try:
        import fetch_econ_data
        fetched = fetch_econ_data.fetch_economic_data(st, cty) or {}
    except Exception as exc:  # noqa: BLE001
        flash(f"Fetch failed: {exc}", "error")
        return redirect(url_for("scale_runs.cu_dashboard", short_name=short_name))

    econ["state"] = st
    econ["county"] = cty
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
    wizard_drafts.save_draft(
        workspace_root, state, active_step="scale_run", model="scale",
    )
    if got:
        flash(
            "Fetched " + ", ".join(got)
            + " from federal sources. Foreclosures has no federal source "
            "— enter it manually if needed.",
            "success",
        )
    else:
        flash(
            "No values could be fetched for "
            f"{st}{(' / ' + cty) if cty else ''}. Check the state name, or "
            "enter the factors manually.",
            "warning",
        )
    return redirect(url_for("scale_runs.cu_dashboard", short_name=short_name))


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
    sc["cu_name_override"] = (request.form.get("cu_name_override") or "").strip()
    # Persist the override + run params before the transient per-run
    # impaired override (below) is applied — that one is never persisted.
    wizard_drafts.save_draft(
        workspace_root, state, active_step="scale_run", model="scale",
    )

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
        # If the user can successfully run reports from the SCALE
        # dashboard, treat setup as completed for Home dashboard grouping.
        try:
            wizard_drafts.save_draft(
                workspace_root, state, active_step="scale_run", model="scale",
            )
            wizard_drafts.mark_completed(
                workspace_root,
                wizard_drafts.draft_key_for_state(state),
                model="scale",
            )
        except Exception as exc:  # noqa: BLE001
            flash(
                f"Report generated, but setup completion status could not be updated: {exc}",
                "warning",
            )
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


def _seed_scale_block(state: dict[str, Any]) -> dict[str, Any]:
    """Ensure ``state['scale']`` exists with every default key present so
    the settings editor can read/write mgmt-adj and life-of-loan fields
    even on older drafts."""
    sc = state.setdefault("scale", {})
    for k, v in _default_scale_block().items():
        sc.setdefault(k, v)
    return sc


@scale_runs_bp.route("/<short_name>/settings", methods=["GET", "POST"])
def edit_settings(short_name: str):
    """Edit per-pool Management Adjustments and ACL (Life of Loan) months
    for an already-configured SCALE credit union — the SCALE analogue of
    the Migration model's ``run.edit_settings`` page.

    Reads/writes the CU's SCALE wizard draft (``mgmt_adj`` and
    ``life_of_loan_overrides``) so the values are re-applied on every
    future quarter's run. SCALE pools are standardized (no per-pool
    risk-rated flag), so that column is intentionally absent.
    """
    workspace_root = _workspace_root()
    state = runs_service.load_state_for_cu(workspace_root, short_name)
    if state is None:
        flash(
            f"No saved wizard draft for {short_name}. Re-open the SCALE "
            "wizard for this CU and save it once before editing settings.",
            "error",
        )
        return redirect(url_for("scale_runs.cu_dashboard",
                                short_name=short_name))

    sc = _seed_scale_block(state)
    ma = sc.setdefault("mgmt_adj", {
        "default_pct": 0.0011, "pool_rows": {}, "portfolio": {},
    })
    overrides = sc.setdefault("life_of_loan_overrides", {})

    try:
        from cecl_ui.services import admin_defaults
        admin = admin_defaults.load() or {}
        admin_default_pct = float(admin.get("default_mgmt_adj") or 0.0011)
    except Exception:  # noqa: BLE001
        admin_default_pct = 0.0011

    tmpl_path = _resolved_template_path(sc)
    pool_names: list[str] = []
    months_by_name: dict[str, Any] = {}
    if tmpl_path:
        try:
            pool_names = mgmt_adj_writer.list_pool_names(tmpl_path)
        except Exception as exc:  # noqa: BLE001
            current_app.logger.warning(
                "scale edit_settings: pool names read failed (%s): %s",
                tmpl_path, exc,
            )
        try:
            months_by_name = {
                r["name"]: r.get("months")
                for r in lol_writer.read_lol_months(tmpl_path)
            }
        except Exception as exc:  # noqa: BLE001
            current_app.logger.warning(
                "scale edit_settings: lol months read failed (%s): %s",
                tmpl_path, exc,
            )

    if request.method == "POST":
        action = request.form.get("action", "save")

        if action == "reset":
            ma["default_pct"] = admin_default_pct
            ma["pool_rows"] = {}
            ma["portfolio"] = {"hard_code_pct": 0.0, "use_default": False}
            sc["life_of_loan_overrides"] = {}
            wizard_drafts.save_draft(
                workspace_root, state,
                active_step="scale_mgmt_adj", model="scale",
            )
            flash("Management Adjustment and ACL Months overrides reset "
                  "to defaults.", "info")
            return redirect(url_for("scale_runs.edit_settings",
                                    short_name=short_name))

        ma["default_pct"] = _coerce_pct_form(
            request.form.get("default_pct", "")
        )

        new_rows: dict[str, dict] = {}
        new_lol: dict[str, int] = {}
        for i, name in enumerate(pool_names):
            new_rows[name] = {
                "hard_code_pct": _coerce_pct_form(
                    request.form.get(f"pool_pct__{i}", "")
                ),
                "use_default": _coerce_bool_form(
                    request.form.get(f"pool_use_default__{i}", "")
                ),
            }
            months = _coerce_months_form(
                request.form.get(f"lol_months__{i}", "")
            )
            template_default = months_by_name.get(name) or 0
            if months > 0 and months != template_default:
                new_lol[name] = months
        ma["pool_rows"] = new_rows
        ma["portfolio"] = {
            "hard_code_pct": _coerce_pct_form(
                request.form.get("portfolio_pct", "")
            ),
            "use_default": _coerce_bool_form(
                request.form.get("portfolio_use_default", "")
            ),
        }
        sc["mgmt_adj"] = ma
        sc["life_of_loan_overrides"] = new_lol

        wizard_drafts.save_draft(
            workspace_root, state,
            active_step="scale_mgmt_adj", model="scale",
        )
        flash("Settings saved.", "success")
        return redirect(url_for("scale_runs.edit_settings",
                                short_name=short_name))

    # Build display rows merging template pool list with saved values.
    saved_rows = ma.get("pool_rows") or {}
    rows: list[dict[str, Any]] = []
    for i, name in enumerate(pool_names):
        if name in saved_rows:
            sv = saved_rows[name] or {}
            use_default = bool(sv.get("use_default"))
        else:
            sv = {}
            use_default = name in _MGMT_ADJ_DEFAULT_ON_POOLS
        default_months = months_by_name.get(name)
        override = overrides.get(name)
        try:
            override_int = int(override) if override else 0
        except (TypeError, ValueError):
            override_int = 0
        rows.append({
            "idx": i,
            "name": name,
            "hard_code_pct": float(sv.get("hard_code_pct") or 0.0),
            "use_default": use_default,
            "default_months": default_months,
            "override_months": override_int,
        })

    port = ma.get("portfolio") or {}
    portfolio = {
        "hard_code_pct": float(port.get("hard_code_pct") or 0.0),
        "use_default": bool(port.get("use_default")),
    }

    return render_template(
        "scale_runs/edit_settings.html",
        short_name=short_name,
        credit_union=state.get("credit_union") or short_name,
        rows=rows,
        portfolio=portfolio,
        default_pct=float(ma.get("default_pct") or 0.0),
        admin_default_pct=admin_default_pct,
        template_path=tmpl_path,
        template_missing=(not tmpl_path),
    )


def _scale_impaired_view(short_name: str, state: dict[str, Any] | None):
    """Render the standalone SCALE impaired-upload page for the current
    draft state."""
    imp = ((state or {}).get("scale") or {}).get("impaired_file") or {}
    parsed = imp.get("parsed") or {}
    preview = (parsed.get("rows") or [])[:25]
    return render_template(
        "scale_runs/impaired_upload.html",
        short_name=short_name,
        credit_union=(state or {}).get("credit_union") or short_name,
        draft_present=state is not None,
        imp=imp,
        parsed=parsed,
        preview=preview,
    )


@scale_runs_bp.route("/<short_name>/impaired", methods=["GET"])
def upload_impaired_form(short_name: str):
    """Standalone Impaired Loans upload page for a SCALE credit union —
    the SCALE analogue of the Migration model's
    ``run.upload_impaired_form``. Shows the impaired workbook currently
    saved in the wizard draft (reused on every future run)."""
    workspace_root = _workspace_root()
    state = runs_service.load_state_for_cu(workspace_root, short_name)
    if state is None:
        flash(
            f"No saved wizard draft for {short_name}. Re-open the SCALE "
            "wizard for this CU and save it once before uploading an "
            "impaired file.",
            "warning",
        )
        return redirect(url_for("scale_runs.cu_dashboard",
                                short_name=short_name))
    return _scale_impaired_view(short_name, state)


@scale_runs_bp.route("/<short_name>/impaired", methods=["POST"])
def upload_impaired(short_name: str):
    """Save (or clear) the persistent Impaired Loans workbook in the CU's
    SCALE wizard draft. Unlike the per-run upload on the dashboard, this
    updates the draft so the file becomes the default for every future
    quarter's run — matching the Migration model's standalone uploader."""
    workspace_root = _workspace_root()
    state = runs_service.load_state_for_cu(workspace_root, short_name)
    if state is None:
        flash(
            f"No saved wizard draft for {short_name}. Re-open the SCALE "
            "wizard for this CU and save it once first.",
            "error",
        )
        return redirect(url_for("scale_runs.cu_dashboard",
                                short_name=short_name))

    sc = _seed_scale_block(state)
    action = request.form.get("action", "upload")

    if action == "clear":
        sc["impaired_file"] = {}
        wizard_drafts.save_draft(
            workspace_root, state,
            active_step="scale_impaired", model="scale",
        )
        flash("Impaired loans file cleared from this CU's draft.", "info")
        return redirect(url_for("scale_runs.upload_impaired_form",
                                short_name=short_name))

    f = request.files.get("impaired_file")
    if not f or not f.filename:
        flash("Choose an Impaired Loans .xlsx file to upload.", "error")
        return redirect(url_for("scale_runs.upload_impaired_form",
                                short_name=short_name))

    ext = Path(f.filename).suffix.lower()
    if ext not in (".xlsx", ".xlsm"):
        flash(
            "Impaired Loans workbook must be an .xlsx or .xlsm file "
            f"(got {ext or 'no extension'}).",
            "error",
        )
        return redirect(url_for("scale_runs.upload_impaired_form",
                                short_name=short_name))

    try:
        target = _save_impaired_upload(f)
    except Exception as exc:  # noqa: BLE001
        flash(f"Save failed: {exc}", "error")
        return redirect(url_for("scale_runs.upload_impaired_form",
                                short_name=short_name))

    parsed = impaired_loader.parse_file(target)
    sc["impaired_file"] = {
        "saved_path": str(target),
        "uploaded_filename": f.filename,
        "parsed": parsed,
    }
    wizard_drafts.save_draft(
        workspace_root, state,
        active_step="scale_impaired", model="scale",
    )
    if parsed.get("ok"):
        flash(
            f"Saved impaired workbook: parsed {parsed['row_count']} "
            f"impaired loan(s) totalling "
            f"${parsed['total_balance']:,.2f}. It will be used by "
            "future report runs for this CU.",
            "success",
        )
    else:
        flash(f"Uploaded, but parse failed: {parsed.get('error')}", "warning")

    return _scale_impaired_view(short_name, state)
