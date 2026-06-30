"""Top-level routes: home (new vs existing) and model picker."""
from __future__ import annotations

from datetime import datetime

from flask import (
    Blueprint, current_app, flash, render_template, request, redirect,
    session, url_for,
)

from cecl_ui.routes.setup import STATE_KEY
from cecl_ui.services import config_service, wizard_drafts


home_bp = Blueprint("home", __name__)


@home_bp.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        choice = request.form.get("choice")
        if choice == "new":
            # Start with a clean slate — drop any in-memory wizard state
            # from a previous CU so the new setup doesn't inherit their
            # name / state / county / pools / etc. Drafts on disk are
            # untouched (user can still "Resume" them below).
            session.pop(STATE_KEY, None)
            session.modified = True
            return redirect(url_for("home.model_select"))
        if choice == "existing":
            return redirect(url_for("run.select_cu"))
    clients = config_service.list_existing_clients(current_app.config["WORKSPACE_ROOT"])
    all_drafts = wizard_drafts.list_drafts(current_app.config["WORKSPACE_ROOT"])
    drafts = [d for d in all_drafts if not d.get("completed_at")]
    completed_drafts = [d for d in all_drafts if d.get("completed_at")]
    # Mark which configured CUs have NO migration draft yet (so the
    # "Adopt as completed setup" button can show only where useful).
    drafted_slugs = {d["key"] for d in all_drafts if d.get("model") == "migration"}
    for c in clients:
        c["has_migration_draft"] = c["short_name"] in drafted_slugs
    return render_template(
        "home.html",
        clients=clients,
        drafts=drafts,
        completed_drafts=completed_drafts,
    )


@home_bp.route("/adopt-config/<short_name>", methods=["POST"])
def adopt_config(short_name: str):
    """Create a completed migration draft from an existing YAML config.

    For credit unions that were configured BEFORE the wizard's
    Completed-setup section existed, this lifts the YAML into a
    draft + stamps it completed so it shows up in the dashboard
    with an Edit-setup button.
    """
    workspace = current_app.config["WORKSPACE_ROOT"]
    try:
        cfg = config_service.load_client_config(workspace, short_name)
    except FileNotFoundError:
        flash(f"No config found for '{short_name}'.", "error")
        return redirect(url_for("home.index"))
    if not cfg:
        flash(f"Config for '{short_name}' is empty.", "error")
        return redirect(url_for("home.index"))
    adopted = adopt_config_to_completed(workspace, short_name, cfg)
    if adopted:
        flash(
            f"Adopted '{cfg.get('credit_union') or short_name}' into Completed setup. "
            "Click Edit setup to revise any step.",
            "success",
        )
    else:
        flash(
            f"'{cfg.get('credit_union') or short_name}' already has a wizard draft.",
            "info",
        )
    return redirect(url_for("home.index"))


def adopt_config_to_completed(
    workspace: str, short_name: str, cfg: dict | None = None,
) -> bool:
    """Lift an existing YAML config into a Completed migration draft.

    Idempotent: when a migration draft already exists (in-progress OR
    completed), this is a no-op and returns False. Returns True on
    successful adoption. Safe to call from any successful-report
    code path so credit unions configured outside the wizard
    automatically show up in the Completed setup table on the home
    page.
    """
    if cfg is None:
        try:
            cfg = config_service.load_client_config(workspace, short_name)
        except FileNotFoundError:
            return False
    if not cfg:
        return False
    # Skip if a migration draft (any state) already exists.
    for d in wizard_drafts.list_drafts(workspace):
        if d.get("key") == short_name and d.get("model") == "migration":
            return False
    state = dict(cfg)
    state["short_name"] = short_name
    state["model"] = "migration"
    # Translate YAML-config schema (``pools`` / ``monthly_balance`` /
    # ``acl`` / ``historical_file_formats``) into wizard-state schema
    # (``pool_settings`` / ``monthly_bal`` / ``acl_balance`` / ``co_columns``)
    # so an adopted draft renders correctly when the user clicks "Edit
    # setup". Without this, Step 2 (Loan Pools) shows an empty table
    # even though the YAML has every pool defined. Best-effort: any
    # translation failure leaves the field untranslated and the
    # wizard's GET-time self-heal will catch it on the next render.
    try:
        from cecl_ui.services import auto_setup as _auto_setup_adopt
        _auto_setup_adopt.selfheal_adopt_yaml_schema_into_state(state)
    except Exception:  # noqa: BLE001
        pass
    state[wizard_drafts.DRAFT_META_KEY] = {
        "model": "migration",
        "active_step": "review",
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "completed_at": datetime.now().isoformat(timespec="seconds"),
        "adopted_from_config": True,
    }
    try:
        wizard_drafts.save_draft(
            workspace, state, active_step="review", model="migration",
        )
        return True
    except Exception:  # noqa: BLE001
        # Adoption is a UX-nicety; never raise into a calling code path.
        return False


@home_bp.route("/model-select", methods=["GET", "POST"])
def model_select():
    if request.method == "POST":
        model = request.form.get("model")
        if model == "migration":
            return redirect(url_for("setup.start_warm_choice"))
        if model == "scale":
            return redirect(url_for("scale_setup.start"))
        # Unknown choice — fall through and re-render.
    return render_template("model_select.html")
