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
    state = dict(cfg)
    state["short_name"] = short_name
    state["model"] = "migration"
    state[wizard_drafts.DRAFT_META_KEY] = {
        "model": "migration",
        "active_step": "review",
        "saved_at": datetime.now().isoformat(timespec="seconds"),
        "completed_at": datetime.now().isoformat(timespec="seconds"),
        "adopted_from_config": True,
    }
    wizard_drafts.save_draft(
        workspace, state, active_step="review", model="migration",
    )
    flash(
        f"Adopted '{cfg.get('credit_union') or short_name}' into Completed setup. "
        "Click Edit setup to revise any step.",
        "success",
    )
    return redirect(url_for("home.index"))


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
