"""Top-level routes: home (new vs existing) and model picker."""
from __future__ import annotations

from flask import Blueprint, current_app, render_template, request, redirect, session, url_for

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
    drafts = wizard_drafts.list_drafts(current_app.config["WORKSPACE_ROOT"])
    return render_template("home.html", clients=clients, drafts=drafts)


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
