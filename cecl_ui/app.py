"""CECL UI - Flask application entry point.

Web-based wizard that walks the user through:
  1. Setting up a new Credit Union (CU) for the CECL Migration Model
  2. Running quarterly reports for an already-configured CU

Launch:
    python run_ui.py
or:
    python -m cecl_ui.app

Then open http://127.0.0.1:5000 in a browser.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

from flask import Flask
from flask_session import Session

# Make sure the workspace root is on sys.path so we can import the existing
# CECL modules (cecl_engine, generate_report, import_data, ...).
WORKSPACE_ROOT = Path(__file__).resolve().parent.parent
if str(WORKSPACE_ROOT) not in sys.path:
    sys.path.insert(0, str(WORKSPACE_ROOT))

from cecl_ui.routes.home import home_bp
from cecl_ui.routes.setup import setup_bp
from cecl_ui.routes.scale_setup import scale_setup_bp
from cecl_ui.routes.scale_runs import scale_runs_bp
from cecl_ui.routes.run import run_bp
from cecl_ui.routes.admin import admin_bp


def create_app() -> Flask:
    app = Flask(
        __name__,
        template_folder="templates",
        static_folder="static",
    )
    # Session secret — for local single-user use.  Override via env for prod.
    app.config["SECRET_KEY"] = os.environ.get(
        "CECL_UI_SECRET", "dev-secret-change-me"
    )
    app.config["WORKSPACE_ROOT"] = str(WORKSPACE_ROOT)
    app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB upload cap

    # ------------------------------------------------------------------
    # Server-side sessions.  The wizard accumulates pool-code lists,
    # sample previews, column maps, etc. — easily exceeding the 4 KB
    # browser-cookie limit.  Storing the state in files keeps the cookie
    # down to a small session id and lets the wizard hold arbitrarily
    # large state.
    #
    # Session files live in the OS-local temp directory (NOT the
    # workspace root) because the workspace often sits on a slow
    # network share (Egnyte).  Filesystem operations during request
    # handling on a network drive add seconds of latency per click.
    # ------------------------------------------------------------------
    import tempfile
    session_dir = Path(tempfile.gettempdir()) / "cecl_ui_sessions"
    session_dir.mkdir(exist_ok=True)
    app.config.update(
        SESSION_TYPE="filesystem",
        SESSION_FILE_DIR=str(session_dir),
        SESSION_PERMANENT=False,
        SESSION_USE_SIGNER=True,
        SESSION_FILE_THRESHOLD=200,  # auto-purge oldest when this many files
    )
    Session(app)

    app.register_blueprint(home_bp)
    app.register_blueprint(setup_bp, url_prefix="/setup")
    app.register_blueprint(scale_setup_bp, url_prefix="/scale")
    app.register_blueprint(scale_runs_bp, url_prefix="/scale-runs")
    app.register_blueprint(run_bp, url_prefix="/run")
    app.register_blueprint(admin_bp, url_prefix="/admin")

    return app


app = create_app()


if __name__ == "__main__":
    app.run(host="127.0.0.1", port=5000, debug=True)
