"""Convenience launcher for the CECL UI Flask app.

Usage:
    python run_ui.py

Opens on http://127.0.0.1:5000.
"""
from cecl_ui.app import app


if __name__ == "__main__":
    print("=" * 60)
    print("  CECL UI starting on http://127.0.0.1:5000")
    print("  Press CTRL+C to stop the server.")
    print("=" * 60)
    # Egnyte/network-drive file events can trigger constant reloads.
    # Keep debug on, but disable auto-reloader for a stable local session.
    app.run(host="127.0.0.1", port=5000, debug=True, use_reloader=False)
