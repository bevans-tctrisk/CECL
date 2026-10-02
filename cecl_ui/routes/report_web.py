"""Browser preview + PDF download for the new HTML-rendered reports.

Renders the generated Vizo workbook straight to a scrollable HTML page
(``/report-web/<short>/<snap>``) or a merged PDF (``.../pdf``) using the
``cecl_report_web`` renderer. Read-only; sits alongside the existing xlsx
downloads while the migration is in progress.
"""
from __future__ import annotations

from flask import Blueprint, Response, abort, current_app, request

from cecl_ui.services import config_service
from cecl_report_web import assembly


report_web_bp = Blueprint("report_web", __name__)


def _resolve(short_name: str, snapshot: str, report_type: str):
    ws = current_app.config["WORKSPACE_ROOT"]
    cfg = config_service.load_client_config(ws, short_name)
    if not cfg:
        abort(404, f"No config for {short_name!r}")
    cu = cfg.get("credit_union") or short_name
    supplemental = report_type in ("supp", "vizo_supp", "supplemental")
    report_path = assembly.find_report(
        ws, cu, snapshot, supplemental=supplemental)
    if not report_path:
        kind = "Supplemental" if supplemental else "Model"
        abort(404, f"No {kind} report found for {cu} at {snapshot}. "
                   f"Generate it first.")
    return cfg, cu, report_path, supplemental


@report_web_bp.route("/<short_name>/<snapshot>")
def preview(short_name: str, snapshot: str):
    """Render the whole report as one scrollable HTML page."""
    report_type = request.args.get("type", "vizo")
    _cfg, cu, report_path, supplemental = _resolve(
        short_name, snapshot, report_type)
    html = assembly.render_report_html(
        report_path, cu, snapshot, supplemental=supplemental)
    return Response(html, mimetype="text/html")


@report_web_bp.route("/<short_name>/<snapshot>/pdf")
def download_pdf(short_name: str, snapshot: str):
    """Render + download the whole report as a single merged PDF."""
    report_type = request.args.get("type", "vizo")
    _cfg, cu, report_path, supplemental = _resolve(
        short_name, snapshot, report_type)
    pdf = assembly.render_report_pdf(
        report_path, cu, snapshot, supplemental=supplemental)
    safe_cu = cu.replace(" ", "_").replace("/", "-")
    kind = "Supplemental" if supplemental else "Vizo_Model"
    fname = f"{snapshot}_CECL_{kind}_{safe_cu}.pdf"
    return Response(
        pdf, mimetype="application/pdf",
        headers={"Content-Disposition": f'inline; filename="{fname}"'},
    )
