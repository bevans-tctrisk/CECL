"""Admin / firm-wide defaults page.

Single-user, on-machine settings. No auth gate by design.
"""
from __future__ import annotations

from flask import Blueprint, flash, redirect, render_template, request, url_for

from cecl_ui.services import admin_defaults


admin_bp = Blueprint("admin", __name__)


def _parse_decimal(raw: str) -> float | None:
    """Accept '0.05' or '5%' style input. Returns decimal float (or None)."""
    s = (raw or "").strip()
    if s == "":
        return None
    pct = s.endswith("%")
    if pct:
        s = s[:-1].strip()
    try:
        v = float(s)
    except ValueError:
        return None
    if pct:
        v = v / 100.0
    return v


def _save_mgmt_adj(values: dict) -> None:
    raw = (request.form.get("default_mgmt_adj_pct") or "").strip()
    try:
        pct = float(raw)
    except ValueError:
        flash("Default Management Adjustment must be a number.", "error")
        return
    if pct < -100 or pct > 100:
        flash(
            "Default Management Adjustment must be between -100% and 100%.",
            "error",
        )
        return
    values["default_mgmt_adj"] = round(pct / 100.0, 6)
    admin_defaults.save(values)
    flash(f"Saved. Default Management Adjustment = {pct:g}%.", "success")


def _save_env_factor_ranges(values: dict) -> None:
    delq_rows: list[list[float]] = []
    bad: list[str] = []
    for i in range(admin_defaults.DELINQUENCY_ROW_COUNT):
        mn = _parse_decimal(request.form.get(f"delq_min_{i}", ""))
        sc = _parse_decimal(request.form.get(f"delq_score_{i}", ""))
        if mn is None or sc is None:
            bad.append(f"Delinquency row {i + 1}")
            continue
        delq_rows.append([round(mn, 6), round(sc, 6)])

    econ_rows: list[list[float]] = []
    for i in range(admin_defaults.ECON_STRESS_ROW_COUNT):
        mn = _parse_decimal(request.form.get(f"econ_min_{i}", ""))
        sc = _parse_decimal(request.form.get(f"econ_score_{i}", ""))
        if mn is None or sc is None:
            bad.append(f"Economic Stress row {i + 1}")
            continue
        econ_rows.append([round(mn, 6), round(sc, 6)])

    if bad:
        flash(
            "Some rows could not be parsed as numbers and were skipped: "
            + ", ".join(bad),
            "error",
        )

    if (
        len(delq_rows) != admin_defaults.DELINQUENCY_ROW_COUNT
        or len(econ_rows) != admin_defaults.ECON_STRESS_ROW_COUNT
    ):
        flash(
            "Environmental Factor Ranges were NOT saved -- every cell "
            "must be filled in.",
            "error",
        )
        return

    values["env_factor_ranges"] = {
        "delinquency": delq_rows,
        "econ_stress": econ_rows,
    }
    admin_defaults.save(values)
    flash(
        "Saved Environmental Factor Ranges. New SCALE reports will use "
        "these values; historical reports are unchanged.",
        "success",
    )


def _reset_env_factor_ranges(values: dict) -> None:
    sys_ranges = admin_defaults.SYSTEM_DEFAULTS["env_factor_ranges"]
    values["env_factor_ranges"] = {
        "delinquency": [list(r) for r in sys_ranges["delinquency"]],
        "econ_stress": [list(r) for r in sys_ranges["econ_stress"]],
    }
    admin_defaults.save(values)
    flash(
        "Environmental Factor Ranges reset to built-in defaults.",
        "success",
    )


@admin_bp.route("/", methods=["GET", "POST"])
def index():
    values = admin_defaults.load()

    if request.method == "POST":
        action = (request.form.get("action") or "save_mgmt_adj").strip()
        if action == "save_env_factor_ranges":
            _save_env_factor_ranges(values)
        elif action == "reset_env_factor_ranges":
            _reset_env_factor_ranges(values)
        else:
            _save_mgmt_adj(values)
        return redirect(url_for("admin.index"))

    # GET: format decimal back to percent for display.
    pct_str = "%g" % (float(values.get("default_mgmt_adj", 0.0)) * 100.0)
    env_ranges = values.get("env_factor_ranges") or {}
    return render_template(
        "admin.html",
        default_mgmt_adj_pct=pct_str,
        default_mgmt_adj_decimal=values.get("default_mgmt_adj", 0.0),
        delinquency_rows=env_ranges.get("delinquency") or [],
        econ_stress_rows=env_ranges.get("econ_stress") or [],
        delinquency_row_count=admin_defaults.DELINQUENCY_ROW_COUNT,
        econ_stress_row_count=admin_defaults.ECON_STRESS_ROW_COUNT,
    )
