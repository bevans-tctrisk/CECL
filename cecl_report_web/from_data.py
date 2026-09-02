"""Compute-time populator: build the ReportModel from report DATA.

The counterpart to :mod:`cecl_report_web.from_workbook`, and the whole point
of the PDF migration: populate the *same* model dataclasses the templates
already consume, but from ``(client_name, snapshot_date, config, grades,
hist, df)`` -- the inputs the report engine itself uses -- instead of by
scraping a generated ``.xlsx``. No openpyxl, no cell coordinates, no reading
business meaning out of fills or fonts.

Pages are added here one archetype at a time; each returns the identical node
``from_workbook`` returns, so ``render.py`` / the Jinja templates are unchanged.
"""

from __future__ import annotations

import base64
import datetime as _dt
import os
from typing import Any

from .model import (
    AclEnvPage,
    AclPoolRow,
    AdjustmentRow,
    ChartSpec,
    CoverPage,
    ImprDeterPage,
    KeyValueRow,
    MatrixCell,
    MatrixRow,
    RiskChangeMatrix,
    RiskChangePage,
    TableCell,
    TablePage,
    TableSection,
)


def _snap_parts(snapshot_date: str) -> tuple[int, int, int]:
    d = _dt.date.fromisoformat(str(snapshot_date)[:10])
    return d.month, d.day, d.year


def _date_text(snapshot_date: str) -> str:
    """Match the workbook cover's ``m/d/yyyy`` display (no leading zeros)."""
    m, d, y = _snap_parts(snapshot_date)
    return f"{m}/{d}/{y}"


def _logo_data_uri(path: str) -> str | None:
    """Base64-encode a PNG asset as a ``data:`` URI for inline <img> use."""
    try:
        if not path or not os.path.isfile(path):
            return None
        with open(path, "rb") as fh:
            raw = fh.read()
        return "data:image/png;base64," + base64.b64encode(raw).decode("ascii")
    except OSError:
        return None


# Verbatim from report_vizo.sheet_cover_vizo (B21) so the from-data cover
# reads identically to the workbook cover.
_DISCLAIMER = (
    "The following analysis and all parts thereof (\u2018analysis\u2019) are based upon "
    "information obtained by Vizo Financial Corporate Credit Union (Vizo Financial) from "
    "the credit union that is subject of this analysis and other sources that Vizo Financial "
    "believes to be reliable and utilized in models using methods and assumptions which Vizo "
    "Financial believes to be reasonable.  However, actual performance compared to estimated "
    "performance of the subject credit union may be different and cannot be guaranteed.  This "
    "analysis is for informational purposes only and is intended only for the use of the "
    "subject credit union.  The analysis does not constitute either legal or tax advice. \n"
    "All reports are confidential."
)


def build_cover(client_name: str, snapshot_date: str, config: dict,
                *, supplemental: bool = False) -> CoverPage:
    """Populate the Vizo cover from config/period alone (no workbook).

    Mirrors ``report_vizo.sheet_cover_vizo``: title, CU name, period date,
    the confidentiality paragraph, the copyright footer, and the two bundled
    logos resolved from ``report_vizo``'s asset paths.
    """
    cu = (config or {}).get("credit_union") or client_name

    # Reuse the report engine's own logo paths so the asset resolves the same
    # way it does for the workbook; import lazily to avoid a heavy import at
    # module load.
    try:
        import report_vizo as _rv
        top = _logo_data_uri(getattr(_rv, "LOGO_VIZO", ""))
        bottom = _logo_data_uri(getattr(_rv, "LOGO_TCT", ""))
    except Exception:  # noqa: BLE001 - logos are optional; cover still renders
        top = bottom = None

    date_text = _date_text(snapshot_date)
    return CoverPage(
        credit_union=cu,
        period_ending=str(snapshot_date)[:10],
        title="CECL Credit Migration Report",
        subtitle="Supplemental Reports" if supplemental else None,
        date_text=date_text,
        paragraph=_DISCLAIMER,
        footer=f"\u00a9 {_snap_parts(snapshot_date)[2]} TCT Risk Solutions",
        top_logo=top,
        bottom_logo=bottom,
    )


def _cell_state(i: int, j: int, cur_grade: str, orig_grade: str,
                no_score: str, n_top: int) -> str:
    """Improved / deteriorated / plain, computed from grade ORDERING.

    Mirrors ``report_vizo._sheet_risk_change`` exactly -- the state that the
    workbook encodes in the cell fill is derived here from position instead:
    a current grade worse than original deteriorates, unless the original is
    among the top ``n_top`` grades and the drop is a single band (the WARM
    "top grades need a 2+ drop" rule). Not-Reported never migrates.
    """
    if cur_grade == no_score or orig_grade == no_score:
        return "plain"
    if i > j:  # current grade ranked below original -> potential deterioration
        if j < n_top and (i - j) < 2:
            return "plain"
        return "deteriorated"
    if i < j:  # current grade ranked above original -> improvement
        return "improved"
    return "plain"


def build_risk_change(client_name: str, snapshot_date: str, df: Any,
                      config: dict, grades: Any,
                      hist: dict | None = None) -> RiskChangePage:
    """Populate the "Risk Change Total" page from the loan frame.

    Reuses the report engine's own compute (``cecl_engine.risk_change_matrix``
    plus report_vizo's grade helpers) so the numbers are identical to the
    workbook; the improved/deteriorated state is computed from grade ordering,
    not read from a cell fill.
    """
    import report_vizo as _rv
    from cecl_engine import risk_change_matrix

    no_score = (config or {}).get("no_score_label", "Not Reported")
    n_top = int((config or {}).get("top_grades_double_drop", 3))
    gl = [g for g in _rv._all_grades(grades, no_score) if not _rv._is_hidden(g)]
    matrix = risk_change_matrix(df, grades, no_score)
    rng = _rv._grade_ranges(grades, no_score)
    total = float(df["current_balance"].sum())
    cu = (config or {}).get("credit_union") or client_name

    def _mv(cur: str, og: str) -> float:
        return float(_rv._matrix_val(matrix, cur, og) or 0.0)

    col_totals = {og: sum(_mv(g2, og) for g2 in gl) for og in gl}

    def _build(is_pct: bool) -> RiskChangeMatrix:
        rows: list[MatrixRow] = []
        for i, g in enumerate(gl):
            cells: list[MatrixCell] = []
            rtotal = 0.0
            for j, og in enumerate(gl):
                v = _mv(g, og)
                rtotal += v
                val = (v / col_totals[og] if col_totals[og] else 0.0) if is_pct else v
                cells.append(MatrixCell(
                    value=val, state=_cell_state(i, j, g, og, no_score, n_top),
                    is_pct=is_pct))
            tot = (rtotal / total if total else 0.0) if is_pct else rtotal
            rows.append(MatrixRow(
                label=g, range_label=rng.get(g, ""), cells=cells,
                total=MatrixCell(value=tot, is_pct=is_pct, bold=True)))
        # Grand Total row: column sums ($) / 100% (pct).
        gt_cells = [MatrixCell(value=(1.0 if is_pct else col_totals[og]),
                               is_pct=is_pct, bold=True) for og in gl]
        rows.append(MatrixRow(
            label="Grand Total", range_label="", cells=gt_cells,
            total=MatrixCell(value=(1.0 if is_pct else total),
                             is_pct=is_pct, bold=True)))
        return RiskChangeMatrix(
            corner=("% Current Grade" if is_pct else "$ Current Grade"),
            col_headers=list(gl), rows=rows, is_pct=is_pct)

    _imp = (hist or {}).get("impaired", {}) or {}
    bal_adj = float(_imp.get("total_balance_adjustment", 0.0) or 0.0)
    tip = _imp.get("total_in_portfolio") or (total + bal_adj)
    summary = [
        KeyValueRow(label="Balance Adjustment", value=bal_adj),
        KeyValueRow(label="Total in Portfolio", value=float(tip)),
    ]
    heading = [
        "Executive Summary Total Loans",
        "Risk Change By Credit Score",
        f"For Quarter Ending {_rv._snap_display(snapshot_date)}",
    ]
    return RiskChangePage(credit_union=cu, heading_lines=heading,
                          matrices=[_build(False), _build(True)],
                          summary=summary)


_ACL_COL_HEADERS = [
    "Current Grade", "Balance", "Specific Identification",
    "Loan Loss Calc Balance", "ACL Base Loss Rate", "Mgmt Adj",
    "Allowance Factor", "Allowance before Env", "Env Factor",
    "Env Allowance", "Total Allowance",
]
_ACL_ADJ_KEYS = [
    ("Total Specifically Identified Allowance", "total_spec_allow"),
    ("Total Allowance Needed", "total_allow_needed"),
    ("Allowance for Credit Loss Balance", "acl_balance"),
    ("Adjustment", "adjustment"),
]


def _acl_data(config: dict, snapshot_date: str, hist: dict | None,
              df: Any, grades: Any) -> tuple[dict, dict, dict]:
    """(acl_pools, acl_summary, acl_impaired): the values report_vizo publishes,
    then the WARM-parsed dicts, then a standalone compute when df/grades exist."""
    import report_vizo as _rv

    _imp = (hist or {}).get("impaired", {}) or {}
    acl_pools = _imp.get("_acl_pools_computed") or _imp.get("acl_pools") or {}
    acl_summary = _imp.get("_acl_summary_computed") or _imp.get("acl_summary") or {}
    acl_impaired = _imp.get("_acl_impaired_computed") or _imp.get("acl_impaired") or {}
    if not acl_pools and df is not None:
        computed = _rv.compute_acl_environmental(df, grades, config, hist, snapshot_date)
        acl_pools = computed.get("acl_pools") or {}
        acl_summary = computed.get("acl_summary") or acl_summary
        acl_impaired = computed.get("acl_impaired") or acl_impaired
    return acl_pools, acl_summary, acl_impaired


def build_acl_env(client_name: str, snapshot_date: str, config: dict,
                  hist: dict | None = None, df: Any = None,
                  grades: Any = None) -> AclEnvPage | None:
    """Populate the "ACL Env by Pool Mgmt Adj" page from the ACL data dicts.

    Prefers the values ``report_vizo._sheet_acl_reserve`` publishes when the
    report was composed (isolated underscore keys), then the WARM-parsed dicts,
    and finally -- when neither is present and ``df``/``grades`` are supplied --
    computes them standalone via ``report_vizo.compute_acl_environmental`` so
    the page renders without the workbook being built at all.
    """
    import report_vizo as _rv

    acl_pools, acl_summary, acl_impaired = _acl_data(
        config, snapshot_date, hist, df, grades)
    if not acl_pools:
        return None
    pool_order = list(acl_pools.keys())
    cu = (config or {}).get("credit_union") or client_name
    cu = (config or {}).get("credit_union") or client_name

    def _match(pool: str) -> dict | None:
        if pool in acl_pools:
            return acl_pools[pool]
        lc = pool.strip().lower()
        return next((v for k, v in acl_pools.items()
                     if k.strip().lower() == lc), None)

    pool_rows: list[AclPoolRow] = []
    for pool in pool_order:
        pdata = _match(pool)
        if not pdata:
            continue
        pool_rows.append(AclPoolRow(pool=pool, kind="header"))
        for g, gv in (pdata.get("grades") or {}).items():
            if str(g).upper().startswith("HIDE"):
                continue
            pool_rows.append(AclPoolRow(
                pool=g, kind="grade",
                balance=gv.get("balance"), specific_id=gv.get("spec_id"),
                llc_balance=gv.get("calc_bal"), base_loss_rate=gv.get("base_rate"),
                mgmt_adj=gv.get("mgmt_adj"), allowance_factor=gv.get("factor"),
                allowance_before_env=gv.get("allow_before")))
        t = pdata.get("total") or {}
        _bal, _spec = t.get("balance"), t.get("spec_id")
        pool_rows.append(AclPoolRow(
            pool="Total", kind="total",
            balance=_bal, specific_id=_spec,
            llc_balance=((_bal or 0) - (_spec or 0)) if _bal is not None else None,
            base_loss_rate=t.get("base_rate"), mgmt_adj=t.get("mgmt_adj"),
            allowance_factor=t.get("factor"),
            allowance_before_env=t.get("allow_before"),
            env_factor=t.get("env_factor"), env_allowance=t.get("env_allow"),
            total_allowance=t.get("total_allow")))

    _pb = acl_summary.get("pooled_balance")
    _ps = acl_summary.get("pooled_spec_id")
    pooled = AclPoolRow(
        pool="Pooled Totals", is_total=True,
        balance=_pb, specific_id=_ps,
        llc_balance=((_pb or 0) - (_ps or 0)) if _pb is not None else None,
        allowance_before_env=acl_summary.get("pooled_allow_before"),
        env_allowance=acl_summary.get("pooled_env_allow"),
        total_allowance=acl_summary.get("pooled_total_allow"))

    impaired_rows = [AdjustmentRow(label=k, value=v)
                     for k, v in acl_impaired.items()]
    adjustment_rows = [
        AdjustmentRow(label=lbl, value=acl_summary.get(key), bold=True)
        for lbl, key in _ACL_ADJ_KEYS if key in acl_summary
    ]
    heading = [
        "Allowance & Provision for Credit Loss Reserve Analysis",
        f"For Quarter Ending {_rv._snap_display(snapshot_date)}",
    ]
    return AclEnvPage(
        credit_union=cu, heading_lines=heading, col_headers=list(_ACL_COL_HEADERS),
        pool_rows=pool_rows, pooled_totals=pooled,
        impaired_rows=impaired_rows, adjustment_rows=adjustment_rows)


def build_impr_deter(client_name: str, snapshot_date: str, config: dict,
                     hist: dict | None = None, df: Any = None,
                     grades: Any = None) -> ImprDeterPage:
    """Populate the "Impr Deter" page's CECL Adjustment box from data.

    The four headline allowance figures are exactly the ACL summary the ACL
    Env page already computes, so the box ties out to that tab. Reuses the
    published/computed ``acl_summary`` (standalone-computes it when absent).
    The improved/deteriorated charts are added once the ChartSpec node lands.
    """
    import report_vizo as _rv

    _imp = (hist or {}).get("impaired", {}) or {}
    acl_summary = _imp.get("_acl_summary_computed") or _imp.get("acl_summary") or {}
    if not acl_summary.get("total_allow_needed") and df is not None:
        acl_summary = (_rv.compute_acl_environmental(
            df, grades, config, hist, snapshot_date).get("acl_summary")
            or acl_summary)

    cu = (config or {}).get("credit_union") or client_name
    adj = float(acl_summary.get("adjustment") or 0.0)
    adj_label = ("Adjustment (Underfunded)" if adj >= 0
                 else "Adjustment (Overfunded)")
    cecl = [
        KeyValueRow(label="Total Specifically Identified Allowance",
                    value=acl_summary.get("total_spec_allow")),
        KeyValueRow(label="Total Allowance Needed",
                    value=acl_summary.get("total_allow_needed")),
        KeyValueRow(label=f"Allowance for Credit Loss Balance as of {snapshot_date}",
                    value=acl_summary.get("acl_balance")),
        KeyValueRow(label=adj_label, value=adj),
    ]
    heading = [
        "CECL Adjustment & Improved/Deteriorated",
        f"For Quarter Ending {_rv._snap_display(snapshot_date)}",
    ]
    return ImprDeterPage(
        credit_union=cu, period_ending=str(snapshot_date)[:10],
        heading_lines=heading, cecl_adjustment=cecl)


# ── Charts (rendered from data via cecl_report_web.charts) ───────────
# Migration-status slice colours, matching report_vizo's DQ pie / CO bar /
# Net-Credit-Change doughnut (olive / maroon / teal / gold).
_MIG_LABELS = ("Improved", "Deteriorated", "Unchanged", "Not Reported")
_MIG_COLORS = ("829901", "873A3A", "0D4D5E", "FFC000")
_NCC_COLORS = ("829901", "873A3A", "0D4D5E")


def _chartspec_to_render_dict(cs: ChartSpec) -> dict:
    """Adapt a semantic ChartSpec to the dict charts.render_chart_svg consumes."""
    lbl_fmt = "0.0%" if cs.value_format == "pct" else None
    if cs.kind in ("pie", "doughnut"):
        s0 = cs.series[0] if cs.series else {}
        ctype = "DoughnutChart" if cs.kind == "doughnut" else "PieChart"
        return {
            "type": ctype, "title": cs.title,
            "series": [{
                "name": s0.get("name"), "values": s0.get("values") or [],
                "cats": list(cs.categories),
                "point_colors": s0.get("colors"),
                "label_fmt": lbl_fmt, "show_labels": True,
            }],
        }
    bar_dir = "bar" if cs.kind in ("bar_h", "diverging_bar") else "col"
    grouping = "stacked" if cs.kind == "diverging_bar" else "clustered"
    return {
        "type": "BarChart", "bar_dir": bar_dir, "grouping": grouping,
        "title": cs.title,
        "series": [{
            "name": s.get("name"), "values": s.get("values") or [],
            "cats": list(cs.categories),
            "point_colors": s.get("colors"),
            "color": (s.get("colors") or [None])[0], "filled": True,
            "label_fmt": lbl_fmt, "show_labels": True,
        } for s in cs.series],
    }


def render_chart_specs(specs: list[ChartSpec]) -> list[str]:
    """Render a list of ChartSpecs to inline SVG strings."""
    from .charts import render_chart_svg
    out: list[str] = []
    for cs in specs:
        try:
            out.append(render_chart_svg(_chartspec_to_render_dict(cs)))
        except Exception:  # noqa: BLE001 - a bad chart must not sink the page
            continue
    return out


def _mig_status_series(data: dict, use_pct: bool = True) -> list[float] | None:
    """[Improved, Deteriorated, Unchanged, Not Reported] from a status dict."""
    if not data:
        return None
    key = "pct" if use_pct else "balance"
    vals = [float((data.get(l) or {}).get(key) or 0.0) for l in _MIG_LABELS]
    return vals if any(vals) else None


def risk_change_charts(hist: dict | None, pool_name: str | None = None) -> list[ChartSpec]:
    """DQ pie + CO bar for a Risk Change tab, from the migration-status dicts."""
    _imp = (hist or {}).get("impaired", {}) or {}
    if pool_name:
        pl = pool_name.strip().lower()
        dq = next((v for k, v in (_imp.get("dq_by_pool") or {}).items()
                   if k.strip().lower() == pl), {})
        co = next((v for k, v in (_imp.get("co_by_pool") or {}).items()
                   if k.strip().lower() == pl), {})
    else:
        dq = _imp.get("dq_by_status") or {}
        co = _imp.get("co_by_status") or {}
    specs: list[ChartSpec] = []
    dq_vals = _mig_status_series(dq)
    if dq_vals is not None:
        specs.append(ChartSpec(
            kind="pie", title="Delinquency by Credit Grade Migration",
            categories=list(_MIG_LABELS),
            series=[{"name": "DQ", "values": dq_vals, "colors": list(_MIG_COLORS)}],
            value_format="pct"))
    co_vals = _mig_status_series(co)
    if co_vals is not None:
        specs.append(ChartSpec(
            kind="bar_h", title="Charge off by Credit Grade Migration",
            categories=list(_MIG_LABELS),
            series=[{"name": "CO", "values": co_vals, "colors": list(_MIG_COLORS)}],
            value_format="pct"))
    return specs


def _ncc_totals(df: Any, grades: Any, config: dict) -> tuple[float, float, float, float]:
    """(improved, deteriorated, unchanged, total) balances from the matrix,
    using the same migration-state rule as the Risk Change grid."""
    import report_vizo as _rv
    from cecl_engine import risk_change_matrix

    no_score = (config or {}).get("no_score_label", "Not Reported")
    n_top = int((config or {}).get("top_grades_double_drop", 3))
    gl = [g for g in _rv._all_grades(grades, no_score) if not _rv._is_hidden(g)]
    matrix = risk_change_matrix(df, grades, no_score)
    total = float(df["current_balance"].sum())
    imp = det = 0.0
    for i, g in enumerate(gl):
        for j, og in enumerate(gl):
            v = float(_rv._matrix_val(matrix, g, og) or 0.0)
            st = _cell_state(i, j, g, og, no_score, n_top)
            if st == "improved":
                imp += v
            elif st == "deteriorated":
                det += v
    return imp, det, max(0.0, total - imp - det), total


def risk_change_ncc_chart(df: Any, grades: Any, config: dict) -> list[ChartSpec]:
    """Net Credit Change doughnut (Improved / Deteriorated / Unchanged)."""
    imp, det, unc, total = _ncc_totals(df, grades, config)
    if total <= 0:
        return []
    return [ChartSpec(
        kind="doughnut", title="Net Credit Change",
        categories=["Improved", "Deteriorated", "Unchanged"],
        series=[{"name": "NCC",
                 "values": [imp / total, det / total, unc / total],
                 "colors": list(_NCC_COLORS)}],
        value_format="pct")]


def impr_deter_charts(df: Any, grades: Any, config: dict,
                      hist: dict | None = None) -> list[ChartSpec]:
    """The four Impr Deter charts: Improved/Deteriorated by grade (%), the
    Improved/Deteriorated diverging bar by pool, and Net Change by pool."""
    import report_vizo as _rv
    from cecl_engine import risk_change_matrix

    no_score = (config or {}).get("no_score_label", "Not Reported")
    _imp = (hist or {}).get("impaired", {}) or {}
    gl = [g for g in _rv._all_grades(grades, no_score) if not _rv._is_hidden(g)]
    chart_grades = [g for g in gl if g != no_score]

    # Grade-level improved/deteriorated: WARM Executive Summary (3) if present,
    # else derived from the migration matrix (same rule as the Risk Change grid).
    es3 = _imp.get("exec_summary_3") or {}
    if es3:
        grade_imp = {g: float((es3.get("improved") or {}).get(g, 0) or 0)
                     for g in chart_grades}
        grade_det = {g: float((es3.get("deteriorated") or {}).get(g, 0) or 0)
                     for g in chart_grades}
    else:
        matrix = risk_change_matrix(df, grades, no_score)
        n_top = int((config or {}).get("top_grades_double_drop", 3))
        grade_imp = {g: 0.0 for g in chart_grades}
        grade_det = {g: 0.0 for g in chart_grades}
        for j, og in enumerate(gl):
            for i, cg in enumerate(gl):
                v = float(_rv._matrix_val(matrix, cg, og) or 0.0)
                if i > j:
                    if not (j < n_top and (i - j) < 2) and og in grade_det:
                        grade_det[og] += v
                elif i < j and og in grade_imp:
                    grade_imp[og] += v
    imp_tot = sum(grade_imp.values()) or 0.0
    det_tot = sum(grade_det.values()) or 0.0
    imp_pct_g = [(grade_imp[g] / imp_tot if imp_tot else 0.0) for g in chart_grades]
    det_pct_g = [(grade_det[g] / det_tot if det_tot else 0.0) for g in chart_grades]

    # Pool-level improved / deteriorated (negative) / net, risk-rated pools only.
    _rr = _imp.get("risk_rated", {}) or {}
    names, p_imp, p_det, p_net = [], [], [], []
    for pool in _rv._ordered_pools(df, hist):
        pdf = df[df["loan_pool"] == pool]
        if _rr.get(pool, True):
            ip, dp, npct = _rv._ncc(pdf, grades, config)
        else:
            ip, dp, npct = 0.0, 0.0, 0.0
        names.append(pool)
        p_imp.append(float(ip)); p_det.append(-float(dp)); p_net.append(float(npct))

    specs: list[ChartSpec] = []
    if chart_grades:
        specs.append(ChartSpec(
            kind="column", title="Improved Loans", categories=chart_grades,
            series=[{"name": "Improved", "values": imp_pct_g, "colors": ["0D4D5E"]}],
            value_format="pct"))
        specs.append(ChartSpec(
            kind="column", title="Deteriorated Loans", categories=chart_grades,
            series=[{"name": "Deteriorated", "values": det_pct_g, "colors": ["873A3A"]}],
            value_format="pct"))
    if names:
        specs.append(ChartSpec(
            kind="diverging_bar", title="Improved / Deteriorated Loans",
            categories=names,
            series=[{"name": "Improved", "values": p_imp, "colors": ["0D4D5E"]},
                    {"name": "Deteriorated", "values": p_det, "colors": ["873A3A"]}],
            value_format="pct"))
        specs.append(ChartSpec(
            kind="bar_h", title="Net Change", categories=names,
            series=[{"name": "Net", "values": p_net, "colors": ["829901"]}],
            value_format="pct"))
    return specs


def build_acl_summary(client_name: str, snapshot_date: str, config: dict,
                      hist: dict | None = None, df: Any = None,
                      grades: Any = None) -> TablePage | None:
    """ACL Summary: one line per pool (the pool Total rows), pooled totals, and
    the impaired/adjustment lines -- a view of the ACL Env data."""
    import report_vizo as _rv

    acl_pools, acl_summary, acl_impaired = _acl_data(
        config, snapshot_date, hist, df, grades)
    if not acl_pools:
        return None
    cu = (config or {}).get("credit_union") or client_name

    cols = ["Portfolio Segment", "Balance", "Specific Identification",
            "Loan Loss Calc Balance", "Allowance before Env Factor",
            "Env Factor", "Env Factor Allowance", "Total Allowance"]
    rows: list[list[TableCell]] = []
    for pool, pdata in acl_pools.items():
        t = pdata.get("total") or {}
        bal, spec = t.get("balance"), t.get("spec_id")
        calc = t.get("calc_bal")
        if calc is None and bal is not None:
            calc = (bal or 0) - (spec or 0)
        rows.append([
            TableCell(pool, "text", align="left"),
            TableCell(bal, "currency"), TableCell(spec, "currency2"),
            TableCell(calc, "currency"), TableCell(t.get("allow_before"), "currency"),
            TableCell(t.get("env_factor"), "pct"),
            TableCell(t.get("env_allow"), "currency"),
            TableCell(t.get("total_allow"), "currency"),
        ])
    pb, ps = acl_summary.get("pooled_balance"), acl_summary.get("pooled_spec_id")
    rows.append([
        TableCell("Pooled Totals", "text", bold=True, align="left"),
        TableCell(pb, "currency", bold=True), TableCell(ps, "currency2", bold=True),
        TableCell(((pb or 0) - (ps or 0)) if pb is not None else None, "currency", bold=True),
        TableCell(acl_summary.get("pooled_allow_before"), "currency", bold=True),
        TableCell(None), TableCell(acl_summary.get("pooled_env_allow"), "currency", bold=True),
        TableCell(acl_summary.get("pooled_total_allow"), "currency", bold=True),
    ])
    sections = [TableSection(columns=cols, rows=rows)]

    adj_rows: list[list[TableCell]] = []
    for k, v in (acl_impaired or {}).items():
        adj_rows.append([TableCell(k, "text", align="left"), TableCell(v, "currency")])
    for lbl, key in _ACL_ADJ_KEYS:
        if key in acl_summary:
            adj_rows.append([TableCell(lbl, "text", bold=True, align="left"),
                             TableCell(acl_summary.get(key), "currency", bold=True)])
    if adj_rows:
        sections.append(TableSection(
            title="Impaired Loans & Adjustment", rows=adj_rows))

    return TablePage(
        credit_union=cu,
        title="Allowance for Credit Loss - Summary by Pool",
        heading_lines=[f"For Quarter Ending {_rv._snap_display(snapshot_date)}"],
        sections=sections)


def build_mgmt_adj_summary(client_name: str, snapshot_date: str, config: dict,
                           hist: dict | None = None, df: Any = None,
                           grades: Any = None) -> TablePage | None:
    """Mgmt Adj Summary: per-pool the grades that carry a management adjustment
    plus the pool's environmental factor -- from the ACL Env data."""
    import report_vizo as _rv

    acl_pools, acl_summary, _ = _acl_data(config, snapshot_date, hist, df, grades)
    if not acl_pools:
        return None
    cu = (config or {}).get("credit_union") or client_name
    heading = [f"For Quarter Ending {_rv._snap_display(snapshot_date)}"]
    title = "Management & Environmental Adjustments"

    def blank() -> TableCell:
        return TableCell(None)

    cols = ["Portfolio Segment", "Grade", "Balance", "ACL Base Loss Rate",
            "Mgmt Adj", "Allowance Factor", "Allowance before Env Factor",
            "Env Factor", "Env Factor Allowance"]
    rows: list[list[TableCell]] = []
    any_adj = False
    for pool, pdata in acl_pools.items():
        grades_d = pdata.get("grades") or {}
        total = pdata.get("total") or {}
        adj_grades = [(g, gv) for g, gv in grades_d.items()
                      if (gv.get("mgmt_adj") or 0)]
        env_factor = total.get("env_factor") or 0
        if not adj_grades and not env_factor:
            continue
        any_adj = True
        rows.append([TableCell(pool, "text", bold=True, align="left")]
                    + [blank() for _ in range(8)])
        for g, gv in adj_grades:
            rows.append([
                blank(), TableCell(g, "text", align="left"),
                TableCell(gv.get("balance"), "currency"),
                TableCell(gv.get("base_rate"), "pct4"),
                TableCell(gv.get("mgmt_adj"), "pct4"),
                TableCell(gv.get("factor"), "pct4"),
                TableCell(gv.get("allow_before"), "currency"), blank(), blank()])
        rows.append([
            blank(), TableCell("Total", "text", bold=True, align="left"),
            TableCell(total.get("balance"), "currency", bold=True),
            blank(), blank(), blank(),
            TableCell(total.get("allow_before"), "currency", bold=True),
            TableCell(total.get("env_factor"), "pct", bold=True),
            TableCell(total.get("env_allow"), "currency", bold=True)])

    if not any_adj:
        return TablePage(
            credit_union=cu, title=title, heading_lines=heading,
            sections=[TableSection(rows=[[TableCell(
                "No management or environmental adjustments were applied "
                "this period.", "text", align="left")]])])

    rows.append([
        TableCell("Pooled Totals", "text", bold=True, align="left"), blank(),
        TableCell(acl_summary.get("pooled_balance"), "currency", bold=True),
        blank(), blank(), blank(),
        TableCell(acl_summary.get("pooled_allow_before"), "currency", bold=True),
        blank(),
        TableCell(acl_summary.get("pooled_env_allow"), "currency", bold=True)])
    return TablePage(credit_union=cu, title=title, heading_lines=heading,
                     sections=[TableSection(columns=cols, rows=rows)])


def build_impaired_loans(client_name: str, snapshot_date: str, config: dict,
                         hist: dict | None = None, df: Any = None,
                         grades: Any = None) -> TablePage | None:
    """Impaired Loans - ASC 310-10: the specifically identified allowance by
    impairment category, from the ACL Env impaired data."""
    import report_vizo as _rv

    acl_pools, acl_summary, acl_impaired = _acl_data(
        config, snapshot_date, hist, df, grades)
    if not acl_impaired:
        return None
    cu = (config or {}).get("credit_union") or client_name
    rows = [[TableCell(k, "text", align="left"), TableCell(v, "currency")]
            for k, v in acl_impaired.items()]
    if "total_spec_allow" in acl_summary:
        rows.append([
            TableCell("Total Specifically Identified Allowance", "text",
                      bold=True, align="left"),
            TableCell(acl_summary.get("total_spec_allow"), "currency", bold=True)])
    return TablePage(
        credit_union=cu, title="Impaired Loans - ASC 310-10",
        heading_lines=[f"For Quarter Ending {_rv._snap_display(snapshot_date)}"],
        sections=[TableSection(columns=["Impairment Category", "Allowance"],
                               rows=rows)])


def build_report_model(client_name: str, snapshot_date: str, config: dict,
                       grades: Any = None, hist: dict | None = None,
                       df: Any = None, *, supplemental: bool = False) -> dict:
    """Build the render-ready page set from report data.

    Returns ``{"cover": CoverPage, "pages": [ (template, ctx, landscape), ... ]}``.
    Archetypes are added here incrementally (cover, Risk Change so far), each
    reusing the report engine's own pure compute functions.
    """
    cover = build_cover(client_name, snapshot_date, config,
                        supplemental=supplemental)
    pages: list[tuple[str, dict, bool]] = [
        ("cover.html", {"cover": cover}, False),
    ]
    if df is not None and not supplemental:
        rc = build_risk_change(client_name, snapshot_date, df, config, grades, hist)
        rc_charts = render_chart_specs(
            risk_change_ncc_chart(df, grades, config) + risk_change_charts(hist))
        pages.append(("risk_change.html", {"page": rc, "charts": rc_charts}, True))
    if not supplemental:
        impd = build_impr_deter(client_name, snapshot_date, config, hist,
                               df=df, grades=grades)
        impd_charts = (render_chart_specs(impr_deter_charts(df, grades, config, hist))
                       if df is not None else [])
        pages.append(("impr_deter.html", {"page": impd, "charts": impd_charts}, False))
        acl = build_acl_env(client_name, snapshot_date, config, hist,
                           df=df, grades=grades)
        if acl is not None:
            pages.append(("acl_env.html", {"page": acl, "charts": []}, True))
        acl_sum = build_acl_summary(client_name, snapshot_date, config, hist,
                                    df=df, grades=grades)
        if acl_sum is not None:
            pages.append(("table_page.html", {"page": acl_sum}, True))
        mgmt = build_mgmt_adj_summary(client_name, snapshot_date, config, hist,
                                      df=df, grades=grades)
        if mgmt is not None:
            pages.append(("table_page.html", {"page": mgmt}, True))
        impaired = build_impaired_loans(client_name, snapshot_date, config, hist,
                                        df=df, grades=grades)
        if impaired is not None:
            pages.append(("table_page.html", {"page": impaired}, False))
    return {"cover": cover, "pages": pages}
