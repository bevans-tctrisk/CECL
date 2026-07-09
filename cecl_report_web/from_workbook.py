"""Adapter: read a generated report workbook into the ReportModel.

Phase-0 data source. Populating the model from the already-generated
xlsx guarantees the PDF numbers match the Excel report exactly, and it
doubles as the "extract a page" capability. Later phases can add a
compute-time populator with the same model as output — templates don't
change.

Scanning is label-anchored (find the row by its text, not a fixed
coordinate) so it survives minor row-position drift between report
versions.
"""

from __future__ import annotations

from pathlib import Path

from openpyxl import load_workbook

from .model import (
    GradeMigrationRow,
    ImprDeterPage,
    KeyValueRow,
    PoolMigrationRow,
)

_IMPR_DETER_TAB_HINTS = ("impr deter", "improved", "deteriorated")
_CECL_LABELS = (
    "Total Specifically Identified Allowance",
    "Total Allowance Needed",
    "Allowance for Credit Loss Balance",
    "Adjustment (Overfunded)",
)


def _find_tab(wb, hints: tuple[str, ...]) -> str | None:
    for sn in wb.sheetnames:
        low = sn.strip().lower()
        if any(h in low for h in hints):
            return sn
    return None


def _num(v):
    if v is None:
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def _cell_str(v) -> str:
    return "" if v is None else str(v).strip()


def load_impr_deter(report_path: str | Path) -> ImprDeterPage:
    """Read the 'Impr Deter' tab of a Vizo Model workbook into the model."""
    wb = load_workbook(report_path, data_only=True)
    tab = _find_tab(wb, _IMPR_DETER_TAB_HINTS)
    if tab is None:
        raise ValueError(
            f"No Improved/Deteriorated tab found in {Path(report_path).name} "
            f"(sheets: {wb.sheetnames})"
        )
    ws = wb[tab]

    # ── Header lines: first 4 populated rows of column A ──
    heading: list[str] = []
    cu = ""
    period = ""
    for r in range(1, 6):
        v = ws.cell(r, 1).value
        if v is None or v == "":
            continue
        if r == 1:
            cu = _cell_str(v)
        elif hasattr(v, "strftime"):
            period = v.strftime("%m-%d-%y")
            heading.append(period)
        else:
            heading.append(_cell_str(v))

    # ── CECL Adjustment box: label in a left column, value ~3 cols right ──
    cecl: list[KeyValueRow] = []
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            label = _cell_str(ws.cell(r, c).value)
            if not label:
                continue
            for want in _CECL_LABELS:
                if label.startswith(want) or want.split(" as of")[0] in label:
                    # value: first numeric cell to the right on this row
                    val = None
                    for cc in range(c + 1, ws.max_column + 1):
                        n = _num(ws.cell(r, cc).value)
                        if n is not None:
                            val = n
                            break
                    cecl.append(KeyValueRow(label=label, value=val))
                    break

    # ── Improved/Deteriorated by pool: anchor on a "Loan Type" header ──
    by_pool: list[PoolMigrationRow] = []
    anchor = _find_label(ws, "Loan Type")
    if anchor:
        hr, hc = anchor
        r = hr + 1
        while r <= ws.max_row:
            pool = _cell_str(ws.cell(r, hc).value)
            if not pool:
                break
            by_pool.append(PoolMigrationRow(
                pool=pool,
                improved=_num(ws.cell(r, hc + 1).value),
                deteriorated=_num(ws.cell(r, hc + 2).value),
                net_change=_num(ws.cell(r, hc + 3).value),
            ))
            r += 1

    # ── Improved/Deteriorated by grade: anchor on a "Grade" header ──
    by_grade: list[GradeMigrationRow] = []
    ganchor = _find_label(ws, "Grade")
    if ganchor:
        hr, hc = ganchor
        r = hr + 1
        while r <= ws.max_row:
            grade = _cell_str(ws.cell(r, hc).value)
            if not grade:
                break
            by_grade.append(GradeMigrationRow(
                grade=grade,
                balance=_num(ws.cell(r, hc + 1).value),
                improved=_num(ws.cell(r, hc + 3).value),
                deteriorated=_num(ws.cell(r, hc + 4).value),
            ))
            r += 1

    return ImprDeterPage(
        credit_union=cu,
        period_ending=period,
        heading_lines=heading,
        cecl_adjustment=cecl,
        by_pool=by_pool,
        by_grade=by_grade,
    )


def _find_label(ws, text: str) -> tuple[int, int] | None:
    """Return (row, col) of the first cell whose stripped value == text."""
    low = text.strip().lower()
    for r in range(1, ws.max_row + 1):
        for c in range(1, ws.max_column + 1):
            if _cell_str(ws.cell(r, c).value).lower() == low:
                return (r, c)
    return None
