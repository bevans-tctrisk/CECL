"""Value formatters shared by templates.

Mirrors the Excel number formats used in the reports so the rendered
HTML matches cell-for-cell. Kept tiny and dependency-free.
"""

from __future__ import annotations


def acct0(v: float | int | None) -> str:
    """Accounting style, 0 decimals: 13157 -> '13,157', -237509 -> '(237,509)',
    0/None -> '-'  (matches ``_(* #,##0_);_(* (#,##0);_(* "-"??_)``).
    """
    if v is None:
        return "-"
    try:
        n = float(v)
    except (TypeError, ValueError):
        return str(v)
    if round(n) == 0:
        return "-"
    if n < 0:
        return f"({abs(n):,.0f})"
    return f"{n:,.0f}"


def acct2(v: float | int | None) -> str:
    """Accounting style, 2 decimals (matches ``_(* #,##0.00...``)."""
    if v is None:
        return "-"
    try:
        n = float(v)
    except (TypeError, ValueError):
        return str(v)
    if round(n, 2) == 0:
        return "-"
    if n < 0:
        return f"({abs(n):,.2f})"
    return f"{n:,.2f}"


def pct0(v: float | int | None) -> str:
    """Whole-percent: 0.0 -> '0%', 0.123 -> '12%' (Excel fraction input)."""
    if v is None:
        return ""
    try:
        return f"{float(v) * 100:.0f}%"
    except (TypeError, ValueError):
        return str(v)


def pct1(v: float | int | None) -> str:
    """One-decimal percent: 0.0 -> '0.0%'."""
    if v is None:
        return ""
    try:
        return f"{float(v) * 100:.1f}%"
    except (TypeError, ValueError):
        return str(v)


def pct2(v: float | int | None) -> str:
    """Two-decimal percent: 0.0 -> '0.00%' (matches Excel 0.00%)."""
    if v is None:
        return ""
    try:
        return f"{float(v) * 100:.2f}%"
    except (TypeError, ValueError):
        return str(v)


def pct4(v: float | int | None) -> str:
    """Four-decimal percent: 0.0011 -> '0.1100%' (matches Excel 0.0000%)."""
    if v is None:
        return ""
    try:
        return f"{float(v) * 100:.4f}%"
    except (TypeError, ValueError):
        return str(v)


def mcell(cell) -> str:
    """Render a MatrixCell's value using its own number format."""
    if cell is None:
        return ""
    return pct2(cell.value) if getattr(cell, "is_pct", False) else acct0(cell.value)


def tcell(cell) -> str:
    """Render a TableCell (generic TablePage) using its ``fmt`` field."""
    if cell is None:
        return ""
    v = getattr(cell, "value", None)
    fmt = getattr(cell, "fmt", "text")
    if v is None or v == "":
        return ""
    return {
        "currency": acct0, "currency2": acct2,
        "pct": pct0, "pct1": pct1, "pct2": pct2,
    }.get(fmt, lambda x: str(x))(v)



import re as _re
from datetime import date as _date, datetime as _datetime

_PCT_RE = _re.compile(r"0(?:\.(0+))?%")


def excel_format(value, numfmt: str | None) -> str:
    """Format a cell value the way Excel would, given its number format.

    Powers the generic faithful-grid renderer: text passes through, dates
    honour mm-dd-yy, and numeric cells map to the accounting / dollar /
    percent styles used across the reports. Good enough for the report
    number formats in use — not a full Excel format engine.
    """
    if value is None or value == "":
        return ""
    if isinstance(value, str):
        return value
    fmt = (numfmt or "General").strip()

    if isinstance(value, (_datetime, _date)):
        if "mm" in fmt.lower() and "dd" in fmt.lower():
            yy = "%Y" if "yyyy" in fmt.lower() else "%y"
            return value.strftime(f"%m-%d-{yy}")
        return value.strftime("%m/%d/%Y")

    try:
        n = float(value)
    except (TypeError, ValueError):
        return str(value)

    low = fmt.lower()
    if "%" in fmt:
        m = _PCT_RE.search(fmt)
        decimals = len(m.group(1)) if (m and m.group(1)) else 0
        return f"{n * 100:.{decimals}f}%"
    if "$" in fmt:
        if round(n) == 0:
            return "-"
        return f"(${abs(n):,.0f})" if n < 0 else f"${n:,.0f}"
    if "#,##0.00" in fmt:
        return acct2(n)
    if "#,##0" in fmt:
        return acct0(n)
    if fmt in ("General", ""):
        # integers show without trailing .0
        return f"{int(n):,}" if float(n).is_integer() else str(n)
    return str(value)


