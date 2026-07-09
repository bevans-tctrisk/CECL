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

