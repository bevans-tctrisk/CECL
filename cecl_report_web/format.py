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
