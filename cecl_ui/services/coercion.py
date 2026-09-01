"""Shared numeric coercion for values read out of client spreadsheets.

Why this module exists
----------------------
Four helpers grew up independently doing "turn a spreadsheet cell into a
number", and three of them stripped every non-digit character. That strips the
DECIMAL POINT too, folding a float's fractional part into the integer:

    60.0            -> 600            (a 6-day delinquency booked as 60+)
    313235957340.0  -> 3132359573400  (a 12-digit account becomes 13)

pandas hands back float64 for any numeric column containing a single blank
cell, so this is the normal case rather than an edge case. The damage was
real: Utah Community FCU recorded $86,351,866 of delinquency against a true
$12,108,522, and its WARM accounts could not be joined at all.

Only ``impaired_parser._digits`` got it right, by dropping a trailing ``.0``
before the strip. This module is that fix, once, for everyone.

Use ``coerce_number`` for a quantity, ``coerce_days`` for a whole-day count,
and ``digits_only`` for an identifier being normalised for a join.
"""
from __future__ import annotations

import math
import re
from typing import Any

_NULLISH = {"", "nan", "none", "null", "-", "na", "n/a"}
_NUM_RX = re.compile(r"-?\d+(?:\.\d+)?")


def coerce_number(val: Any) -> float | None:
    """Parse a spreadsheet cell to ``float``. Empty / non-numeric -> ``None``.

    Numeric input is taken as-is, so ``60.0`` stays ``60.0``. Text is parsed
    with currency and thousands separators removed, then by finding the first
    number embedded in it, so ``"30 days"`` and ``"$1,200.50"`` both work.
    """
    if val is None or isinstance(val, bool):
        return None
    if isinstance(val, (int, float)):
        try:
            f = float(val)
        except (TypeError, ValueError):
            return None
        return None if math.isnan(f) else f
    s = str(val).strip()
    if not s or s.lower() in _NULLISH:
        return None
    try:
        return float(s.replace(",", "").replace("$", ""))
    except ValueError:
        pass
    m = _NUM_RX.search(s)
    if not m:
        return None
    try:
        return float(m.group())
    except ValueError:
        return None


def coerce_days(val: Any) -> int | None:
    """Whole days past due. ``60.0`` -> ``60``, never ``600``."""
    f = coerce_number(val)
    if f is None or math.isinf(f):
        return None
    return int(f)


def digits_only(val: Any) -> str:
    """Digit string for an identifier, dropping a trailing Excel ``.0``.

    ``19238``, ``19238.0`` and ``'19238-'`` all collapse to ``'19238'`` so an
    account or member number survives a round trip through a float column.
    """
    if val is None:
        return ""
    s = str(val).strip()
    if s.endswith(".0"):
        s = s[:-2]
    return "".join(ch for ch in s if ch.isdigit())
