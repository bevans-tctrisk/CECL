"""NCUA 5300 ACL (Allowance for Credit Losses on Loans) fetcher.

Companion to ``solr_5300_backfill`` and friends. Used by the report
runtime as a *fallback* source when the credit union's monthly-balance
file does not include an ACL row for a given quarter end (or includes
$0). Pre-CECL adopters report the value in ``A719`` ("ALLL"); CECL
adopters use ``A718A3`` (ACL on Loans) and ``A718A5`` (Total ACL =
loans + other assets). This module probes all three and returns the
first non-zero match so it works for both reporting regimes.

Public API:
    ``fetch_acl_for_period(charter, period_iso, *, solr_url, core)``
    -> ``{"value": float, "field": "A718A3"|"A718A5"|"A719", ...}`` or
    ``None`` when no doc / all candidate fields are zero/missing.

This is read-only and idempotent. Callers should treat any transport
error as a soft miss (the report engine continues with whatever
non-Solr sources it already has).
"""
from __future__ import annotations

from datetime import date
from typing import Any

from .solr_5300_backfill import (  # type: ignore[import-not-found]
    _coerce_number,
    fetch_solr_doc,
)

# Default Solr coordinates — match the rest of the 5300 backfill suite.
DEFAULT_SOLR_URL = "http://searchserver1.tctrisk.com:8983/solr"
DEFAULT_CORE = "ncua"

# Field probe order:
#   1. A718A3 — ACL on Loans (CECL, post-2023). Matches the 5300 line
#      labeled "Allowance for Credit Losses on Loans".
#   2. A718A5 — Total ACL (loans + other assets). For CECL adopters
#      whose A718A3 is rolled into A718A5 only.
#   3. A719   — Legacy ALLL (pre-CECL adopters / state-charter CUs
#      that still file under the old form).
# First non-zero positive match wins.
_ACL_FIELD_ORDER: tuple[str, ...] = ("A718A3", "A718A5", "A719")


def _quarter_end_iso(period_iso: str) -> str | None:
    """Snap an arbitrary YYYY-MM-DD to the *containing* quarter end.

    The Solr ``ncua`` core only carries quarter-end docs (3/31, 6/30,
    9/30, 12/31), so callers may pass a month-end string and we will
    map to the matching quarter. Returns None on parse failure.
    """
    try:
        y, m, _d = period_iso.split("-")
        yi = int(y)
        mi = int(m)
    except (ValueError, AttributeError):
        return None
    if mi <= 3:
        return f"{yi:04d}-03-31"
    if mi <= 6:
        return f"{yi:04d}-06-30"
    if mi <= 9:
        return f"{yi:04d}-09-30"
    if mi <= 12:
        return f"{yi:04d}-12-31"
    return None


def fetch_acl_for_period(
    charter: int | str,
    period_iso: str,
    *,
    solr_url: str = DEFAULT_SOLR_URL,
    core: str = DEFAULT_CORE,
    username: str | None = None,
    password: str | None = None,
    timeout: int = 20,
    fields: tuple[str, ...] = _ACL_FIELD_ORDER,
) -> dict[str, Any] | None:
    """Return the ACL on Loans value from NCUA 5300 for the period.

    ``period_iso`` may be any month-end inside the target quarter; it
    is snapped to the quarter-end before querying.

    Returns ``{"value": float, "field": <field>, "period": <qe_iso>,
    "charter": int}`` when a non-zero positive value is found,
    otherwise ``None``.

    Network/transport errors propagate; callers wanting "soft miss"
    semantics should wrap in try/except.
    """
    try:
        charter_int = int(charter)
    except (TypeError, ValueError):
        return None
    qe_iso = _quarter_end_iso(period_iso)
    if not qe_iso:
        return None
    doc = fetch_solr_doc(
        solr_url,
        core,
        charter_int,
        qe_iso,
        username=username,
        password=password,
        timeout=timeout,
    )
    if not doc:
        return None
    for field in fields:
        raw = doc.get(field)
        if raw is None:
            continue
        v = abs(_coerce_number(raw))
        if v > 0:
            return {
                "value": v,
                "field": field,
                "period": qe_iso,
                "charter": charter_int,
            }
    return None


def fetch_acl_history(
    charter: int | str,
    periods_iso: list[str],
    *,
    solr_url: str = DEFAULT_SOLR_URL,
    core: str = DEFAULT_CORE,
    username: str | None = None,
    password: str | None = None,
    timeout: int = 20,
) -> dict[str, dict[str, Any]]:
    """Bulk variant — returns ``{quarter_end_iso: result_dict}``.

    Quarter-ends are deduped after snapping. Misses (None) are NOT
    included in the output. Transport errors during one period skip
    that period and continue.
    """
    out: dict[str, dict[str, Any]] = {}
    seen: set[str] = set()
    for p in periods_iso:
        qe = _quarter_end_iso(p)
        if not qe or qe in seen:
            continue
        seen.add(qe)
        try:
            res = fetch_acl_for_period(
                charter,
                qe,
                solr_url=solr_url,
                core=core,
                username=username,
                password=password,
                timeout=timeout,
            )
        except Exception:  # noqa: BLE001
            continue
        if res is not None:
            out[qe] = res
    return out
