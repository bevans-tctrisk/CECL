"""Solr fetch helpers for the SCALE model.

Lifted from ``fill_5300_from_solr_to_excel.py`` and made importable so
the wizard can call them directly without spawning a subprocess.
"""
from __future__ import annotations

import re
from typing import Any, Dict, Tuple

import requests


VALID_Q_MONTHS = (3, 6, 9, 12)


def end_of_quarter(ym: str) -> Tuple[int, int, int]:
    """Parse ``YYYY-MM`` -> ``(month, day, year)`` for that quarter end."""
    try:
        y_str, m_str = ym.split("-")
        y, m = int(y_str), int(m_str)
    except Exception as exc:  # noqa: BLE001
        raise ValueError("Invalid period. Use YYYY-MM (e.g., 2025-12).") from exc
    if m not in VALID_Q_MONTHS:
        raise ValueError("Period month must be a quarter end: 03, 06, 09, 12.")
    return {3: (3, 31, y), 6: (6, 30, y), 9: (9, 30, y), 12: (12, 31, y)}[m]


def quarter_date_string(charter: int, ym: str) -> str:
    m, d, y = end_of_quarter(ym)
    return f"{charter}-{m:02d}/{d:02d}/{y}"


def coerce_numeric(value: Any) -> Any:
    """Normalize Solr string values into numbers when possible.

    Mirrors the legacy script's behavior: handles currency markers,
    parentheses for negatives, percentages, and thousands separators.
    """
    if value is None:
        return None
    if isinstance(value, (int, float)):
        return value
    s = str(value).strip()
    if s == "":
        return None
    if s.endswith("%"):
        neg = False
        core = s[:-1].strip()
        if core.startswith("(") and core.endswith(")"):
            neg = True
            core = core[1:-1].strip()
        if core.startswith("+"):
            core = core[1:].strip()
        if core.startswith("-"):
            neg = True
            core = core[1:].strip()
        core = core.replace(",", "")
        try:
            num = float(core) / 100.0
            return -num if neg else num
        except ValueError:
            return value
    s_clean = re.sub(r"[$\s]", "", s)
    neg = False
    if s_clean.startswith("(") and s_clean.endswith(")"):
        neg = True
        s_clean = s_clean[1:-1]
    s_clean = s_clean.replace(",", "")
    if s_clean.startswith("+"):
        s_clean = s_clean[1:]
    elif s_clean.startswith("-"):
        neg = not neg
        s_clean = s_clean[1:]
    try:
        num = float(s_clean)
        return -num if neg else num
    except ValueError:
        return value


def fetch_doc(
    solr_url: str,
    core: str,
    charter: int,
    period: str,
    *,
    charter_field: str = "charter",
    charterdate_field: str = "charterdate",
    username: str | None = None,
    password: str | None = None,
    timeout: int = 20,
) -> Dict[str, Any]:
    """Fetch a single 5300 doc for ``charter`` at quarter-end ``period``.

    Raises ``LookupError`` when no doc matches.
    """
    target = quarter_date_string(charter, period)
    base = solr_url.rstrip("/")
    url = f"{base}/{core}/select"
    q = f'{charter_field}:{charter} AND {charterdate_field}:"{target}"'
    auth = (username, password) if username and password else None
    r = requests.get(
        url,
        params={"q": q, "rows": 1, "wt": "json"},
        auth=auth,
        timeout=timeout,
    )
    r.raise_for_status()
    docs = r.json().get("response", {}).get("docs", [])
    if not docs:
        raise LookupError(f"No Solr doc found for q={q}")
    return docs[0]


def test_connection(
    solr_url: str,
    core: str,
    *,
    timeout: int = 5,
) -> Dict[str, Any]:
    """Ping the Solr core. Returns ``{ok, status, message}``."""
    base = solr_url.rstrip("/")
    url = f"{base}/{core}/admin/ping"
    try:
        r = requests.get(url, params={"wt": "json"}, timeout=timeout)
        r.raise_for_status()
        body = r.json()
        return {
            "ok": body.get("status") == "OK",
            "status": body.get("status", ""),
            "message": "",
        }
    except Exception as exc:  # noqa: BLE001
        return {"ok": False, "status": "", "message": str(exc)}


_CHARTERDATE_RE = re.compile(
    r"^\d+-(?P<m>\d{2})/(?P<d>\d{2})/(?P<y>\d{4})$"
)


def _charterdate_to_period(value: str) -> str | None:
    """Convert a Solr ``charterdate`` field value (``"<n>-MM/DD/YYYY"``)
    into a ``"YYYY-MM"`` period string. Returns None when it doesn't
    parse or the month isn't a quarter end.
    """
    if not isinstance(value, str):
        return None
    m = _CHARTERDATE_RE.match(value.strip())
    if not m:
        return None
    month = int(m["m"])
    if month not in VALID_Q_MONTHS:
        return None
    return f"{m['y']}-{month:02d}"


def list_charter_periods(
    solr_url: str,
    core: str,
    charter: int,
    *,
    charter_field: str = "charter",
    charterdate_field: str = "charterdate",
    username: str | None = None,
    password: str | None = None,
    timeout: int = 10,
    max_rows: int = 500,
) -> Dict[str, Any]:
    """Return the set of ``YYYY-MM`` quarter-end periods for which Solr
    has a 5300 doc for ``charter``.

    Returns ``{"ok": bool, "periods": set[str], "error": str}``. On
    network/auth errors ``ok=False`` and ``periods`` is empty so the
    caller can fall back to its unfiltered period list.
    """
    result: Dict[str, Any] = {"ok": False, "periods": set(), "error": ""}
    base = solr_url.rstrip("/")
    url = f"{base}/{core}/select"
    auth = (username, password) if username and password else None
    try:
        r = requests.get(
            url,
            params={
                "q": f"{charter_field}:{charter}",
                "fl": charterdate_field,
                "rows": max_rows,
                "wt": "json",
            },
            auth=auth,
            timeout=timeout,
        )
        r.raise_for_status()
        docs = r.json().get("response", {}).get("docs", [])
    except Exception as exc:  # noqa: BLE001
        result["error"] = str(exc)
        return result

    periods: set[str] = set()
    for doc in docs:
        raw = doc.get(charterdate_field)
        # Solr may return multi-valued fields as a list.
        values = raw if isinstance(raw, list) else [raw]
        for v in values:
            p = _charterdate_to_period(v)
            if p:
                periods.add(p)
    result["ok"] = True
    result["periods"] = periods
    return result

