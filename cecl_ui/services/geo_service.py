"""State + county lookup service for the new-CU wizard.

* States come from a static dict (the same 50 states + DC that
  ``fetch_econ_data.STATE_INFO`` already supports).
* Counties are fetched on demand from the Census Bureau public API and cached
  to a local JSON file (one per state FIPS) so subsequent lookups are
  offline.

Network access is only required the very first time a particular state's
counties are requested.  Cache files live under
``cecl_ui/data/counties_cache/<fips>.json``.
"""
from __future__ import annotations

import json
from pathlib import Path

# Re-export the state list maintained in fetch_econ_data so we have one
# source of truth.
from fetch_econ_data import STATE_INFO  # type: ignore

CACHE_DIR = Path(__file__).resolve().parent.parent / "data" / "counties_cache"

_TIMEOUT = 15


def states() -> list[dict[str, str]]:
    """Return [{'name': 'New York', 'fips': '36', 'abbr': 'NY'}, ...] sorted."""
    out = [
        {"name": name, "fips": fips, "abbr": abbr}
        for name, (fips, abbr) in STATE_INFO.items()
    ]
    out.sort(key=lambda s: s["name"])
    return out


def fips_for_state(state_name: str) -> str | None:
    s = (state_name or "").strip().lower()
    for name, (fips, _abbr) in STATE_INFO.items():
        if name.lower() == s:
            return fips
    return None


def counties_for_state(state_name: str) -> list[str]:
    """Return county names for ``state_name``.  Cached to disk after first fetch.

    Returns ``[]`` if the state is unknown or the network fetch fails and no
    cache exists yet.
    """
    fips = fips_for_state(state_name)
    if not fips:
        return []

    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    cache_file = CACHE_DIR / f"{fips}.json"
    if cache_file.exists():
        try:
            return json.loads(cache_file.read_text(encoding="utf-8"))
        except Exception:
            pass  # fall through and refetch

    counties = _fetch_from_census(fips)
    if counties:
        try:
            cache_file.write_text(
                json.dumps(counties, indent=0), encoding="utf-8"
            )
        except Exception:
            pass
    return counties


def _fetch_from_census(state_fips: str) -> list[str]:
    """Pull county names for ``state_fips`` from the Census Bureau.

    Uses the public, keyless 2020 county gazetteer file
    (``national_county2020.txt``).  The previous JSON Data API endpoint
    started requiring an API key, which would silently return zero rows
    for any uncached state.  The gazetteer is a small pipe-delimited text
    file with one row per county nationwide, so a single fetch primes any
    state we ask about.
    """
    import requests  # deferred -- heavy import on slow drives

    url = (
        "https://www2.census.gov/geo/docs/reference/codes2020/"
        "national_county2020.txt"
    )
    try:
        resp = requests.get(url, timeout=_TIMEOUT)
        resp.raise_for_status()
        text = resp.text
    except Exception:
        return []

    out: list[str] = []
    for line in text.splitlines():
        # Skip header and blank lines.
        if not line or line.startswith("STATE|"):
            continue
        parts = line.split("|")
        # Expected layout: STATE|STATEFP|COUNTYFP|COUNTYNS|COUNTYNAME|...
        if len(parts) < 5:
            continue
        if parts[1] != state_fips:
            continue
        name = (parts[4] or "").strip()
        if name:
            out.append(name)
    out.sort(key=str.lower)
    return out
