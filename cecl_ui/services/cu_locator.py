"""NCUA charter-based State + County auto-detect.

Public API:
    ``lookup_cu_geography(charter, *, solr_url=None, core=None,
        workspace_root=None) -> dict``

Returns::

    {
      "ok": True,
      "charter": 5641,
      "cu_name": "CENSUS",
      "city": "SUITLAND",
      "state_abbr": "MD",
      "state": "Maryland",
      "state_fips": "24",
      "county": "Prince George's County",
      "county_fips": "24033",
      "sources": {
        "state":  "solr:2020-12-31",
        "county": "foicu:2025-09",
      },
      "errors": [],
    }

State is sourced from the Solr 5300 ``ncua`` core's ``custate`` field
(walked back through prior quarter-ends when the most recent quarter
indexes as ``NA``).

County is *not* in the 5300; it lives in NCUA's quarterly FOICU.txt
(inside ``call-report-data-YYYY-MM.zip``). The download is fetched
on cache miss and cached under ``<workspace_root>/.cecl_cache/ncua/``.

This module is read-only. Any transport / parse failure is converted
to a soft miss + a string in ``errors`` so the wizard never blocks
on it.
"""
from __future__ import annotations

import csv
import io
import json
import os
import tempfile
import zipfile
from datetime import date
from pathlib import Path
from typing import Any

# Default Solr coordinates — match the rest of the 5300 backfill suite.
DEFAULT_SOLR_URL = "http://searchserver1.tctrisk.com:8983/solr"
DEFAULT_CORE = "ncua"

# NCUA quarterly call-report data download URL pattern.
# Example: https://ncua.gov/files/publications/analysis/call-report-data-2025-09.zip
_NCUA_ZIP_URL = (
    "https://ncua.gov/files/publications/analysis/call-report-data-{ym}.zip"
)

# Reverse map: 2-letter state abbr -> full name. Built lazily from
# fetch_econ_data.STATE_INFO so we have one source of truth for the
# wizard's state list.
_ABBR_TO_NAME: dict[str, str] | None = None
_ABBR_TO_FIPS: dict[str, str] | None = None


def _state_maps() -> tuple[dict[str, str], dict[str, str]]:
    global _ABBR_TO_NAME, _ABBR_TO_FIPS
    if _ABBR_TO_NAME is None or _ABBR_TO_FIPS is None:
        from fetch_econ_data import STATE_INFO  # type: ignore[import-not-found]

        _ABBR_TO_NAME = {abbr: name for name, (_fips, abbr) in STATE_INFO.items()}
        _ABBR_TO_FIPS = {abbr: fips for _name, (fips, abbr) in STATE_INFO.items()}
    return _ABBR_TO_NAME, _ABBR_TO_FIPS


# ---------------------------------------------------------------- cache dirs

def _cache_dir(workspace_root: str | Path | None) -> Path:
    """Return the directory we cache NCUA downloads under.

    Prefer ``<workspace_root>/.cecl_cache/ncua/`` so team members share
    the cache off Egnyte. Falls back to a process-local TEMP dir if the
    workspace root is unwritable.
    """
    if workspace_root:
        wr = Path(workspace_root)
        try:
            d = wr / ".cecl_cache" / "ncua"
            d.mkdir(parents=True, exist_ok=True)
            # Quick writability probe — touch a marker file.
            probe = d / ".write_probe"
            probe.write_text("ok", encoding="utf-8")
            probe.unlink(missing_ok=True)
            return d
        except OSError:
            pass
    d = Path(tempfile.gettempdir()) / "cecl_ui_ncua"
    d.mkdir(parents=True, exist_ok=True)
    return d


# ---------------------------------------------------------------- state via Solr

def _walk_back_quarter_ends(start_iso: str | None = None, max_quarters: int = 8) -> list[str]:
    """Yield ``YYYY-MM-DD`` quarter-end ISO dates newest-first."""
    if start_iso:
        try:
            y, m, _ = start_iso.split("-")
            y, m = int(y), int(m)
        except ValueError:
            y, m = date.today().year, date.today().month
    else:
        today = date.today()
        y, m = today.year, today.month
    # Snap to most recent quarter end <= today.
    qm = ((m - 1) // 3) * 3 + 3  # 3, 6, 9, 12
    if qm > m or (qm == m and date.today().day < 1):
        # Roll back one quarter (rare branch — keeps API safe).
        qm -= 3
        if qm < 3:
            qm = 12
            y -= 1
    out: list[str] = []
    for _ in range(max_quarters):
        last_day = {3: 31, 6: 30, 9: 30, 12: 31}[qm]
        out.append(f"{y:04d}-{qm:02d}-{last_day:02d}")
        qm -= 3
        if qm < 3:
            qm = 12
            y -= 1
    return out


def lookup_cu_state(
    charter: int,
    *,
    solr_url: str = DEFAULT_SOLR_URL,
    core: str = DEFAULT_CORE,
    max_quarters: int = 8,
) -> tuple[str | None, str | None]:
    """Return ``(state_abbr, source_period_iso)`` for ``charter`` from Solr.

    Walks back through quarter ends until a doc with a real (non-blank,
    non-``NA``) ``custate`` is found. Returns ``(None, None)`` on miss
    or any transport failure.
    """
    try:
        # Defer the heavy import so module-load is cheap.
        from cecl_ui.services.solr_5300_backfill import fetch_solr_doc
    except Exception:
        return None, None
    for qe in _walk_back_quarter_ends(max_quarters=max_quarters):
        try:
            doc = fetch_solr_doc(solr_url, core, int(charter), qe)
        except Exception:
            return None, None
        if not doc:
            continue
        raw = doc.get("custate")
        if isinstance(raw, list):
            raw = raw[0] if raw else None
        if raw is None:
            continue
        s = str(raw).strip().upper()
        if not s or s in {"NA", "N/A", "NONE", "NULL"}:
            continue
        if len(s) == 2 and s.isalpha():
            return s, qe
    return None, None


# ---------------------------------------------------------------- county via FOICU

def _latest_quarter_yms(max_quarters: int = 6) -> list[str]:
    """Yield ``YYYY-MM`` quarter-end markers newest-first for FOICU URLs."""
    out: list[str] = []
    for iso in _walk_back_quarter_ends(max_quarters=max_quarters):
        y, m, _d = iso.split("-")
        out.append(f"{y}-{m}")
    return out


def _download_foicu(workspace_root: str | Path | None) -> tuple[Path | None, str | None, str | None]:
    """Download (or load cached) FOICU.txt for the latest available quarter.

    Returns ``(path, quarter_ym, error)`` — ``path`` is None on failure.
    Cached file path: ``<cache>/foicu_<YYYY-MM>.txt``.
    """
    cache = _cache_dir(workspace_root)
    # Prefer the newest cached FOICU file if it covers the latest
    # quarter we'd otherwise re-fetch.
    candidates = _latest_quarter_yms()
    for ym in candidates:
        cached = cache / f"foicu_{ym}.txt"
        if cached.exists() and cached.stat().st_size > 1000:
            return cached, ym, None

    # No cache — download the newest available quarter we can find.
    try:
        import requests  # deferred — heavy
    except Exception as e:
        return None, None, f"requests not installed: {e}"

    last_err: str | None = None
    for ym in candidates:
        url = _NCUA_ZIP_URL.format(ym=ym)
        try:
            r = requests.get(url, timeout=120)
        except Exception as e:
            last_err = f"download {ym}: {e}"
            continue
        if r.status_code != 200 or len(r.content) < 100_000:
            last_err = f"download {ym}: HTTP {r.status_code} / {len(r.content)} bytes"
            continue
        # Unzip in-memory and pull FOICU.txt.
        try:
            z = zipfile.ZipFile(io.BytesIO(r.content))
            name = next((n for n in z.namelist() if n.upper() == "FOICU.TXT"), None)
            if not name:
                last_err = f"zip {ym}: FOICU.txt not found"
                continue
            data = z.read(name)
        except Exception as e:
            last_err = f"unzip {ym}: {e}"
            continue
        cached = cache / f"foicu_{ym}.txt"
        try:
            cached.write_bytes(data)
        except OSError as e:
            last_err = f"cache write {cached}: {e}"
            continue
        return cached, ym, None
    return None, None, last_err or "no FOICU file could be obtained"


def _parse_foicu_for_charter(foicu_path: Path, charter: int) -> dict[str, str] | None:
    """Scan FOICU.txt for the row matching ``charter`` (CU_NUMBER).

    Returns a dict of the relevant columns or ``None`` on miss.
    Uses ``csv.DictReader`` since FOICU.txt is a quoted CSV (not the
    pipe-delimited format some older NCUA files use).
    """
    try:
        with foicu_path.open("r", encoding="latin-1", newline="") as f:
            rdr = csv.DictReader(f)
            target = str(int(charter))
            for row in rdr:
                if (row.get("CU_NUMBER") or "").strip() == target:
                    return {
                        "cu_number": target,
                        "cu_name": (row.get("CU_NAME") or "").strip(),
                        "city": (row.get("CITY") or "").strip(),
                        "state": (row.get("STATE") or "").strip().upper(),
                        "county_code": (row.get("COUNTY_CODE") or "").strip(),
                        "zip_code": (row.get("ZIP_CODE") or "").strip(),
                    }
    except Exception:
        return None
    return None


# ---------------------------------------------------------------- county FIPS → name

_GAZETTEER_URL = (
    "https://www2.census.gov/geo/docs/reference/codes2020/"
    "national_county2020.txt"
)


def _gazetteer_cache_path(workspace_root: str | Path | None) -> Path:
    return _cache_dir(workspace_root) / "national_county2020.txt"


def _load_gazetteer(workspace_root: str | Path | None) -> dict[tuple[str, str], str]:
    """Build ``{(state_fips, county_fips_3digit): county_name}`` lookup.

    Cached to ``<cache>/national_county2020.txt`` after first fetch so
    the network call only happens once per workspace.
    """
    cache = _gazetteer_cache_path(workspace_root)
    text: str | None = None
    if cache.exists() and cache.stat().st_size > 1000:
        try:
            text = cache.read_text(encoding="utf-8")
        except Exception:
            text = None
    if text is None:
        try:
            import requests  # deferred
            r = requests.get(_GAZETTEER_URL, timeout=30)
            r.raise_for_status()
            text = r.text
            try:
                cache.write_text(text, encoding="utf-8")
            except OSError:
                pass
        except Exception:
            return {}
    out: dict[tuple[str, str], str] = {}
    for line in text.splitlines():
        if not line or line.startswith("STATE|"):
            continue
        parts = line.split("|")
        # Layout: STATE|STATEFP|COUNTYFP|COUNTYNS|COUNTYNAME|CLASSFP
        if len(parts) < 5:
            continue
        state_fips = parts[1].strip()
        county_fips = parts[2].strip()  # 3-digit, e.g. "033"
        name = (parts[4] or "").strip()
        if state_fips and county_fips and name:
            out[(state_fips, county_fips)] = name
    return out


def _resolve_county_name(
    state_abbr: str,
    county_code_3digit: str,
    workspace_root: str | Path | None,
) -> tuple[str | None, str | None]:
    """Return ``(county_name, county_fips_5digit)`` or ``(None, None)``."""
    _, abbr_to_fips = _state_maps()
    state_fips = abbr_to_fips.get((state_abbr or "").upper())
    if not state_fips:
        return None, None
    cc = (county_code_3digit or "").strip()
    # Normalise to zero-padded 3 digits.
    try:
        cc = f"{int(cc):03d}"
    except (TypeError, ValueError):
        return None, None
    gaz = _load_gazetteer(workspace_root)
    name = gaz.get((state_fips, cc))
    if not name:
        return None, None
    return name, f"{state_fips}{cc}"


# ---------------------------------------------------------------- combined entry point


def lookup_cu_geography(
    charter: int | str,
    *,
    solr_url: str = DEFAULT_SOLR_URL,
    core: str = DEFAULT_CORE,
    workspace_root: str | Path | None = None,
) -> dict[str, Any]:
    """Look up state + county for ``charter``. Never raises.

    See module docstring for the return shape.
    """
    out: dict[str, Any] = {
        "ok": False,
        "charter": None,
        "cu_name": None,
        "city": None,
        "state_abbr": None,
        "state": None,
        "state_fips": None,
        "county": None,
        "county_fips": None,
        "sources": {},
        "errors": [],
    }
    try:
        charter_int = int(str(charter).strip())
    except (TypeError, ValueError):
        out["errors"].append("charter must be an integer")
        return out
    out["charter"] = charter_int

    # ---- State (preferred: Solr; fallback: FOICU) -----------------
    state_abbr, period = lookup_cu_state(
        charter_int, solr_url=solr_url, core=core,
    )
    if state_abbr:
        out["state_abbr"] = state_abbr
        out["sources"]["state"] = f"solr:{period}"
    else:
        out["errors"].append("Solr returned no usable custate")

    # ---- County (FOICU.txt) ---------------------------------------
    foicu_path, foicu_ym, foicu_err = _download_foicu(workspace_root)
    foicu_row: dict[str, str] | None = None
    if foicu_path:
        foicu_row = _parse_foicu_for_charter(foicu_path, charter_int)
        if not foicu_row:
            out["errors"].append(
                f"Charter {charter_int} not found in FOICU {foicu_ym}"
            )
    elif foicu_err:
        out["errors"].append(foicu_err)

    if foicu_row:
        out["cu_name"] = foicu_row.get("cu_name") or None
        out["city"] = foicu_row.get("city") or None
        # Prefer FOICU's STATE column if Solr gave us nothing.
        if not out["state_abbr"]:
            fs = foicu_row.get("state") or ""
            if len(fs) == 2 and fs.isalpha():
                out["state_abbr"] = fs.upper()
                out["sources"]["state"] = f"foicu:{foicu_ym}"
        county_name, county_5 = _resolve_county_name(
            out["state_abbr"] or foicu_row.get("state", ""),
            foicu_row.get("county_code") or "",
            workspace_root,
        )
        if county_name:
            out["county"] = county_name
            out["county_fips"] = county_5
            out["sources"]["county"] = f"foicu:{foicu_ym}"
        else:
            out["errors"].append(
                "Could not resolve county FIPS "
                f"{foicu_row.get('state')}-{foicu_row.get('county_code')} "
                "via gazetteer"
            )

    # ---- Resolve full state name + FIPS for downstream display ----
    if out["state_abbr"]:
        abbr_to_name, abbr_to_fips = _state_maps()
        out["state"] = abbr_to_name.get(out["state_abbr"]) or out["state_abbr"]
        out["state_fips"] = abbr_to_fips.get(out["state_abbr"])
        out["ok"] = True
    return out


if __name__ == "__main__":  # pragma: no cover — manual smoke test
    import json as _json
    import sys
    ch = int(sys.argv[1]) if len(sys.argv) > 1 else 5641
    wr = os.environ.get("CECL_WORKSPACE_ROOT")
    print(_json.dumps(
        lookup_cu_geography(ch, workspace_root=wr),
        indent=2, default=str,
    ))
