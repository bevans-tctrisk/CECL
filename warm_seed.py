"""One-time WARM -> seed extractor (Phase 1 of the "WARM at setup only" migration).

A "WARM-sourced" credit union (its config has ``credit_pull.fallback_report_folder``
pointing at a folder of analyst-built ``CECL-Migration-WARM - <CU>.xlsx`` workbooks)
currently has the report engine RE-READ that workbook every quarter for its loss-rate
history, non-risk-rated base rates, non-extract balances, environmental ranges and ACL
months. The goal is to consult the WARM only ONCE, at onboarding, freeze everything it
supplies into a durable per-CU seed, and thereafter drive the report from the raw data
plus that seed (rolled forward each quarter).

This module performs the one-time capture. It is purely additive: it reads the most
recent WARM workbook via the engine's own parser (``generate_report.load_impaired_data``)
and writes a JSON seed to ``<workspace>/client_configs/_warm_seeds/<short>.json``. It does
NOT modify any config, database row, or report output — wiring the report engine to
consume the seed is a later phase.

Usage:
    python warm_seed.py <short_name> [YYYY-MM-DD]
    # snapshot optional; defaults to the newest WARM workbook found for the CU
"""

from __future__ import annotations

import copy
import datetime as _dt
import json
import os
import re
from calendar import monthrange

import yaml


# Keys that load_impaired_data returns, classified by how a future quarter should
# treat them once the report is seed-driven. Recorded in the manifest so the
# consumption phase knows what to roll forward vs. recompute from raw data.
_SEED_KEY_ROLES = {
    # Historical rolling series — seed once, then APPEND each quarter's raw actuals.
    "hist_bal_data": "historical_series",
    "warm_co_monthly": "historical_series",
    "warm_rc_monthly": "historical_series",
    # Semi-static — frozen at setup from the WARM; analyst edits only when they change.
    "acl_pools": "semi_static",          # carries NRR per-grade base rates
    "acl_months": "semi_static",
    "risk_rated": "semi_static",
    "env_ranges": "semi_static",
    "balance_adjustments": "semi_static",  # non-extract balances (mortgages, participations)
    "pool_bal_detail": "semi_static",
    "total_balance_adjustment": "semi_static",
    "total_in_portfolio": "semi_static",
    "economic_data": "semi_static",
    "pool_order": "semi_static",
    # Current-quarter — captured for reference only; a seeded report recomputes these
    # from the current raw loan extract / impaired file / 5300, NOT from the seed.
    "acl_balance": "current_quarter",
    "pooled_total_allowance": "current_quarter",
    "spec_id_by_pool": "current_quarter",
    "acl_impaired": "current_quarter",
    "acl_summary": "current_quarter",
    "exec_summary_3": "current_quarter",
    "co_by_status": "current_quarter",
    "co_by_pool": "current_quarter",
    "dq_by_status": "current_quarter",
    "dq_by_pool": "current_quarter",
    "items": "current_quarter",
    "total_spec_id": "current_quarter",
}


def _jsonable(obj):
    """Recursively convert a value into a JSON-serialisable form.

    Handles the datetime/pandas/numpy scalars that appear in the parsed WARM
    structures (e.g. ``hist_bal_data[pool]['dates']`` holds Timestamps).
    """
    # Scalars first
    if obj is None or isinstance(obj, (bool, int, str)):
        return obj
    if isinstance(obj, float):
        return None if obj != obj else obj  # NaN -> None
    if isinstance(obj, (_dt.datetime, _dt.date)):
        return obj.isoformat()
    # numpy / pandas scalars expose .item(); Timestamps expose .isoformat()
    if hasattr(obj, "isoformat") and callable(obj.isoformat):
        try:
            return obj.isoformat()
        except Exception:  # noqa: BLE001
            pass
    if hasattr(obj, "item") and callable(getattr(obj, "item")):
        try:
            return _jsonable(obj.item())
        except Exception:  # noqa: BLE001
            pass
    if isinstance(obj, dict):
        # JSON object keys must be strings
        return {str(_jsonable_key(k)): _jsonable(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [_jsonable(v) for v in obj]
    if isinstance(obj, set):
        return sorted(_jsonable(v) for v in obj)
    return str(obj)


def _jsonable_key(k):
    if isinstance(k, (_dt.datetime, _dt.date)):
        return k.isoformat()
    # (year, month) tuple keys (warm_co_monthly / warm_rc_monthly) -> "YYYY-MM"
    if (isinstance(k, tuple) and len(k) == 2
            and all(isinstance(x, int) for x in k)
            and 1900 <= k[0] <= 2100 and 1 <= k[1] <= 12):
        return f"{k[0]:04d}-{k[1]:02d}"
    if hasattr(k, "isoformat") and callable(k.isoformat):
        try:
            return k.isoformat()
        except Exception:  # noqa: BLE001
            return str(k)
    return k


def _load_config(short_name: str, workspace_root: str) -> dict:
    path = os.path.join(workspace_root, "client_configs", f"{short_name}.yaml")
    if not os.path.isfile(path):
        raise FileNotFoundError(f"No config found at {path}")
    with open(path, encoding="utf-8") as fh:
        return yaml.safe_load(fh) or {}


def _newest_warm_snapshot(config: dict) -> str | None:
    """Return the newest 'YYYY-MM-DD' month-end for which a WARM workbook exists."""
    cu = config.get("credit_union", "")
    first = cu.lower().split()[0] if cu.strip() else ""
    safe_cu = cu.replace(" ", "_").replace("/", "-").lower()
    dirs = []
    for key in (config.get("data_directory", ""),
                (config.get("credit_pull") or {}).get("fallback_report_folder", "")):
        if key and key not in dirs:
            dirs.append(key)
    pat = re.compile(r"(\d{4})-(\d{2}).*CECL[\s_\-]+Migration[\s_\-]+WARM.*\.xls[xm]?$",
                     re.IGNORECASE)
    best = None  # (year, month)
    for d in dirs:
        if not os.path.isdir(d):
            continue
        for root, _sub, files in os.walk(d):
            for f in files:
                if f.startswith("~$") or f.upper().startswith("DNU"):
                    continue
                m = pat.match(f)
                if not m:
                    continue
                norm = f.lower().replace(" ", "_")
                if cu and safe_cu not in norm and not (len(first) >= 4 and first in norm):
                    continue
                ym = (int(m.group(1)), int(m.group(2)))
                if best is None or ym > best:
                    best = ym
    if not best:
        return None
    y, mo = best
    return f"{y:04d}-{mo:02d}-{monthrange(y, mo)[1]:02d}"


def build_seed(short_name: str, snap: str | None = None,
               workspace_root: str | None = None) -> dict:
    """Capture the WARM data for *short_name* into a durable JSON seed.

    Returns the manifest dict. Writes ``client_configs/_warm_seeds/<short>.json``.
    """
    workspace_root = workspace_root or os.environ.get(
        "CECL_WORKSPACE_ROOT", os.getcwd())
    os.environ.setdefault("CECL_WORKSPACE_ROOT", workspace_root)

    config = _load_config(short_name, workspace_root)

    if snap is None:
        snap = _newest_warm_snapshot(config)
        if snap is None:
            raise RuntimeError(
                f"No CECL-Migration-WARM workbook found for '{short_name}' in its "
                f"data_directory or credit_pull.fallback_report_folder — nothing to seed.")

    # Import here so CECL_WORKSPACE_ROOT is set before the module resolves BASE.
    import generate_report as gr

    warm = gr.load_impaired_data(config, snap)
    if not warm:
        raise RuntimeError(
            f"load_impaired_data returned no data for '{short_name}' at {snap} — "
            f"the WARM workbook was not found or not parseable.")

    warm_json = _jsonable(warm)

    # Build a small validation/summary block so the capture can be eyeballed.
    hbd = warm.get("hist_bal_data") or {}
    months = 0
    for pdata in hbd.values():
        if isinstance(pdata, dict) and pdata.get("dates"):
            months = max(months, len(pdata["dates"]))
    summary = {
        "pools_in_acl": len(warm.get("acl_pools") or {}),
        "hist_bal_pools": len(hbd),
        "hist_bal_months": months,
        "warm_co_monthly_pools": len(warm.get("warm_co_monthly") or {}),
        "warm_rc_monthly_pools": len(warm.get("warm_rc_monthly") or {}),
        "pooled_total_allowance": warm.get("pooled_total_allowance"),
        "acl_balance": warm.get("acl_balance"),
        "total_balance_adjustment": warm.get("total_balance_adjustment"),
        "env_range_groups": len(warm.get("env_ranges") or {}),
        "acl_months_pools": len(warm.get("acl_months") or {}),
    }
    roles = {k: _SEED_KEY_ROLES.get(k, "uncategorized") for k in warm.keys()}

    manifest = {
        "schema": "warm_seed/v1",
        "short_name": short_name,
        "credit_union": config.get("credit_union", ""),
        "seed_snapshot": snap,
        "generated_at": _dt.datetime.now().isoformat(timespec="seconds"),
        "source": "generate_report.load_impaired_data (CECL-Migration-WARM workbook)",
        "summary": summary,
        "key_roles": roles,
        "warm_data": warm_json,
    }

    out_dir = os.path.join(workspace_root, "client_configs", "_warm_seeds")
    os.makedirs(out_dir, exist_ok=True)
    out_path = os.path.join(out_dir, f"{short_name}.json")
    with open(out_path, "w", encoding="utf-8") as fh:
        json.dump(manifest, fh, indent=2, ensure_ascii=False)
    manifest["_path"] = out_path
    return manifest


def _rehydrate(warm: dict) -> dict:
    """Inverse of the JSON sanitising done at capture time, so the seed is a
    drop-in for ``load_impaired_data``'s live return value.

    - ``hist_bal_data[pool]['dates']`` ISO strings -> ``pd.Timestamp`` (the ACL
      reserve compute does ``d.year`` on them).
    - ``warm_co_monthly`` / ``warm_rc_monthly`` ``"YYYY-MM"`` keys -> ``(year,
      month)`` tuples (the history overlay indexes ``ym[0]``/``ym[1]``).
    """
    import pandas as pd

    warm = copy.deepcopy(warm)
    hbd = warm.get("hist_bal_data")
    if isinstance(hbd, dict):
        for pdata in hbd.values():
            if isinstance(pdata, dict) and isinstance(pdata.get("dates"), list):
                pdata["dates"] = [pd.Timestamp(d) for d in pdata["dates"]]
    for mk in ("warm_co_monthly", "warm_rc_monthly"):
        series = warm.get(mk)
        if isinstance(series, dict):
            rebuilt = {}
            for k, v in series.items():
                if isinstance(k, str) and len(k) == 7 and k[4] == "-":
                    rebuilt[(int(k[:4]), int(k[5:7]))] = v
                else:
                    rebuilt[k] = v
            warm[mk] = rebuilt
    return warm


def _find_seed_path(config: dict, workspace_root: str) -> str | None:
    """Locate this CU's seed file by matching the manifest credit_union name."""
    seed_dir = os.path.join(workspace_root, "client_configs", "_warm_seeds")
    if not os.path.isdir(seed_dir):
        return None
    cu = (config.get("credit_union") or "").strip().lower()
    lookup = (config.get("_lookup_credit_union") or "").strip().lower()
    for fn in os.listdir(seed_dir):
        if not fn.endswith(".json"):
            continue
        path = os.path.join(seed_dir, fn)
        try:
            with open(path, encoding="utf-8") as fh:
                m = json.load(fh)
        except Exception:  # noqa: BLE001
            continue
        seed_cu = (m.get("credit_union") or "").strip().lower()
        if seed_cu and seed_cu in (cu, lookup):
            return path
    return None


def load_seed(config: dict, snap: str | None = None,
              workspace_root: str | None = None) -> dict | None:
    """Return the rehydrated WARM data captured for this CU, or None.

    Only active when the config opts in via ``warm_seed_driven: true`` (so no
    other CU's behaviour changes). Substitutes for the live-WARM read in
    ``load_impaired_data``: it supplies the deep historical loss series and the
    semi-static per-pool settings, but STRIPS the seed-quarter's frozen
    allowance so the engine recomputes it against the CURRENT quarter's loan
    balances (verified to reproduce the WARM to the cent at the seed quarter).
    """
    if not config.get("warm_seed_driven"):
        return None
    workspace_root = workspace_root or os.environ.get(
        "CECL_WORKSPACE_ROOT", os.getcwd())
    path = _find_seed_path(config, workspace_root)
    if not path:
        return None
    with open(path, encoding="utf-8") as fh:
        manifest = json.load(fh)
    warm = manifest.get("warm_data")
    if not warm:
        return None
    return _prepare_for_recompute(_rehydrate(warm), config)


# Seed-quarter values the engine must RECOMPUTE from the current loan balances
# rather than reuse frozen (otherwise every future quarter shows the seed
# quarter's numbers). Impaired/spec-id come from the current uploaded impaired
# file; ACL balance from the CU's 719001-00 GL entry (acl.history).
_FROZEN_KEYS = (
    "acl_pools", "acl_summary", "pooled_total_allowance", "spec_id_by_pool",
    "acl_impaired", "exec_summary_3", "pool_order", "acl_balance",
)


def _seed_rate_overrides(acl_pools: dict, risk_rated: dict) -> dict:
    """Build ``base_loss_rate_by_pool_grade`` from the seed's per-grade effective
    ``factor`` (= WARM base loss rate + its management adjustment). Applying that
    frozen effective rate to the CURRENT balances reproduces the WARM's pooled
    allowance and rolls forward as balances change. Risk-rated pools carry a rate
    per grade; non-risk-rated pools carry a single ``Total`` rate."""
    ovr = {}
    for pool, pdata in (acl_pools or {}).items():
        if risk_rated.get(pool, True):
            gm = {g: (gv or {}).get("factor")
                  for g, gv in (pdata.get("grades") or {}).items()
                  if (gv or {}).get("factor") is not None}
            if gm:
                ovr[pool] = gm
        else:
            f = (pdata.get("total") or {}).get("factor")
            if f is not None:
                ovr[pool] = {"Total": f}
    return ovr


def _prepare_for_recompute(warm: dict, config: dict | None = None) -> dict:
    """Strip frozen seed-quarter values, drop the WARM parse's junk per-pool keys
    (it captures ~30 spurious ``risk_rated``/``acl_months`` entries such as
    'Credit Grade Deteriorated Type 1'), and publish the WARM's frozen effective
    loss rates into ``config['base_loss_rate_by_pool_grade']`` so the engine
    recomputes each pool's allowance as those rates x the CURRENT balances."""
    warm = dict(warm)
    acl_pools = warm.get("acl_pools") or {}
    real_pools = set(acl_pools.keys())
    risk_rated = warm.get("risk_rated") or {}
    if config is not None and acl_pools and not config.get("base_loss_rate_by_pool_grade"):
        ovr = _seed_rate_overrides(acl_pools, risk_rated)
        if ovr:
            config["base_loss_rate_by_pool_grade"] = ovr
    for k in _FROZEN_KEYS:
        warm.pop(k, None)
    if real_pools:
        for mk in ("risk_rated", "acl_months"):
            if isinstance(warm.get(mk), dict):
                warm[mk] = {p: v for p, v in warm[mk].items() if p in real_pools}
    return warm


def _print_manifest(m: dict) -> None:
    print(f"  Seed written: {m.get('_path')}")
    print(f"  CU: {m['credit_union']}  |  seed snapshot: {m['seed_snapshot']}")
    print("  Captured summary:")
    for k, v in m["summary"].items():
        if isinstance(v, float):
            print(f"    {k:28} {v:,.2f}")
        else:
            print(f"    {k:28} {v}")
    hist = [k for k, r in m["key_roles"].items() if r == "historical_series"]
    semi = [k for k, r in m["key_roles"].items() if r == "semi_static"]
    print(f"  historical_series keys: {', '.join(sorted(hist)) or '(none)'}")
    print(f"  semi_static keys:       {', '.join(sorted(semi)) or '(none)'}")


if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("usage: python warm_seed.py <short_name> [YYYY-MM-DD]")
        raise SystemExit(2)
    _short = sys.argv[1]
    _snap = sys.argv[2] if len(sys.argv) > 2 else None
    _m = build_seed(_short, _snap)
    _print_manifest(_m)
