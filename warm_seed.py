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


def _read_bs_co_dq_pools(warm_path: str | None) -> dict:
    """Read the WARM ``BS CO DQ Data Enter`` tab -> authoritative per-pool config.

    That tab is the analyst's data-entry sheet: col A = Loan Pools, col B = Risk
    Rated Yes/No, col G = ACL Months, col I = Pool Order. It is the source of
    truth for which pools are risk-rated (per-grade FICO curve) vs flat
    non-risk-rated (e.g. off-extract credit-card / solar balances). Returns
    ``{pool_name: {'risk_rated': bool, 'acl_months': int|None, 'pool_order':
    int|None}}``; empty dict on any failure (older WARMs may lack the tab)."""
    out: dict = {}
    if not warm_path or not os.path.isfile(warm_path):
        return out
    try:
        import openpyxl
        wb = openpyxl.load_workbook(warm_path, read_only=True, data_only=True)
    except Exception:
        return out
    try:
        ws = None
        for name in wb.sheetnames:
            if name.strip().lower() == "bs co dq data enter":
                ws = wb[name]
                break
        if ws is None:
            return out
        rows = list(ws.iter_rows(values_only=True))
        hdr_i = None
        for i, r in enumerate(rows):
            if (r and isinstance(r[0], str) and "loan pool" in r[0].lower()
                    and len(r) > 1 and isinstance(r[1], str)
                    and "risk rated" in r[1].lower()):
                hdr_i = i
                break
        if hdr_i is None:
            return out
        for r in rows[hdr_i + 1:]:
            if not r or r[0] is None or not str(r[0]).strip():
                continue
            pool = str(r[0]).strip()
            low = pool.lower()
            if "grand total" in low:
                break  # end of the pool table; sections below are not pools
            if low.startswith("hide") or low in ("exclude", "total"):
                continue
            rr = (str(r[1]).strip().lower() if len(r) > 1 and r[1] is not None else "")
            acl_m = None
            if len(r) > 6 and r[6] is not None:
                try:
                    acl_m = int(float(r[6]))
                except (TypeError, ValueError):
                    acl_m = None
            order = None
            if len(r) > 8 and r[8] is not None:
                try:
                    order = int(float(r[8]))
                except (TypeError, ValueError):
                    order = None
            out[pool] = {"risk_rated": rr == "yes",
                         "acl_months": acl_m, "pool_order": order}
    finally:
        wb.close()
    return out


def _apply_bs_co_dq_to_config(config: dict, bs_pools: dict) -> None:
    """Set ``not_risk_rated`` (and pool_order / acl_months_by_pool) on *config*
    from the WARM's ``BS CO DQ Data Enter`` tab. Authoritative for risk-rated:
    any pool the tab marks 'No' becomes flat non-risk-rated, so a CU never needs
    the flags hand-maintained. Existing NRR entries (e.g. 'Ignore') are kept."""
    if not bs_pools:
        return
    nrr = {p for p, v in bs_pools.items() if not v.get("risk_rated", True)}
    nrr.update(config.get("not_risk_rated") or [])
    config["not_risk_rated"] = sorted(nrr)
    order = [p for p, _ in sorted(
        ((p, v.get("pool_order")) for p, v in bs_pools.items()
         if v.get("pool_order") is not None), key=lambda kv: kv[1])]
    if order:
        config["pool_order"] = order
    acl_m = {p: v["acl_months"] for p, v in bs_pools.items()
             if v.get("acl_months") is not None}
    if acl_m:
        config.setdefault("acl_months_by_pool", {}).update(acl_m)


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

    # Capture the authoritative per-pool config from the WARM's data-entry tab
    # (risk-rated Yes/No, ACL months, pool order) so forward quarters don't need
    # the flags hand-maintained.
    _warm_path = None
    try:
        _warm_path = gr._find_prior_warm_xlsx(config, snap)
    except Exception:  # noqa: BLE001
        _warm_path = None
    bs_co_dq_pools = _read_bs_co_dq_pools(_warm_path)
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
        "bs_co_dq_pools": bs_co_dq_pools,
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
    # Authoritative risk-rated / pool-order / ACL-months from the WARM's
    # 'BS CO DQ Data Enter' tab (captured at seed-build). Applied to the live
    # config so the model determines these from the WARM, not hand-config.
    _apply_bs_co_dq_to_config(config, manifest.get("bs_co_dq_pools") or {})
    return _prepare_for_recompute(_rehydrate(warm), config)


# Seed-quarter values the engine must RECOMPUTE from the current loan balances
# rather than reuse frozen (otherwise every future quarter shows the seed
# quarter's numbers). Impaired/spec-id come from the current uploaded impaired
# file; ACL balance from the CU's 719001-00 GL entry (acl.history).
_FROZEN_KEYS = (
    "acl_pools", "acl_summary", "pooled_total_allowance", "spec_id_by_pool",
    "acl_impaired", "exec_summary_3", "pool_order", "acl_balance",
)


def _seed_mgmt_adj_overrides(acl_pools: dict, risk_rated: dict) -> dict:
    """Per-(pool, grade) management adjustment captured from the seed.

    The management adjustment is the qualitative overlay the analyst sets in the
    WARM; it is semi-static and must persist across quarters. The ACL *base loss
    rate*, by contrast, is recomputed every quarter from the rolled-forward
    charge-off/recovery history — so we freeze ONLY the mgmt adjustment here (not
    the combined ``factor``) and let the engine recompute the base rate. Injected
    as ``prior_mgmt_adj`` so ``_resolve_mgmt_adj_grade`` applies it on top of the
    fresh base rate. Only risk-rated pools carry per-grade values; NRR pools take
    their adjustment through the pool-total resolver (config-driven)."""
    ovr = {}
    for pool, pdata in (acl_pools or {}).items():
        if not risk_rated.get(pool, True):
            continue
        gm = {g: (gv or {}).get("mgmt_adj")
              for g, gv in (pdata.get("grades") or {}).items()
              if (gv or {}).get("mgmt_adj") is not None}
        if gm:
            ovr[pool] = gm
    return ovr


def _prepare_for_recompute(warm: dict, config: dict | None = None) -> dict:
    """Strip frozen seed-quarter values, drop the WARM parse's junk per-pool keys
    (it captures ~30 spurious ``risk_rated``/``acl_months`` entries such as
    'Credit Grade Deteriorated Type 1'), and carry the seed's per-grade
    management adjustment forward as ``prior_mgmt_adj`` so the engine recomputes
    each pool's ACL base loss rate from the current charge-off history while
    keeping the qualitative overlay stable."""
    warm = dict(warm)
    acl_pools = warm.get("acl_pools") or {}
    real_pools = set(acl_pools.keys())
    risk_rated = warm.get("risk_rated") or {}
    if acl_pools:
        madj = _seed_mgmt_adj_overrides(acl_pools, risk_rated)
        if madj:
            existing = dict(warm.get("prior_mgmt_adj") or {})
            for _p, _gm in madj.items():
                existing.setdefault(_p, {}).update(_gm)
            warm["prior_mgmt_adj"] = existing
    # Preserve the frozen seed allowance for small NRR "specific reserve" pools
    # the firm-wide model cannot recompute without their GL balance (e.g. Erie's
    # Courtesy Pay / Negative Share / Business Credit Card). Their seed-quarter
    # allowance passes through via ``warm_allowance_pools``; balances are stable.
    _frozen_allow = {
        str(p).strip().lower()
        for p in ((config or {}).get("seed_frozen_allowance_pools") or [])
        if str(p).strip()
    }
    _preserved = {p: pd for p, pd in acl_pools.items()
                  if p.strip().lower() in _frozen_allow} if _frozen_allow else {}
    for k in _FROZEN_KEYS:
        warm.pop(k, None)
    if _preserved:
        warm["acl_pools"] = _preserved
    # Drop the seed-quarter's frozen life-of-loan net charge-off so the base loss
    # rate recomputes from the rolled-forward charge-off/recovery history.
    warm.pop("warm_net_co", None)
    # Drop the frozen total-in-portfolio so the exec summary recomputes it from
    # the current grand balance + the (semi-static) non-extract adjustment,
    # rather than showing the seed quarter's stale total.
    warm.pop("total_in_portfolio", None)
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
