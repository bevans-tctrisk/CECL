"""Read/write client_configs YAML files."""
from __future__ import annotations

import re
from pathlib import Path
from typing import Any

import yaml


CONFIG_DIR_NAME = "client_configs"
RAW_UPLOADS_DIR_NAME = "Raw_Uploads"


def configs_dir(workspace_root: str | Path) -> Path:
    return Path(workspace_root) / CONFIG_DIR_NAME


def raw_uploads_dir(workspace_root: str | Path) -> Path:
    return Path(workspace_root) / RAW_UPLOADS_DIR_NAME


def list_existing_clients(workspace_root: str | Path) -> list[dict[str, str]]:
    """Return [{short_name, credit_union}] for every YAML file (excluding _template)."""
    out: list[dict[str, str]] = []
    cdir = configs_dir(workspace_root)
    if not cdir.exists():
        return out
    for yml in sorted(cdir.glob("*.yaml")):
        if yml.stem.startswith("_"):
            continue
        try:
            data = yaml.safe_load(yml.read_text(encoding="utf-8")) or {}
        except Exception:
            data = {}
        out.append(
            {
                "short_name": yml.stem,
                "credit_union": data.get("credit_union", yml.stem),
            }
        )
    return out


def slugify(name: str) -> str:
    """Turn 'Sample Credit Union' into 'sample_credit_union' (safe filename)."""
    s = re.sub(r"[^A-Za-z0-9]+", "_", name.strip().lower()).strip("_")
    return s or "client"


def short_name_available(workspace_root: str | Path, short_name: str) -> bool:
    return not (configs_dir(workspace_root) / f"{short_name}.yaml").exists()


def save_client_config(
    workspace_root: str | Path,
    short_name: str,
    config: dict[str, Any],
    overwrite: bool = False,
) -> Path:
    cdir = configs_dir(workspace_root)
    cdir.mkdir(parents=True, exist_ok=True)
    target = cdir / f"{short_name}.yaml"
    if target.exists() and not overwrite:
        raise FileExistsError(target)
    # Pretty-dump preserving order, no anchor aliases.
    text = yaml.safe_dump(
        config,
        sort_keys=False,
        default_flow_style=False,
        allow_unicode=True,
    )
    target.write_text(text, encoding="utf-8")
    # Also create the Raw_Uploads/<short_name>/ folder.
    upload_folder = raw_uploads_dir(workspace_root) / short_name
    upload_folder.mkdir(parents=True, exist_ok=True)
    return target


def load_client_config(workspace_root: str | Path, short_name: str) -> dict[str, Any]:
    target = configs_dir(workspace_root) / f"{short_name}.yaml"
    return yaml.safe_load(target.read_text(encoding="utf-8")) or {}


# ---------- Canonical pools registry ----------

def get_pools(cfg: dict[str, Any]) -> list[dict[str, Any]]:
    """Return the canonical list of pools for a CU.

    Each pool is a dict with keys::

        {"name": str,
         "risk_rated": bool,
         "acl_months": int | None,
         "excluded": bool,
         "use_default_mgmt_adj": bool}

    If ``cfg`` already has a top-level ``pools`` block (the new schema), that
    is returned verbatim with defaults filled in. Otherwise the list is
    synthesized from the legacy ``pool_order`` / ``not_risk_rated`` /
    ``excluded_pools`` / ``acl_months_by_pool`` fields so older configs keep
    working until they're re-saved through the editor.
    """
    raw = cfg.get("pools")
    if isinstance(raw, list) and raw:
        return [_normalize_pool(p) for p in raw if isinstance(p, dict) and p.get("name")]

    # Legacy fallback — synthesize from the old keys.
    nrr = set(cfg.get("not_risk_rated") or [])
    excl = set(cfg.get("excluded_pools") or [])
    acl_map = cfg.get("acl_months_by_pool") or {}
    order = list(cfg.get("pool_order") or [])

    # Union of every name we can find, preserving pool_order first.
    seen: set[str] = set()
    names: list[str] = []
    for n in order:
        if n and n not in seen:
            seen.add(n); names.append(n)
    for src in ((cfg.get("pool_map") or {}).values(),
                acl_map.keys(), nrr, excl):
        for n in src:
            if n and n not in seen:
                seen.add(n); names.append(n)

    out: list[dict[str, Any]] = []
    for n in names:
        out.append({
            "name": n,
            "risk_rated": (n not in nrr) and (n not in excl),
            "acl_months": int(acl_map[n]) if n in acl_map else None,
            "excluded": n in excl,
            "use_default_mgmt_adj": False,
        })
    return out


def _normalize_pool(p: dict[str, Any]) -> dict[str, Any]:
    name = str(p.get("name") or "").strip()
    acl = p.get("acl_months")
    try:
        acl_int: int | None = int(acl) if acl not in (None, "", 0) else None
    except (TypeError, ValueError):
        acl_int = None
    return {
        "name": name,
        "risk_rated": bool(p.get("risk_rated", True)) and not bool(p.get("excluded", False)),
        "acl_months": acl_int,
        "excluded": bool(p.get("excluded", False)),
        "use_default_mgmt_adj": bool(p.get("use_default_mgmt_adj", False)),
    }


def set_pools(cfg: dict[str, Any], pools: list[dict[str, Any]]) -> dict[str, Any]:
    """Write ``pools`` into ``cfg`` and keep the legacy fields in sync so the
    report engine, importer, and any older code path that still reads
    ``pool_order`` / ``excluded_pools`` / ``not_risk_rated`` / ``acl_months_by_pool``
    continues to work without a flag day.

    Returns the mutated ``cfg`` for chaining.
    """
    norm = [_normalize_pool(p) for p in pools if (p or {}).get("name")]
    cfg["pools"] = norm

    cfg["pool_order"] = [p["name"] for p in norm]
    nrr = [p["name"] for p in norm if not p["risk_rated"] and not p["excluded"]]
    if nrr:
        cfg["not_risk_rated"] = nrr
    else:
        cfg.pop("not_risk_rated", None)
    excl = [p["name"] for p in norm if p["excluded"]]
    if excl:
        cfg["excluded_pools"] = excl
    else:
        cfg.pop("excluded_pools", None)
    acl = {p["name"]: p["acl_months"] for p in norm
           if p["acl_months"] and not p["excluded"]}
    if acl:
        cfg["acl_months_by_pool"] = acl
    else:
        cfg.pop("acl_months_by_pool", None)
    if not cfg["pool_order"]:
        cfg.pop("pool_order", None)
    return cfg


# ---------- Defaults / templates for the wizard ----------

DEFAULT_CREDIT_GRADES: list[dict[str, Any]] = [
    {"label": "A+", "min_score": 720, "max_score": 900, "reserve_rate": 0.0011},
    {"label": "A",  "min_score": 680, "max_score": 719, "reserve_rate": 0.0025},
    {"label": "B",  "min_score": 640, "max_score": 679, "reserve_rate": 0.0050},
    {"label": "C",  "min_score": 620, "max_score": 639, "reserve_rate": 0.0116},
    {"label": "D",  "min_score": 600, "max_score": 619, "reserve_rate": 0.0250},
    {"label": "E",  "min_score": 0,   "max_score": 599, "reserve_rate": 0.0500},
]


def build_yaml_from_wizard(state: dict[str, Any]) -> dict[str, Any]:
    """Convert the wizard's session state dict into a final YAML-shaped dict."""
    cfg: dict[str, Any] = {
        "credit_union": state["credit_union"],
    }
    # NCUA charter number — only emit when present.
    if state.get("charter_number"):
        cfg["charter_number"] = str(state["charter_number"])
    cfg.update({
        "file_pattern": state["file_pattern"],
        "date_pattern": state["date_pattern"],
        "account_suffix_length": int(state.get("account_suffix_length", 3)),
        "member_account": dict(state.get("member_account") or {
            "mode": "fixed_suffix",
            "suffix_length": int(state.get("account_suffix_length", 3)),
            "delimiter": "-",
        }),
        "has_header": bool(state.get("has_header", True)),
        "column_mappings": state["column_mappings"],
        "credit_pull": state["credit_pull"],
        "balance_format": {
            "remove_chars": state.get("balance_remove_chars", ["$", ","]),
            "accounting_negatives": bool(
                state.get("accounting_negatives", True)
            ),
        },
        "pool_code_split": state.get("pool_code_split", "/"),
        "pool_map": state["pool_map"],
        "default_pool": state.get("default_pool", "Other/Uncategorized"),
        "credit_grades": [
            g for g in state["credit_grades"]
            if (g.get("label") or "").strip()
            != (state.get("no_score_label") or "Not Reported").strip()
        ],
        "no_score_label": state.get("no_score_label", "Not Reported"),
        "reports": state["reports"],
        "economic_data": state["economic_data"],
        "mgmt_adj": state["mgmt_adj"],
    })
    # Per-pool management-adjustment overlay — only emitted when non-empty so
    # the YAML stays clean for CUs that don't use it.
    overlay = state.get("mgmt_adj_by_pool") or {}
    if overlay:
        cfg["mgmt_adj_by_pool"] = overlay

    # Per-file loan-data extracts. When the CU has multiple loan-data
    # files with different column layouts (e.g. mortgages vs. autos) or
    # different fixed pool codes, each gets its own block here. The
    # importer routes monthly files to the matching extract by
    # ``file_pattern`` regex. Top-level ``column_mappings`` /
    # ``member_account`` / ``has_header`` / ``file_pattern`` are kept as
    # a back-compat fallback (mirror of the first extract).
    loan_files = (
        (state.get("sample_uploads") or {}).get("loan_data_files") or []
    )
    extracts_block: list[dict[str, Any]] = []
    for lf in loan_files:
        cm = dict(lf.get("column_mappings") or {})
        ma_lf = lf.get("member_account") or {}
        if not cm:
            continue
        entry: dict[str, Any] = {
            "label": lf.get("name") or "",
            "file_pattern": lf.get("file_pattern") or "",
            "column_mappings": cm,
            "member_account": dict(ma_lf) if ma_lf else dict(
                cfg.get("member_account") or {}
            ),
            "has_header": bool(lf.get("has_header")),
        }
        extracts_block.append(entry)
    if extracts_block:
        cfg["loan_data_extracts"] = extracts_block

    ps_list = state.get("pool_settings") or []
    # Promote the wizard's per-pool table into the canonical ``pools`` block
    # (and let set_pools mirror it into the legacy pool_order /
    # not_risk_rated / excluded_pools / acl_months_by_pool fields so the
    # report engine and importer continue to work unchanged).
    pools_block = []
    for p in ps_list:
        if not p.get("name"):
            continue
        pools_block.append({
            "name": p["name"],
            "risk_rated": bool(p.get("risk_rated")),
            "acl_months": int(p["acl_months"]) if p.get("acl_months") else None,
            "excluded": bool(p.get("excluded")),
            "use_default_mgmt_adj": bool(p.get("use_default_mgmt_adj")),
        })
    if pools_block:
        set_pools(cfg, pools_block)

    try:
        acl_bal = float(state.get("acl_balance") or 0)
    except (TypeError, ValueError):
        acl_bal = 0.0
    if acl_bal:
        cfg["acl_balance"] = acl_bal

    # ACL settings sourced from the Monthly Balance step (Step 5). The
    # wizard offers three sources: "monthly_file" (auto-detected row in the
    # monthly balance file), "separate" (uploaded standalone file), and
    # "manual" (user-entered values). For the latter two, we fall back to
    # writing the snap-month value into ``acl_balance`` so the report
    # engine's existing ``config.get('acl_balance', 0)`` lookup picks it up.
    mb_state = state.get("monthly_bal") or {}
    acl_state = mb_state.get("acl") or {}
    acl_src = acl_state.get("source") or ""
    acl_history = acl_state.get("history") or {}
    acl_block: dict[str, Any] = {}
    if acl_src:
        acl_block["source"] = acl_src
    if acl_src == "monthly_file" and acl_state.get("row"):
        acl_block["row"] = int(acl_state["row"])
        if acl_state.get("label"):
            acl_block["label"] = acl_state["label"]
    elif acl_src == "separate":
        sep = acl_state.get("separate_file") or {}
        if sep.get("filename"):
            acl_block["separate_file"] = {
                "filename": sep.get("filename", ""),
                "sheet": sep.get("sheet", ""),
                "cell": sep.get("cell", ""),
            }
            if sep.get("value") is not None:
                try:
                    cfg["acl_balance"] = float(sep["value"])
                except (TypeError, ValueError):
                    pass
    elif acl_src == "manual":
        manual = acl_state.get("manual") or {}
        manual_clean: dict[str, float] = {}
        for k in ("month1", "month2", "month3"):
            d = (manual.get(f"{k}_date") or "").strip()
            v = manual.get(f"{k}_value")
            if d and v is not None:
                try:
                    manual_clean[d] = float(v)
                except (TypeError, ValueError):
                    pass
        if manual_clean:
            acl_block["manual"] = manual_clean
            # Use the latest month-end value as the current acl_balance.
            latest_date = max(manual_clean.keys())
            cfg["acl_balance"] = manual_clean[latest_date]
            # Merge into history so it can be carried forward.
            acl_history = {**acl_history, **manual_clean}
    if acl_history:
        # Persist sorted by date for stable YAML diffs.
        acl_block["history"] = {
            d: float(v) for d, v in sorted(acl_history.items())
        }
    if acl_block:
        cfg["acl"] = acl_block

    if state.get("include_other_allowance"):
        oac = []
        for r in (state.get("other_allowance_considerations") or []):
            try:
                bal = float(r.get("balance") or 0)
                pct = float(r.get("percentage") or 0)
            except (TypeError, ValueError):
                continue
            oac.append({
                "title": str(r.get("title") or "").strip() or "(untitled)",
                "balance": bal,
                "percentage": pct,
                "amount": round(bal * pct / 100.0, 2),
            })
        if oac:
            cfg["other_allowance_considerations"] = oac

    if state.get("data_directory"):
        cfg["data_directory"] = state["data_directory"]

    # Balance-title -> pool mapping (Step 3 in the WARM path). Persist only
    # the non-ignored entries so the YAML stays clean; an empty/missing
    # value at runtime means "row was intentionally excluded".
    bt_map = state.get("balance_title_map") or {}
    bt_clean = {
        str(k).strip(): str(v).strip()
        for k, v in bt_map.items()
        if str(k).strip() and str(v).strip()
    }
    if bt_clean:
        cfg["balance_title_map"] = bt_clean

    # Charge-off / Recovery file column mapping (Step "co_recov" in the
    # wizard). The report engine reads these out of
    # ``historical_file_formats`` so it knows which columns hold the
    # account, code, amount and date in each CU's quarterly file.
    hff: dict[str, Any] = {}
    for src_key, dst_key in (("co_columns", "chargeoff"),
                              ("recov_columns", "recovery")):
        src = state.get(src_key) or {}
        if src.get("code_col") in (None, "") or src.get("amount_col") in (None, ""):
            continue
        block: dict[str, Any] = {
            "has_header": bool(src.get("has_header")),
            "skip_rows": int(src.get("skip_rows") or 0),
            "account_col": int(src.get("account_col") or 0),
            "code_col": int(src["code_col"]),
            "amount_col": int(src["amount_col"]),
        }
        if src.get("date_col") not in (None, ""):
            block["date_col"] = int(src["date_col"])
        if src.get("member_col") not in (None, ""):
            block["member_col"] = int(src["member_col"])
        ma_src = src.get("member_account") or {}
        if ma_src:
            block["member_account"] = {
                "mode": ma_src.get("mode") or "split",
                "suffix_length": int(ma_src.get("suffix_length") or 3),
                "delimiter": ma_src.get("delimiter") or "-",
            }
        hff[dst_key] = block
    if hff:
        existing = cfg.get("historical_file_formats") or {}
        existing.update(hff)
        cfg["historical_file_formats"] = existing

    # Impaired-loans configuration (Step "impaired"). Persist the
    # editable impairment-type / provision-percentage list and the
    # DQ-range table so the report engine can apply per-type provisions
    # without having to re-open the upload.
    imp = state.get("impaired") or {}
    imp_types = imp.get("types") or []
    if imp_types:
        block_imp: dict[str, Any] = {
            "types": [
                {
                    "name": str(t.get("name") or "").strip(),
                    "provision_pct": t.get("provision_pct"),
                }
                for t in imp_types
                if (t.get("name") or "").strip()
            ],
        }
        if imp.get("dq_ranges"):
            block_imp["dq_ranges"] = [
                {
                    "label": str(r.get("label") or "").strip(),
                    "min_days": r.get("min_days"),
                    "provision_pct": r.get("provision_pct"),
                }
                for r in imp["dq_ranges"]
                if (r.get("label") or "").strip()
            ]
        if imp.get("period_ending"):
            block_imp["period_ending"] = imp["period_ending"]

        # Persist the resolved per-loan rows so the report engine can
        # consume them without re-running the wizard's parser/lookup.
        # We only emit rows that have a non-zero balance_removed AND a
        # resolved loan_pool (so the engine knows where to attribute
        # the specific-ID balance). The wizard's impaired_parser
        # populates balance_removed (= current_balance), loan_pool, and
        # credit_grade via lookup_from_loan_data; provision_amount /
        # pct_at_risk come from compute_calculations.
        data_rows_raw = imp.get("data_rows") or []
        data_rows_out: list[dict[str, Any]] = []
        for r in data_rows_raw:
            if not isinstance(r, dict):
                continue
            bal_rem = r.get("balance_removed")
            try:
                bal_rem_f = float(bal_rem) if bal_rem is not None else 0.0
            except (TypeError, ValueError):
                bal_rem_f = 0.0
            pool = (r.get("loan_pool") or "").strip()
            if bal_rem_f <= 0 or not pool:
                continue
            data_rows_out.append({
                "impairment_type": (r.get("impairment_type") or "").strip(),
                "member": (str(r.get("member") or "")).strip(),
                "suffix": (str(r.get("suffix") or "")).strip(),
                "loan_pool": pool,
                "credit_grade": (r.get("credit_grade") or "").strip(),
                "balance_removed": bal_rem_f,
                "provision_amount": (None if r.get("provision_amount") is None
                                     else float(r["provision_amount"])),
            })
        if data_rows_out:
            block_imp["data_rows"] = data_rows_out
        cfg["impaired_loans"] = block_imp

    # Original FICO baseline (one-time upload step). Stores a flat list of
    # (member, suffix, score) rows that the importer consults to fill in
    # ``original_fico_score`` for loans whose monthly extract is missing
    # that field (typical for VISA / credit-card extracts).
    osb = state.get("orig_score_baseline") or {}
    osb_rows = osb.get("rows") or []
    if osb_rows:
        cleaned: list[dict[str, Any]] = []
        for r in osb_rows:
            if not isinstance(r, dict):
                continue
            mem = str(r.get("member") or "").strip()
            if not mem:
                continue
            try:
                score = int(r.get("score") or 0)
            except (TypeError, ValueError):
                continue
            if score <= 0:
                continue
            cleaned.append({
                "member": mem,
                "suffix": str(r.get("suffix") or "").strip(),
                "score": score,
            })
        if cleaned:
            block_osb = {
                "source": osb.get("uploaded_filename") or "",
                "rows": cleaned,
            }
            scoped_pools = [
                str(p).strip() for p in (osb.get("pools") or [])
                if str(p).strip()
            ]
            if scoped_pools:
                block_osb["pools"] = scoped_pools
            cfg["original_fico_baseline"] = block_osb

    # Monthly Balance file (Step 5). Three delivery modes:
    #   single    — one wide quarterly workbook
    #   per_month — one balance-sheet style file per month
    #   manual    — pool × month grid entered by hand
    mb = mb_state  # already resolved above for ACL
    if mb:
        mb_source = mb.get("source") or "single"
        mb_block: dict[str, Any] = {"source": mb_source}

        if mb_source == "single":
            for src_key, dst_key in (
                ("sheet", "sheet"),
                ("header_row", "header_row"),
                ("pool_name_col", "pool_name_col"),
                ("first_date_col", "first_date_col"),
                ("file_pattern", "file_pattern"),
                ("filename", "filename"),
                ("saved_path", "saved_path"),
            ):
                v = mb.get(src_key)
                if v not in (None, "", 0):
                    mb_block[dst_key] = v
            pool_map = {
                str(k).strip(): str(v).strip()
                for k, v in (mb.get("pool_map") or {}).items()
                if str(k).strip() and str(v).strip()
            }
            if pool_map:
                mb_block["pool_map"] = pool_map

        elif mb_source == "per_month":
            layout = mb.get("per_month_layout") or {}
            mb_block["layout"] = {
                "sheet": layout.get("sheet") or "",
                "header_row": int(layout.get("header_row") or 1),
                "label_col": (layout.get("label_col") or "A").upper(),
                "balance_col": (layout.get("balance_col") or "B").upper(),
            }
            files_out: list[dict[str, Any]] = []
            for entry in (mb.get("monthly_files") or []):
                period = (entry.get("period") or "").strip()
                fn = (entry.get("filename") or "").strip()
                sp = (entry.get("saved_path") or "").strip()
                if not period or not fn:
                    continue
                files_out.append({
                    "period": period,
                    "filename": fn,
                    "saved_path": sp,
                })
            files_out.sort(key=lambda e: e.get("period") or "")
            if files_out:
                mb_block["files"] = files_out
            if mb.get("file_pattern"):
                mb_block["file_pattern"] = mb["file_pattern"]
            pool_map = {
                str(k).strip(): str(v).strip()
                for k, v in (mb.get("pool_map") or {}).items()
                if str(k).strip() and str(v).strip()
            }
            if pool_map:
                mb_block["pool_map"] = pool_map

        elif mb_source == "manual":
            months = [
                (m or "").strip()
                for m in (mb.get("manual_months") or [])
                if (m or "").strip()
            ]
            if months:
                mb_block["months"] = months
            entries_clean: dict[str, dict[str, float]] = {}
            for pool, row in (mb.get("manual_entries") or {}).items():
                if not pool or not isinstance(row, dict):
                    continue
                clean_row: dict[str, float] = {}
                for d, v in row.items():
                    if not d:
                        continue
                    try:
                        clean_row[str(d)] = float(v)
                    except (TypeError, ValueError):
                        continue
                if clean_row:
                    entries_clean[str(pool)] = clean_row
            if entries_clean:
                mb_block["entries"] = entries_clean

        if mb.get("notes"):
            mb_block["notes"] = mb["notes"]
        cfg["monthly_balance"] = mb_block

    return cfg


def get_balance_title_map(cfg: dict[str, Any]) -> dict[str, str]:
    """Return the saved balance-title → pool-name mapping for ``cfg``.

    Each key is a CU-supplied balance title that appears in the monthly
    balance-sheet feed (e.g. ``"New Autos"``) and each value is one of
    the canonical loan-pool names declared in the ``pools`` block.
    Titles the user marked « ignore » at wizard time are intentionally
    NOT persisted, so any title not in this dict should be skipped by
    the historical-balance ingestion.
    """
    raw = cfg.get("balance_title_map") or {}
    if not isinstance(raw, dict):
        return {}
    out: dict[str, str] = {}
    for k, v in raw.items():
        ks = str(k).strip()
        vs = str(v).strip()
        if ks and vs:
            out[ks] = vs
    return out
