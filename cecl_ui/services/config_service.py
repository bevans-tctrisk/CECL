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
            "brr": False,
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
        # ``brr`` is True when the pool is broken out by Business Risk
        # Ratings instead of Credit Grade bands. Only meaningful when
        # ``risk_rated`` is True. Defaults to False so legacy configs
        # (which don't carry this field) keep using credit grades.
        "brr": bool(p.get("brr", False)) and bool(p.get("risk_rated", True))
                 and not bool(p.get("excluded", False)),
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
    # Report period (YYYY-MM) — used as a snapshot fallback when a
    # loan-data filename has a month name but no year.
    if state.get("report_period"):
        cfg["report_period"] = str(state["report_period"]).strip()
    cfg.update({
        "file_pattern": state["file_pattern"],
        "date_pattern": state["date_pattern"],
        "date_format": state.get("date_format") or "YYYY-MM",
        "account_suffix_length": int(state.get("account_suffix_length", 3)),
        "member_account": dict(state.get("member_account") or {
            "mode": "fixed_suffix",
            "suffix_length": int(state.get("account_suffix_length", 3)),
            "delimiter": "-",
        }),
        "has_header": bool(state.get("has_header", True)),
        "column_mappings": state["column_mappings"],
        "credit_pull": state.get("credit_pull") or {},
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
        "mgmt_adj": state.get("mgmt_adj") or {"ltv_baseline": 0.9, "probability_factor": 0.35},
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
        # Per-file header_row (1-indexed). Some AIRES-style extracts use
        # row 1 as column position numbers and row 2 as the real headers;
        # the wizard's Sample step stores the right value per file. Only
        # emit when > 1 so the YAML stays clean for the common case.
        try:
            hr_lf = int(lf.get("header_row") or 0)
        except (TypeError, ValueError):
            hr_lf = 0
        if hr_lf > 1:
            entry["header_row"] = hr_lf
        # Phase 9.22: per-extract ``pool_code_split`` override. When the
        # wizard's Sample step set this on the entry, faithfully emit it
        # (including the empty string, which signals "do not split" — used
        # for CUMA mortgage files whose loan codes legitimately contain
        # ``/``, e.g. ``15/15 ARM``). Absent key inherits the CU-level
        # ``pool_code_split``.
        if "pool_code_split" in lf:
            entry["pool_code_split"] = lf.get("pool_code_split") or ""
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
            "brr": bool(p.get("brr")),
        })
    if pools_block:
        set_pools(cfg, pools_block)

    # Business Risk Ratings (optional CU-wide registry — only emitted
    # when the user enabled BRR on the Credit Grades step). Mirrors the
    # wizard's ``state.uses_brr`` flag plus the ``business_risk_ratings``
    # list of {label, criteria} rows. Downstream the report engine looks
    # at each pool's ``brr`` flag to decide whether to bucket loans by
    # credit grade or BRR label; the loan-extract column that carries
    # the raw rating value rides in ``column_mappings.business_risk_rating``
    # (per extract) as already mapped on the Column Mappings step.
    brr_rows_raw = state.get("business_risk_ratings") or []
    brr_rows: list[dict[str, Any]] = []
    for r in brr_rows_raw:
        lbl = str((r or {}).get("label") or "").strip()
        crit = str((r or {}).get("criteria") or "").strip()
        if not lbl and not crit:
            continue
        brr_rows.append({"label": lbl, "criteria": crit})
    if state.get("uses_brr") or brr_rows:
        cfg["uses_brr"] = bool(state.get("uses_brr"))
        if brr_rows:
            cfg["business_risk_ratings"] = brr_rows

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
    # NCUA 5300 fallback flag — emitted unconditionally so the report
    # engine has a clear True/False toggle. Default True for any wizard
    # state that didn't explicitly set it (matches the new-CU default).
    if acl_state:
        acl_block["use_5300_fallback"] = bool(
            acl_state.get("use_5300_fallback", True)
        )
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

    # Data-derived "Negative Share Provision" OAC. Emitted as an
    # other_allowance_considerations entry with source: negative_share so the
    # report engine recomputes balance / rate / amount each quarter.
    if state.get("include_negative_share"):
        ns = state.get("negative_share") or {}
        try:
            life_months = int(ns.get("life_of_loan_months") or 12)
        except (TypeError, ValueError):
            life_months = 12
        ns_entry = {
            "title": str(ns.get("title") or "").strip() or "Negative Share Provision",
            "source": "negative_share",
            "life_of_loan_months": life_months,
            "source_folder": str(ns.get("source_folder") or "").strip(),
            "balance_column": str(ns.get("balance_column") or "").strip()
            or "Current Balance",
            "balance_pattern": str(ns.get("balance_pattern") or "").strip()
            or r"(?i)Negative Share File",
            "co_summary_pattern": str(ns.get("co_summary_pattern") or "").strip()
            or r"(?i)Negative Shares Charge Off and Recovery",
            "co_quarterly_pattern": str(ns.get("co_quarterly_pattern") or "").strip()
            or r"(?i)Share COs?\s*-\s*Recoveries",
            "percentage": 0.0,
            "balance": 0.0,
            "amount": 0.0,
        }
        cfg.setdefault("other_allowance_considerations", []).append(ns_entry)

    # Data-derived "Unfunded Commitments" OAC. Emitted with
    # source: unfunded_commitment; the report engine sums undrawn credit for
    # the configured codes grouped by pool and applies each pool's ACL rate.
    if state.get("include_unfunded_commitment"):
        uc = state.get("unfunded_commitment") or {}
        codes = [str(c).strip() for c in (uc.get("loan_type_codes") or [])
                 if str(c).strip()]
        if codes:
            uc_entry = {
                "title": str(uc.get("title") or "").strip() or "Unfunded Commitments",
                "source": "unfunded_commitment",
                "loan_type_codes": codes,
                "percentage": 0.0,
                "balance": 0.0,
                "amount": 0.0,
            }
            cfg.setdefault("other_allowance_considerations", []).append(uc_entry)

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
            # Honour explicit suffix_length=0 (CUs whose member-number
            # column already holds the full account, no separate suffix).
            _sl_src = ma_src.get("suffix_length")
            try:
                _sl_int = int(_sl_src) if _sl_src is not None else 3
            except (TypeError, ValueError):
                _sl_int = 3
            block["member_account"] = {
                "mode": ma_src.get("mode") or "split",
                "suffix_length": _sl_int,
                "delimiter": ma_src.get("delimiter") or "-",
            }
        hff[dst_key] = block
    if hff:
        existing = cfg.get("historical_file_formats") or {}
        existing.update(hff)
        cfg["historical_file_formats"] = existing

    # Multi-format CO/Recovery: an optional list of named formats, each
    # routing files by ``file_pattern`` to its own chargeoff/recovery column
    # wiring. Lets one CU mix several CO/recovery layouts (e.g. consumer-loan,
    # credit-card and overdraft files whose columns differ). The report engine
    # reads ``historical_file_formats.formats``; when present it takes
    # precedence over the single top-level chargeoff/recovery block. Each
    # per-side block carries ``strict_columns: true`` so the engine trusts the
    # explicit indices (a combined file's CO and recovery amount columns
    # differ, which the header-text heuristic cannot disambiguate). A side may
    # use ``code_static`` instead of ``code_col`` when the file has no loan-
    # code column (every row is implicitly one pool, e.g. a credit-card file).
    formats_state = state.get("co_recov_formats") or []
    formats_out: list[dict[str, Any]] = []
    for fmt in formats_state:
        fp = str((fmt or {}).get("file_pattern") or "").strip()
        if not fp:
            continue
        entry: dict[str, Any] = {
            "name": str(fmt.get("name") or "").strip() or "(unnamed)",
            "file_pattern": fp,
        }
        for src_key, dst_key in (("co_columns", "chargeoff"),
                                  ("recov_columns", "recovery")):
            src = fmt.get(src_key) or {}
            has_amount = src.get("amount_col") not in (None, "")
            has_code = (src.get("code_col") not in (None, "")
                        or str(src.get("code_static") or "").strip())
            if not (has_amount and has_code):
                continue
            block2: dict[str, Any] = {
                "has_header": bool(src.get("has_header")),
                "skip_rows": int(src.get("skip_rows") or 0),
                "account_col": int(src.get("account_col") or 0),
                "amount_col": int(src["amount_col"]),
                "strict_columns": True,
            }
            if src.get("code_col") not in (None, ""):
                block2["code_col"] = int(src["code_col"])
            if str(src.get("code_static") or "").strip():
                block2["code_static"] = str(src["code_static"]).strip()
            if src.get("date_col") not in (None, ""):
                block2["date_col"] = int(src["date_col"])
            entry[dst_key] = block2
        if entry.get("chargeoff") or entry.get("recovery"):
            formats_out.append(entry)
    if formats_out:
        existing = cfg.get("historical_file_formats") or {}
        existing["formats"] = formats_out
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

            # Pass through the supplemental monthly-balance mechanisms
            # verbatim so advanced wiring survives wizard regeneration.
            # A CU that starts delivering a fresh rolling workbook
            # (``supplemental_wide``) or a per-month / formatted balance-
            # sheet snapshot (``monthly_file_pattern`` family) alongside
            # the historical workbook keeps that config across reruns
            # instead of silently losing the recent month's balances.
            sw = mb.get("supplemental_wide")
            if sw:
                mb_block["supplemental_wide"] = sw
            for _k in ("monthly_file_pattern", "monthly_sheet",
                       "monthly_label_col", "monthly_balance_col",
                       "monthly_header_row", "monthly_start_marker"):
                _v = mb.get(_k)
                if _v not in (None, "", 0):
                    mb_block[_k] = _v
            if mb.get("monthly_strict_pool_map"):
                mb_block["monthly_strict_pool_map"] = True

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

        elif mb_source == "per_year":
            layout = mb.get("per_year_layout") or {}
            mb_block["layout"] = {
                "sheet": layout.get("sheet") or "",
                "header_row": int(layout.get("header_row") or 1),
                "label_col": (layout.get("label_col") or "B").upper(),
                "period_columns": [
                    {
                        "col": str(pc.get("col") or "").upper(),
                        "period": str(pc.get("period") or ""),
                    }
                    for pc in (layout.get("period_columns") or [])
                    if pc.get("col") and pc.get("period")
                ],
            }
            files_out: list[dict[str, Any]] = []
            for entry in (mb.get("year_files") or []):
                fn = (entry.get("filename") or "").strip()
                sp = (entry.get("saved_path") or "").strip()
                yr = entry.get("year")
                try:
                    yr_int = int(yr) if yr is not None else None
                except (TypeError, ValueError):
                    yr_int = None
                if not fn or yr_int is None:
                    continue
                files_out.append({
                    "year": yr_int,
                    "filename": fn,
                    "saved_path": sp,
                })
            files_out.sort(key=lambda e: e.get("year") or 0)
            if files_out:
                mb_block["files"] = files_out
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

    # Historical-balance provenance — a display-only record of HOW the monthly
    # historical balances were compiled, so a reviewer of the saved report can
    # see and validate the source (e.g. "derived from the imported loan
    # extracts" vs. an uploaded balance workbook). Persisted verbatim from the
    # wizard state; only emitted when a method was recorded.
    prov = state.get("hist_balance_provenance") or {}
    if isinstance(prov, dict) and prov.get("method"):
        prov_block: dict[str, Any] = {"method": str(prov["method"])}
        for k in ("label", "summary", "generated_by", "validation_hint",
                  "generated_at"):
            if prov.get(k):
                prov_block[k] = str(prov[k])
        inputs = [str(x) for x in (prov.get("inputs") or []) if str(x).strip()]
        if inputs:
            prov_block["inputs"] = inputs
        cov = prov.get("coverage") or {}
        if isinstance(cov, dict) and cov:
            cov_clean: dict[str, Any] = {}
            for ck in ("start", "end"):
                if cov.get(ck):
                    cov_clean[ck] = str(cov[ck])
            for ck in ("months", "pools"):
                try:
                    if cov.get(ck) is not None:
                        cov_clean[ck] = int(cov[ck])
                except (TypeError, ValueError):
                    pass
            if cov_clean:
                prov_block["coverage"] = cov_clean
        cfg["hist_balance_provenance"] = prov_block

    # Charge-off / recovery provenance — display-only record of HOW the
    # historical charge-offs and recoveries were compiled (CU files, NCUA
    # 5300 backfill, or a mix), so the saved report carries a validator-facing
    # trail. Passed through verbatim from wizard state when present.
    crp = state.get("co_recov_provenance") or {}
    if isinstance(crp, dict) and crp:
        crp_block: dict[str, Any] = {}
        for side in ("chargeoff", "recovery"):
            rec = crp.get(side) or {}
            if not isinstance(rec, dict) or not (rec.get("method") or rec.get("no_recoveries")):
                continue
            side_block: dict[str, Any] = {}
            for k in ("method", "label", "summary", "validation_hint"):
                if rec.get(k):
                    side_block[k] = str(rec[k])
            if rec.get("no_recoveries"):
                side_block["no_recoveries"] = True
            cu_files = [
                {"name": str(f.get("name") or ""),
                 "file_pattern": str(f.get("file_pattern") or "")}
                for f in (rec.get("cu_files") or []) if f.get("name")
            ]
            if cu_files:
                side_block["cu_files"] = cu_files
            db_src = []
            for s in (rec.get("db_sources") or []):
                entry = {"label": str(s.get("label") or "")}
                if s.get("years"):
                    entry["years"] = str(s["years"])
                try:
                    if s.get("total") is not None:
                        entry["total"] = round(float(s["total"]), 2)
                except (TypeError, ValueError):
                    pass
                if entry["label"]:
                    db_src.append(entry)
            if db_src:
                side_block["db_sources"] = db_src
            if side_block:
                crp_block[side] = side_block
        if crp_block:
            cfg["co_recov_provenance"] = crp_block

    # ---- WARM auto-derive: reserve pinning + monthly book ----
    # Populated by the setup wizard's WARM upload (the ``warm_autoderive``
    # bridge) with the validated resolver-path derivations that this function
    # does not otherwise produce: col-G per-grade base-loss pins, warm-allowance
    # pools, balance-only ``not_risk_rated`` (detected from the WARM ACL tab, not
    # the coarse risk-rated flag), and the WARM's manual monthly book. These are
    # what make a WARM-sourced config tie to the analyst's WARM. Overlaid only
    # when present, so non-WARM CUs are unaffected.
    wr = state.get("warm_reserve") or {}
    if wr.get("base_loss_rate_by_pool_grade"):
        cfg["base_loss_rate_by_pool_grade"] = wr["base_loss_rate_by_pool_grade"]
    if wr.get("warm_allowance_pools"):
        cfg["warm_allowance_pools"] = wr["warm_allowance_pools"]
    if wr.get("not_risk_rated"):
        cfg["not_risk_rated"] = wr["not_risk_rated"]
    if wr.get("monthly_balance") and not cfg.get("monthly_balance"):
        cfg["monthly_balance"] = wr["monthly_balance"]

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
