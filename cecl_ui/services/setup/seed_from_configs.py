"""Seed the setup pattern registry from the existing production configs
(Setup Router, Phase 0 completion).

Reconstructs, for each of the ~38 client configs, the setup decisions that
were made and records them as ``source='mined', trust='trusted'`` patterns
(they're proven in production). Aggregated by support so a decision common to
K configs is one pattern with ``support=K``.

Decision kinds seeded:
  * member_account        — signal ``loan_suffix_present`` -> mode (the 10/10 rule)
  * pool_code_column      — library of column NAMES that mean "loan pool code"
  * pool_map_code         — code -> pool priors (bootstrap for new clients)
  * credit_pull_keys      — member/score column pairs
  * score_source          — original/current FICO column names
  * monthly_balance_source— single | per_month | per_year | manual archetype
  * participation_extract — static-pool second-extract templates
"""

from __future__ import annotations

import glob
import os
import re
from collections import Counter
from typing import Any

import yaml
from sqlalchemy.engine import Engine

from . import setup_registry as reg


def _norm_name(s: Any) -> str:
    s = str(s or "").strip().lower()
    s = re.sub(r"[_\-]+", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def _norm_code(x: Any) -> str:
    x = str(x).strip()
    if x.replace(".", "", 1).isdigit():
        try:
            return str(int(float(x)))
        except (TypeError, ValueError):
            return x
    return x


def _iter_configs(configs_dir: str):
    for fp in glob.glob(os.path.join(configs_dir, "*.yaml")):
        if ".bak" in os.path.basename(fp):
            continue
        try:
            cfg = yaml.safe_load(open(fp, encoding="utf-8"))
        except Exception:
            continue
        if isinstance(cfg, dict):
            yield os.path.splitext(os.path.basename(fp))[0], cfg


def _extract_scopes(cfg: dict) -> list[dict]:
    """Return the mapping scopes to mine: each loan_data_extracts entry, else
    the top-level config."""
    exts = cfg.get("loan_data_extracts") or []
    scopes = []
    for e in exts:
        scopes.append({
            "column_mappings": e.get("column_mappings") or {},
            "member_account": e.get("member_account") or cfg.get("member_account") or {},
            "row_filter": e.get("row_filter"),
            "header_row": e.get("header_row"),
        })
    if not exts:
        scopes.append({
            "column_mappings": cfg.get("column_mappings") or {},
            "member_account": cfg.get("member_account") or {},
            "row_filter": None, "header_row": None,
        })
    return scopes


def seed(engine: Engine, configs_dir: str) -> dict[str, int]:
    reg.ensure_tables(engine)
    counts: Counter = Counter()

    for short, cfg in _iter_configs(configs_dir):
        # ---- member_account (loan_suffix_present -> mode) ----------------
        for sc in _extract_scopes(cfg):
            cm = sc["column_mappings"]
            ma = sc["member_account"]
            mode = (ma or {}).get("mode")
            if not mode:
                continue
            has_suffix = bool(cm.get("loan_suffix"))
            reg.save_pattern(
                engine, "member_account",
                {"loan_suffix_present": has_suffix},
                {"mode": mode},
                evidence={"suffix_lengths": [(ma or {}).get("suffix_length")]},
                cu=short)
            counts["member_account"] += 1

            # ---- pool_code_column name library ---------------------------
            if cm.get("loan_pool_code"):
                col = cm["loan_pool_code"]
                reg.save_pattern(
                    engine, "pool_code_column",
                    {"normalized": _norm_name(col)},
                    {"role": "loan_pool_code", "column": col},
                    cu=short)
                counts["pool_code_column"] += 1

            # ---- score sources -------------------------------------------
            for role, key in (("original", "original_fico_score"),
                              ("current", "current_fico_score")):
                if cm.get(key):
                    reg.save_pattern(
                        engine, "score_source",
                        {"role": role, "normalized": _norm_name(cm[key])},
                        {"role": role, "column": cm[key]}, cu=short)
                    counts["score_source"] += 1

            # ---- participation extract template --------------------------
            if cm.get("loan_pool_code_static"):
                reg.save_pattern(
                    engine, "participation_extract",
                    {"static_pool": cm["loan_pool_code_static"]},
                    {"static_pool": cm["loan_pool_code_static"],
                     "balance_column": cm.get("current_balance"),
                     "original_score": cm.get("original_fico_score"),
                     "current_score": cm.get("current_fico_score"),
                     "header_row": sc["header_row"],
                     "row_filter": sc["row_filter"]},
                    cu=short)
                counts["participation_extract"] += 1

        # ---- pool_map code -> pool priors --------------------------------
        for code, pool in (cfg.get("pool_map") or {}).items():
            reg.save_pattern(
                engine, "pool_map_code",
                {"code": _norm_code(code)},
                {"pool": str(pool)}, cu=short)
            counts["pool_map_code"] += 1

        # ---- credit_pull keys --------------------------------------------
        cp = cfg.get("credit_pull") or {}
        if cp.get("member_column") or cp.get("score_column"):
            reg.save_pattern(
                engine, "credit_pull_keys", {},
                {"member_column": cp.get("member_column"),
                 "score_column": cp.get("score_column")}, cu=short)
            counts["credit_pull_keys"] += 1

        # ---- monthly balance archetype -----------------------------------
        mb = cfg.get("monthly_balance") or {}
        if isinstance(mb, dict) and mb.get("source"):
            reg.save_pattern(
                engine, "monthly_balance_source", {},
                {"source": mb["source"]}, cu=short)
            counts["monthly_balance_source"] += 1

    return dict(counts)
