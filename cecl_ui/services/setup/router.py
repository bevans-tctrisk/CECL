"""Setup Router orchestrator (Setup Router, capstone).

Ties the pieces together: profile a client's delivered files, propose each
setup decision from the learned registry + detectors, assemble a candidate
config, and validate it with the ground-truth probe. Emits a
:class:`SetupProposal` — per-decision value + tier + evidence — for human
approval. Nothing is written to a live config; the router only proposes.

Modes:
- Re-onboard / audit: pass an existing ``base_cfg`` (supplies the client
  givens — folders, pool_map, monthly_balance, credit_pull); the router
  re-derives the STRUCTURAL decisions (member_account, pool_code,
  participation) from the files + registry and validates. Also catches drift
  when a CU changes file layouts.
- New client: ``base_cfg`` minimal / no pool_map -> the router proposes
  candidates and flags the pool_map bootstrap (key off code descriptions).
"""

from __future__ import annotations

import copy
import glob
import os
from dataclasses import dataclass, field
from typing import Any

import pandas as pd
from sqlalchemy.engine import Engine

from . import profiler as prof
from . import probe_validator as pvv
from . import resolvers as rsv


@dataclass
class DecisionProposal:
    kind: str
    value: Any
    tier: str                       # green | yellow | red
    source: str                     # registry | detector | profile | needs_input
    evidence: str = ""
    alternatives: list[Any] = field(default_factory=list)


@dataclass
class SetupProposal:
    short: str
    files: dict[str, str] = field(default_factory=dict)     # archetype -> path
    decisions: list[DecisionProposal] = field(default_factory=list)
    config_draft: dict = field(default_factory=dict)
    validation: pvv.ValidationReport | None = None
    overall_tier: str = "red"
    messages: list[str] = field(default_factory=list)


_TIER_RANK = {"green": 2, "yellow": 1, "red": 0}


class SetupRouter:
    def __init__(self, engine: Engine):
        self.engine = engine

    # -- file discovery / profiling ------------------------------------
    def _profile_delivery(self, folder: str):
        """Return (archetype->FileProfile [biggest per archetype], all_profiles)."""
        out: dict[str, prof.FileProfile] = {}
        allp: list[prof.FileProfile] = []

        def _size(fp: prof.FileProfile) -> int:
            return max((sp.n_rows * sp.n_cols for sp in fp.sheets), default=0)

        for path in sorted(glob.glob(os.path.join(folder, "*.xls*"))):
            if os.path.basename(path).startswith("~$"):
                continue
            try:
                fp = prof.profile_file(path)
            except Exception:
                continue
            allp.append(fp)
            a = fp.archetype
            # Keep the biggest file per archetype: when several files share an
            # archetype (esp. the catch-all 'unknown'), the real extract dwarfs
            # any small report/history that lands in the same bucket.
            if a not in out or _size(fp) > _size(out[a]):
                out[a] = fp
        return out, allp

    def _find_loan_profile(self, profiles: dict[str, prof.FileProfile],
                           balance_col: str | None):
        """Return the loan-extract profile. Falls back to the largest profiled
        file whose columns include the known balance column when the profiler's
        keyword archetype test misses it (e.g. Symitar LND* headers)."""
        lp = profiles.get("loan_extract")
        if lp is not None:
            return lp
        if not balance_col:
            return None
        best = None
        for fp in profiles.values():
            for sp in fp.sheets:
                names = {str(c.name).strip() for c in sp.columns}
                if balance_col in names and sp.n_rows >= 50:
                    if best is None or sp.n_rows > best[0]:
                        best = (sp.n_rows, fp)
        return best[1] if best else None

    # -- orchestration --------------------------------------------------
    def propose(self, *, short: str, delivery_folder: str, snapshot: str,
                base_cfg: dict, validate: bool = True,
                warm: dict | None = None,
                warm_folder: str | None = None) -> SetupProposal:
        proposal = SetupProposal(short=short)
        profiles, all_profiles = self._profile_delivery(delivery_folder)
        proposal.files = {a: p.path for a, p in profiles.items()}

        # Balance column comes from the WARM (Symitar LNDCBAL etc.) or an
        # existing config — used both to drive the DAG and to recognize the
        # loan extract when the profiler's keyword test doesn't (Symitar LND*
        # headers don't match 'balance'/'account').
        main_ext = (base_cfg.get("loan_data_extracts") or [{}])[0]
        main_cm = main_ext.get("column_mappings") or base_cfg.get("column_mappings") or {}
        balance_col = main_cm.get("current_balance")
        if not balance_col and warm:
            balance_col = (warm.get("column_mappings") or {}).get("current_balance")
        if not balance_col:
            balance_col = "CURRENT_BALANCE"

        loan_prof = self._find_loan_profile(profiles, balance_col)
        if loan_prof is None:
            proposal.messages.append("No loan extract detected in delivery folder.")
            return proposal
        # Register the resolved extract canonically so every specialist finds it
        # (the profiler may have bucketed a non-keyword extract as 'unknown').
        profiles["loan_extract"] = loan_prof
        proposal.files["loan_extract"] = loan_prof.path
        header_row = loan_prof.sheets[0].header_row

        # monthly_by_pool for the snapshot: main (non-NRR, non-static) pools.
        # Use the engine's universal loader so ALL monthly formats work
        # (single / per_month / per_year / manual) — not just Cottonwood's
        # manual grid. Falls back to the raw manual entries if that fails.
        pool_map = base_cfg.get("pool_map")
        nrr = {str(p).strip() for p in (base_cfg.get("not_risk_rated") or [])}
        static_pools = {
            (e.get("column_mappings") or {}).get("loan_pool_code_static")
            for e in (base_cfg.get("loan_data_extracts") or [])
        } - {None}
        monthly_by_pool = None
        try:
            import generate_report as _gr
            bdf, _alll = _gr.load_monthly_balances(base_cfg)
            if bdf is not None and not bdf.empty:
                snap_ts = pd.Timestamp(snapshot)
                sub = bdf[bdf["date"] == snap_ts]
                m = {}
                for _, r in sub.iterrows():
                    pool = str(r["pool"]).strip()
                    if pool in nrr or pool in static_pools:
                        continue
                    m[pool] = float(r["balance"])
                if m:
                    monthly_by_pool = m
        except Exception:
            monthly_by_pool = None
        if not monthly_by_pool:
            entries = ((base_cfg.get("monthly_balance") or {}).get("entries")) or {}
            if entries:
                monthly_by_pool = {
                    p: v.get(snapshot) for p, v in entries.items()
                    if p not in nrr and p not in static_pools
                    and isinstance(v, dict) and v.get(snapshot) is not None
                }

        # --- run the specialist DAG ---
        # Each decision is owned by a specialist resolver (member_account,
        # pool_code, pool_map, score_source, monthly_balance,
        # secondary_extracts, reserve_config), sequenced by dependency. A
        # parsed WARM (when supplied) is a high-trust source the specialists
        # read from; otherwise they fall back to detectors + the registry.
        ctx = rsv.ResolverContext(
            engine=self.engine, short=short, delivery_folder=delivery_folder,
            profiles=profiles, all_profiles=all_profiles, base_cfg=base_cfg,
            monthly_by_pool=monthly_by_pool, warm=warm, warm_folder=warm_folder,
            snapshot_date=snapshot, loan_path=loan_prof.path,
            header_row=header_row, balance_col=balance_col)
        results = rsv.run_dag(ctx, rsv.build_default_resolvers())
        for r in results.values():
            proposal.decisions.append(DecisionProposal(
                kind=r.kind, value=r.value, tier=r.tier, source=r.source,
                evidence=r.evidence, alternatives=list(r.alternatives)))

        # --- assemble draft: overlay resolved decisions onto base_cfg ---
        draft = self._assemble_draft(base_cfg, results)
        proposal.config_draft = draft

        # --- validate via probe (needs a pool_map + book to tie against) ---
        if validate and monthly_by_pool and draft.get("pool_map"):
            proposal.validation = pvv.probe_validate(
                self.engine, draft, short=short, snapshot=snapshot,
                scan_folder=delivery_folder)

        # --- overall tier ---
        base_tier = _TIER_RANK[rsv.overall_tier(results)]
        if proposal.validation is not None:
            base_tier = min(base_tier, _TIER_RANK[proposal.validation.verdict])
        proposal.overall_tier = {2: "green", 1: "yellow", 0: "red"}[base_tier]
        return proposal

    # -- draft assembly ------------------------------------------------
    def _assemble_draft(self, base_cfg: dict,
                        results: dict[str, "rsv.Resolution"]) -> dict:
        """Overlay the resolved decisions onto a copy of ``base_cfg``."""
        draft = copy.deepcopy(base_cfg)
        draft_ext = (draft.get("loan_data_extracts") or [None])[0]
        target_cm = (draft_ext or draft).setdefault("column_mappings", {})

        pc = results.get("pool_code")
        if pc and pc.value:
            col = pc.value.get("column") if isinstance(pc.value, dict) else pc.value
            if col:
                target_cm["loan_pool_code"] = col

        pm = results.get("pool_map")
        if pm and pm.value:
            mapping = (pm.value.get("pool_map")
                       if isinstance(pm.value, dict) else pm.value)
            if mapping:
                draft["pool_map"] = mapping

        dm = results.get("member_account")
        if dm and dm.value:
            ma = {"mode": dm.value["mode"]}
            if dm.value["mode"] == "split":
                ma["suffix_length"] = dm.value.get("suffix_length", 3)
                if dm.value.get("loan_suffix_col"):
                    target_cm["loan_suffix"] = dm.value["loan_suffix_col"]
            elif dm.value.get("suffix_length") is not None:
                ma["suffix_length"] = dm.value["suffix_length"]
            if draft_ext:
                draft_ext["member_account"] = ma
            draft["member_account"] = ma

        rc = results.get("reserve_config")
        if rc and isinstance(rc.value, dict):
            for k in ("not_risk_rated", "warm_allowance_pools",
                      "base_loss_rate_by_pool_grade"):
                if rc.value.get(k):
                    draft[k] = rc.value[k]

        # Score sources: original/current FICO columns + an optional SSN-style
        # credit-pull join key (loans whose pull is keyed by SSN join on their
        # SSN column instead of the member number).
        ss = results.get("score_source")
        if ss and isinstance(ss.value, dict):
            if ss.value.get("original_fico_score"):
                target_cm["original_fico_score"] = ss.value["original_fico_score"]
            if ss.value.get("current_fico_score"):
                target_cm["current_fico_score"] = ss.value["current_fico_score"]
            if ss.value.get("loan_join_column"):
                cp = draft.setdefault("credit_pull", {})
                cp["loan_join_column"] = ss.value["loan_join_column"]
                if ss.value.get("pull_member_column"):
                    cp["member_column"] = ss.value["pull_member_column"]

        # Secondary extracts (credit-card / participation files) become
        # additional loan_data_extracts with a static pool code.
        se = results.get("secondary_extracts")
        if se and isinstance(se.value, dict):
            for e in (se.value.get("extracts") or []):
                if not e.get("path") or not e.get("balance_column"):
                    continue
                cm = {"current_balance": e["balance_column"]}
                if e.get("member_column"):
                    cm["member_number"] = e["member_column"]
                if e.get("available_credit_column"):
                    cm["total_available_credit"] = e["available_credit_column"]
                if e.get("static_pool"):
                    cm["loan_pool_code_static"] = e["static_pool"]
                if e.get("original_score"):
                    cm["original_fico_score"] = e["original_score"]
                if e.get("current_score"):
                    cm["current_fico_score"] = e["current_score"]
                entry = {
                    "label": e.get("label", "Secondary"),
                    "column_mappings": cm,
                    "member_account": {"mode": "fixed_suffix", "suffix_length": 0},
                    "has_header": True,
                }
                if e.get("file_pattern"):
                    entry["file_pattern"] = e["file_pattern"]
                # Secondary extracts (esp. credit cards) are scored by the
                # analyst from a source outside the delivery, so a CU-level
                # SSN pull-join is the wrong proxy for them. Disable the join
                # for this file; it falls back to member/WARM/previous-snapshot.
                if draft.get("credit_pull", {}).get("loan_join_column"):
                    entry["credit_pull_join_column"] = ""
                draft.setdefault("loan_data_extracts", []).append(entry)
        return draft
