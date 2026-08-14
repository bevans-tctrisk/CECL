"""Concrete specialist resolvers for the Setup Router.

Each class owns exactly one ``decision_kind`` and wraps the existing detector /
registry logic behind the uniform :class:`~.base.Resolver` contract. A parsed
WARM (``ctx.warm``) is treated as a high-trust source: when present, a
specialist reads its value (``source='warm'``) but still exposes it for
independent validation by the final probe.

Split further into one-file-per-specialist later; grouped here while the set is
small.
"""

from __future__ import annotations

import os
import re
from typing import Optional

import pandas as pd

from .. import config_assembler as ca
from .. import pool_code_detector as pcd
from .. import pool_map_bootstrap as pmb
from .. import setup_registry as reg
from .base import Resolution, Resolver, ResolverContext


def _loan_profile(ctx: ResolverContext):
    return ctx.profiles.get("loan_extract")


# ---------------------------------------------------------------------------
# Loan-file resolver — establishes loan_path / header_row / balance_col.
# Prerequisite for the pool-code / pool-map specialists.
# ---------------------------------------------------------------------------
class LoanFileResolver(Resolver):
    kind = "loan_file"
    depends_on = ()

    def applies(self, ctx: ResolverContext) -> bool:
        return _loan_profile(ctx) is not None or bool(ctx.loan_path)

    def resolve_local(self, ctx: ResolverContext) -> Optional[Resolution]:
        prof = _loan_profile(ctx)
        if prof is not None:
            ctx.loan_path = getattr(prof, "path", ctx.loan_path)
            sheet = prof.sheets[0] if prof.sheets else None
            ctx.header_row = getattr(sheet, "header_row", 1) or 1
        # balance column: prefer the WARM mapping, else an existing config
        # mapping (re-onboard), else a balance-like header.
        bal = None
        if ctx.warm:
            bal = (ctx.warm.get("column_mappings") or {}).get("current_balance")
        if not bal and ctx.base_cfg:
            main_ext = (ctx.base_cfg.get("loan_data_extracts") or [{}])[0] or {}
            bal = (main_ext.get("column_mappings")
                   or ctx.base_cfg.get("column_mappings") or {}).get("current_balance")
        if not bal and prof is not None and prof.sheets:
            bal = next((c.name for c in prof.sheets[0].columns
                        if re.search(r"balance|\bupb\b|LNDCBAL", c.name, re.I)), None)
        ctx.balance_col = bal
        if not ctx.loan_path:
            return Resolution(self.kind, None, "red", "needs_input",
                              "no loan-extract file profiled")
        tier = "green" if bal else "yellow"
        return Resolution(self.kind,
                          {"path": ctx.loan_path, "header_row": ctx.header_row,
                           "balance_col": bal},
                          tier, "warm" if (ctx.warm and bal) else "profile",
                          f"loan file resolved; balance_col={bal!r}")


# ---------------------------------------------------------------------------
# member_account — fixed_suffix / split / delimiter (+ suffix length)
#   Oracle: member/impaired join rate (deferred to the final probe).
# ---------------------------------------------------------------------------
class MemberAccountResolver(Resolver):
    kind = "member_account"
    depends_on = ()

    def resolve_local(self, ctx: ResolverContext) -> Optional[Resolution]:
        if ctx.warm:
            ma = ctx.warm.get("member_account") or {}
            if ma.get("mode"):
                val = {k: ma[k] for k in ("mode", "suffix_length") if k in ma}
                return Resolution(
                    self.kind, val, "green", "warm",
                    f"WARM Data-tab formula → {ma.get('mode')} "
                    f"(suffix_length {ma.get('suffix_length')})")
        prof = _loan_profile(ctx)
        if prof is None or not prof.sheets:
            return None
        cols = prof.sheets[0].columns
        suffix_col = next(
            (c.name for c in cols if re.search(r"suffix", c.name, re.I)), None)
        has_suffix = suffix_col is not None
        pats = reg.get_patterns(
            ctx.engine, "member_account",
            fingerprint_filter={"loan_suffix_present": has_suffix})
        if pats:
            mode = pats[0]["fragment"]["mode"]
            support = pats[0]["support"]
            val = {"mode": mode}
            if mode == "split" and suffix_col:
                val["loan_suffix_col"] = suffix_col
                val["suffix_length"] = 3
            return Resolution(
                self.kind, val, "green", "registry",
                f"loan_suffix {'present ('+suffix_col+')' if has_suffix else 'absent'}"
                f" → {mode} (registry support {support} CUs)")
        return Resolution(self.kind, {"mode": "fixed_suffix"}, "yellow",
                          "profile", "no registry pattern; default fixed_suffix")


# ---------------------------------------------------------------------------
# pool_code — which column is the loan-pool code.
#   Oracle: tie-to-book (self-validated via the detector).
# ---------------------------------------------------------------------------
class PoolCodeResolver(Resolver):
    kind = "pool_code"
    depends_on = ("loan_file",)

    def resolve_local(self, ctx: ResolverContext) -> Optional[Resolution]:
        if ctx.warm:
            col = (ctx.warm.get("column_mappings") or {}).get("loan_pool_code")
            if col:
                return Resolution(
                    self.kind, {"column": col}, "green", "warm",
                    f"WARM Data-tab formula → loan_pool_code = {col!r}",
                    self_validated=False)
        if not ctx.loan_path or not ctx.balance_col:
            return None
        pool_map = (ctx.base_cfg.get("pool_map") if ctx.base_cfg else None)
        if not pool_map or not ctx.monthly_by_pool:
            names = [p["fragment"]["column"]
                     for p in reg.get_patterns(ctx.engine, "pool_code_column")]
            return Resolution(
                self.kind, None, "red", "needs_input",
                "no pool_map/book yet → pool-map bootstrap needed. Known "
                "code columns: " + ", ".join(names[:6]), alternatives=names[:8])
        try:
            df = pd.read_excel(ctx.loan_path, header=ctx.header_row - 1)
        except Exception as exc:  # noqa: BLE001
            return Resolution(self.kind, None, "red", "profile",
                              f"could not read loan file: {exc}")
        if ctx.balance_col not in df.columns:
            return Resolution(
                self.kind, None, "red", "profile",
                f"balance column {ctx.balance_col!r} not in loan file "
                f"(header/format mismatch)")
        res = pcd.detect(df, ctx.balance_col, pool_map, ctx.monthly_by_pool)
        return Resolution(
            self.kind, res.winner, res.decision_tier, "detector",
            res.evidence, alternatives=[c.column for c in res.candidates],
            self_validated=True)


# ---------------------------------------------------------------------------
# score_source — original / current FICO columns + credit-pull keys.
#   Oracle: pull-join rate + movement sanity (deferred to the probe).
# ---------------------------------------------------------------------------
class ScoreSourceResolver(Resolver):
    kind = "score_source"
    depends_on = ()

    def resolve_local(self, ctx: ResolverContext) -> Optional[Resolution]:
        if ctx.warm:
            cm = ctx.warm.get("column_mappings") or {}
            ss = ctx.warm.get("score_sources") or {}
            cp = ctx.warm.get("credit_pull") or {}
            val = {
                "original_fico_score": cm.get("original_fico_score"),
                "current_fico_score": cm.get("current_fico_score"),
                "current_source": ss.get("current"),
            }
            # SSN-keyed pull: loans join the pull on their SSN column rather
            # than the member number. warm_parser flags this from the pull's
            # key-column header + the Data tab's raw SSN column.
            if cp.get("ssn_keyed") and cp.get("loan_join_column"):
                val["loan_join_column"] = cp["loan_join_column"]
                val["pull_member_column"] = cp.get("member_column") or "SSN"
            if val["original_fico_score"] or val["current_source"]:
                jc = (f", pull-join ← {val['loan_join_column']}"
                      if val.get("loan_join_column") else "")
                return Resolution(
                    self.kind, val, "green", "warm",
                    f"WARM: original ← {val['original_fico_score']}, "
                    f"current ← {ss.get('current')}{jc}")
        # Registry: most-supported score-source column names.
        orig = reg.get_patterns(ctx.engine, "score_source",
                                fingerprint_filter={"role": "original"})
        if orig:
            return Resolution(
                self.kind,
                {"original_fico_score": orig[0]["fragment"]["column"]},
                "yellow", "registry",
                f"registry original-FICO column {orig[0]['fragment']['column']!r} "
                f"(support {orig[0]['support']})")
        return Resolution(self.kind, None, "red", "needs_input",
                          "no score-source signal")


# ---------------------------------------------------------------------------
# pool_map — code → pool mapping.
#   Oracle: per-pool tie-to-book (self-validated via the bootstrap).
# ---------------------------------------------------------------------------
class PoolMapResolver(Resolver):
    kind = "pool_map"
    depends_on = ("pool_code",)

    def resolve_local(self, ctx: ResolverContext) -> Optional[Resolution]:
        if ctx.warm and ctx.warm.get("pool_map"):
            pm = ctx.warm["pool_map"]
            return Resolution(
                self.kind, pm, "green", "warm",
                f"WARM Grade Ranges & Loan Codes → {len(pm)} codes",
                self_validated=False)
        if ctx.base_cfg.get("pool_map"):
            pm = ctx.base_cfg["pool_map"]
            return Resolution(self.kind, pm, "green", "profile",
                              f"existing config pool_map ({len(pm)} codes)")
        pc = ctx.resolved.get("pool_code")
        code_col = None
        if pc and isinstance(pc.value, dict):
            code_col = pc.value.get("column") or pc.value.get("code_column")
        if not (ctx.loan_path and ctx.balance_col and ctx.monthly_by_pool):
            return Resolution(self.kind, None, "red", "needs_input",
                              "pool_map bootstrap needs loan file + book")
        try:
            df = pd.read_excel(ctx.loan_path, header=ctx.header_row - 1)
        except Exception as exc:  # noqa: BLE001
            return Resolution(self.kind, None, "red", "profile",
                              f"could not read loan file: {exc}")
        if ctx.balance_col not in df.columns:
            return Resolution(self.kind, None, "red", "profile",
                              f"balance column {ctx.balance_col!r} not in loan file")
        # Candidate code columns: the resolved one first, then every numeric
        # low-cardinality column. Bootstrapping each and keeping the best tie
        # picks the right code column by ground truth (the U-vs-V resolution).
        candidates: list[str] = []
        if code_col in df.columns:
            candidates.append(code_col)
        for c in df.columns:
            if c in candidates or c == ctx.balance_col:
                continue
            s = df[c].dropna()
            if s.empty:
                continue
            if (pd.to_numeric(s, errors="coerce").notna().mean() > 0.9
                    and 5 <= s.astype(str).nunique() <= 60):
                candidates.append(c)
        candidates.sort(key=lambda c: 0 if re.search(r"type|code", str(c), re.I) else 1)
        best = best_res = best_col = None
        for c in candidates[:8]:
            r = pmb.bootstrap(df, c, ctx.balance_col, ctx.monthly_by_pool,
                              exclude_pools=set())
            hi = sum(p.balance for p in r.code_proposals if p.confidence == "high")
            key = (r.pools_tied, hi)
            if best is None or key > best:
                best, best_res, best_col = key, r, c
        if best_res is None:
            return Resolution(self.kind, None, "red", "needs_input",
                              "no candidate pool-code column found")
        amb = [p for p in best_res.code_proposals if p.confidence != "high"]
        tier = "green" if not amb else ("yellow" if best_res.pools_tied else "red")
        # write the chosen code column back so the pool_code decision reflects
        # the ground-truth pick when it was greenfield/unresolved.
        if pc is None or pc.value is None:
            ctx.resolved["pool_code"] = Resolution(
                "pool_code", {"column": best_col}, tier, "detector",
                f"chosen by pool_map bootstrap tie ({best_res.pools_tied}"
                f"/{best_res.pools_total})", self_validated=True)
        return Resolution(
            self.kind,
            {"pool_map": best_res.pool_map, "code_column": best_col,
             "ambiguous_codes": [p.code for p in amb]},
            tier, "detector",
            f"picked {best_col!r}; bootstrapped {len(best_res.pool_map)} codes; "
            f"ties {best_res.pools_tied}/{best_res.pools_total}; {len(amb)} ambiguous",
            alternatives=best_res.ambiguous, self_validated=True)


# ---------------------------------------------------------------------------
# monthly_balance — the book source (single / per_month / per_year / manual).
#   Oracle: book totals reconcile.
# ---------------------------------------------------------------------------
class MonthlyBalanceResolver(Resolver):
    kind = "monthly_balance"
    depends_on = ()

    def resolve_local(self, ctx: ResolverContext) -> Optional[Resolution]:
        if ctx.warm and ctx.warm.get("monthly_balance"):
            mb = ctx.warm["monthly_balance"]
            n = len(mb.get("months") or [])
            return Resolution(
                self.kind, {"source": mb.get("source", "manual"), "n_months": n},
                "green", "warm",
                f"WARM BS Data → {n} months (source=manual)")
        if ctx.base_cfg.get("monthly_balance"):
            src = (ctx.base_cfg["monthly_balance"] or {}).get("source")
            return Resolution(self.kind, {"source": src}, "green", "profile",
                              f"existing config monthly_balance source={src!r}")
        if "monthly_balances" in ctx.profiles:
            return Resolution(self.kind, {"source": "single"}, "yellow",
                              "profile", "monthly-balances file present")
        pats = reg.get_patterns(ctx.engine, "monthly_balance_source")
        if pats:
            return Resolution(self.kind, {"source": pats[0]["fragment"]["source"]},
                              "yellow", "registry",
                              f"registry archetype {pats[0]['fragment']['source']!r}")
        return Resolution(self.kind, None, "red", "needs_input",
                          "no monthly-balance source detected")


# ---------------------------------------------------------------------------
# secondary_extracts — participation / credit-card files beyond the core.
#   Oracle: the pool balance gap closes (deferred to the probe).
# ---------------------------------------------------------------------------
def _all_pool_names(ctx: ResolverContext) -> set:
    names: set = set()
    for src in ((ctx.warm or {}).get("pool_map"), ctx.base_cfg.get("pool_map")):
        if src:
            names |= {str(v) for v in src.values()}
    mb = (ctx.warm or {}).get("monthly_balance") or {}
    names |= {str(p) for p in (mb.get("entries") or {})}
    names |= {str(p) for p in ((ctx.warm or {}).get("acl_reserve_rates") or {})}
    return names


def _detect_credit_card_extract(ctx: ResolverContext) -> Optional[dict]:
    """Detect a separate credit-card extract (many cores export CC loans in a
    file of their own, so the main loan file is 'without Credit Cards'). Mapped
    as a static-pool second extract."""
    profs = ctx.all_profiles or list(ctx.profiles.values())
    cc_pool = next((p for p in _all_pool_names(ctx)
                    if re.search(r"visa|credit\s*card|master", p, re.I)), None)
    for fp in profs:
        if fp.path == ctx.loan_path or not fp.sheets:
            continue
        cols = [str(c.name).strip() for c in fp.sheets[0].columns]
        bal = next((c for c in cols
                    if re.search(r"stmt\s*bal|cc.*bal|card.*bal|last.*bal", c, re.I)), None)
        limit = next((c for c in cols if re.search(r"credit\s*limit|cc.*limit", c, re.I)), None)
        cc_ind = any(re.search(r"\bcc\b|credit\s*card|card\s*(nbr|number|note)", c, re.I)
                     for c in cols)
        if bal and (limit or cc_ind):
            member = (
                next((c for c in cols
                      if re.search(r"acct\s*nbr|account\s*(number|nbr)|\bacct\b", c, re.I)), None)
                or next((c for c in cols
                         if re.search(r"\bmember\b", c, re.I)
                         and not re.search(r"date|name|close|address|city|state|zip|ssn",
                                           c, re.I)), None))
            base = os.path.basename(fp.path)
            prefix = re.split(r"\d", base)[0].strip()  # descriptive part before the date
            file_pattern = (r"(?i)^" + re.escape(prefix) + r".*\.(xlsx|xls|csv)$"
                            if prefix else None)
            return {"label": "Credit Cards", "path": fp.path,
                    "file_pattern": file_pattern,
                    "static_pool": cc_pool, "balance_column": bal,
                    "member_column": member, "available_credit_column": limit}
    return None


class SecondaryExtractsResolver(Resolver):
    kind = "secondary_extracts"
    depends_on = ("pool_map",)

    def applies(self, ctx: ResolverContext) -> bool:
        return bool(ctx.warm) or bool(ctx.all_profiles) or "participation" in ctx.profiles

    def resolve_local(self, ctx: ResolverContext) -> Optional[Resolution]:
        extracts: list = []
        bal_only: list = []
        # Balance-only participation pools (carried by the monthly book).
        if ctx.warm:
            acl = ctx.warm.get("acl_reserve_rates") or {}
            bal_only = [p for p, d in acl.items()
                        if d.get("balance_only") and "partic" in p.lower()]
        # Credit-card extract (a separate file with a static CC pool).
        cc = _detect_credit_card_extract(ctx)
        if cc:
            extracts.append(cc)
        # Participation file with loan-level detail.
        pp = ctx.profiles.get("participation")
        if pp is not None:
            cols = [c.name for sp in pp.sheets for c in sp.columns]

            def _find(patterns):
                for p in patterns:
                    for c in cols:
                        if re.search(p, c, re.I):
                            return c
                return None
            extracts.append({
                "label": "Participation", "path": pp.path,
                "balance_column": _find([r"investor upb", r"current upb", r"upb"]),
                "original_score": _find([r"orig.*risk score", r"orig.*score"]),
                "current_score": _find([r"cur.*risk score", r"cur.*score"]),
                "static_pool": None})

        # Tier: yellow when a detected extract still needs a static pool name.
        tier = "green"
        for e in extracts:
            if e.get("static_pool") is None:
                tier = "yellow"
        parts = []
        if cc:
            parts.append(f"credit-card file {os.path.basename(cc['path'])} "
                         f"→ static pool {cc.get('static_pool')!r}")
        if pp is not None:
            parts.append(f"participation file {os.path.basename(pp.path)}")
        if bal_only:
            parts.append(f"{len(bal_only)} balance-only participation pool(s)")
        return Resolution(
            self.kind, {"extracts": extracts, "balance_only_pools": bal_only},
            tier, "warm" if ctx.warm else "detector",
            "; ".join(parts) or "no secondary extracts")


# ---------------------------------------------------------------------------
# reserve_config — not_risk_rated / warm_allowance_pools / base loss rates.
#   Oracle: the generated reserve reconciles to the WARM.
# ---------------------------------------------------------------------------
class ReserveConfigResolver(Resolver):
    kind = "reserve_config"
    depends_on = ("pool_map",)

    def resolve_local(self, ctx: ResolverContext) -> Optional[Resolution]:
        acl = (ctx.warm or {}).get("acl_reserve_rates")
        if acl:
            reserve = ca.derive_reserve_config(acl)
            n_nrr = len(reserve.get("not_risk_rated") or [])
            n_wa = len(reserve.get("warm_allowance_pools") or [])
            n_blr = len(reserve.get("base_loss_rate_by_pool_grade") or {})
            return Resolution(
                self.kind, reserve, "green", "warm",
                f"WARM ACL tab → {n_nrr} NRR pool(s), {n_wa} warm-allowance "
                f"pool(s), {n_blr} pinned rate table(s)", self_validated=False)
        return Resolution(self.kind, None, "red", "needs_input",
                          "no WARM ACL rates — reserve computed from CO history")


def build_default_resolvers() -> list[Resolver]:
    """The default specialist set + their DAG edges."""
    return [
        LoanFileResolver(),
        MemberAccountResolver(),
        PoolCodeResolver(),
        ScoreSourceResolver(),
        PoolMapResolver(),
        MonthlyBalanceResolver(),
        SecondaryExtractsResolver(),
        ReserveConfigResolver(),
    ]
