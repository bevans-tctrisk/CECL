"""Pure-function extraction of the ACL environmental calculation.

``report_vizo._sheet_acl_reserve`` (report_vizo.py:2223) both *computes* the
ACL environmental reserve and *renders* it onto the
``ACL Env by Pool Mgmt Adj`` worksheet.  Because the computation only ever
existed inside the renderer, four other tabs (``ACL Summary``,
``Mgmt Adj Summary``, ``Impaired Loans``, ``Summary Variance``) recover their
numbers by screen-scraping that worksheet, and ``change_analysis`` applies the
same screen-scrape to the prior quarter's ``.xlsx`` on disk.

This module lifts the arithmetic out -- and *only* the arithmetic -- into
``compute_acl_environmental``.  It is a transcription of the computation in
``_sheet_acl_reserve``: every branch, every fallback and every precedence rule
is preserved verbatim, including the ones that look like bugs.  It deliberately
does **not** "improve" anything.  Numeric equality with the rendered worksheet
is the only acceptance criterion (see ``scripts/verify_acl_model.py``).

Step 1a of ``docs/pdf_migration/00_plan.md``.  This module is *additive*:
nothing in ``report_vizo`` imports it yet, so no generated report changes.

The helper functions are imported from ``report_vizo`` rather than copied, so
the two implementations cannot silently drift apart while both exist.
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field, asdict
from typing import Any, Dict, List, Optional

__all__ = [
    'SCHEMA_VERSION',
    'AclGradeRow',
    'AclPool',
    'ImpairedCategory',
    'OtherAllowanceConsideration',
    'AclEnvironmental',
    'compute_acl_environmental',
]

SCHEMA_VERSION = 1

# The six impairment categories rendered by _sheet_acl_reserve, in order
# (report_vizo.py:2695-2696).
IMPAIRED_LABELS = [
    "Delinquent Loans", "Known Losses", "Repossessions",
    "Foreclosed Real Estate", "Deceased", "Bankruptcy",
]


def _f(v):
    """Coerce numpy scalars / None to plain Python floats for serialization."""
    if v is None:
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


# ==================================================================
# Dataclasses
# ==================================================================

@dataclass
class AclGradeRow:
    """One row of an ACL pool block -- columns A..H of the worksheet.

    Used both for the per-grade detail rows and for a pool's ``Total`` row.
    On a risk-rated pool's Total row the three rate columns (E, F, G) are
    left blank on the sheet; they are ``None`` here to match.
    """
    grade: str                                  # col A
    balance: Optional[float]                    # col B
    specific_id: Optional[float]                # col C
    loan_loss_calc_balance: Optional[float]     # col D
    acl_base_loss_rate: Optional[float]         # col E
    mgmt_adj: Optional[float]                   # col F
    allowance_factor: Optional[float]           # col G
    allowance_before_env: Optional[float]       # col H

    def to_dict(self) -> Dict[str, Any]:
        d = asdict(self)
        for k in list(d):
            if k != 'grade':
                d[k] = _f(d[k])
        return d

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'AclGradeRow':
        return cls(
            grade=d['grade'],
            balance=d.get('balance'),
            specific_id=d.get('specific_id'),
            loan_loss_calc_balance=d.get('loan_loss_calc_balance'),
            acl_base_loss_rate=d.get('acl_base_loss_rate'),
            mgmt_adj=d.get('mgmt_adj'),
            allowance_factor=d.get('allowance_factor'),
            allowance_before_env=d.get('allowance_before_env'),
        )


@dataclass
class AclPool:
    """One pool block on ``ACL Env by Pool Mgmt Adj``.

    ``grade_rows`` is empty for a non-risk-rated pool, which renders only its
    Total row (and, unlike a risk-rated pool, does fill columns E/F/G on it).
    """
    name: str
    risk_rated: bool
    brr: bool
    grade_rows: List[AclGradeRow]
    totals: AclGradeRow                 # the pool "Total" row (cols A..H)
    env_factor: float                   # col I
    env_allowance: float                # col J
    total_allowance: float              # col K

    # -- provenance: not rendered on this tab, but it is what
    #    "Env Factor by Pool" shows and what makes the numbers auditable --
    has_db_data: bool = True
    effective_rate: float = 0.0         # total_allowance / calc balance
    life_loss_rate: float = 0.0
    ncc_pct: float = 0.0
    ncc_score: float = 0.0
    dq_variance: float = 0.0
    dq_score: float = 0.0
    econ_stress: float = 0.0
    es_score: float = 0.0
    acl_months: int = 36

    def to_dict(self) -> Dict[str, Any]:
        return {
            'name': self.name,
            'risk_rated': bool(self.risk_rated),
            'brr': bool(self.brr),
            'grade_rows': [g.to_dict() for g in self.grade_rows],
            'totals': self.totals.to_dict(),
            'env_factor': _f(self.env_factor),
            'env_allowance': _f(self.env_allowance),
            'total_allowance': _f(self.total_allowance),
            'has_db_data': bool(self.has_db_data),
            'effective_rate': _f(self.effective_rate),
            'life_loss_rate': _f(self.life_loss_rate),
            'ncc_pct': _f(self.ncc_pct),
            'ncc_score': _f(self.ncc_score),
            'dq_variance': _f(self.dq_variance),
            'dq_score': _f(self.dq_score),
            'econ_stress': _f(self.econ_stress),
            'es_score': _f(self.es_score),
            'acl_months': int(self.acl_months or 0),
        }

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'AclPool':
        return cls(
            name=d['name'],
            risk_rated=bool(d.get('risk_rated', True)),
            brr=bool(d.get('brr', False)),
            grade_rows=[AclGradeRow.from_dict(g) for g in d.get('grade_rows', [])],
            totals=AclGradeRow.from_dict(d['totals']),
            env_factor=d.get('env_factor') or 0.0,
            env_allowance=d.get('env_allowance') or 0.0,
            total_allowance=d.get('total_allowance') or 0.0,
            has_db_data=bool(d.get('has_db_data', True)),
            effective_rate=d.get('effective_rate') or 0.0,
            life_loss_rate=d.get('life_loss_rate') or 0.0,
            ncc_pct=d.get('ncc_pct') or 0.0,
            ncc_score=d.get('ncc_score') or 0.0,
            dq_variance=d.get('dq_variance') or 0.0,
            dq_score=d.get('dq_score') or 0.0,
            econ_stress=d.get('econ_stress') or 0.0,
            es_score=d.get('es_score') or 0.0,
            acl_months=int(d.get('acl_months') or 0),
        )


@dataclass
class ImpairedCategory:
    """One line of the ``Impaired Loans`` block (label in A, allowance in K)."""
    label: str
    allowance: float

    def to_dict(self) -> Dict[str, Any]:
        return {'label': self.label, 'allowance': _f(self.allowance)}

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'ImpairedCategory':
        return cls(label=d['label'], allowance=d.get('allowance') or 0.0)


@dataclass
class OtherAllowanceConsideration:
    """One ``Other Allowance Considerations`` row.

    ``balance`` and ``percentage`` are not rendered on the ACL Env tab (only
    ``title`` in A and ``amount`` in K) but they are the audit trail for the
    amount and are carried so the PDF renderer can show the derivation.
    """
    title: str
    amount: float
    balance: float = 0.0
    percentage: float = 0.0

    def to_dict(self) -> Dict[str, Any]:
        return {
            'title': self.title,
            'amount': _f(self.amount),
            'balance': _f(self.balance),
            'percentage': _f(self.percentage),
        }

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'OtherAllowanceConsideration':
        return cls(
            title=d['title'],
            amount=d.get('amount') or 0.0,
            balance=d.get('balance') or 0.0,
            percentage=d.get('percentage') or 0.0,
        )


@dataclass
class AclEnvironmental:
    """Everything ``ACL Env by Pool Mgmt Adj`` renders -- the hub object.

    Every number the four summary tabs currently screen-scrape out of that
    worksheet is a field here.
    """
    credit_union: str
    snap: str

    pools: List[AclPool] = field(default_factory=list)

    # -- "Pooled Totals" line --
    pooled_balance: float = 0.0                 # col B
    pooled_specific_id: float = 0.0             # col C
    pooled_calc_balance: float = 0.0            # col D
    pooled_allowance_before_env: float = 0.0    # col H
    pooled_env_allowance: float = 0.0           # col J
    pooled_total_allowance: float = 0.0         # col K

    # -- Impaired Loans block --
    impaired: List[ImpairedCategory] = field(default_factory=list)
    total_specifically_identified_allowance: float = 0.0

    # -- Other Allowance Considerations block --
    other_allowance_considerations: List[OtherAllowanceConsideration] = \
        field(default_factory=list)
    total_other_allowance_considerations: float = 0.0

    # -- the four closing lines --
    total_allowance_needed: float = 0.0
    acl_balance: float = 0.0
    adjustment: float = 0.0
    adjustment_label: str = 'Adjustment (Underfunded)'
    balance_label: str = ''

    # -- the side-effect stash, made explicit ------------------------
    # ``_sheet_acl_reserve`` writes this value onto ``hist['impaired']`` as
    # ``_computed_pooled_total_allow`` (report_vizo.py:2680) so that
    # ``_compute_acl_totals`` (report_vizo.py:599) -- and therefore the
    # ``Impr Deter`` tab -- agrees with the ACL Env tab.  That stash forces
    # the out-of-order sheet build in ``compose_vizo_main`` and the
    # ``wb.move_sheet`` repair afterwards.  Exposing it as a field is what
    # lets the stash eventually be retired: the caller passes this value
    # explicitly instead of mutating its own input dict.
    # It is always equal to ``pooled_total_allowance``.
    computed_pooled_total_allow: float = 0.0
    # The two sibling stashes (report_vizo.py:2681-2682) are write-only
    # today; their values are ``pooled_allowance_before_env`` and
    # ``pooled_env_allowance`` above.

    # -- derived / provenance --
    pool_effective_rates: Dict[str, float] = field(default_factory=dict)

    @property
    def acl_over_total_loans(self) -> Optional[float]:
        """``Total Allowance Needed / Pooled Balance`` -- the ratio the
        Summary Variance tab computes as an Excel IFERROR division."""
        if not self.pooled_balance:
            return None
        return self.total_allowance_needed / self.pooled_balance

    # -- serialization ----------------------------------------------
    def to_dict(self) -> Dict[str, Any]:
        return {
            'schema_version': SCHEMA_VERSION,
            'credit_union': self.credit_union,
            'snap': str(self.snap),
            'pools': [p.to_dict() for p in self.pools],
            'pooled_balance': _f(self.pooled_balance),
            'pooled_specific_id': _f(self.pooled_specific_id),
            'pooled_calc_balance': _f(self.pooled_calc_balance),
            'pooled_allowance_before_env': _f(self.pooled_allowance_before_env),
            'pooled_env_allowance': _f(self.pooled_env_allowance),
            'pooled_total_allowance': _f(self.pooled_total_allowance),
            'impaired': [i.to_dict() for i in self.impaired],
            'total_specifically_identified_allowance':
                _f(self.total_specifically_identified_allowance),
            'other_allowance_considerations':
                [o.to_dict() for o in self.other_allowance_considerations],
            'total_other_allowance_considerations':
                _f(self.total_other_allowance_considerations),
            'total_allowance_needed': _f(self.total_allowance_needed),
            'acl_balance': _f(self.acl_balance),
            'adjustment': _f(self.adjustment),
            'adjustment_label': self.adjustment_label,
            'balance_label': self.balance_label,
            'computed_pooled_total_allow': _f(self.computed_pooled_total_allow),
            'pool_effective_rates':
                {k: _f(v) for k, v in (self.pool_effective_rates or {}).items()},
        }

    @classmethod
    def from_dict(cls, d: Dict[str, Any]) -> 'AclEnvironmental':
        ver = d.get('schema_version', SCHEMA_VERSION)
        if ver > SCHEMA_VERSION:
            raise ValueError(
                f'AclEnvironmental schema_version {ver} is newer than this '
                f'code understands ({SCHEMA_VERSION})')
        return cls(
            credit_union=d.get('credit_union', ''),
            snap=d.get('snap', ''),
            pools=[AclPool.from_dict(p) for p in d.get('pools', [])],
            pooled_balance=d.get('pooled_balance') or 0.0,
            pooled_specific_id=d.get('pooled_specific_id') or 0.0,
            pooled_calc_balance=d.get('pooled_calc_balance') or 0.0,
            pooled_allowance_before_env=d.get('pooled_allowance_before_env') or 0.0,
            pooled_env_allowance=d.get('pooled_env_allowance') or 0.0,
            pooled_total_allowance=d.get('pooled_total_allowance') or 0.0,
            impaired=[ImpairedCategory.from_dict(i) for i in d.get('impaired', [])],
            total_specifically_identified_allowance=(
                d.get('total_specifically_identified_allowance') or 0.0),
            other_allowance_considerations=[
                OtherAllowanceConsideration.from_dict(o)
                for o in d.get('other_allowance_considerations', [])
            ],
            total_other_allowance_considerations=(
                d.get('total_other_allowance_considerations') or 0.0),
            total_allowance_needed=d.get('total_allowance_needed') or 0.0,
            acl_balance=d.get('acl_balance') or 0.0,
            adjustment=d.get('adjustment') or 0.0,
            adjustment_label=d.get('adjustment_label', ''),
            balance_label=d.get('balance_label', ''),
            computed_pooled_total_allow=d.get('computed_pooled_total_allow') or 0.0,
            pool_effective_rates=dict(d.get('pool_effective_rates') or {}),
        )

    def to_json(self, **kwargs) -> str:
        kwargs.setdefault('indent', 2)
        return json.dumps(self.to_dict(), **kwargs)

    @classmethod
    def from_json(cls, text: str) -> 'AclEnvironmental':
        return cls.from_dict(json.loads(text))


# ==================================================================
# The computation
# ==================================================================

def compute_acl_environmental(df, config, grades, hist, snap) -> AclEnvironmental:
    """Compute the ACL environmental reserve.

    A verbatim extraction of the arithmetic in
    ``report_vizo._sheet_acl_reserve`` (report_vizo.py:2223-2776), with the
    worksheet writes removed.  Every branch and fallback is preserved exactly;
    the acceptance criterion is that the numbers equal the ones the existing
    builder renders (``scripts/verify_acl_model.py``).

    Args:
        df:     the loan DataFrame (post ``calculate_cecl``)
        config: the client config dict
        grades: ``config['credit_grades']``
        hist:   the historical / WARM dict
        snap:   ``'YYYY-MM-DD'`` snapshot date string

    Returns:
        ``AclEnvironmental``

    Side effects (inherited from the original, deliberately preserved so the
    numbers match):
      * ``generate_report._expand_unfunded_commitment_oac`` mutates
        ``config['other_allowance_considerations']`` in place, expanding any
        ``source: unfunded_commitment`` template into one resolved row per
        pool.  It is idempotent.  Pass a copy of ``config`` if you need the
        caller's dict left alone.

    It does NOT write ``_computed_pooled_total_allow`` back onto
    ``hist['impaired']`` -- that value is returned as a field instead.
    """
    # Imported lazily so report_vizo can later import this module without a
    # circular import at module load.
    import report_vizo as rv

    no_score = config.get('no_score_label', 'Not Reported')
    mgmt_adj_by_pool = config.get('mgmt_adj_by_pool', {})
    pool_use_default = rv._build_pool_use_default_map(config)
    admin_default_mgmt_adj = rv._load_admin_default_mgmt_adj()
    gl = rv._all_grades(grades, no_score)
    visible_gl = [g for g in gl if not rv._is_hidden(g)]
    brr_labels = rv._brr_grade_labels(config, no_score)
    brr_pool_lcs = rv._brr_pools_set(config) if brr_labels else set()

    _imp = hist.get('impaired', {}) if hist else {}
    econ_stress = rv._eco_stress(config, ed_override=_imp.get('economic_data'))
    _ncc_r, _dq_r, _es_r = rv._env_ranges(hist)

    pools = rv._ordered_pools(df, hist)
    dq_var = rv._pool_dq_variance(pools, hist, snap)
    acl_pools_data = _imp.get('acl_pools', {})
    acl_impaired = _imp.get('acl_impaired', {})
    acl_summary = _imp.get('acl_summary', {})
    prior_mgmt_adj = _imp.get('prior_mgmt_adj', {})
    spec_id_by_pool = _imp.get('spec_id_by_pool', {})

    # -- Per-pool Life Loss Rate (matches Display Hist Bal) --
    co_data = hist.get('chargeoffs', {}) if hist else {}
    rc_data = hist.get('recoveries', {}) if hist else {}
    avg_bals = hist.get('avg_balances', {}) if hist else {}
    years = hist.get('years', []) if hist else []
    acl_months_map = _imp.get('acl_months', {})
    snap_year = int(snap[:4])
    snap_month = int(snap[5:7])
    warm_net_co = _imp.get('warm_net_co', {})
    hbd = _imp.get('hist_bal_data', {})
    annual_grade_avg = {}
    for _pk, pdata in hbd.items():
        _dates = pdata.get('dates', [])
        _grades_data = pdata.get('grades', {})
        annual_grade_avg[_pk] = {}
        for _gk, _vals in _grades_data.items():
            if _gk.upper().startswith('HIDE'):
                continue
            yr_sums = {}
            yr_cnts = {}
            for _i, _d in enumerate(_dates):
                if _i < len(_vals) and _vals[_i] > 0:
                    yr_sums[_d.year] = yr_sums.get(_d.year, 0) + _vals[_i]
                    yr_cnts[_d.year] = yr_cnts.get(_d.year, 0) + 1
            for _y in yr_sums:
                annual_grade_avg[_pk].setdefault(_y, {})
                annual_grade_avg[_pk][_y][_gk] = yr_sums[_y] / yr_cnts[_y]

    life_loss = {}
    for pool in pools:
        pool_acl = acl_months_map.get(pool, 36)
        abs_first = (snap_year * 12 + snap_month) - pool_acl + 1
        pe = (abs_first - 1) // 12
        pa = annual_grade_avg.get(pool, {})
        yr_tots = []
        for y in years:
            if y < pe:
                continue
            yt = sum(pa.get(y, {}).values())
            if not yt:
                yt = avg_bals.get(y, {}).get(pool, 0)
            if yt:
                yr_tots.append(yt)
        avg_tot = sum(yr_tots) / len(yr_tots) if yr_tots else 0
        pool_stripped = pool.strip()
        net_co_match = warm_net_co.get(pool_stripped, warm_net_co.get(pool, None))
        if net_co_match is not None:
            total_net = net_co_match
        else:
            total_net = 0
            for y in years:
                if y < pe:
                    continue
                total_net += abs(co_data.get(y, {}).get(pool, 0) or 0) \
                             - abs(rc_data.get(y, {}).get(pool, 0) or 0)
        life_loss[pool] = total_net / avg_tot if avg_tot > 0 else 0

    grand_allowance = 0
    grand_allow_before = 0
    grand_env_allow = 0
    pool_eff_rate = {}
    grand_balance = 0
    grand_spec_id = 0
    _bal_detail = _imp.get('pool_bal_detail', {})

    # -- Unified pool list in WARM order, including WARM-only pools --
    risk_rated_flags = _imp.get('risk_rated', {})
    warm_order = _imp.get('pool_order', [])
    db_pools = set(df['loan_pool'].unique())
    if not warm_order:
        warm_order = list(pools)
    all_acl_pools = []
    seen = set()
    extra_pools = list((_imp.get('hist_bal_data') or {}).keys()) \
                  + list(_bal_detail.keys()) \
                  + list(acl_pools_data.keys())
    candidates = list(warm_order) + list(pools) + extra_pools
    nrr_set = set(config.get('not_risk_rated', []) or [])
    for p in candidates:
        if not p:
            continue
        s = str(p).strip()
        if not s or s == 'Exclude' or s.upper().startswith('HIDE'):
            continue
        if s.lower() in ('total', 'grand total', 'excluded'):
            continue
        if p in seen:
            continue
        seen.add(p)
        all_acl_pools.append(p)

    out_pools: List[AclPool] = []

    for pool in all_acl_pools:
        pdf = df[df['loan_pool'] == pool]
        pool_total = pdf['current_balance'].sum()
        has_db_data = pool in db_pools

        _pool_lc = pool.strip().lower()
        warm_pool = next((v for k, v in acl_pools_data.items()
                          if k.strip().lower() == _pool_lc), None)
        warm_grades = warm_pool['grades'] if warm_pool else {}
        warm_total = warm_pool['total'] if warm_pool else {}

        if pool in nrr_set:
            is_rr = False
        else:
            is_rr = risk_rated_flags.get(pool, has_db_data)

        # Env factor -- same inputs and the same scoring ranges the
        # "Env Factor by Pool" tab uses, so the two tabs agree.
        ncc_pct = 0.0
        dq_v = 0.0
        ncc_score = dq_score = es_score = 0.0
        if has_db_data:
            if is_rr:
                _, _, ncc_pct = rv._ncc(pdf, grades, config)
            else:
                ncc_pct = 0.0
            dq_v = dq_var.get(pool, 0)
            ncc_score = rv._score(ncc_pct * 100, _ncc_r)
            dq_score = rv._score(dq_v * 100, _dq_r)
            es_score = rv._score(econ_stress, _es_r)
            env_factor = (ncc_score + dq_score + es_score) / 100.0
        else:
            env_factor = 0

        pool_ll = life_loss.get(pool, 0)
        _is_brr = rv._is_brr_pool(pool, brr_pool_lcs)
        grade_rows: List[AclGradeRow] = []

        if is_rr:
            # -- Risk-rated pool: per-grade detail --
            pool_grade_labels = (
                brr_labels if _is_brr else visible_gl
            )
            pool_allow_before = 0
            pool_grade_balance_sum = 0
            pool_grade_spec_id_sum = 0
            for gi, g in enumerate(pool_grade_labels):
                wg = warm_grades.get(g, {})
                _pd = _bal_detail.get(pool, {})
                _gd = _pd.get(g, {})
                bst = _gd.get('balance_sheet_total') if _gd else None
                if wg:
                    balance = bst if bst is not None else wg.get('balance', 0)
                    specific_id = wg.get('spec_id', 0)
                    if specific_id == 0 and pool in spec_id_by_pool:
                        specific_id = spec_id_by_pool[pool].get(g, 0)
                    calc_bal = balance - specific_id
                    dist = (rv._dist_factor(len(rv.DIST_FACTORS) - 1)
                            if g == no_score else rv._dist_factor(gi))
                    base_rate = max(0, pool_ll * dist)
                    mgmt_adj = rv._resolve_mgmt_adj_grade(
                        pool, g, gi, no_score,
                        pool_use_default, mgmt_adj_by_pool,
                        admin_default_mgmt_adj, prior_mgmt_adj,
                        base_rate=base_rate,
                    )
                    factor = base_rate + mgmt_adj
                    allow_before = calc_bal * factor
                elif has_db_data:
                    if _gd and _gd.get('balance_sheet_total', 0):
                        balance = _gd['balance_sheet_total']
                    else:
                        g_df = pdf[pdf['current_grade'] == g]
                        balance = g_df['current_balance'].sum()
                    specific_id = spec_id_by_pool.get(pool, {}).get(g, 0)
                    calc_bal = balance - specific_id
                    dist = (rv._dist_factor(len(rv.DIST_FACTORS) - 1)
                            if g == no_score else rv._dist_factor(gi))
                    base_rate = max(0, pool_ll * dist)
                    mgmt_adj = rv._resolve_mgmt_adj_grade(
                        pool, g, gi, no_score,
                        pool_use_default, mgmt_adj_by_pool,
                        admin_default_mgmt_adj, prior_mgmt_adj,
                        base_rate=base_rate,
                    )
                    factor = base_rate + mgmt_adj
                    allow_before = calc_bal * factor
                else:
                    balance = specific_id = calc_bal = 0
                    base_rate = mgmt_adj = factor = allow_before = 0
                pool_allow_before += allow_before

                grade_rows.append(AclGradeRow(
                    grade=g,
                    balance=balance,
                    specific_id=specific_id,
                    loan_loss_calc_balance=calc_bal,
                    acl_base_loss_rate=base_rate,
                    mgmt_adj=mgmt_adj,
                    allowance_factor=factor,
                    allowance_before_env=allow_before,
                ))
                pool_grade_balance_sum += balance or 0
                pool_grade_spec_id_sum += specific_id or 0

            # Pool total row
            _ptd = _bal_detail.get(pool, {}).get('Total', {})
            if _is_brr:
                total_balance = pool_grade_balance_sum
            elif _ptd and _ptd.get('balance_sheet_total'):
                total_balance = _ptd['balance_sheet_total']
            elif warm_total:
                total_balance = warm_total.get('balance', pool_total)
            else:
                total_balance = pool_total
            pool_allow_before_out = pool_allow_before
            env_allow = pool_allow_before_out * env_factor
            total_allow = pool_allow_before_out + env_allow
            grand_allowance += total_allow
            grand_allow_before += pool_allow_before_out
            grand_env_allow += env_allow

            if _is_brr:
                total_spec_id = pool_grade_spec_id_sum
            else:
                total_spec_id = warm_total.get('spec_id', 0) if warm_total else 0
                if total_spec_id == 0 and pool in spec_id_by_pool:
                    total_spec_id = sum(spec_id_by_pool[pool].values())
            total_calc_bal = total_balance - total_spec_id
            grand_balance += total_balance or 0
            grand_spec_id += total_spec_id or 0
            pool_eff_rate[pool] = (total_allow / total_calc_bal) \
                if total_calc_bal else 0

            totals_row = AclGradeRow(
                grade='Total',
                balance=total_balance,
                specific_id=total_spec_id,
                loan_loss_calc_balance=total_calc_bal,
                # Columns E/F/G are intentionally left blank on a risk-rated
                # pool's Total row by the renderer.
                acl_base_loss_rate=None,
                mgmt_adj=None,
                allowance_factor=None,
                allowance_before_env=pool_allow_before_out,
            )
            out_pools.append(AclPool(
                name=pool, risk_rated=True, brr=bool(_is_brr),
                grade_rows=grade_rows, totals=totals_row,
                env_factor=env_factor, env_allowance=env_allow,
                total_allowance=total_allow,
                has_db_data=has_db_data,
                effective_rate=pool_eff_rate[pool],
                life_loss_rate=pool_ll,
                ncc_pct=ncc_pct, ncc_score=ncc_score,
                dq_variance=dq_v, dq_score=dq_score,
                econ_stress=econ_stress, es_score=es_score,
                acl_months=acl_months_map.get(pool, 36),
            ))
        else:
            # -- Non-risk-rated pool: Total row only, with rate columns --
            _ptd_nrr = _bal_detail.get(pool, {}).get('Total', {})
            if _ptd_nrr and _ptd_nrr.get('balance_sheet_total') is not None:
                nrr_balance = _ptd_nrr['balance_sheet_total']
            elif warm_total:
                nrr_balance = warm_total.get('balance', pool_total)
            else:
                nrr_balance = pool_total
            nrr_spec_id = warm_total.get('spec_id', 0)
            if nrr_spec_id == 0 and pool in spec_id_by_pool:
                nrr_spec_id = sum(spec_id_by_pool[pool].values())
            nrr_calc_bal = nrr_balance - nrr_spec_id
            nrr_base_rate = max(0, pool_ll)
            nrr_mgmt_adj = rv._resolve_mgmt_adj_total(
                pool, pool_use_default, mgmt_adj_by_pool,
                admin_default_mgmt_adj,
                base_rate=nrr_base_rate,
            )
            nrr_factor = nrr_base_rate + nrr_mgmt_adj
            nrr_allow_before = nrr_calc_bal * nrr_factor
            nrr_env_factor = env_factor
            nrr_env_allow = nrr_allow_before * nrr_env_factor
            nrr_total_allow = nrr_allow_before + nrr_env_allow
            grand_allowance += nrr_total_allow
            grand_allow_before += nrr_allow_before
            grand_env_allow += nrr_env_allow
            grand_balance += nrr_balance or 0
            grand_spec_id += nrr_spec_id or 0
            pool_eff_rate[pool] = (nrr_total_allow / nrr_calc_bal) \
                if nrr_calc_bal else 0

            totals_row = AclGradeRow(
                grade='Total',
                balance=nrr_balance,
                specific_id=nrr_spec_id,
                loan_loss_calc_balance=nrr_calc_bal,
                acl_base_loss_rate=nrr_base_rate,
                mgmt_adj=nrr_mgmt_adj,
                allowance_factor=nrr_factor,
                allowance_before_env=nrr_allow_before,
            )
            out_pools.append(AclPool(
                name=pool, risk_rated=False, brr=bool(_is_brr),
                grade_rows=[], totals=totals_row,
                env_factor=nrr_env_factor, env_allowance=nrr_env_allow,
                total_allowance=nrr_total_allow,
                has_db_data=has_db_data,
                effective_rate=pool_eff_rate[pool],
                life_loss_rate=pool_ll,
                ncc_pct=ncc_pct, ncc_score=ncc_score,
                dq_variance=dq_v, dq_score=dq_score,
                econ_stress=econ_stress, es_score=es_score,
                acl_months=acl_months_map.get(pool, 36),
            ))

    # -- Grand totals --
    pooled_balance = grand_balance if grand_balance else \
        acl_summary.get('pooled_balance', df['current_balance'].sum())
    pooled_total_allow = grand_allowance

    pooled_spec_id = grand_spec_id if grand_balance else \
        acl_summary.get('pooled_spec_id', 0)
    if pooled_spec_id == 0 and spec_id_by_pool:
        pooled_spec_id = sum(sum(g.values()) for g in spec_id_by_pool.values())
    pooled_calc_bal = pooled_balance - pooled_spec_id

    # -- Impaired Loans block --
    impaired_rows: List[ImpairedCategory] = []
    for lbl in IMPAIRED_LABELS:
        imp_val = acl_impaired.get(lbl, 0)
        if lbl.upper().startswith('HIDE'):
            continue
        impaired_rows.append(ImpairedCategory(label=lbl, allowance=imp_val))
    total_spec_allow = acl_summary.get('total_spec_allow',
                                       sum(acl_impaired.values()))

    # Expand any Unfunded-Commitment OAC templates now that every pool's
    # effective ACL rate is known.  Mutates config in place; idempotent.
    try:
        import generate_report as _gr
        _gr._expand_unfunded_commitment_oac(config, pool_eff_rate, snap)
    except Exception as _e:  # noqa: BLE001
        print(f"  OAC unfunded expansion skipped: {_e}")
    oac_raw = rv._other_allowance_considerations(config)
    oac_total = sum(o['amount'] for o in oac_raw)
    oac_rows = [OtherAllowanceConsideration(
        title=o['title'], amount=o['amount'],
        balance=o.get('balance', 0.0), percentage=o.get('percentage', 0.0),
    ) for o in oac_raw]

    total_allow_needed = pooled_total_allow + total_spec_allow + oac_total
    acl_bal = acl_summary.get('acl_balance', config.get('acl_balance', 0))
    adjustment = total_allow_needed - acl_bal

    # Positive adjustment (needed > balance) => CU is UNDERfunded.
    adj_label = (
        "Adjustment (Underfunded)" if adjustment >= 0
        else "Adjustment (Overfunded)"
    )

    return AclEnvironmental(
        credit_union=config.get('credit_union', ''),
        snap=str(snap),
        pools=out_pools,
        pooled_balance=pooled_balance,
        pooled_specific_id=pooled_spec_id,
        pooled_calc_balance=pooled_calc_bal,
        pooled_allowance_before_env=grand_allow_before,
        pooled_env_allowance=grand_env_allow,
        pooled_total_allowance=pooled_total_allow,
        impaired=impaired_rows,
        total_specifically_identified_allowance=total_spec_allow,
        other_allowance_considerations=oac_rows,
        total_other_allowance_considerations=oac_total,
        total_allowance_needed=total_allow_needed,
        acl_balance=acl_bal,
        adjustment=adjustment,
        adjustment_label=adj_label,
        # The renderer writes the raw snap string here, not the mm/dd/yyyy
        # display form used elsewhere (report_vizo.py:2743).
        balance_label=f"Allowance for Credit Loss Balance as of {snap}",
        computed_pooled_total_allow=pooled_total_allow,
        pool_effective_rates=dict(pool_eff_rate),
    )
