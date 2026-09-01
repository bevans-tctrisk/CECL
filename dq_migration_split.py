"""Derive delinquency split by credit-grade migration status from loan-level data.

Background
----------
The "Delinquency by Credit Grade Migration" pie on every ``Risk Change`` /
``Risk Chg <pool>`` tab is fed from ``hist['impaired']['dq_by_status']`` and
``hist['impaired']['dq_by_pool']``.  Historically those two keys had exactly
one producer -- ``generate_report.load_impaired_data``, which reads the
``DQ Data Entry`` tab out of a legacy ~80-tab CECL-Migration-WARM workbook.
Credit unions onboarded through the wizard never had such a workbook, so the
dict stayed empty and ``_sheet_risk_change`` silently wrote literal zeros
(see ``docs/pdf_migration/04_blank_charts.md``).

This module reconstructs the same split from data the pipeline already reads:

* **migration status** -- from the loan-level frame (``monthly_loan_data`` +
  ``cecl_engine.calculate_cecl``), bucketed with *exactly* the rule
  ``report_vizo._sheet_risk_change`` uses to colour the matrix and to compute
  the ``Net Credit Change`` block, so the pie and the table beside it agree.
* **delinquency** -- ``days_delinquent`` read off the configured loan
  extract(s) at report time and joined to the loan rows on the full account
  string, reusing ``generate_impdet_report._load_extract_enrichment`` (the
  same join the Improved/Deteriorated detail report already relies on).

Nothing here writes to the database and nothing here overrides a WARM-supplied
value: ``generate_report`` only calls this when the WARM did not supply the
keys.
"""

from __future__ import annotations

import math
import os
import re

STATUS_LABELS = ("Improved", "Deteriorated", "Unchanged", "Not Reported")

# WARM convention: a loan counts as delinquent at 60+ days past due.
DEFAULT_DQ_THRESHOLD = 60


def _coerce_days(val):
    """Parse a 'days delinquent' cell to a float. Empty/non-numeric -> None.

    Deliberately *not* ``dq_extract_parser._coerce_days``: that helper strips
    every non-digit character, which turns a pandas float cell (``60.0``)
    into ``600``.  Loan extracts routinely come back as float64 columns
    whenever a single row is blank, so a digits-only strip is unsafe here.
    """
    from cecl_ui.services.coercion import coerce_number
    return coerce_number(val)


def _label_sets(config, grades, no_score):
    """Return ``(fico_labels, brr_labels_or_None, brr_pool_lcs)``.

    Mirrors the grade-list resolution at the top of
    ``report_vizo._sheet_risk_change`` so the buckets line up with the matrix
    on both the FICO tabs and the Business-Risk-Rating pool tabs.
    """
    import report_vizo as _rv

    fico = [g for g in _rv._all_grades(grades, no_score) if not _rv._is_hidden(g)]
    brr_full = _rv._brr_grade_labels(config, no_score)
    if brr_full:
        brr = [g for g in brr_full if not _rv._is_hidden(g)]
        brr_lcs = _rv._brr_pools_set(config)
    else:
        brr, brr_lcs = None, set()
    return fico, brr, brr_lcs


def classify_migration(current_grade, original_grade, grade_index, n_top, no_score):
    """Bucket one loan into Improved / Deteriorated / Unchanged / Not Reported.

    ``grade_index`` maps grade label -> ordinal position in the matrix's
    row/column order.  The rule is copied verbatim from the per-original-grade
    loop in ``report_vizo._sheet_risk_change`` (the one that produces
    ``grand_det`` / ``grand_imp`` / ``grand_unc``):

    * either side unscored              -> ``Not Reported``
    * current worse than original (i>j) -> ``Deteriorated``, except the
      top-grade exception ``j < n_top and (i - j) < 2`` which stays Unchanged
    * current better than original (i<j) -> ``Improved``
    * same grade                        -> ``Unchanged``

    Returns ``None`` when either grade is outside the matrix (e.g. a BRR
    rating evaluated against the FICO label set) so the caller can decide.
    """
    if current_grade == no_score or original_grade == no_score:
        return 'Not Reported'
    i = grade_index.get(current_grade)
    j = grade_index.get(original_grade)
    if i is None or j is None:
        return None
    if i > j:
        if j < n_top and (i - j) < 2:
            return 'Unchanged'
        return 'Deteriorated'
    if i < j:
        return 'Improved'
    return 'Unchanged'


# Other column_mappings fields a ``days_delinquent`` reference must not
# duplicate. The wizard's auto-suggest occasionally lands ``days_delinquent``
# on an unrelated column (e.g. Destinations CU maps it to ``Interest Rate`` on
# one extract and to the account-number column ``ACCTBS`` on another). Those
# values are not days past due, and silently bucketing them would replace an
# empty chart with a wrong one.
_DQ_ALIAS_FIELDS = (
    'member_number', 'loan_suffix', 'current_balance', 'loan_pool_code',
    'interest_rate', 'open_date', 'original_loan_amount',
    'total_available_credit', 'original_fico_score', 'current_fico_score',
    'business_risk_rating',
)

# Refuse the derivation when an implausible share of the matched loans comes
# back delinquent -- the near-certain sign of a mis-mapped column rather than
# a genuinely distressed portfolio.
MAX_PLAUSIBLE_DQ_SHARE = 0.50


def _strip_bogus_dq_mappings(config):
    """Return a shallow copy of ``config`` whose ``days_delinquent`` mappings
    are dropped wherever they merely alias another mapped column."""
    def _clean(col_map, label):
        ref = (col_map or {}).get('days_delinquent')
        if ref is None or (isinstance(ref, str) and not ref.strip()):
            return col_map, False
        key = str(ref).strip().lower()
        for other in _DQ_ALIAS_FIELDS:
            oref = col_map.get(other)
            if oref is None:
                continue
            if str(oref).strip().lower() == key:
                out = dict(col_map)
                out.pop('days_delinquent', None)
                print(f"    DQ migration split: ignoring 'days_delinquent' on "
                      f"{label} -- it maps to {ref!r}, the same column as "
                      f"'{other}', so it is not a days-past-due field.")
                return out, True
        return col_map, False

    cfg = dict(config)
    changed = False
    top, hit = _clean(cfg.get('column_mappings') or {}, 'the top-level mapping')
    if hit:
        cfg['column_mappings'] = top
        changed = True
    extracts = config.get('loan_data_extracts') or []
    if extracts:
        new_ex = []
        touched = False
        for ex in extracts:
            cm, hit = _clean(ex.get('column_mappings') or {},
                             f"extract {ex.get('label')!r}")
            if hit:
                ex = dict(ex)
                ex['column_mappings'] = cm
                touched = True
            new_ex.append(ex)
        if touched:
            cfg['loan_data_extracts'] = new_ex
    return cfg


def load_days_delinquent_by_account(config, snapshot_date, workspace_root):
    """``{full_account_string: days_delinquent}`` read from the loan extract(s).

    Returns ``{}`` when no extract can be located or none of them map
    ``days_delinquent``.  Reuses the impdet enrichment loader so the file
    resolution, header normalisation and member/account derivation are
    identical to the Improved/Deteriorated report's.
    """
    if not os.getenv('DATABASE_URL'):
        # generate_impdet_report builds its SQLAlchemy engine at import time.
        try:
            from cecl_credentials import get_database_url as _gdu
            os.environ['DATABASE_URL'] = _gdu()
        except Exception:  # noqa: BLE001 - fall through to the import error
            pass
    try:
        from generate_impdet_report import _load_extract_enrichment
    except Exception as exc:  # noqa: BLE001
        print(f"    DQ migration split: enrichment loader unavailable ({exc}).")
        return {}

    # Only worth reading the extracts when at least one of them maps
    # days_delinquent; otherwise we would pay the I/O for nothing.
    safe_cfg = _strip_bogus_dq_mappings(config)
    extracts = list(safe_cfg.get('loan_data_extracts') or [])
    if extracts:
        mapped = any((ex.get('column_mappings') or {}).get('days_delinquent')
                     for ex in extracts)
    else:
        mapped = bool((safe_cfg.get('column_mappings') or {}).get('days_delinquent'))
    if not mapped:
        print("    DQ migration split: no extract maps 'days_delinquent'; skipping.")
        return {}

    try:
        enrich = _load_extract_enrichment(safe_cfg, workspace_root, snapshot_date)
    except Exception as exc:  # noqa: BLE001
        print(f"    DQ migration split: extract read failed ({exc}).")
        return {}

    out = {}
    for acct, row in (enrich or {}).items():
        days = _coerce_days(row.get('days_delinquent'))
        if days is None:
            continue
        out[str(acct).strip()] = days
    return out


def derive_dq_by_migration(config, snapshot_date, loan_df, grades,
                           no_score=None, workspace_root=None,
                           days_by_account=None):
    """Derive ``(dq_by_status, dq_by_pool)`` from loan-level data + extracts.

    ``loan_df`` must be the post-``calculate_cecl`` frame (it needs
    ``current_grade``, ``original_grade``, ``loan_pool``, ``current_balance``
    and ``member_number``).

    Returns two dicts shaped exactly like ``load_impaired_data``'s
    ``_read_migration_blocks`` output::

        dq_by_status = {'Improved': {'balance': float, 'pct': float}, ...}
        dq_by_pool   = {'<pool name>': {'Improved': {...}, ...}, ...}

    ``pct`` is the bucket's share of the *delinquent* total, matching the WARM
    ``DQ Data Entry`` convention.  Returns ``({}, {})`` whenever the split
    cannot be produced, so the caller can leave the existing behaviour alone.
    """
    if loan_df is None or getattr(loan_df, 'empty', True):
        return {}, {}
    required = {'current_grade', 'original_grade', 'loan_pool',
                'current_balance', 'member_number'}
    missing = required - set(loan_df.columns)
    if missing:
        print(f"    DQ migration split: loan frame missing {sorted(missing)}; skipping.")
        return {}, {}

    dq_cfg = config.get('delinquency') or {}
    if dq_cfg.get('derive_migration_split') is False:
        print("    DQ migration split: disabled by config.")
        return {}, {}
    try:
        threshold = int(dq_cfg.get('dq_threshold') or DEFAULT_DQ_THRESHOLD)
    except (TypeError, ValueError):
        threshold = DEFAULT_DQ_THRESHOLD

    if no_score is None:
        no_score = config.get('no_score_label', 'Not Reported')
    n_top = config.get('top_grades_double_drop', 3)
    try:
        n_top = int(n_top)
    except (TypeError, ValueError):
        n_top = 3

    if days_by_account is None:
        root = workspace_root or os.environ.get('CECL_WORKSPACE_ROOT') \
            or os.path.dirname(os.path.abspath(__file__))
        days_by_account = load_days_delinquent_by_account(
            config, snapshot_date, root)
    if not days_by_account:
        return {}, {}

    fico_labels, brr_labels, brr_pool_lcs = _label_sets(config, grades, no_score)
    fico_idx = {g: i for i, g in enumerate(fico_labels)}
    brr_idx = {g: i for i, g in enumerate(brr_labels)} if brr_labels else None

    def _index_for(pool):
        if brr_idx and str(pool or '').strip().lower() in brr_pool_lcs:
            return brr_idx
        return fico_idx

    def _blank():
        return {lbl: 0.0 for lbl in STATUS_LABELS}

    grand = _blank()
    by_pool = {}
    matched = 0
    n_dq = 0
    unclassified = 0.0

    for acct, pool, cg, og, bal in zip(
        loan_df['member_number'], loan_df['loan_pool'],
        loan_df['current_grade'], loan_df['original_grade'],
        loan_df['current_balance'],
    ):
        pool_key = str(pool).strip() if pool is not None else ''
        if pool_key and pool_key.lower() != 'nan':
            by_pool.setdefault(pool_key, _blank())
        days = days_by_account.get(str(acct).strip())
        if days is None:
            continue
        matched += 1
        if days < threshold:
            continue
        try:
            b = float(bal)
        except (TypeError, ValueError):
            continue
        if b != b:  # NaN
            continue
        n_dq += 1
        status = classify_migration(cg, og, _index_for(pool_key), n_top, no_score)
        if status is None:
            # Grade outside this pool's matrix (e.g. a BRR rating on a
            # non-BRR pool). ``_sheet_risk_change`` folds anything the
            # matrix does not carry into the Unchanged residual
            # (``unc_bal = total - imp_bal - det_bal``); do the same so
            # every delinquent dollar lands in a slice.
            status = 'Unchanged'
            unclassified += b
        grand[status] += b
        if pool_key and pool_key.lower() != 'nan':
            by_pool[pool_key][status] += b

    total_loans = len(loan_df)
    if not matched:
        print(f"    DQ migration split: no loan matched an extract row "
              f"(0 of {total_loans:,}); leaving DQ charts unpopulated.")
        return {}, {}
    total_dq = sum(grand.values())
    if total_dq <= 0:
        print(f"    DQ migration split: {matched:,} of {total_loans:,} loan(s) "
              f"matched an extract row, none >= {threshold} days delinquent.")
        return {}, {}

    share = (n_dq / matched) if matched else 0.0
    if share > MAX_PLAUSIBLE_DQ_SHARE:
        print(f"    *** DQ migration split REFUSED: {n_dq:,} of {matched:,} "
              f"matched loan(s) ({share*100:.1f}%) came back >= {threshold} days "
              f"delinquent. That is implausible and almost always means an "
              f"extract's 'days_delinquent' column is mis-mapped. Leaving the "
              f"Risk Change DQ charts unpopulated rather than publishing "
              f"wrong numbers.")
        return {}, {}

    cover = 100.0 * matched / total_loans if total_loans else 0.0
    print(f"    DQ migration split derived from loan extract(s): {n_dq:,} loan(s) "
          f">= {threshold} days, ${total_dq:,.2f} delinquent "
          f"({matched:,}/{total_loans:,} loans matched, {cover:.1f}% coverage).")
    if cover < 95.0:
        print(f"    *** WARNING: {total_loans - matched:,} loan(s) matched no "
              f"extract row, so their delinquency is treated as $0 in the "
              f"Risk Change DQ pie. Check the 'using <file>' lines above for a "
              f"stale / wrong-month extract.")
    if unclassified:
        print(f"    DQ migration split: ${unclassified:,.2f} of delinquent "
              f"balance carried a grade outside its pool's matrix and was "
              f"counted as Unchanged.")

    def _shape(counts):
        tot = sum(counts.values())
        return {
            lbl: {
                'balance': round(counts[lbl], 2),
                'pct': (counts[lbl] / tot) if tot else 0.0,
            }
            for lbl in STATUS_LABELS
        }

    return _shape(grand), {p: _shape(c) for p, c in by_pool.items()}


def fill_missing_dq_migration(hist, config, snapshot_date, loan_df, grades,
                              no_score=None, workspace_root=None):
    """Populate ``hist['impaired']['dq_by_status'|'dq_by_pool']`` if absent.

    A WARM-supplied value always wins: when either key already holds a
    non-empty dict this is a no-op, so credit unions whose analyst-maintained
    ``DQ Data Entry`` tab is reachable are completely unaffected.

    Returns ``True`` when it filled something.
    """
    if hist is None:
        return False
    imp = hist.get('impaired') or {}
    have_status = bool(imp.get('dq_by_status'))
    have_pool = bool(imp.get('dq_by_pool'))
    if have_status and have_pool:
        return False

    status, by_pool = derive_dq_by_migration(
        config, snapshot_date, loan_df, grades,
        no_score=no_score, workspace_root=workspace_root)
    if not status and not by_pool:
        return False

    filled = []
    if not have_status and status:
        imp['dq_by_status'] = status
        filled.append('dq_by_status')
    if not have_pool and by_pool:
        imp['dq_by_pool'] = by_pool
        filled.append(f'dq_by_pool ({len(by_pool)} pools)')
    if not filled:
        return False
    imp['dq_source'] = 'derived_from_loan_extract'
    hist['impaired'] = imp
    print(f"    DQ migration split: filled {', '.join(filled)} "
          f"(no WARM 'DQ Data Entry' source for this snapshot).")
    return True
