"""
CECL Calculation Engine
Handles credit grading, migration analysis, and expected loss calculations.
"""
import pandas as pd
import numpy as np
import re
from datetime import date


def assign_credit_grade(fico_score, grades, no_score_label="Not Reported"):
    """Assign a credit grade label based on FICO score and grade config."""
    if pd.isna(fico_score) or fico_score == 0:
        return no_score_label
    score = int(fico_score)
    for g in grades:
        if g['min_score'] <= score <= g['max_score']:
            return g['label']
    return no_score_label


# ── Business Risk Rating (BRR) bucket evaluator ──────────────────────
#
# Some CUs run their commercial / business pools off an analyst-assigned
# Risk Rating instead of a FICO score. The wizard captures a list of
# ``business_risk_ratings`` rows (``{label, criteria}``) on the Credit
# Grades step plus a per-pool ``brr: true`` flag on the Loan Pools step,
# and writes the per-loan raw rating value to ``monthly_loan_data
# .business_risk_rating`` via the Column Mappings step. At grade time we
# walk each rule's ``criteria`` string and return the first matching
# label. Supported criteria forms (all whitespace-stripped):
#
#   ``<=N``, ``>=N``, ``<N``, ``>N``, ``=N``  numeric comparison
#   ``N``                                       exact numeric equality
#   ``A-B``                                     inclusive numeric range
#   ``Pass``                                    case-insensitive text
#                                               equality (any non-numeric)
#
# Rules with empty / unparseable criteria are skipped silently.
_BRR_NUM_RE = re.compile(r"^([<>]=?|=)?\s*(-?\d+(?:\.\d+)?)$")
_BRR_RANGE_RE = re.compile(r"^(-?\d+(?:\.\d+)?)\s*-\s*(-?\d+(?:\.\d+)?)$")


def _brr_value_as_number(value):
    if value is None:
        return None
    if isinstance(value, (int, float)) and not (
        isinstance(value, float) and pd.isna(value)
    ):
        return float(value)
    s = str(value).strip()
    if not s:
        return None
    try:
        return float(s)
    except (TypeError, ValueError):
        return None


def _brr_matches(value, criteria):
    """Return True when ``value`` satisfies the BRR ``criteria`` string."""
    crit = str(criteria or "").strip()
    if not crit:
        return False
    num = _brr_value_as_number(value)
    m = _BRR_NUM_RE.match(crit)
    if m:
        if num is None:
            return False
        op = m.group(1) or "="
        target = float(m.group(2))
        if op == "<=":
            return num <= target
        if op == ">=":
            return num >= target
        if op == "<":
            return num < target
        if op == ">":
            return num > target
        return num == target  # "=" or bare number
    m = _BRR_RANGE_RE.match(crit)
    if m:
        if num is None:
            return False
        lo, hi = float(m.group(1)), float(m.group(2))
        if lo > hi:
            lo, hi = hi, lo
        return lo <= num <= hi
    # Fall back to case-insensitive text equality.
    if value is None:
        return False
    if isinstance(value, float) and pd.isna(value):
        return False
    return str(value).strip().lower() == crit.lower()


def assign_business_risk_grade(value, brr_rules,
                               no_score_label="Not Reported"):
    """Assign a BRR bucket label by walking ``brr_rules`` in order.

    ``brr_rules`` is the list of ``{label, criteria}`` rows persisted by
    the wizard. The first rule whose ``criteria`` accepts ``value`` wins;
    if none match (or the row's value is missing/blank), return
    ``no_score_label`` so the loan still surfaces on the report under a
    well-known bucket.
    """
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return no_score_label
    if isinstance(value, str) and not value.strip():
        return no_score_label
    for rule in (brr_rules or []):
        lbl = str((rule or {}).get('label') or '').strip()
        if not lbl:
            continue
        if _brr_matches(value, (rule or {}).get('criteria')):
            return lbl
    return no_score_label


def get_reserve_rate(grade_label, grades):
    """Get the reserve rate for a given grade label."""
    for g in grades:
        if g['label'] == grade_label:
            return g['reserve_rate']
    return 0.0


def determine_migration_status(original_grade, current_grade, grade_order):
    """
    Determine if a loan's credit has Improved, Deteriorated, or is Unchanged.
    grade_order: dict mapping grade_label -> rank (lower = better, e.g., A+=1, A=2, ...)
    """
    if original_grade not in grade_order or current_grade not in grade_order:
        return "Unchanged"
    orig_rank = grade_order[original_grade]
    curr_rank = grade_order[current_grade]
    if curr_rank < orig_rank:
        return "Improved"
    elif curr_rank > orig_rank:
        return "Deteriorated"
    return "Unchanged"


def build_grade_order(grades, brr_rules=None):
    """Build a rank ordering dict from grade config list (first = best).

    When ``brr_rules`` is provided, Business Risk Rating labels are
    appended to the ordering so ``determine_migration_status`` can
    compare BRR original/current grades for BRR-flagged pools. BRR
    labels rank after FICO labels (the two domains never coexist on
    the same loan, so absolute rank values don't matter — only
    relative order within each domain).
    """
    order = {g['label']: i + 1 for i, g in enumerate(grades)}
    if brr_rules:
        # Dedupe BRR labels case-insensitively while preserving order.
        seen_lc: set[str] = set()
        next_rank = len(order) + 1
        for rule in brr_rules:
            lbl = (rule or {}).get('label')
            if not lbl:
                continue
            key = str(lbl).strip().lower()
            if not key or key in seen_lc:
                continue
            seen_lc.add(key)
            order.setdefault(lbl, next_rank)
            next_rank += 1
    return order


def calculate_cecl(df, grades, no_score_label="Not Reported",
                   brr_rules=None, brr_pools=None,
                   prior_brr_lookup=None):
    """
    Apply full CECL calculations to a loan DataFrame.

    Input df must have: member_number, current_balance, current_fico_score,
                        original_fico_score, loan_pool

    Optional ``business_risk_rating`` column is consumed when ``brr_pools``
    is non-empty: rows whose ``loan_pool`` is in that set are graded by
    matching their BRR value against ``brr_rules`` (a list of
    ``{label, criteria}`` dicts) instead of by FICO.

    ``prior_brr_lookup``: optional ``{member_number_str: prior_brr_value}``
    map. When provided, BRR-graded loans whose ``member_number`` is found
    in the lookup receive ``original_grade`` derived from that prior raw
    BRR value (run through the same ``brr_rules`` so label normalization
    matches). Loans not in the lookup (e.g. originated since the prior
    snapshot, or BRR-pool loans on the very first report for a CU) keep
    ``original_grade == current_grade``. When the param is None or empty,
    BRR-graded loans receive identical ``original_grade`` (baseline mode).

    The reserve_rate for BRR labels is looked up against ``grades`` first
    by label, falling back to the median rate when the label isn't
    represented in the FICO grade list.

    Returns df with added columns: current_grade, original_grade, migration_status,
                                    reserve_rate, expected_loss_amount
    """
    grade_order = build_grade_order(grades, brr_rules=brr_rules)

    df = df.copy()
    brr_pools = set(brr_pools or [])
    brr_rules = list(brr_rules or [])
    prior_brr_lookup = prior_brr_lookup or {}

    # FICO path (always computed — keeps non-BRR pools and BRR pools
    # that happen to also have a FICO score on file rendering normally).
    df['current_grade'] = df['current_fico_score'].apply(
        lambda s: assign_credit_grade(s, grades, no_score_label)
    )
    df['original_grade'] = df['original_fico_score'].apply(
        lambda s: assign_credit_grade(s, grades, no_score_label)
    )

    # BRR override for the configured business pools. Done as a row-wise
    # mask so we only touch loans whose ``loan_pool`` is in ``brr_pools``;
    # missing ``business_risk_rating`` column on legacy DBs is handled by
    # treating every value as None (→ no_score_label).
    if brr_pools and brr_rules and 'loan_pool' in df.columns:
        brr_col = (
            df['business_risk_rating']
            if 'business_risk_rating' in df.columns
            else pd.Series([None] * len(df), index=df.index)
        )
        mask = df['loan_pool'].isin(brr_pools)
        if mask.any():
            current_labels = brr_col[mask].apply(
                lambda v: assign_business_risk_grade(
                    v, brr_rules, no_score_label
                )
            )
            df.loc[mask, 'current_grade'] = current_labels
            if prior_brr_lookup and 'member_number' in df.columns:
                # Per-loan prior BRR derived from the most recent prior
                # snapshot in the DB. Loans not present in the lookup
                # (newly originated or first-report baseline) fall back
                # to the current label so migration shows Unchanged.
                def _prior_label(row):
                    key = row.get('member_number')
                    if key is None:
                        prior_raw = None
                    else:
                        prior_raw = prior_brr_lookup.get(str(key))
                        if prior_raw is None:
                            # Some DBs may store member_number as int;
                            # try the numeric form too.
                            try:
                                prior_raw = prior_brr_lookup.get(str(int(float(key))))
                            except (TypeError, ValueError):
                                prior_raw = None
                    if prior_raw is None:
                        return None
                    return assign_business_risk_grade(
                        prior_raw, brr_rules, no_score_label
                    )
                prior_series = df.loc[mask].apply(_prior_label, axis=1)
                # Where prior label is None (loan not in lookup) keep
                # current_grade as original_grade so migration is
                # Unchanged for new loans.
                fallback = current_labels
                resolved = prior_series.where(
                    prior_series.notna(), fallback
                )
                df.loc[mask, 'original_grade'] = resolved
            else:
                df.loc[mask, 'original_grade'] = current_labels

    df['migration_status'] = df.apply(
        lambda row: determine_migration_status(
            row['original_grade'], row['current_grade'], grade_order
        ), axis=1
    )
    df['reserve_rate'] = df['current_grade'].apply(
        lambda g: get_reserve_rate(g, grades)
    )
    # For "Not Reported" scores, use the median reserve rate as a conservative estimate
    median_rate = np.median([g['reserve_rate'] for g in grades])
    df.loc[df['current_grade'] == no_score_label, 'reserve_rate'] = median_rate

    df['expected_loss_amount'] = df['current_balance'] * df['reserve_rate']

    return df


def risk_change_matrix(df, grades, no_score_label="Not Reported", labels=None):
    """
    Build a credit migration matrix: rows=current grade, cols=original grade.
    Values are sum of current_balance.

    ``labels`` (optional): explicit ordered list of grade/rating labels. When
    provided, these replace the FICO labels derived from ``grades``. This is
    used by BRR-flagged pools so the matrix indexes by Business Risk Rating
    labels (Pass / Watch / Substandard / ...) instead of FICO grades.
    """
    # Dedupe while preserving order — some configs include the no-score label
    # both in the grade list and as no_score_label, which would otherwise
    # produce duplicate index/column entries (and break .loc scalar access).
    seen: set[str] = set()
    grade_labels: list[str] = []
    if labels is not None:
        source_labels = list(labels)
        if no_score_label not in source_labels:
            source_labels.append(no_score_label)
    else:
        source_labels = [g['label'] for g in grades] + [no_score_label]
    for label in source_labels:
        if label not in seen:
            seen.add(label)
            grade_labels.append(label)
    matrix = pd.DataFrame(0.0, index=grade_labels, columns=grade_labels)
    for _, row in df.iterrows():
        og = row.get('original_grade', no_score_label)
        cg = row.get('current_grade', no_score_label)
        if og in matrix.columns and cg in matrix.index:
            matrix.loc[cg, og] += float(row['current_balance'])
    return matrix


def pool_summary(df):
    """Summarize by loan pool: count, total balance, total reserve, avg rate."""
    summary = df.groupby('loan_pool').agg(
        loan_count=('member_number', 'count'),
        total_balance=('current_balance', 'sum'),
        total_reserve=('expected_loss_amount', 'sum'),
    ).reset_index()
    summary['reserve_pct'] = np.where(
        summary['total_balance'] > 0,
        summary['total_reserve'] / summary['total_balance'],
        0
    )
    return summary.sort_values('total_balance', ascending=False)


def migration_summary_by_pool(df):
    """For each pool, summarize improved/deteriorated/unchanged counts and balances."""
    summary = df.groupby(['loan_pool', 'migration_status']).agg(
        count=('member_number', 'count'),
        balance=('current_balance', 'sum'),
    ).reset_index()
    return summary


def grade_distribution(df):
    """Distribution of loans by current credit grade."""
    dist = df.groupby('current_grade').agg(
        loan_count=('member_number', 'count'),
        total_balance=('current_balance', 'sum'),
        total_reserve=('expected_loss_amount', 'sum'),
    ).reset_index()
    dist['reserve_pct'] = np.where(
        dist['total_balance'] > 0,
        dist['total_reserve'] / dist['total_balance'],
        0
    )
    return dist


def trend_data(df_all, credit_union):
    """
    Given a DataFrame with all periods for a CU, compute per-period totals.
    Returns a DataFrame with one row per snapshot_date.
    """
    cu_data = df_all[df_all['credit_union'] == credit_union].copy()
    trend = cu_data.groupby('snapshot_date').agg(
        total_balance=('current_balance', 'sum'),
        total_reserve=('expected_loss_amount', 'sum'),
        loan_count=('member_number', 'count'),
        improved_count=('migration_status', lambda x: (x == 'Improved').sum()),
        deteriorated_count=('migration_status', lambda x: (x == 'Deteriorated').sum()),
        unchanged_count=('migration_status', lambda x: (x == 'Unchanged').sum()),
    ).reset_index()
    trend['reserve_pct'] = np.where(
        trend['total_balance'] > 0,
        trend['total_reserve'] / trend['total_balance'],
        0
    )
    trend = trend.sort_values('snapshot_date')
    return trend


def years_on_books(open_date_str, snapshot_date):
    """Calculate years a loan has been on books."""
    try:
        if pd.isna(open_date_str):
            return 0.0
        open_dt = pd.to_datetime(open_date_str)
        snap_dt = pd.to_datetime(snapshot_date)
        return max(0, (snap_dt - open_dt).days / 365.25)
    except Exception:
        return 0.0


def principal_paid(current_balance, original_amount):
    """Calculate dollar and percent of principal paid."""
    try:
        orig = float(original_amount) if pd.notna(original_amount) else 0
        curr = float(current_balance) if pd.notna(current_balance) else 0
        paid = max(0, orig - curr)
        pct = paid / orig if orig > 0 else 0
        return paid, pct
    except Exception:
        return 0.0, 0.0
