"""
CECL Calculation Engine
Handles credit grading, migration analysis, and expected loss calculations.
"""
import pandas as pd
import numpy as np
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


def build_grade_order(grades):
    """Build a rank ordering dict from grade config list (first = best)."""
    return {g['label']: i + 1 for i, g in enumerate(grades)}


def calculate_cecl(df, grades, no_score_label="Not Reported"):
    """
    Apply full CECL calculations to a loan DataFrame.

    Input df must have: member_number, current_balance, current_fico_score,
                        original_fico_score, loan_pool

    Returns df with added columns: current_grade, original_grade, migration_status,
                                    reserve_rate, expected_loss_amount
    """
    grade_order = build_grade_order(grades)

    df = df.copy()
    df['current_grade'] = df['current_fico_score'].apply(
        lambda s: assign_credit_grade(s, grades, no_score_label)
    )
    df['original_grade'] = df['original_fico_score'].apply(
        lambda s: assign_credit_grade(s, grades, no_score_label)
    )
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


def risk_change_matrix(df, grades, no_score_label="Not Reported"):
    """
    Build a credit migration matrix: rows=current grade, cols=original grade.
    Values are sum of current_balance.
    """
    # Dedupe while preserving order — some configs include the no-score label
    # both in the grade list and as no_score_label, which would otherwise
    # produce duplicate index/column entries (and break .loc scalar access).
    seen: set[str] = set()
    grade_labels: list[str] = []
    for label in [g['label'] for g in grades] + [no_score_label]:
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
