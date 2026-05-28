"""
CECL Report Generator - Three Report Types

Generates TCT-format and Vizo-format CECL reports.

Report Types:
  tct      - Full CECL-Migration-WARM report (Franklin Trust style, single file)
  vizo     - CECL Credit Migration Report (Credit Union B style, main report)
  vizo_supp - CECL Supplemental Report (Credit Union B style, supplemental)

Usage:
    python generate_report.py --client franklin --date 2025-12-31
    python generate_report.py --client franklin --reports tct
    python generate_report.py --client franklin --reports tct vizo vizo_supp
    python generate_report.py --all --date 2025-12-31
    python generate_report.py --list
"""
import os, re, argparse, glob
from datetime import datetime, date
import numpy as np
import pandas as pd
import yaml
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill, numbers
from openpyxl.utils import get_column_letter
from sqlalchemy import create_engine, text
from dotenv import load_dotenv
from cecl_engine import (
    calculate_cecl, risk_change_matrix, pool_summary,
    migration_summary_by_pool, grade_distribution,
    years_on_books, principal_paid, build_grade_order,
)
from report_tct import compose_tct as compose_tct_new
from report_vizo import compose_vizo_main as compose_vizo_main_new, compose_vizo_supp as compose_vizo_supp_new, patch_impdet_charts, patch_drawing_onecell_to_twocell, patch_dq_pie_zero_labels, patch_remove_chart_borders_and_axis_lines
from fetch_econ_data import fetch_economic_data
from cecl_audit_log import get_audit_logger, log_report_generation, log_session_start, log_session_end
from cecl_credentials import get_database_url

load_dotenv()
# When the code clone lives on a local SSD (e.g. C:\Dev\CECL) and the
# analyst data (client_configs/, Reports/, Raw_Uploads/, ...) lives on a
# shared drive (e.g. Egnyte), CECL_WORKSPACE_ROOT points at the data root.
# Falls back to the historical 'code dir == data dir' layout when unset.
BASE = os.environ.get('CECL_WORKSPACE_ROOT') or os.path.dirname(os.path.abspath(__file__))
CFG_DIR = os.path.join(BASE, 'client_configs')
RPT_DIR = os.path.join(BASE, 'Reports')
engine = create_engine(get_database_url())


def resolve_path(path_value, base=BASE):
    """Resolve configured paths: keep absolute paths, join relative paths to base."""
    if not path_value:
        return ''
    return path_value if os.path.isabs(path_value) else os.path.join(base, path_value)

# ── Styling Constants ──────────────────────────────────────────────
TITLE_FONT = Font(name='Calibri', bold=True, size=18, color='1B4F72')
HDR_FONT = Font(name='Calibri', bold=True, size=10, color='FFFFFF')
SUB_FONT = Font(name='Calibri', bold=True, size=11)
NORM = Font(name='Calibri', size=10)
MONEY = '#,##0'
MONEY2 = '#,##0.00'
PCT = '0.00%'
PCT4 = '0.0000%'
HDR_FILL = PatternFill('solid', fgColor='1B4F72')
ALT_FILL = PatternFill('solid', fgColor='D6EAF8')
IMP_FILL = PatternFill('solid', fgColor='D5F5E3')
DET_FILL = PatternFill('solid', fgColor='FADBD8')
THIN = Border(left=Side('thin'), right=Side('thin'), top=Side('thin'), bottom=Side('thin'))
TEAL_FILL = PatternFill('solid', fgColor='1A5276')
DARK_FILL = PatternFill('solid', fgColor='2C3E50')

# ── Environmental Factor Score Tables ──────────────────────────────
# Net Credit Change scoring (same in TCT and Vizo)
NCC_RANGES = [
    (-999, -18.00, 7), (-18.00, -16.00, 6), (-16.00, -14.00, 5),
    (-14.00, -11.00, 4), (-11.00, -8.00, 3), (-8.00, -6.00, 2),
    (-6.00, -4.00, 1), (-4.00, 4.00, 0), (4.00, 6.00, -1),
    (6.00, 8.00, -2), (8.00, 9.00, -3), (9.00, 11.00, -4),
    (11.00, 13.00, -5), (13.00, 15.00, -6), (15.00, 999, -7),
]
# Delinquency scoring
DQ_RANGES = [
    (5.00, 999, 20), (4.00, 5.00, 17), (3.00, 4.00, 12),
    (2.50, 3.00, 8), (2.00, 2.50, 4), (1.50, 2.00, 2.5),
    (1.00, 1.50, 1.5), (0.50, 1.00, 0.75), (-0.50, 0.50, 0),
    (-1.00, -0.50, -0.75), (-1.50, -1.00, -1.5), (-2.00, -1.50, -2.5),
    (-2.50, -2.00, -4), (-3.00, -2.50, -8), (-4.00, -3.00, -12),
    (-5.00, -4.00, -17), (-999, -5.00, -20),
]
# Economic Stress scoring
ES_RANGES = [
    (25.00, 999, 10), (24.00, 25.00, 8), (22.00, 24.00, 7),
    (20.00, 22.00, 6), (18.00, 20.00, 5), (16.00, 18.00, 4),
    (14.00, 16.00, 3.5), (12.00, 14.00, 3), (10.00, 12.00, 2),
    (8.00, 10.00, 1), (6.00, 8.00, 0), (4.00, 6.00, 0),
    (2.00, 4.00, -1), (0.00, 2.00, -2),
]
# Standard TCT Distribution Factors per grade position
DIST_FACTORS = [10.52, 22.93, 45.15, 116.10, 141.17, 152.04, 160.21]


def score_from_ranges(value, ranges):
    """Look up a score from a range table."""
    v = value * 100 if abs(value) < 1 else value  # handle both 0.05 and 5.0
    for lo, hi, score in ranges:
        if lo <= v < hi:
            return score
    return 0


# ── Data Loading ───────────────────────────────────────────────────
def load_config(client):
    with open(os.path.join(CFG_DIR, f'{client}.yaml'), 'r', encoding='utf-8') as f:
        cfg = yaml.safe_load(f)
    # Normalize the 'Ignore' sentinel: any pool_map value (or default_pool)
    # set to 'Ignore' is rewritten to the existing 'Exclude' sentinel, which
    # is already dropped everywhere downstream (per-pool sheets,
    # _apply_excluded_pools, HIDE/Exclude prefix filters).
    pm = cfg.get('pool_map') or {}
    if any(v == 'Ignore' for v in pm.values()):
        cfg['pool_map'] = {
            k: ('Exclude' if v == 'Ignore' else v) for k, v in pm.items()
        }
    if cfg.get('default_pool') == 'Ignore':
        cfg['default_pool'] = 'Exclude'
    # Honor excluded_pools by remapping any pool_map value matching an
    # excluded pool name to the existing 'Exclude' sentinel. All downstream
    # filters (HIDE/Exclude prefix checks throughout report_tct/generate_report)
    # already drop 'Exclude' rows from balances, charge-offs, recoveries, and
    # per-pool sheets, so this single rewrite excludes them from the entire
    # analysis without touching every individual filter site.
    excl = set((cfg.get('excluded_pools') or []))
    if excl:
        pm = cfg.get('pool_map') or {}
        cfg['pool_map'] = {
            k: ('Exclude' if v in excl else v) for k, v in pm.items()
        }
        if cfg.get('default_pool') in excl:
            cfg['default_pool'] = 'Exclude'
        # Also remove excluded pools from the pool registries so the report
        # engine's pool enumerators (which iterate ``cfg['pools']`` /
        # ``risk_rated`` / ``not_risk_rated`` directly) don't render empty
        # columns/rows/sheets for them. Without this, _apply_excluded_pools
        # drops the *loan rows* but the *pool name* still shows up as a
        # zero-balance bucket on sheets like "ACL Env by Pool Mgmt Adj".
        pools_list = cfg.get('pools')
        if isinstance(pools_list, list):
            cfg['pools'] = [
                p for p in pools_list
                if not (
                    (isinstance(p, dict) and p.get('name') in excl)
                    or (isinstance(p, str) and p in excl)
                )
            ]
        for key in ('risk_rated', 'not_risk_rated', 'pool_order'):
            val = cfg.get(key)
            if isinstance(val, list):
                cfg[key] = [p for p in val if p not in excl]
    return cfg

def list_clients():
    return sorted(os.path.splitext(f)[0] for f in os.listdir(CFG_DIR)
                  if f.endswith('.yaml') and not f.startswith('_'))

def load_loans(cu, snap=None, config=None):
    if snap:
        df = pd.read_sql(text("SELECT * FROM monthly_loan_data WHERE credit_union=:c AND snapshot_date=:s"),
                         engine, params={"c": cu, "s": snap})
    else:
        df = pd.read_sql(text("SELECT * FROM monthly_loan_data WHERE credit_union=:c"), engine, params={"c": cu})
    return _apply_excluded_pools(df, config)


def _apply_excluded_pools(df, config):
    """Drop any loan rows whose ``loan_pool`` is listed in
    ``cfg['excluded_pools']`` (and any rows already tagged with the legacy
    'Exclude' sentinel from import-time pool_map remapping).

    Existing DB rows were categorized at import time using the original
    pool_map, so configs that excluded pools *after* import would otherwise
    still include those rows in every per-pool sheet that just iterates
    ``df['loan_pool'].unique()``. Dropping them here makes the runtime
    config authoritative across the whole report.
    """
    if df is None or df.empty or 'loan_pool' not in df.columns:
        return df
    excl = set((config.get('excluded_pools') or [])) if config else set()
    excl.add('Exclude')
    mask = df['loan_pool'].isin(excl)
    if not mask.any():
        return df
    return df.loc[~mask].copy()

def latest_date(cu):
    with engine.connect() as c:
        r = c.execute(text("SELECT MAX(snapshot_date) FROM monthly_loan_data WHERE credit_union=:c"), {"c": cu}).fetchone()
    return str(r[0]) if r and r[0] else None

def all_dates(cu):
    with engine.connect() as c:
        rows = c.execute(text("SELECT DISTINCT snapshot_date FROM monthly_loan_data WHERE credit_union=:c ORDER BY snapshot_date DESC"), {"c": cu}).fetchall()
    return [str(r[0]) for r in rows]


# ── Historical Data Loading ───────────────────────────────────────

def _find_quarter_folders(data_dir):
    """Find all quarterly data folders under the data directory.
    Returns list of (folder_path, quarter_label) sorted by date."""
    quarters = []
    for root, dirs, files in os.walk(data_dir):
        folder = os.path.basename(root)
        # Match patterns like 2024-03, 2024-06, 2022-10, 2022-12, 2023-03, etc.
        m = re.match(r'^(\d{4})-(\d{2})$', folder)
        if m:
            quarters.append((root, f"{m.group(1)}-{m.group(2)}"))
    return sorted(quarters, key=lambda x: x[1])


def _read_data_file(filepath):
    """Read an Excel or CSV file, returning a DataFrame with no header.
    For Excel files with multiple sheets, concatenates all sheets that share
    the maximum column count (so multi-sheet quarterly files are fully read).
    """
    ext = os.path.splitext(filepath)[1].lower()
    if ext == '.csv':
        return pd.read_csv(filepath, header=None)
    xl = pd.ExcelFile(filepath)
    parts = []
    for s in xl.sheet_names:
        d = pd.read_excel(xl, sheet_name=s, header=None)
        if not d.empty:
            parts.append(d)
    if not parts:
        return pd.DataFrame()
    if len(parts) == 1:
        return parts[0]
    target_cols = max(d.shape[1] for d in parts)
    parts = [d for d in parts if d.shape[1] == target_cols]
    return pd.concat(parts, ignore_index=True)


def _looks_like_loan_code(val):
    """Return True if val looks like a 1-5 char loan code (has at least one letter, not pure digits)."""
    if not isinstance(val, str):
        return False
    s = val.strip()
    if not (1 <= len(s) <= 5):
        return False
    if s.isdigit():
        return False
    return any(ch.isalpha() for ch in s)


def _coerce_mixed_dates(values):
    """Parse a column that mixes ``datetime`` objects and Excel serial
    numbers (e.g. 41091 -> 2012-06-30) into a tz-naive Timestamp Series.

    Honolulu's CO / Recovery tracking workbooks have both formats
    interleaved on the same column, which ``pd.to_datetime`` alone
    misinterprets (it treats raw ints as nanoseconds and produces 1970).
    """
    s = pd.Series(values)
    # Identify cells that are numeric in the Excel-serial range
    # (~1927..~2119) up-front and route them through the Excel-origin
    # parser. Everything else goes through the standard parser.
    nums = pd.to_numeric(s, errors='coerce')
    serial_mask = nums.between(10000, 80000) & ~s.apply(
        lambda v: isinstance(v, (pd.Timestamp, datetime)))
    out = pd.to_datetime(s.mask(serial_mask), errors='coerce')
    if serial_mask.any():
        converted = pd.to_datetime(
            nums[serial_mask], unit='D', origin='1899-12-30', errors='coerce')
        out.loc[converted.index] = converted
    return out


def _is_numeric_or_date(v):
    """True if ``v`` is a non-null number, datetime, or numeric-looking
    string. Used by the account-column filter so that rows whose
    account_col is actually a date (e.g. recovery files with no
    account number) are not dropped.
    """
    if pd.isna(v):
        return False
    if isinstance(v, (pd.Timestamp, datetime)):
        return True
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        return v.replace('-', '').replace(' ', '').isdigit()
    return False


def _parse_chargeoff_file(filepath, parse_config=None):
    """Parse a charge-off file (varying formats). Returns DataFrame with [code, amount, date]."""
    df = _read_data_file(filepath)
    if df.empty:
        return pd.DataFrame(columns=['code', 'amount', 'date'])

    if parse_config:
        cfg_df = df.copy()
        if parse_config.get('has_header', False):
            cfg_df = cfg_df.iloc[1:]

        skip_rows = int(parse_config.get('skip_rows', 0) or 0)
        if skip_rows > 0:
            cfg_df = cfg_df.iloc[skip_rows:]

        account_col = parse_config.get('account_col', 0)
        code_col = parse_config.get('code_col')
        amount_col = parse_config.get('amount_col')
        date_col = parse_config.get('date_col')

        ncols = cfg_df.shape[1]
        code_valid = code_col is not None and 0 <= int(code_col) < ncols
        amount_valid = amount_col is not None and 0 <= int(amount_col) < ncols
        date_valid = date_col is not None and 0 <= int(date_col) < ncols
        account_valid = account_col is not None and 0 <= int(account_col) < ncols

        if code_valid and amount_valid:
            if account_valid:
                acct_series = cfg_df.iloc[:, account_col]
                acct_numeric = acct_series.apply(_is_numeric_or_date)
                cfg_df = cfg_df[acct_numeric]

            if not cfg_df.empty:
                result = pd.DataFrame({
                    'code': cfg_df.iloc[:, code_col].values,
                    'amount': pd.to_numeric(cfg_df.iloc[:, amount_col], errors='coerce').values,
                })
                if date_valid:
                    result['date'] = _coerce_mixed_dates(
                        cfg_df.iloc[:, date_col].values).values
                else:
                    result['date'] = pd.NaT
                result = result.dropna(subset=['amount'])
                if not result.empty:
                    return result

    # Check if first row is a header.  Require col 0 to NOT be numeric (a real
    # header has 'Account Number'-style text in col 0; a data row starts with
    # an account number).  This prevents trailing comment cells like
    # "...use Col D for charge off amounts." from triggering a false header.
    first_row = df.iloc[0]
    first_a = first_row.iloc[0] if len(first_row) > 0 else None
    col0_is_numeric = isinstance(first_a, (int, float)) and pd.notna(first_a)
    first_vals = [str(v).lower() if pd.notna(v) else '' for v in first_row]
    has_header = (not col0_is_numeric) and any(
        'account' in v or 'charge' in v or 'loan' in v or 'acct' in v
        for v in first_vals
    )

    # Use header keywords to identify columns when header is present
    code_col = amount_col = date_col = None
    if has_header:
        for c, hdr in enumerate(first_vals):
            # Skip the FICO score column (contains 'sc'/'score' but is not a loan code)
            if 'fico' in hdr or 'score' in hdr:
                continue
            if any(k in hdr for k in ('security', 'code')) and 'account' not in hdr and 'sub' not in hdr:
                code_col = c
            elif any(k in hdr for k in ('amount', 'chg off am', 'chargeoff am', 'pymt', 'principal')):
                amount_col = c
            elif 'date' in hdr or 'effective' in hdr:
                date_col = c
        df = df.iloc[1:]

    # Drop rows with NaN in the first column (empty/total rows)
    df = df.dropna(subset=[0])
    # Drop total rows (where account is not a number)
    # Handle accounts with hyphens like '65670-029'
    acct_numeric = df[0].apply(lambda v: pd.notna(v) and (
        str(v).replace('-', '').replace(' ', '').isdigit() if isinstance(v, str)
        else isinstance(v, (int, float)) and not pd.isna(v)))
    df = df[acct_numeric]

    if df.empty:
        return pd.DataFrame(columns=['code', 'amount', 'date'])

    # Identify columns by heuristic if header-based detection missed any
    ncols = df.shape[1]

    if code_col is None or amount_col is None or date_col is None:
        for c in range(ncols):
            sample = df[c].dropna()
            if sample.empty:
                continue
            first_val = sample.iloc[0]
            if code_col is None and _looks_like_loan_code(first_val):
                code_col = c
            elif date_col is None and isinstance(first_val, (pd.Timestamp, datetime)):
                date_col = c
            elif isinstance(first_val, (int, float)) and c > 1 and amount_col is None:
                amount_col = c

    if code_col is None or amount_col is None:
        # Fallback: try common layouts
        if ncols >= 5:
            if code_col is None:
                code_col = ncols - 1  # last column is often the code
            # Check if col 3 is numeric (amount) or datetime (date)
            s3 = df[3].dropna().iloc[0] if len(df[3].dropna()) > 0 else None
            if isinstance(s3, (pd.Timestamp, datetime)):
                if date_col is None:
                    date_col = 3
                if amount_col is None:
                    amount_col = 4
            else:
                if amount_col is None:
                    amount_col = 3
                if date_col is None:
                    date_col = 4
        elif ncols == 4:
            # Common 4-column layout: Account, Date, Amount, Code
            if code_col is None:
                code_col = 3
            if amount_col is None:
                amount_col = 2
            if date_col is None:
                date_col = 1

    result = pd.DataFrame({
        'code': df[code_col].values if code_col is not None else '',
        'amount': pd.to_numeric(df[amount_col], errors='coerce').values if amount_col is not None else 0,
    })
    if date_col is not None:
        result['date'] = pd.to_datetime(df[date_col].values, errors='coerce')
    else:
        result['date'] = pd.NaT

    return result.dropna(subset=['amount'])


def _parse_recovery_file(filepath, parse_config=None):
    """Parse a recovery file (varying formats). Returns DataFrame with [code, amount, date]."""
    df = _read_data_file(filepath)
    if df.empty:
        return pd.DataFrame(columns=['code', 'amount', 'date'])

    if parse_config:
        cfg_df = df.copy()
        if parse_config.get('has_header', False):
            cfg_df = cfg_df.iloc[1:]

        skip_rows = int(parse_config.get('skip_rows', 0) or 0)
        if skip_rows > 0:
            cfg_df = cfg_df.iloc[skip_rows:]

        account_col = parse_config.get('account_col', 0)
        code_col = parse_config.get('code_col')
        amount_col = parse_config.get('amount_col')
        date_col = parse_config.get('date_col')

        ncols = cfg_df.shape[1]
        code_valid = code_col is not None and 0 <= int(code_col) < ncols
        amount_valid = amount_col is not None and 0 <= int(amount_col) < ncols
        date_valid = date_col is not None and 0 <= int(date_col) < ncols
        account_valid = account_col is not None and 0 <= int(account_col) < ncols

        if code_valid and amount_valid:
            if account_valid:
                acct_series = cfg_df.iloc[:, account_col]
                acct_numeric = acct_series.apply(_is_numeric_or_date)
                cfg_df = cfg_df[acct_numeric]

            if not cfg_df.empty:
                result = pd.DataFrame({
                    'code': cfg_df.iloc[:, code_col].values,
                    'amount': pd.to_numeric(cfg_df.iloc[:, amount_col], errors='coerce').values,
                })
                if date_valid:
                    result['date'] = _coerce_mixed_dates(
                        cfg_df.iloc[:, date_col].values).values
                else:
                    result['date'] = pd.NaT
                result = result.dropna(subset=['amount'])
                if not result.empty:
                    return result

    # Check if first row is a header (see _parse_chargeoff_file for rationale)
    first_row = df.iloc[0]
    first_a = first_row.iloc[0] if len(first_row) > 0 else None
    col0_is_numeric = isinstance(first_a, (int, float)) and pd.notna(first_a)
    first_vals = [str(v).lower() if pd.notna(v) else '' for v in first_row]
    has_header = (not col0_is_numeric) and any(
        'account' in v or 'recov' in v or 'loan' in v or 'acct' in v
        for v in first_vals
    )

    # Use header keywords to identify columns when header is present
    code_col = amount_col = date_col = None
    if has_header:
        for c, hdr in enumerate(first_vals):
            # Skip the FICO score column (contains 'sc'/'score' but is not a loan code)
            if 'fico' in hdr or 'score' in hdr:
                continue
            if any(k in hdr for k in ('security', 'code')) and 'account' not in hdr and 'sub' not in hdr:
                code_col = c
            elif any(k in hdr for k in ('amount', 'pymt', 'payment', 'principal')):
                amount_col = c
            elif 'date' in hdr or 'effective' in hdr:
                date_col = c
        df = df.iloc[1:]

    df = df.dropna(subset=[0])
    # Handle accounts with hyphens like '51930-27'
    acct_numeric = df[0].apply(lambda v: pd.notna(v) and (
        str(v).replace('-', '').replace(' ', '').isdigit() if isinstance(v, str)
        else isinstance(v, (int, float)) and not pd.isna(v)))
    df = df[acct_numeric]

    if df.empty:
        return pd.DataFrame(columns=['code', 'amount', 'date'])

    ncols = df.shape[1]

    # Heuristic fallback if header-based detection missed columns
    if code_col is None or amount_col is None or date_col is None:
        for c in range(ncols):
            sample = df[c].dropna()
            if sample.empty:
                continue
            first_val = sample.iloc[0]
            if code_col is None and _looks_like_loan_code(first_val):
                code_col = c
            elif date_col is None and isinstance(first_val, (pd.Timestamp, datetime)):
                date_col = c
            elif isinstance(first_val, (int, float)) and c > 1 and amount_col is None:
                amount_col = c

    if code_col is None or amount_col is None:
        if ncols >= 5:
            if code_col is None:
                code_col = ncols - 1
            s3 = df[3].dropna().iloc[0] if len(df[3].dropna()) > 0 else None
            if isinstance(s3, (pd.Timestamp, datetime)):
                if date_col is None:
                    date_col = 3
                if amount_col is None:
                    amount_col = 4
            else:
                if amount_col is None:
                    amount_col = 3
                if date_col is None:
                    date_col = 4
        elif ncols == 4:
            if code_col is None:
                code_col = 3
            if amount_col is None:
                amount_col = 2
            if date_col is None:
                date_col = 1

    result = pd.DataFrame({
        'code': df[code_col].values if code_col is not None else '',
        'amount': pd.to_numeric(df[amount_col], errors='coerce').values if amount_col is not None else 0,
    })
    if date_col is not None:
        result['date'] = pd.to_datetime(df[date_col].values, errors='coerce')
    else:
        result['date'] = pd.NaT

    return result.dropna(subset=['amount'])


def load_chargeoff_recovery_history(config):
    """Load all historical charge-off and recovery data.
    Returns dict: {'chargeoffs': {year: {pool: amount}}, 'recoveries': {year: {pool: amount}}}"""
    data_dir = resolve_path(config.get('data_directory', ''))
    if not data_dir or not os.path.isdir(data_dir):
        return {'chargeoffs': {}, 'recoveries': {}, 'years': []}

    historical_parse_cfg = config.get('historical_file_formats', {})
    chargeoff_parse_cfg = historical_parse_cfg.get('chargeoff')
    recovery_parse_cfg = historical_parse_cfg.get('recovery')
    pool_map = config.get('pool_map', {})
    quarters = _find_quarter_folders(data_dir)

    chargeoffs = {}  # {year: {pool: amount}}
    recoveries = {}
    co_monthly = {}   # {(year, month): {pool: amount}}
    rc_monthly = {}   # {(year, month): {pool: amount}}

    # --- Detect cumulative charge-off files (Ontario-style) ---
    # These have sheets "C-Offs 3 Years" and "Recoveries 3 Years" with all data
    cumulative_files = []
    for root, dirs, files in os.walk(data_dir):
        for f in files:
            fl = f.lower()
            if ('charge' in fl and 'off' in fl) and fl.endswith('.xlsx'):
                filepath = os.path.join(root, f)
                try:
                    xl = pd.ExcelFile(filepath)
                    if 'C-Offs 3 Years' in xl.sheet_names:
                        cumulative_files.append(filepath)
                except Exception:
                    pass

    if cumulative_files:
        # Use the largest cumulative file (most recent/complete data)
        cumulative_files.sort(key=os.path.getsize, reverse=True)
        filepath = cumulative_files[0]
        print(f"    Using cumulative charge-off file: {os.path.basename(filepath)}")

        def _extract_pool_code(raw_code, pool_map):
            """Extract pool from Ontario-style loan codes like '99 / Sig', 'Visa', 99, 'LP #1'."""
            code = str(raw_code).strip()
            # Try the raw value first (handles 'Visa', 'LP #1', etc.)
            pool = pool_map.get(code) or pool_map.get(code.upper()) or pool_map.get(code.lower())
            if pool:
                return pool
            # Try numeric extraction: "99 / Sig" -> "99", "11 / New Car" -> "11"
            if ' / ' in code:
                num_part = code.split(' / ')[0].strip()
                pool = pool_map.get(num_part) or pool_map.get(num_part.upper())
                if pool:
                    return pool
            # Try as-is for pure integers
            try:
                int_code = str(int(float(code)))
                pool = pool_map.get(int_code)
                if pool:
                    return pool
            except (ValueError, TypeError):
                pass
            return None

        # Parse charge-offs sheet
        try:
            df = pd.read_excel(filepath, sheet_name='C-Offs 3 Years', header=None)
            # Find header row (contains "Charge off Amount")
            hdr_idx = 0
            data_start = 0
            for i in range(min(10, len(df))):
                row_vals = [str(v).lower() if pd.notna(v) else '' for v in df.iloc[i]]
                if any('charge off amount' in v for v in row_vals):
                    hdr_idx = i
                    data_start = i + 1
                    break
            # Skip blank rows after header
            while data_start < len(df) and df.iloc[data_start].isna().all():
                data_start += 1

            # Find column positions from header
            hdr = df.iloc[hdr_idx]
            # Determine which columns have data - skip leading NaN columns
            first_data_col = 0
            for c in range(df.shape[1]):
                if pd.notna(hdr.iloc[c]):
                    first_data_col = c
                    break

            # Columns relative to first_data_col: Member#, Suffix, Code, Amount, Date, [FICO]
            code_col = first_data_col + 2
            amount_col = first_data_col + 3
            date_col = first_data_col + 4

            for i in range(data_start, len(df)):
                raw_code = df.iloc[i, code_col]
                if pd.isna(raw_code):
                    continue
                code_str = str(raw_code).strip()
                if code_str.upper() in ('TOTAL', ''):
                    continue
                pool = _extract_pool_code(raw_code, pool_map)
                amount = pd.to_numeric(df.iloc[i, amount_col], errors='coerce')
                date_val = pd.to_datetime(df.iloc[i, date_col], errors='coerce')
                if pool and pd.notna(amount) and pd.notna(date_val) and 2000 <= date_val.year <= 2099:
                    yr = date_val.year
                    chargeoffs.setdefault(yr, {})
                    chargeoffs[yr][pool] = chargeoffs[yr].get(pool, 0) + amount
                    ym = (yr, date_val.month)
                    co_monthly.setdefault(ym, {})
                    co_monthly[ym][pool] = co_monthly[ym].get(pool, 0) + amount
        except Exception as e:
            print(f"    Warning: Could not parse charge-offs from {filepath}: {e}")

        # Parse recoveries sheet
        try:
            xl = pd.ExcelFile(filepath)
            if 'Recoveries 3 Years' in xl.sheet_names:
                df = pd.read_excel(filepath, sheet_name='Recoveries 3 Years', header=None)
                hdr_idx = 0
                data_start = 0
                for i in range(min(10, len(df))):
                    row_vals = [str(v).lower() if pd.notna(v) else '' for v in df.iloc[i]]
                    if any('recovery amount' in v for v in row_vals):
                        hdr_idx = i
                        data_start = i + 1
                        break
                while data_start < len(df) and df.iloc[data_start].isna().all():
                    data_start += 1

                hdr = df.iloc[hdr_idx]
                first_data_col = 0
                for c in range(df.shape[1]):
                    if pd.notna(hdr.iloc[c]):
                        first_data_col = c
                        break

                code_col = first_data_col + 2
                amount_col = first_data_col + 3
                date_col = first_data_col + 4

                for i in range(data_start, len(df)):
                    raw_code = df.iloc[i, code_col]
                    if pd.isna(raw_code):
                        continue
                    code_str = str(raw_code).strip()
                    if code_str.lower() in ('total', ''):
                        continue
                    pool = _extract_pool_code(raw_code, pool_map)
                    amount = pd.to_numeric(df.iloc[i, amount_col], errors='coerce')
                    date_val = pd.to_datetime(df.iloc[i, date_col], errors='coerce')
                    if pool and pd.notna(amount) and pd.notna(date_val) and 2000 <= date_val.year <= 2099:
                        yr = date_val.year
                        recoveries.setdefault(yr, {})
                        recoveries[yr][pool] = recoveries[yr].get(pool, 0) + amount
                        ym = (yr, date_val.month)
                        rc_monthly.setdefault(ym, {})
                        rc_monthly[ym][pool] = rc_monthly[ym].get(pool, 0) + amount
        except Exception as e:
            print(f"    Warning: Could not parse recoveries from {filepath}: {e}")

    else:
        # --- Franklin-style: per-quarter charge-off/recovery files ---
        # Build a string-keyed pool map that handles numeric codes
        str_pool_map = {str(k).strip(): v for k, v in pool_map.items()}

        def _lookup_pool(raw_code):
            """Look up pool from a code value (numeric or text)."""
            code = str(raw_code).strip()
            # Try as-is
            pool = str_pool_map.get(code) or str_pool_map.get(code.upper()) or str_pool_map.get(code.lower())
            if pool:
                return pool
            # Try integer conversion (handles 28.0 -> "28")
            try:
                int_code = str(int(float(code)))
                pool = str_pool_map.get(int_code)
                if pool:
                    return pool
            except (ValueError, TypeError):
                pass
            # Try matching against pool names (for text codes like "visa" -> pool "VISA")
            code_upper = code.upper()
            for v in set(str_pool_map.values()):
                if v.upper() == code_upper:
                    return v
            return None

        # Fallback for flat folders (no YYYY-MM quarter subfolders):
        # treat ``data_dir`` itself as a single bucket so wizard-style
        # CECL setups (Raw_Uploads/<short>/*.xlsx) get picked up. The
        # per-row date column (configured in
        # ``historical_file_formats``) supplies the actual year/month;
        # ``qlabel`` is only used as a fallback when row-level date
        # parsing fails.
        if not quarters:
            from datetime import datetime as _dt
            quarters = [(data_dir, f"{_dt.today().year}-12")]
            print(f"    No YYYY-MM quarter subfolders under {data_dir}; "
                  f"scanning top-level for charge-off / recovery files.")

        for folder, qlabel in quarters:
            year = int(qlabel[:4])

            for f in os.listdir(folder):
                fl = f.lower()
                if ('charge' in fl and 'off' in fl) and (fl.endswith('.xlsx') or fl.endswith('.csv')):
                    if 'proposed' in fl or '3yr' in fl or 'recov' in fl:
                        continue
                    filepath = os.path.join(folder, f)
                    try:
                        df = _parse_chargeoff_file(filepath, parse_config=chargeoff_parse_cfg)
                        for _, row in df.iterrows():
                            pool = _lookup_pool(row['code'])
                            if pool and pd.notna(row['amount']):
                                row_year = year
                                row_month = int(qlabel[5:7]) if len(qlabel) >= 7 else 12
                                if pd.notna(row.get('date')):
                                    try:
                                        dt = pd.to_datetime(row['date'])
                                        y = int(dt.year)
                                        if 2000 <= y <= 2099:
                                            row_year = y
                                            row_month = int(dt.month)
                                    except Exception:
                                        pass
                                chargeoffs.setdefault(row_year, {})
                                chargeoffs[row_year][pool] = chargeoffs[row_year].get(pool, 0) + row['amount']
                                ym = (row_year, row_month)
                                co_monthly.setdefault(ym, {})
                                co_monthly[ym][pool] = co_monthly[ym].get(pool, 0) + row['amount']
                    except Exception as e:
                        print(f"    Warning: Could not parse {filepath}: {e}")

                if ('recov' in fl) and (fl.endswith('.xlsx') or fl.endswith('.csv')):
                    if '3yr' in fl:
                        continue
                    filepath = os.path.join(folder, f)
                    try:
                        df = _parse_recovery_file(filepath, parse_config=recovery_parse_cfg)
                        for _, row in df.iterrows():
                            pool = _lookup_pool(row['code'])
                            if pool and pd.notna(row['amount']):
                                row_year = year
                                row_month = int(qlabel[5:7]) if len(qlabel) >= 7 else 12
                                if pd.notna(row.get('date')):
                                    try:
                                        dt = pd.to_datetime(row['date'])
                                        y = int(dt.year)
                                        if 2000 <= y <= 2099:
                                            row_year = y
                                            row_month = int(dt.month)
                                    except Exception:
                                        pass
                                recoveries.setdefault(row_year, {})
                                recoveries[row_year][pool] = recoveries[row_year].get(pool, 0) + row['amount']
                                ym = (row_year, row_month)
                                rc_monthly.setdefault(ym, {})
                                rc_monthly[ym][pool] = rc_monthly[ym].get(pool, 0) + row['amount']
                    except Exception as e:
                        print(f"    Warning: Could not parse {filepath}: {e}")

        # Also check for 3yr file (covers 2019-2022 Q3)
        for root, dirs, files in os.walk(data_dir):
            for f in files:
                if '3yr' in f.lower() and f.endswith('.xlsx'):
                    filepath = os.path.join(root, f)
                    try:
                        df = _parse_chargeoff_file(filepath)
                        for _, row in df.iterrows():
                            code = str(row['code']).strip().upper()
                            pool = pool_map.get(code, pool_map.get(code.lower(), None))
                            if pool and pd.notna(row['amount']) and pd.notna(row['date']):
                                yr = row['date'].year
                                chargeoffs.setdefault(yr, {})
                                chargeoffs[yr][pool] = chargeoffs[yr].get(pool, 0) + row['amount']
                                ym = (yr, int(row['date'].month))
                                co_monthly.setdefault(ym, {})
                                co_monthly[ym][pool] = co_monthly[ym].get(pool, 0) + row['amount']
                    except Exception as e:
                        print(f"    Warning: Could not parse 3yr file {filepath}: {e}")

        # Also check the standalone recovery file in 2022-10
        for root, dirs, files in os.walk(data_dir):
            for f in files:
                if f.lower() == 'recovery.xlsx':
                    filepath = os.path.join(root, f)
                    try:
                        df = _parse_recovery_file(filepath)
                        for _, row in df.iterrows():
                            code = str(row['code']).strip().upper()
                            pool = pool_map.get(code, pool_map.get(code.lower(), None))
                            if pool and pd.notna(row['amount']) and pd.notna(row['date']):
                                yr = row['date'].year
                                recoveries.setdefault(yr, {})
                                recoveries[yr][pool] = recoveries[yr].get(pool, 0) + row['amount']
                                ym = (yr, int(row['date'].month))
                                rc_monthly.setdefault(ym, {})
                                rc_monthly[ym][pool] = rc_monthly[ym].get(pool, 0) + row['amount']
                    except Exception as e:
                        print(f"    Warning: Could not parse recovery file {filepath}: {e}")

    all_years = sorted(set(list(chargeoffs.keys()) + list(recoveries.keys())))
    return {
        'chargeoffs': chargeoffs, 'recoveries': recoveries, 'years': all_years,
        'co_monthly': co_monthly, 'rc_monthly': rc_monthly,
    }


def _col_letter_to_idx(letter):
    """Convert an Excel column letter (A, B, ..., Z, AA, AB, ...) to a 0-based
    index. Accepts already-numeric values as well."""
    if letter is None:
        return 0
    if isinstance(letter, (int, float)):
        return int(letter)
    s = str(letter).strip().upper()
    if not s:
        return 0
    if s.isdigit():
        return int(s)
    n = 0
    for ch in s:
        if not ('A' <= ch <= 'Z'):
            return 0
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n - 1


def _load_monthly_balances_manual(mb_cfg):
    """Build (df, alll_by_date) from a wizard-entered pool × month grid."""
    entries = mb_cfg.get('entries') or {}
    records = []
    for pool, row in entries.items():
        if not pool or not isinstance(row, dict):
            continue
        for d, v in row.items():
            try:
                dt = pd.to_datetime(d, errors='coerce')
            except Exception:
                continue
            if pd.isna(dt):
                continue
            try:
                bal = float(v)
            except (TypeError, ValueError):
                continue
            records.append({'pool': str(pool).strip(),
                            'date': dt,
                            'balance': bal})
    return pd.DataFrame(records, columns=['pool', 'date', 'balance']), {}


def _load_monthly_balances_per_month(mb_cfg):
    """Read one balance-sheet style file per month and emit (pool, date, bal)
    rows. Each file is opened on ``layout.sheet`` (or the first sheet) and
    the label/balance columns are pulled from the configured letters,
    skipping ``header_row`` rows. Labels are mapped to wizard pool names via
    ``pool_map`` (case-insensitive, falls back to the raw label).

    Supported file types: .xlsx / .xls / .csv (via pandas) and .pdf (via
    pdfplumber's table extraction). For PDFs the ``sheet`` field is
    interpreted as a 1-based page number; blank means scan every page and
    concatenate rows.
    """
    layout = mb_cfg.get('layout') or {}
    sheet = (layout.get('sheet') or '').strip()
    header_row = int(layout.get('header_row') or 1)
    label_idx = _col_letter_to_idx(layout.get('label_col') or 'A')
    balance_idx = _col_letter_to_idx(layout.get('balance_col') or 'B')
    raw_map = mb_cfg.get('pool_map') or {}
    pool_map = {str(k).strip().lower(): str(v).strip()
                for k, v in raw_map.items() if str(k).strip() and str(v).strip()}

    records = []
    for entry in (mb_cfg.get('files') or []):
        period = (entry.get('period') or '').strip()
        path = (entry.get('saved_path') or '').strip()
        if not path:
            # Fall back to looking up filename inside data_directory if it
            # was copied there post-save.
            continue
        try:
            dt = pd.to_datetime(period, errors='coerce')
        except Exception:
            continue
        if pd.isna(dt) or not os.path.isfile(path):
            continue
        ext = os.path.splitext(path)[1].lower()
        try:
            if ext == '.pdf':
                df = _read_pdf_balance_table(path, sheet)
            elif ext == '.csv':
                df = pd.read_csv(path, header=None, dtype=str)
            else:
                if sheet:
                    df = pd.read_excel(path, sheet_name=sheet, header=None)
                else:
                    df = pd.read_excel(path, header=None)
        except Exception as e:
            print(f"    Warning: Could not read monthly-balance file {path}: {e}")
            continue
        if df is None or df.empty or df.shape[1] <= max(label_idx, balance_idx):
            continue
        for i in range(header_row, df.shape[0]):
            label = df.iat[i, label_idx]
            bal = df.iat[i, balance_idx]
            if pd.isna(label) or str(label).strip() == '':
                continue
            if pd.isna(bal):
                continue
            bal_f = _coerce_balance(bal)
            if bal_f is None:
                continue
            key = str(label).strip().lower()
            pool = pool_map.get(key, str(label).strip())
            if not pool:
                continue
            records.append({'pool': pool, 'date': dt, 'balance': bal_f})
    return pd.DataFrame(records, columns=['pool', 'date', 'balance']), {}


def _coerce_balance(v):
    """Convert a cell value to float, stripping $, commas, parens for
    negatives, and whitespace. Returns None on failure."""
    if v is None:
        return None
    if isinstance(v, (int, float)):
        try:
            f = float(v)
        except (TypeError, ValueError):
            return None
        if pd.isna(f):
            return None
        return f
    s = str(v).strip()
    if not s:
        return None
    neg = False
    if s.startswith('(') and s.endswith(')'):
        neg = True
        s = s[1:-1]
    s = s.replace('$', '').replace(',', '').replace(' ', '')
    if s in ('', '-', '–', '—'):
        return None
    try:
        f = float(s)
    except ValueError:
        return None
    return -f if neg else f


def _read_pdf_balance_table(path, sheet):
    """Open a PDF via pdfplumber, extract tables on the requested page (or
    all pages when ``sheet`` is blank), and return a single pandas
    DataFrame whose rows are the concatenated table rows. Returns an empty
    DataFrame if pdfplumber is unavailable or no tables are detected."""
    try:
        import pdfplumber  # type: ignore
    except ImportError:
        print(f"    Warning: pdfplumber not installed; cannot read {path}")
        return pd.DataFrame()

    page_filter = None
    sht = (sheet or '').strip()
    if sht:
        try:
            page_filter = int(sht) - 1  # convert 1-based to 0-based
        except ValueError:
            page_filter = None

    all_rows = []
    max_width = 0
    try:
        with pdfplumber.open(path) as pdf:
            pages = pdf.pages
            if page_filter is not None and 0 <= page_filter < len(pages):
                pages = [pages[page_filter]]
            for page in pages:
                try:
                    tables = page.extract_tables() or []
                except Exception:
                    tables = []
                for tbl in tables:
                    for row in tbl:
                        if not row:
                            continue
                        cleaned = [('' if c is None else str(c).strip())
                                   for c in row]
                        if not any(cleaned):
                            continue
                        all_rows.append(cleaned)
                        if len(cleaned) > max_width:
                            max_width = len(cleaned)
    except Exception as e:
        print(f"    Warning: pdfplumber failed on {path}: {e}")
        return pd.DataFrame()

    if not all_rows:
        return pd.DataFrame()
    # Normalize ragged rows so iat[i, col] is always valid.
    normalized = [r + [''] * (max_width - len(r)) for r in all_rows]
    return pd.DataFrame(normalized)


def _merge_acl_history(alll_by_date: dict, config: dict) -> dict:
    """Merge wizard-entered/manually-extracted ACL history from
    ``cfg["acl"]["history"]`` into the ALLL-by-date map loaded from the
    monthly balances file. YAML history values take precedence on key
    collision (they are the user's audited / latest source of truth).
    """
    acl_cfg = (config or {}).get('acl') or {}
    hist_map = acl_cfg.get('history') or {}
    if not hist_map:
        return alll_by_date
    out = dict(alll_by_date or {})
    n = 0
    for k, v in hist_map.items():
        try:
            d = pd.Timestamp(k)
            out[d] = abs(float(v))
            n += 1
        except (ValueError, TypeError):
            continue
    if n:
        print(f"    Merged {n} ACL history entries from YAML cfg['acl']['history']")
    return out


def _load_monthly_balances_from_wizard(config):
    """Load monthly balances using wizard-provided cfg['monthly_balance']
    metadata (``saved_path`` + ``sheet`` + ``pool_name_col`` +
    ``first_date_col`` + ``header_row``) plus optional
    ``cfg['balance_title_map']`` for label→pool translation and
    ``cfg['acl']['row']`` / ``cfg['acl']['label']`` for locating the
    ACL row.

    Returns ``(df, alll_by_date)`` or ``(None, None)`` if the wizard
    metadata is missing / file is unreadable.
    """
    mb_cfg = (config or {}).get('monthly_balance') or {}
    saved_path = mb_cfg.get('saved_path')
    if not saved_path or not os.path.isfile(saved_path):
        return None, None

    sheet = mb_cfg.get('sheet')
    header_row = mb_cfg.get('header_row')  # 1-based
    pool_col_letter = mb_cfg.get('pool_name_col')
    date_col_letter = mb_cfg.get('first_date_col')

    def _col_letter_to_idx(letter):
        if not letter or not isinstance(letter, str):
            return None
        letter = letter.strip().upper()
        if not letter.isalpha():
            return None
        n = 0
        for ch in letter:
            n = n * 26 + (ord(ch) - ord('A') + 1)
        return n - 1  # 0-based

    pool_col = _col_letter_to_idx(pool_col_letter)
    date_start_col = _col_letter_to_idx(date_col_letter)

    try:
        if sheet:
            df_raw = pd.read_excel(saved_path, sheet_name=sheet, header=None)
        else:
            df_raw = pd.read_excel(saved_path, header=None)
    except Exception as e:
        print(f"    Warning: could not read monthly_balance saved_path "
              f"{saved_path}: {e}")
        return None, None

    # Resolve header row: prefer wizard's value, else autodetect (same
    # rules as legacy loader — first row with ≥3 datetime cells).
    hdr_idx = None
    if header_row is not None:
        try:
            hdr_idx = max(0, int(header_row) - 1)
        except (TypeError, ValueError):
            hdr_idx = None
    if hdr_idx is None:
        for i in range(min(8, len(df_raw))):
            row = df_raw.iloc[i]
            dt_count = sum(1 for v in row
                           if isinstance(v, (pd.Timestamp, datetime)))
            if dt_count >= 3:
                hdr_idx = i
                break
    if hdr_idx is None:
        return None, None

    # Resolve column anchors: prefer wizard letters, else autodetect.
    if date_start_col is None or pool_col is None:
        hdr_row = df_raw.iloc[hdr_idx]
        for c in range(df_raw.shape[1]):
            val = hdr_row.iloc[c]
            if (date_start_col is None
                    and isinstance(val, (pd.Timestamp, datetime))):
                date_start_col = c
            if (pool_col is None and isinstance(val, str)
                    and 'pool' in val.lower()):
                pool_col = c
        if date_start_col is None:
            return None, None
        if pool_col is None:
            pool_col = max(0, date_start_col - 1)

    dates = pd.to_datetime(
        df_raw.iloc[hdr_idx, date_start_col:].values, errors='coerce')

    # Optional label→pool translation map; when present, only rows whose
    # label appears in the map are included (translated to the pool
    # name). When absent, fall through to the legacy "use label as-is
    # with a few hard-coded skips" behavior.
    title_map = (config or {}).get('balance_title_map') or {}
    use_title_map = bool(title_map)

    # ACL row resolution: cfg['acl']['row'] is a 1-based row number on
    # the same sheet; cfg['acl']['label'] is an alternate text match.
    acl_cfg = (config or {}).get('acl') or {}
    acl_row_1based = acl_cfg.get('row')
    acl_label = (acl_cfg.get('label') or '').strip().lower()
    try:
        acl_row_idx = (int(acl_row_1based) - 1) if acl_row_1based else None
    except (TypeError, ValueError):
        acl_row_idx = None

    records = []
    alll_by_date: dict = {}

    # Extract ACL row first (by explicit row index when given).
    if acl_row_idx is not None and 0 <= acl_row_idx < df_raw.shape[0]:
        for j in range(len(dates)):
            if pd.notna(dates[j]):
                aval = df_raw.iloc[acl_row_idx, date_start_col + j]
                if pd.notna(aval):
                    try:
                        alll_by_date[dates[j]] = abs(float(aval))
                    except (ValueError, TypeError):
                        pass

    for i in range(hdr_idx + 1, df_raw.shape[0]):
        if i == acl_row_idx:
            continue  # already handled above
        raw_label = df_raw.iloc[i, pool_col]
        if pd.isna(raw_label) or str(raw_label).strip() == '':
            continue
        label = str(raw_label).strip()
        label_lc = label.lower()

        # ACL row by label match (when no explicit row given).
        if acl_row_idx is None and (
                (acl_label and label_lc == acl_label)
                or label in ('ALLL Balance', 'ACL Balance')):
            for j in range(len(dates)):
                if pd.notna(dates[j]):
                    aval = df_raw.iloc[i, date_start_col + j]
                    if pd.notna(aval):
                        try:
                            alll_by_date[dates[j]] = abs(float(aval))
                        except (ValueError, TypeError):
                            pass
            continue

        # Pool resolution.
        if use_title_map:
            pool_name = title_map.get(label)
            if not pool_name:
                continue  # label not opted-in; skip silently
        else:
            if label in ('Total', 'Total Loans'):
                continue
            if len(label) > 35 or label.startswith('In ') or label.startswith('Before'):
                continue
            pool_name = label

        for j in range(len(dates)):
            if pd.notna(dates[j]):
                bal = df_raw.iloc[i, date_start_col + j]
                if pd.notna(bal):
                    try:
                        records.append({
                            'pool': pool_name,
                            'date': dates[j],
                            'balance': float(bal),
                        })
                    except (ValueError, TypeError):
                        pass

    out_df = pd.DataFrame(records)
    if out_df.empty and not alll_by_date:
        return None, None
    print(f"    Loaded monthly balances from wizard saved_path "
          f"({os.path.basename(saved_path)}): "
          f"{len(out_df)} rows, {len(alll_by_date)} ACL dates")
    return out_df, alll_by_date


def load_monthly_balances(config):
    """Load monthly loan balances by pool from the most recent file available.
    Returns (DataFrame with columns [pool, date, balance],
            dict mapping date -> ALLL balance (absolute value))."""
    # New (May 2026): the wizard can declare three delivery modes in
    # ``config["monthly_balance"]``. Honor per_month / manual modes first;
    # fall through to the legacy data_directory scan when no block is set
    # or source == "single".
    mb_cfg = config.get('monthly_balance') or {}
    mb_source = (mb_cfg.get('source') or '').strip().lower()
    if mb_source == 'manual':
        return _load_monthly_balances_manual(mb_cfg)
    if mb_source == 'per_month':
        df, alll = _load_monthly_balances_per_month(mb_cfg)
        if not df.empty:
            return df, _merge_acl_history(alll, config)
        # If per_month failed (no files / unreadable), fall through to the
        # legacy scan so the user at least gets the historical context.

    # Preferred path for "single" mode: use the wizard's saved_path +
    # sheet metadata directly. Honors cfg['balance_title_map'] (label
    # → pool translation) and cfg['acl']['row'/'label'] for ACL row
    # discovery. Falls through to the legacy data_directory scan when
    # the wizard metadata is missing or the file can't be read.
    wiz_df, wiz_alll = _load_monthly_balances_from_wizard(config)
    if wiz_df is not None:
        return wiz_df, _merge_acl_history(wiz_alll or {}, config)

    data_dir = resolve_path(config.get('data_directory', ''))
    if not data_dir or not os.path.isdir(data_dir):
        return (pd.DataFrame(columns=['pool', 'date', 'balance']),
                _merge_acl_history({}, config))

    # Find balance files - match various naming conventions
    balance_files = []
    for root, dirs, files in os.walk(data_dir):
        for f in files:
            fl = f.lower()
            if not fl.endswith('.xlsx'):
                continue
            # Match: "monthly*balance*", "Loan Balances*", "*BalancesGrades*", "*LoanDataBalances*", "*cecl book*"
            if (('monthly' in fl and 'balance' in fl) or
                ('loan' in fl and 'balance' in fl) or
                ('balancesgrades' in fl) or
                ('loandatabalances' in fl) or
                ('cecl book' in fl or 'cecl_book' in fl)):
                balance_files.append(os.path.join(root, f))

    if not balance_files:
        return (pd.DataFrame(columns=['pool', 'date', 'balance']),
                _merge_acl_history({}, config))

    # Use the most recently modified file
    balance_files.sort(key=os.path.getmtime, reverse=True)
    filepath = balance_files[0]

    try:
        # Try to find a balance sheet (various naming conventions)
        xl = pd.ExcelFile(filepath)
        balance_sheet = None
        for sname in xl.sheet_names:
            sl = sname.lower()
            if 'balances by pool' in sl or ('loan' in sl and 'balance' in sl):
                balance_sheet = sname
                break
        if balance_sheet:
            df = pd.read_excel(filepath, sheet_name=balance_sheet, header=None)
        else:
            df = pd.read_excel(filepath, header=None)

        # Detect layout: find the header row with dates
        hdr_row_idx = None
        pool_col = None
        date_start_col = None
        for i in range(min(5, len(df))):
            row = df.iloc[i]
            # Look for a row that has datetime values
            date_count = sum(1 for v in row if isinstance(v, (pd.Timestamp, datetime)))
            if date_count >= 3:
                hdr_row_idx = i
                # Find where dates start and where pool names are
                for c in range(df.shape[1]):
                    val = row.iloc[c]
                    if isinstance(val, (pd.Timestamp, datetime)) and date_start_col is None:
                        date_start_col = c
                    if isinstance(val, str) and 'pool' in val.lower() and pool_col is None:
                        pool_col = c
                break

        if hdr_row_idx is None:
            # Fallback: assume row 1 has dates
            hdr_row_idx = 1
            date_start_col = 1
            pool_col = 0

        # If pool column wasn't labeled "Pool", it's the column just before dates
        if pool_col is None:
            pool_col = max(0, date_start_col - 1)

        dates = pd.to_datetime(df.iloc[hdr_row_idx, date_start_col:].values, errors='coerce')

        records = []
        alll_by_date = {}
        for i in range(hdr_row_idx + 1, df.shape[0]):
            pool_name = df.iloc[i, pool_col]
            if pd.isna(pool_name) or str(pool_name).strip() == '':
                continue
            pool_name = str(pool_name).strip()
            # Skip notes/metadata rows
            if len(pool_name) > 35 or pool_name.startswith('In ') or pool_name.startswith('Before'):
                continue
            if pool_name in ('ALLL Balance', 'Total', 'Total Loans', 'ACL Balance'):
                if pool_name in ('ALLL Balance', 'ACL Balance'):
                    for j in range(len(dates)):
                        if pd.notna(dates[j]):
                            aval = df.iloc[i, date_start_col + j]
                            if pd.notna(aval):
                                try:
                                    alll_by_date[dates[j]] = abs(float(aval))
                                except (ValueError, TypeError):
                                    pass
                continue
            for j in range(len(dates)):
                if pd.notna(dates[j]):
                    bal = df.iloc[i, date_start_col + j]
                    if pd.notna(bal):
                        try:
                            records.append({
                                'pool': pool_name,
                                'date': dates[j],
                                'balance': float(bal)
                            })
                        except (ValueError, TypeError):
                            pass
        return pd.DataFrame(records), _merge_acl_history(alll_by_date, config)
    except Exception as e:
        print(f"    Warning: Could not parse monthly balances {filepath}: {e}")
        return (pd.DataFrame(columns=['pool', 'date', 'balance']),
                _merge_acl_history({}, config))


def _compute_balance_adjustments(df, hist, config, snapshot_date):
    """Compare loan-file balances with monthly-balance-file totals per pool.

    If a difference exists, populate hist['impaired'] with balance_adjustments,
    total_balance_adjustment, and total_in_portfolio so the Pool_Balance Adjust
    sheet and migration matrix reflect the gap.
    """
    snap_dt = pd.Timestamp(snapshot_date)
    snap_ym = snap_dt.to_period('M')

    # Primary source: monthly balance file on disk.
    monthly_bals = {}
    monthly_df = hist.get('monthly_balances')
    if monthly_df is not None and not monthly_df.empty:
        mb = monthly_df.copy()
        mb['ym'] = mb['date'].dt.to_period('M')
        month_data = mb[mb['ym'] == snap_ym]
        for _, row in month_data.iterrows():
            pool = str(row['pool']).strip()
            monthly_bals[pool] = float(row['balance'])

    # Fallback: WARM template's per-pool monthly balance series. Use the
    # snapshot month's value (or the most recent value at/before snapshot).
    _SKIP_POOLS = {'grand total', 'total', 'exclude', 'excluded'}
    if not monthly_bals:
        # Prefer the pre-extension snapshot captured at load time; the live
        # hist_bal_data 'total' for snapshot month may have been overwritten
        # with loan-extract sums by extend_hist_bal_with_db.
        warm_snap = (hist.get('impaired') or {}).get('warm_snapshot_balances') or {}
        for pool, val in warm_snap.items():
            pname = str(pool).strip()
            if pname.lower() in _SKIP_POOLS:
                continue
            try:
                monthly_bals[pname] = float(val)
            except (TypeError, ValueError):
                continue

    if not monthly_bals:
        hbd = (hist.get('impaired') or {}).get('hist_bal_data') or {}
        for pool, pdata in hbd.items():
            pname = str(pool).strip()
            if pname.lower() in _SKIP_POOLS:
                continue
            dates = pdata.get('dates') or []
            tots = pdata.get('total') or []
            best_idx = None
            for i, d in enumerate(dates):
                try:
                    d_ym = pd.Timestamp(d).to_period('M')
                except Exception:
                    continue
                if d_ym <= snap_ym:
                    best_idx = i
                if d_ym == snap_ym:
                    break
            if best_idx is not None and best_idx < len(tots):
                try:
                    monthly_bals[pname] = float(tots[best_idx])
                except (TypeError, ValueError):
                    continue

    if not monthly_bals:
        return

    # Loan-file balances by pool
    loan_bals = df.groupby('loan_pool')['current_balance'].sum().to_dict()

    # Map monthly pool names to loan-file pool names (fuzzy match)
    pool_order = config.get('pool_order', [])

    def _match_pool(monthly_name):
        """Return the loan-file pool name that best matches a monthly name."""
        mn = monthly_name.lower().strip()
        # Direct match
        for lp in loan_bals:
            if lp.lower() == mn:
                return lp
        # Match against pool_order (canonical names)
        for pn in pool_order:
            if pn.lower() == mn:
                return pn
        # Prefix match (e.g. "Re-write(RW TM)" -> "Re-write")
        for lp in list(loan_bals.keys()) + pool_order:
            if mn.startswith(lp.lower()) or lp.lower().startswith(mn[:6]):
                return lp
        return None

    balance_adjustments = {}  # {pool: total_adj}
    total_adj = 0.0
    grand_loan = 0.0

    all_pools = set(loan_bals.keys())
    matched_loan_pools = set()
    nrr_set = set(config.get('not_risk_rated', []) or [])
    # Configured loan pools — used to drop balance-sheet line items that
    # aren't loan pools (e.g. "ACH Clearing", "Accrued Interest"). Anything
    # not in this set AND not already a loan-extract pool is ignored.
    configured_pools = set(pool_order)
    configured_pools.update(p.get('name') for p in (config.get('pools') or [])
                            if p and p.get('name'))
    configured_pools.update(nrr_set)
    skipped_non_pool: list[str] = []

    for mp, mb_val in monthly_bals.items():
        lp = _match_pool(mp)
        if lp is None:
            # Monthly-balance pool with no loan-extract counterpart (e.g.
            # non-risk-rated pools like Loan Participation, Repo/Foreclosed).
            # Treat the full monthly balance as an adjustment so it shows up
            # in the per-pool "Loans Not Risk Rated and Adjustments" row —
            # but ONLY when the name is a configured loan pool. Balance-sheet
            # line items like ACH Clearing have no place on the ACL tabs.
            mp_clean = str(mp).strip()
            if configured_pools and mp_clean not in configured_pools:
                # Try case-insensitive match before giving up.
                lc = mp_clean.lower()
                hit = next((c for c in configured_pools
                            if str(c).strip().lower() == lc), None)
                if hit is None:
                    if mb_val and abs(mb_val) > 0.005:
                        skipped_non_pool.append(f"{mp_clean} (${mb_val:,.2f})")
                    continue
                mp_clean = hit
            if mb_val and abs(mb_val) > 0.005:
                balance_adjustments[mp_clean] = round(float(mb_val), 2)
                total_adj += float(mb_val)
            continue
        matched_loan_pools.add(lp)
        lb = loan_bals.get(lp, 0)
        diff = mb_val - lb
        if abs(diff) > 0.005:  # ignore sub-penny rounding
            balance_adjustments[lp] = round(diff, 2)
            total_adj += diff

    # Include loan pools with no monthly match (balance goes unreported)
    # These don't need adjustments — they're just in the loan file

    total_adj = round(total_adj, 2)
    grand_loan = sum(loan_bals.values())
    total_in_portfolio = round(grand_loan + total_adj, 2)

    if not balance_adjustments:
        return

    # Store in hist['impaired']
    imp = hist.setdefault('impaired', {})
    imp['balance_adjustments'] = balance_adjustments
    imp['total_balance_adjustment'] = total_adj
    imp['total_in_portfolio'] = total_in_portfolio

    # Build pool_bal_detail for the Vizo Balance Adjust sheet.
    # Per-grade detail uses loan-file balances; adjustment is pool-level.
    grades_cfg = config.get('credit_grades', config.get('grades', []))
    no_score = config.get('no_score_label', 'Not Reported')
    grade_labels = [g['label'] for g in grades_cfg] + [no_score]
    pool_bal_detail = {}
    for pool in set(list(loan_bals.keys()) + list(balance_adjustments.keys())):
        pool_df = df[df['loan_pool'] == pool]
        adj = balance_adjustments.get(pool, 0)
        gd = {}
        pool_loan_total = 0
        grade_bals = {}
        for g in grade_labels:
            g_bal = pool_df[pool_df['current_grade'] == g]['current_balance'].sum() \
                    if not pool_df.empty else 0
            grade_bals[g] = g_bal
            pool_loan_total += g_bal
        # Distribute pool adjustment across grades proportionally
        adj_remaining = adj
        for i, g in enumerate(grade_labels):
            g_bal = grade_bals[g]
            if pool_loan_total and adj:
                if i == len(grade_labels) - 1:
                    g_adj = adj_remaining  # last grade gets remainder to avoid rounding drift
                else:
                    g_adj = round(adj * g_bal / pool_loan_total, 2)
                    adj_remaining = round(adj_remaining - g_adj, 2)
            else:
                g_adj = 0
            gd[g] = {
                'loan_report_bal': g_bal,
                'bal_adj': g_adj,
                'balance_sheet_total': g_bal + g_adj,
            }
        gd['Total'] = {
            'loan_report_bal': pool_loan_total,
            'bal_adj': adj,
            'balance_sheet_total': pool_loan_total + adj,
        }
        pool_bal_detail[pool] = gd
    imp['pool_bal_detail'] = pool_bal_detail

    # Also store in config for the TCT Pool_Balance Adjust detail sheet
    config['balance_adjustments'] = {pool: {'_pool_total': adj}
                                     for pool, adj in balance_adjustments.items()}

    n_adj = len(balance_adjustments)
    print(f"    Balance adjustments: {n_adj} pools, "
          f"total adj: ${total_adj:,.2f}, "
          f"total in portfolio: ${total_in_portfolio:,.2f}")
    if skipped_non_pool:
        print(f"    Skipped {len(skipped_non_pool)} balance-sheet line "
              f"item(s) not mapped to any loan pool: "
              f"{', '.join(skipped_non_pool[:8])}"
              f"{'...' if len(skipped_non_pool) > 8 else ''}")


def load_delinquency_history(config):
    """Load delinquency data from all available quarterly files.
    Returns dict: {quarter_label: {pool: dq_balance}}"""
    data_dir = resolve_path(config.get('data_directory', ''))
    if not data_dir or not os.path.isdir(data_dir):
        return {}

    pool_map = config.get('pool_map', {})
    quarters = _find_quarter_folders(data_dir)
    dq_data = {}  # {quarter_label: {pool: dq_balance}}

    for folder, qlabel in quarters:
        for f in os.listdir(folder):
            fl = f.lower()
            if 'delq' in fl and fl.endswith('.xlsx'):
                filepath = os.path.join(folder, f)
                try:
                    df = pd.read_excel(filepath, header=None)
                    # Check for header row
                    first_vals = [str(v).lower() if pd.notna(v) else '' for v in df.iloc[0]]
                    if any('account' in v or 'delinq' in v for v in first_vals):
                        df = df.iloc[1:]
                    df = df.dropna(subset=[0])
                    df = df[pd.to_numeric(df[0], errors='coerce').notna()]

                    if df.empty:
                        continue

                    # Format: [account, suffix, code, balance, days]
                    # code col = 2, balance col = 3
                    for _, row in df.iterrows():
                        code = str(row[2]).strip().upper() if pd.notna(row[2]) else ''
                        pool = pool_map.get(code, pool_map.get(code.lower(), None))
                        bal = pd.to_numeric(row[3], errors='coerce')
                        if pool and pd.notna(bal):
                            dq_data.setdefault(qlabel, {})
                            dq_data[qlabel][pool] = dq_data[qlabel].get(pool, 0) + bal
                except Exception as e:
                    print(f"    Warning: Could not parse {filepath}: {e}")

    return dq_data


# ─────────────────────────────────────────────────────────────────────
# Manual-WARM template loader
#
# Some "prior reports" in Reports/ are not previously-generated TCTs but
# the manual CECL-Migration-WARM workbook the credit-union staff used to
# produce, imported as a baseline during new-CU setup.  Those workbooks
# carry the *same* historical data we want, but in differently-named
# sheets and with a one-column offset in the historical-balance grid:
#
#   Generated TCT sheet        Manual WARM sheet
#   ─────────────────────────  ─────────────────────────────────
#   > Detail_HIst Balances     HIst Bal Data        (col B blank, dates start col C)
#   Display CO-Recov-DQ        Display CO-Recov -DQ (note the stray space)
#   >Detail_Charge off Hist    Charge off History
#   ACL Env by Pool Mgmt Adj   ACL Env by Pool Mgmt Adj  (same name)
#
# These helpers detect a manual-WARM workbook and extract the same
# result dict shape that load_prior_tct_hist_bal produces from a
# generated TCT.  The WARM hist-bal grid is column-mapped (we record the
# actual column index of every date instead of assuming contiguous
# columns) so a leading blank column does not shift values.
# ─────────────────────────────────────────────────────────────────────

# Grade labels in HIst Bal Data that are not real grades and must be
# excluded from the synthesised hist_bal_data result.
_WARM_HIDDEN_GRADE_PREFIXES = ('hide', 'hide-', 'minimum', 'maximum',
                               'max for', 'min for')


def _is_warm_template_workbook(sheet_names):
    """Return True if the workbook looks like a manual WARM (vs. one of our
    generated TCT outputs)."""
    sn = set(sheet_names)
    has_warm_hist = 'HIst Bal Data' in sn
    has_tct_hist = '> Detail_HIst Balances' in sn
    return has_warm_hist and not has_tct_hist


def _warm_resolve_sheet(sheet_names, *candidates):
    """Return the first matching sheet name from candidates, or None."""
    sn = set(sheet_names)
    for c in candidates:
        if c in sn:
            return c
    return None


def _read_sheet_rows(ws):
    """Materialise a worksheet to a list[list] for fast random access."""
    return [list(r) for r in ws.iter_rows(values_only=True)]


def _warm_parse_hist_bal(rows):
    """Parse a manual-WARM 'HIst Bal Data' sheet.

    Returns (hist_bal_data, pool_order, risk_rated). Each pool block is:
        Row N:   pool name in col A
        Row N+1: 'Current Grade' in col A; dates in cols C..  (col B blank)
        Row N+2..: grade label in col A, balance values in same cols as dates
        Row M:   'Total' in col A, balance values
        Row M+1: blank spacer
    """
    hist_bal_data = {}
    pool_order = []
    risk_rated = {}

    r = 0
    n = len(rows)
    while r < n:
        row = rows[r] or []
        a = row[0] if row else None
        if a is None or str(a).strip() == '':
            r += 1
            continue
        a_s = str(a).strip()
        # Skip header / metadata rows
        low = a_s.lower()
        if low in ('current grade', 'total', 'balance', 'grand total',
                   'excluded', 'exclude') \
           or low.startswith(('for period', 'loss factor', 'allowance',
                              'charge off', 'tongass', 'siskiyou')) \
           or any(low.startswith(p) for p in ('hide', 'minimum', 'maximum',
                                              'max for')):
            r += 1
            continue

        # Need a 'Current Grade' row immediately below to qualify as a pool
        if r + 1 >= n:
            break
        next_a = (rows[r + 1] or [None])[0]
        if not next_a or str(next_a).strip() != 'Current Grade':
            r += 1
            continue

        pool_name = a_s
        hdr_row = rows[r + 1]

        # Capture date columns by their actual indices
        date_cols = []  # list[(col_idx, pd.Timestamp)]
        for ci in range(1, len(hdr_row)):
            v = hdr_row[ci]
            if v is None:
                continue
            try:
                date_cols.append((ci, pd.Timestamp(v)))
            except Exception:
                continue

        if not date_cols:
            r += 2
            continue

        dates = [d for _, d in date_cols]
        pool_grades = {}
        pool_total = []
        gr = r + 2
        while gr < n:
            grow = rows[gr] or []
            ga = grow[0] if grow else None
            if ga is None or str(ga).strip() == '':
                break
            glabel = str(ga).strip()
            glow = glabel.lower()

            vals = []
            for ci, _ in date_cols:
                v = grow[ci] if ci < len(grow) else 0
                try:
                    vals.append(float(v) if v is not None else 0.0)
                except (ValueError, TypeError):
                    vals.append(0.0)

            if glow == 'total':
                pool_total = vals
                gr += 1
                break
            # Filter out 'Hide-*', 'Minimum', etc.
            if any(glow.startswith(p) for p in _WARM_HIDDEN_GRADE_PREFIXES):
                gr += 1
                continue
            pool_grades[glabel] = vals
            gr += 1

        if pool_grades or pool_total:
            pool_order.append(pool_name)
            risk_rated[pool_name] = bool(pool_grades)
            # Ensure pool_total length always matches dates so downstream
            # code that does `pool_total[idx] = ...` is safe even when the
            # WARM block omitted a 'Total' row. Reconstruct from grade
            # rows when missing; pad with zeros as a last resort.
            if len(pool_total) != len(dates):
                if pool_grades:
                    rebuilt = []
                    for ci in range(len(dates)):
                        s = 0.0
                        for vals in pool_grades.values():
                            if ci < len(vals):
                                s += vals[ci]
                        rebuilt.append(s)
                    pool_total = rebuilt
                else:
                    pool_total = [0.0] * len(dates)
            hist_bal_data[pool_name] = {
                'dates': dates,
                'grades': pool_grades,
                'total': pool_total,
            }
        r = gr if gr > r + 2 else r + 2

    return hist_bal_data, pool_order, risk_rated


def _load_hist_from_warm_template(wb, snap):
    """Build the same result dict as load_prior_tct_hist_bal from a manual
    WARM workbook (HIst Bal Data + Display CO-Recov -DQ + Charge off History
    + ACL Env by Pool Mgmt Adj).
    """
    result = {}
    sheets = wb.sheetnames

    # ── Hist balances (synthetic '> Detail_HIst Balances') ──
    hb_sheet = _warm_resolve_sheet(sheets, 'HIst Bal Data')
    if hb_sheet:
        rows = _read_sheet_rows(wb[hb_sheet])
        hbd, pord, rrated = _warm_parse_hist_bal(rows)
        if hbd:
            result['hist_bal_data'] = hbd
            result['pool_order'] = pord
            result['risk_rated'] = rrated
            try:
                n_dates = max(len(d.get('dates', [])) for d in hbd.values())
            except ValueError:
                n_dates = 0
            print(f"    WARM template hist bal: {len(hbd)} pools, "
                  f"{n_dates} months")

    # ── CO/RC/Net/DQ year totals (Display CO-Recov -DQ) ──
    co_sheet = _warm_resolve_sheet(sheets, 'Display CO-Recov -DQ',
                                   'Display CO-Recov-DQ')
    if co_sheet:
        co_rows = _read_sheet_rows(wb[co_sheet])
        # Reuse the same section parser used for generated TCTs by
        # shelling out to a local copy (kept inline to avoid threading
        # the helper through module scope).
        warm_co, warm_co_tot, _ = _warm_parse_co_section(co_rows, 'Charge offs')
        warm_rc, warm_rc_tot, _ = _warm_parse_co_section(co_rows, 'Recoveries')
        warm_net, warm_net_tot, _ = _warm_parse_co_section(co_rows,
                                                            'Net Charge offs')
        warm_dq, _, _ = _warm_parse_co_section(co_rows, 'DQ %')
        # Recoveries are stored negative in WARM — flip to positive.
        for yr in warm_rc:
            for p in warm_rc[yr]:
                warm_rc[yr][p] = abs(warm_rc[yr][p])
        warm_rc_tot = {p: abs(v) for p, v in warm_rc_tot.items()}

        if warm_co:
            result['warm_co'] = warm_co
            result['warm_rc'] = warm_rc
            result['warm_net'] = warm_net
            result['warm_co_totals'] = warm_co_tot
            result['warm_rc_totals'] = warm_rc_tot
            result['warm_net_co'] = warm_net_tot
            if warm_dq:
                result['warm_dq_pct'] = warm_dq
            n_pools = len(set(p for yr in warm_co.values() for p in yr))
            n_years = len(warm_co)
            print(f"    WARM template CO/RC: {n_pools} pools, {n_years} years"
                  f" (CO totals: {len(warm_co_tot)} pools)")

    # ── Monthly CO / Recovery detail (Charge off History) ──
    mo_sheet = _warm_resolve_sheet(sheets, 'Charge off History',
                                   '>Detail_Charge off Hist')
    if mo_sheet:
        mo_rows = _read_sheet_rows(wb[mo_sheet])
        warm_co_mo = _warm_parse_monthly_co(mo_rows, 'Charge offs')
        warm_rc_mo = _warm_parse_monthly_co(mo_rows, 'Recoveries')
        # Recoveries again may be negative
        for ym in warm_rc_mo:
            for p in warm_rc_mo[ym]:
                warm_rc_mo[ym][p] = abs(warm_rc_mo[ym][p])
        if warm_co_mo:
            result['warm_co_monthly'] = warm_co_mo
            print(f"    WARM template monthly CO: {len(warm_co_mo)} months")
        if warm_rc_mo:
            result['warm_rc_monthly'] = warm_rc_mo

    return result


def _warm_parse_co_section(rows, start_label):
    """Parse a CO/RC/Net/DQ section of a Display CO-Recov-DQ-style sheet.

    Same shape & tolerances as the inline _parse_co_section inside
    load_prior_tct_hist_bal — kept separate so the WARM-template loader can
    be invoked outside that function.
    """
    import re as _re
    year_data = {}
    totals = {}
    header_row = None
    pool_start = None
    col_years = []
    acl_col = None

    for ri, row in enumerate(rows):
        c0 = str((row[0] if row else '') or '').strip().lower()
        if c0 == start_label.lower():
            header_row = ri
            for ci in range(1, len(row)):
                val = row[ci]
                if val is None:
                    continue
                sv = str(val).strip()
                m = _re.match(r'(?:YTD\s+)?(\d{4})', sv)
                if m:
                    col_years.append((ci, int(m.group(1))))
                elif 'acl' in sv.lower() or 'net charge' in sv.lower():
                    acl_col = ci
            pool_start = ri + 1
            break

    if header_row is None:
        return year_data, totals, []

    for ri in range(pool_start, len(rows)):
        row = rows[ri] or []
        pool_name = str((row[0] if row else '') or '').strip()
        if not pool_name:
            break
        pl = pool_name.lower()
        if any(kw in pl for kw in ['recoveries', 'net charge', 'dq %',
                                    'charge offs']):
            break
        # Skip 'Hide-*' / 'Exclude' rows
        if pl.startswith(('hide', 'exclude')):
            continue

        for ci, yr in col_years:
            val = row[ci] if ci < len(row) else None
            if val is not None and val != 0:
                try:
                    fval = float(val)
                except (ValueError, TypeError):
                    continue
                year_data.setdefault(yr, {})[pool_name] = fval

        if acl_col is not None and acl_col < len(row):
            aval = row[acl_col]
            if aval is not None:
                try:
                    totals[pool_name] = float(aval)
                except (ValueError, TypeError):
                    pass

    return year_data, totals, [yr for _, yr in col_years]


def _warm_parse_monthly_co(rows, start_label):
    """Parse the WARM 'Charge off History' sheet's monthly section.

    Date headers are on the same row as the section label (e.g. row 9 for
    Charge offs, with col B = 2019-01-31, col C = 2019-02-28, ...).  Pool
    rows follow until a blank or 'Recoveries' label.
    """
    import datetime as _dt
    monthly = {}

    header_row = None
    for ri, row in enumerate(rows):
        c0 = str((row[0] if row else '') or '').strip().lower()
        if c0 == start_label.lower():
            header_row = ri
            break
    if header_row is None:
        return monthly

    date_cols = []
    for ci in range(1, len(rows[header_row])):
        v = rows[header_row][ci]
        if isinstance(v, _dt.datetime):
            date_cols.append((ci, v))

    if not date_cols:
        return monthly

    for ri in range(header_row + 1, len(rows)):
        row = rows[ri] or []
        pool_name = str((row[0] if row else '') or '').strip()
        if not pool_name:
            break
        pl = pool_name.lower()
        if pl in ('recoveries', 'charge offs', 'net charge offs', 'dq %') \
           or pl.startswith('total'):
            break
        if pl.startswith(('hide', 'exclude')):
            continue
        for ci, dt_val in date_cols:
            val = row[ci] if ci < len(row) else None
            if val is not None and val != 0:
                try:
                    fval = float(val)
                except (ValueError, TypeError):
                    continue
                ym = (dt_val.year, dt_val.month)
                monthly.setdefault(ym, {})[pool_name] = fval

    return monthly


def load_prior_tct_hist_bal(config, snap):
    """Load hist_bal_data from the most recent prior TCT report.

    When no WARM file exists for the current snapshot (i.e. the TCT report IS
    the new WARM replacement), this function reads the previous TCT report's
    '> Detail_HIst Balances' sheet so historical months carry forward.

    Returns dict compatible with hist['impaired'] keys:
      'hist_bal_data': {pool: {dates, grades, total}},
      'pool_order': [...],
      'acl_months': {pool: n},
      'risk_rated': {pool: bool},
    or {} if no prior report found.
    """
    from openpyxl import load_workbook as _load_wb

    cu = config['credit_union']
    safe_cu = cu.replace(' ', '_').replace('/', '-')

    # Search Reports/ for prior TCT files for this credit union
    rpt_dir = os.path.join(BASE, 'Reports')
    if not os.path.isdir(rpt_dir):
        return {}

    pattern = re.compile(
        rf'\d{{4}}-\d{{2}}-\d{{2}}_CECL_Migration_{re.escape(safe_cu)}_TCT_Model\.xlsx$',
        re.IGNORECASE,
    )
    candidates = []
    for root, dirs, files in os.walk(rpt_dir):
        for f in files:
            if f.startswith('~$'):
                continue
            if pattern.match(f):
                # Extract date from filename
                date_str = f[:10]
                # Allow equal-date matches only when the file lives in the
                # dedicated WARM-baseline subdir (so a WARM workbook
                # uploaded as the baseline for the current snapshot still
                # supplies historical hist_bal data). Exclude same-date
                # top-level files because those are this run's own output.
                in_warm_baselines = '_warm_baselines' in os.path.normpath(
                    os.path.relpath(root, rpt_dir)
                ).split(os.sep)
                if date_str < snap or (date_str == snap and in_warm_baselines):
                    candidates.append((date_str, in_warm_baselines,
                                       os.path.join(root, f)))

    if not candidates:
        return {}

    # Sort by date desc; for ties, prefer WARM-baseline (in_warm_baselines=True).
    candidates.sort(key=lambda c: (c[0], c[1]), reverse=True)
    prior_date, _, prior_path = candidates[0]
    print(f"    Loading prior TCT hist balances from: {os.path.basename(prior_path)}")

    try:
        wb = _load_wb(prior_path, read_only=True, data_only=True)
    except Exception as e:
        print(f"    Warning: Could not open prior report: {e}")
        return {}

    # Branch: manual-WARM workbook (uploaded as baseline during new-CU
    # setup) uses different sheet names than our generated TCTs.  Use the
    # WARM-template loader for everything except the ACL Env sheet, which
    # has the same name in both formats and is parsed below.
    if _is_warm_template_workbook(wb.sheetnames):
        print(f"    Detected manual-WARM workbook layout in "
              f"{os.path.basename(prior_path)}; using WARM template loader.")
        result = _load_hist_from_warm_template(wb, snap)

        # Continue into the existing 'ACL Env by Pool Mgmt Adj' parser
        # below so management adjustments + env factors carry forward.
        acl_rows_data = []
        if 'ACL Env by Pool Mgmt Adj' in wb.sheetnames:
            ws_acl = wb['ACL Env by Pool Mgmt Adj']
            for row in ws_acl.iter_rows(min_row=1, max_row=ws_acl.max_row,
                                        max_col=ws_acl.max_column,
                                        values_only=True):
                acl_rows_data.append(list(row))
        wb.close()

        if acl_rows_data:
            prior_mgmt_adj = {}
            prior_env_factor = {}
            current_pool = None
            for row in acl_rows_data:
                a_val = row[0] if row else None
                if a_val is None:
                    continue
                label = str(a_val).strip()
                if not label:
                    continue
                e_val = row[4] if len(row) > 4 else None
                f_val = row[5] if len(row) > 5 else None
                i_val = row[8] if len(row) > 8 else None
                if label == 'Current Grade':
                    continue
                if label == 'Total':
                    if current_pool and i_val is not None:
                        try:
                            prior_env_factor[current_pool] = float(i_val)
                        except (ValueError, TypeError):
                            pass
                    current_pool = None
                    continue
                if current_pool and e_val is not None:
                    try:
                        mgmt = float(f_val) if f_val is not None else 0.0
                    except (ValueError, TypeError):
                        mgmt = 0.0
                    if mgmt != 0:
                        prior_mgmt_adj.setdefault(current_pool, {})[label] = mgmt
                    continue
                if label not in ('Current Grade', 'Total') and e_val is None:
                    current_pool = label
            if prior_mgmt_adj or prior_env_factor:
                result['prior_mgmt_adj'] = prior_mgmt_adj
                result['prior_env_factor'] = prior_env_factor
                print(f"    Prior ACL adjustments (WARM): "
                      f"{len(prior_mgmt_adj)} pools with mgmt adj, "
                      f"{len(prior_env_factor)} pools with env factor")
        return result

    if '> Detail_HIst Balances' not in wb.sheetnames:
        wb.close()
        return {}

    ws = wb['> Detail_HIst Balances']

    # Parse structure:
    # Row 5+: pool blocks, each has:
    #   pool_name row  (col A = name)
    #   "Current Grade" row  (col A = "Current Grade", cols B+ = dates)
    #   grade rows  (col A = grade label, cols B+ = values)
    #   "Total" row  (col A = "Total", cols B+ = values)
    #   blank spacer row

    result = {}
    hist_bal_data = {}
    pool_order = []
    acl_months = {}
    risk_rated = {}

    # Read all cell values into memory for fast random access
    rows_data = []
    for row in ws.iter_rows(min_row=1, max_row=ws.max_row,
                            max_col=ws.max_column, values_only=True):
        rows_data.append(list(row))

    # Also read CO-Recov-DQ sheet if present
    co_rows_data = []
    if 'Display CO-Recov-DQ' in wb.sheetnames:
        ws_co = wb['Display CO-Recov-DQ']
        for row in ws_co.iter_rows(min_row=1, max_row=ws_co.max_row,
                                   max_col=ws_co.max_column, values_only=True):
            co_rows_data.append(list(row))

    # Read monthly CO detail sheet if present
    co_monthly_rows = []
    if '>Detail_Charge off Hist' in wb.sheetnames:
        ws_co_det = wb['>Detail_Charge off Hist']
        for row in ws_co_det.iter_rows(min_row=1, max_row=ws_co_det.max_row,
                                       max_col=ws_co_det.max_column, values_only=True):
            co_monthly_rows.append(list(row))

    # Read ACL sheet to carry forward management adjustments & env factors
    acl_rows_data = []
    if 'ACL Env by Pool Mgmt Adj' in wb.sheetnames:
        ws_acl = wb['ACL Env by Pool Mgmt Adj']
        for row in ws_acl.iter_rows(min_row=1, max_row=ws_acl.max_row,
                                    max_col=ws_acl.max_column, values_only=True):
            acl_rows_data.append(list(row))

    wb.close()

    if not rows_data:
        return {}

    max_col = max(len(r) for r in rows_data)
    r = 4  # 0-indexed, start at row 5 (Excel row 5)
    while r < len(rows_data):
        row_vals = rows_data[r]
        a_val = row_vals[0] if row_vals else None

        if a_val is None or str(a_val).strip() == '':
            r += 1
            continue

        pool_name = str(a_val).strip()
        if pool_name in ('Current Grade', 'Total', 'Balance', '% of Loans',
                         'WARM\nMonths', 'Loss Factor Historical Detail'):
            r += 1
            continue
        # Skip header rows
        if pool_name.startswith('For Quarter') or pool_name == cu:
            r += 1
            continue

        # This should be a pool name; next row should be "Current Grade" or "Balance"
        if r + 1 >= len(rows_data):
            break
        next_a = rows_data[r + 1][0] if rows_data[r + 1] else None
        next_label = str(next_a).strip() if next_a else ''

        if next_label not in ('Current Grade', 'Balance'):
            r += 1
            continue

        pool_order.append(pool_name)
        is_rr = (next_label == 'Current Grade')
        risk_rated[pool_name] = is_rr

        # Read dates from the header row (row r+1)
        hdr_row = rows_data[r + 1]
        dates = []
        date_start_col = 1  # 0-indexed col B
        for ci in range(date_start_col, len(hdr_row)):
            v = hdr_row[ci]
            if v is None:
                continue
            if isinstance(v, str):
                vs = v.strip()
                if 'WARM' in vs:
                    # ACL months column — read value from first data row
                    for gr in range(r + 2, min(r + 20, len(rows_data))):
                        acl_val = rows_data[gr][ci] if ci < len(rows_data[gr]) else None
                        if acl_val is not None:
                            try:
                                acl_months[pool_name] = int(acl_val)
                            except (ValueError, TypeError):
                                pass
                            break
                    break
                # Skip non-date string headers like "% of Loans"
                continue
            try:
                dates.append(pd.Timestamp(v))
            except Exception:
                continue

        if not is_rr:
            # Non-risk-rated: just has Balance header + Total row
            r += 2  # skip to Total row
            total_row = rows_data[r] if r < len(rows_data) else []
            total_vals = []
            for ci in range(date_start_col, date_start_col + len(dates)):
                v = total_row[ci] if ci < len(total_row) else 0
                try:
                    total_vals.append(float(v) if v is not None else 0.0)
                except (ValueError, TypeError):
                    total_vals.append(0.0)
            hist_bal_data[pool_name] = {
                'dates': dates,
                'grades': {},
                'total': total_vals,
            }
            r += 1
            continue

        # Risk-rated pool: read grade rows
        pool_grades = {}
        pool_total = []
        gr_idx = r + 2  # first grade row
        while gr_idx < len(rows_data):
            ga = rows_data[gr_idx][0] if rows_data[gr_idx] else None
            if ga is None or str(ga).strip() == '':
                break
            glabel = str(ga).strip()
            vals = []
            for ci in range(date_start_col, date_start_col + len(dates)):
                v = rows_data[gr_idx][ci] if ci < len(rows_data[gr_idx]) else 0
                try:
                    vals.append(float(v) if v is not None else 0.0)
                except (ValueError, TypeError):
                    vals.append(0.0)
            if glabel == 'Total':
                pool_total = vals
                gr_idx += 1
                break
            else:
                pool_grades[glabel] = vals
            gr_idx += 1

        hist_bal_data[pool_name] = {
            'dates': dates,
            'grades': pool_grades,
            'total': pool_total,
        }
        r = gr_idx
        continue

    if hist_bal_data:
        result['hist_bal_data'] = hist_bal_data
        result['pool_order'] = pool_order
        if acl_months:
            result['acl_months'] = acl_months
        if risk_rated:
            result['risk_rated'] = risk_rated
        n_dates = max(len(d.get('dates', [])) for d in hist_bal_data.values())
        print(f"    Prior TCT hist bal: {len(hist_bal_data)} pools, {n_dates} months (from {prior_date})")

    # ── Parse CO/Recovery data from prior report's Display CO-Recov-DQ ──
    if co_rows_data:
        warm_co, warm_rc, warm_net = {}, {}, {}
        warm_co_totals, warm_rc_totals, warm_net_co = {}, {}, {}
        warm_dq_pct = {}

        def _parse_co_section(rows, start_label):
            """Parse a CO/RC/Net/DQ section from the Display CO-Recov-DQ sheet.

            Returns (year_data, totals, acl_col_years).
            year_data = {year(int): {pool: value}}
            totals = {pool: total}
            acl_col_years = list of year ints from header row
            """
            import re
            year_data = {}
            totals = {}
            header_row = None
            pool_start = None
            col_years = []
            acl_col = None  # column index of the ACL totals column

            for ri, row in enumerate(rows):
                c0 = str(row[0] or '').strip().lower()
                if c0 == start_label.lower():
                    header_row = ri
                    # Parse year columns from header
                    for ci in range(1, len(row)):
                        val = row[ci]
                        if val is None:
                            continue
                        sv = str(val).strip()
                        # Match year (e.g. "2019") or "YTD 2026"
                        m = re.match(r'(?:YTD\s+)?(\d{4})', sv)
                        if m:
                            col_years.append((ci, int(m.group(1))))
                        elif 'acl' in sv.lower() or 'net charge' in sv.lower():
                            acl_col = ci
                    pool_start = ri + 1
                    break

            if header_row is None:
                return year_data, totals, []

            # Read pool rows until blank or next section header
            for ri in range(pool_start, len(rows)):
                row = rows[ri]
                pool_name = str(row[0] or '').strip()
                if not pool_name:
                    break
                # Stop at next section header
                pl = pool_name.lower()
                if any(kw in pl for kw in ['recoveries', 'net charge', 'dq %',
                                            'charge offs']):
                    break

                for ci, yr in col_years:
                    val = row[ci] if ci < len(row) else None
                    if val is not None and val != 0:
                        try:
                            fval = float(val)
                        except (ValueError, TypeError):
                            continue
                        year_data.setdefault(yr, {})[pool_name] = fval

                # ACL total column
                if acl_col is not None and acl_col < len(row):
                    aval = row[acl_col]
                    if aval is not None:
                        try:
                            totals[pool_name] = float(aval)
                        except (ValueError, TypeError):
                            pass

            return year_data, totals, [yr for _, yr in col_years]

        # Parse Charge offs section
        co_data, co_tots, co_years = _parse_co_section(co_rows_data,
                                                        'Charge offs')
        if co_data:
            warm_co = co_data
            warm_co_totals = co_tots

        # Parse Recoveries section — values may be negative in prior WARM
        rc_data, rc_tots, _ = _parse_co_section(co_rows_data, 'Recoveries')
        if rc_data:
            # Ensure recovery values are positive (WARM stores them negative)
            for yr in rc_data:
                for pool in rc_data[yr]:
                    rc_data[yr][pool] = abs(rc_data[yr][pool])
            for pool in rc_tots:
                rc_tots[pool] = abs(rc_tots[pool])
            warm_rc = rc_data
            warm_rc_totals = rc_tots

        # Parse Net Charge offs section
        net_data, net_tots, _ = _parse_co_section(co_rows_data,
                                                   'Net Charge offs')
        if net_data:
            warm_net = net_data
            warm_net_co = net_tots

        # Parse DQ % section
        dq_data, _, _ = _parse_co_section(co_rows_data, 'DQ %')
        if dq_data:
            warm_dq_pct = dq_data

        if warm_co:
            result['warm_co'] = warm_co
            result['warm_rc'] = warm_rc
            result['warm_net'] = warm_net
            result['warm_co_totals'] = warm_co_totals
            result['warm_rc_totals'] = warm_rc_totals
            result['warm_net_co'] = warm_net_co
            if warm_dq_pct:
                result['warm_dq_pct'] = warm_dq_pct
            n_pools = len(set(p for yr in warm_co.values() for p in yr))
            n_years = len(warm_co)
            print(f"    Prior TCT CO/RC: {n_pools} pools, {n_years} years"
                  f" (CO totals: {len(warm_co_totals)} pools)")

    # ── Parse monthly CO detail from prior report ──
    if co_monthly_rows:
        import datetime as _dt
        warm_co_monthly = {}
        warm_rc_monthly = {}

        def _parse_monthly_section(rows, start_label, pool_list):
            """Parse a monthly CO or RC section.

            Returns {(year, month): {pool: amount}}.
            """
            monthly = {}
            header_row = None
            date_cols = []  # [(col_idx, datetime)]

            for ri, row in enumerate(rows):
                c0 = str(row[0] or '').strip().lower()
                if c0 == start_label.lower():
                    header_row = ri
                    # Header is one row above the pool rows; dates in cols B+
                    # Actually the header with dates is this row itself
                    # But sometimes the dates are in an earlier header row
                    break

            if header_row is None:
                return monthly

            # Find the date header row (row 5 for CO, check the previous
            # occurrence of dates)
            # Dates are in the same row as the section label or in a prior row
            # Check if this row has dates
            for ci in range(1, len(rows[header_row])):
                val = rows[header_row][ci]
                if isinstance(val, _dt.datetime):
                    date_cols.append((ci, val))

            # If no dates in header row, check above
            if not date_cols:
                # The dates are typically in row 5 (index 4)
                for ri in range(header_row - 1, -1, -1):
                    for ci in range(1, len(rows[ri])):
                        val = rows[ri][ci]
                        if isinstance(val, _dt.datetime):
                            date_cols.append((ci, val))
                    if date_cols:
                        break

            if not date_cols:
                return monthly

            # Read pool rows
            for pi, pool in enumerate(pool_list):
                ri = header_row + 1 + pi
                if ri >= len(rows):
                    break
                row = rows[ri]
                row_label = str(row[0] or '').strip()
                # Verify pool name matches (or just read in order)
                for ci, dt_val in date_cols:
                    val = row[ci] if ci < len(row) else None
                    if val is not None and val != 0:
                        try:
                            fval = float(val)
                        except (ValueError, TypeError):
                            continue
                        ym = (dt_val.year, dt_val.month)
                        monthly.setdefault(ym, {})[pool] = fval

            return monthly

        # Get pool list from CO section (row labels)
        monthly_pools = []
        for ri, row in enumerate(co_monthly_rows):
            c0 = str(row[0] or '').strip().lower()
            if c0 == 'charge offs':
                # Read pool names from subsequent rows until "Total"
                for pi in range(ri + 1, len(co_monthly_rows)):
                    prow = co_monthly_rows[pi]
                    pname = str(prow[0] or '').strip()
                    if not pname or pname.lower().startswith('total'):
                        break
                    monthly_pools.append(pname)
                break

        if monthly_pools:
            warm_co_monthly = _parse_monthly_section(
                co_monthly_rows, 'Charge offs', monthly_pools)
            warm_rc_monthly = _parse_monthly_section(
                co_monthly_rows, 'Recoveries', monthly_pools)

            if warm_co_monthly:
                result['warm_co_monthly'] = warm_co_monthly
            if warm_rc_monthly:
                result['warm_rc_monthly'] = warm_rc_monthly
            n_mo = len(warm_co_monthly)
            print(f"    Prior TCT monthly CO: {n_mo} months, "
                  f"{len(monthly_pools)} pools")

    # ── Parse management adjustments & env factors from prior ACL sheet ──
    if acl_rows_data:
        prior_mgmt_adj = {}   # {pool: {grade: float}}
        prior_env_factor = {}  # {pool: float}
        current_pool = None
        for row in acl_rows_data:
            a_val = row[0] if row else None
            if a_val is None:
                continue
            label = str(a_val).strip()
            if not label:
                continue
            # Detect pool header: next row-ish will have 'Current Grade'
            # Grade rows have data in columns B-H; pool headers don't have col E data
            e_val = row[4] if len(row) > 4 else None
            f_val = row[5] if len(row) > 5 else None
            i_val = row[8] if len(row) > 8 else None
            if label == 'Current Grade':
                continue
            if label == 'Total':
                # Pool total row — read env factor (col I, index 8)
                if current_pool and i_val is not None:
                    try:
                        prior_env_factor[current_pool] = float(i_val)
                    except (ValueError, TypeError):
                        pass
                current_pool = None
                continue
            # If this row has a base_rate in col E but no 'Total' label,
            # it's a grade data row
            if current_pool and e_val is not None:
                try:
                    mgmt = float(f_val) if f_val is not None else 0.0
                except (ValueError, TypeError):
                    mgmt = 0.0
                if mgmt != 0:
                    prior_mgmt_adj.setdefault(current_pool, {})[label] = mgmt
                continue
            # Otherwise this might be a pool header
            if label not in ('Current Grade', 'Total') and e_val is None:
                current_pool = label

        if prior_mgmt_adj or prior_env_factor:
            result['prior_mgmt_adj'] = prior_mgmt_adj
            result['prior_env_factor'] = prior_env_factor
            n_pools_ma = len(prior_mgmt_adj)
            n_pools_ef = len(prior_env_factor)
            print(f"    Prior ACL adjustments: {n_pools_ma} pools with mgmt adj, "
                  f"{n_pools_ef} pools with env factor")

    return result


def build_hist_bal_from_monthly(monthly_balances, df, snap, grades, config):
    """Build a fresh hist_bal_data dict from monthly pool balances + current snapshot.

    Used when no WARM file and no prior TCT report exist for the credit union.
    Pool-level monthly totals come from the monthly balances workbook; grade-
    level distribution is allocated proportionally using the current snapshot's
    grade mix per pool.

    Returns dict shaped like load_prior_tct_hist_bal output:
      {'hist_bal_data': {...}, 'pool_order': [...], 'risk_rated': {...}}
    or {} if no monthly balance data is available.
    """
    if monthly_balances is None or monthly_balances.empty:
        return {}

    no_score = config.get('no_score_label', 'Not Reported')
    all_gl = [g['label'] for g in grades] + [no_score]
    snap_ts = pd.Timestamp(snap) + pd.offsets.MonthEnd(0)

    # Map monthly-file pool names to DB pool names (case-insensitive, strip parens)
    db_pools = list(df['loan_pool'].dropna().unique())
    pool_norm = {}
    for p in db_pools:
        pool_norm[p.strip().lower()] = p
        clean = re.sub(r'\s*\(.*\)\s*$', '', str(p)).strip()
        if clean.lower() != str(p).strip().lower():
            pool_norm[clean.lower()] = p

    hist_bal_data = {}
    pool_order = []
    risk_rated = {}

    for pool_key, grp in monthly_balances.groupby('pool'):
        pk = str(pool_key).strip()
        mapped = pool_norm.get(pk.lower())
        if not mapped:
            clean = re.sub(r'\s*\(.*\)\s*$', '', pk).strip()
            mapped = pool_norm.get(clean.lower())
        if not mapped:
            # Pool exists in monthly file but not in current DB snapshot — keep
            # it as a non-risk-rated pool so its history still displays.
            mapped = pk
            is_rr = False
        else:
            is_rr = True

        # Build current-snapshot grade percentages for this pool
        pcts = {}
        if is_rr:
            pdf = df[df['loan_pool'] == mapped]
            ptotal = pdf['current_balance'].sum()
            if ptotal > 0:
                for g in all_gl:
                    gbal = pdf[pdf['current_grade'] == g]['current_balance'].sum()
                    pcts[g] = gbal / ptotal
            else:
                is_rr = False

        # Sorted, unique month-end dates from monthly file
        sorted_grp = grp.sort_values('date')
        dates = []
        totals = []
        seen = set()
        for _, row in sorted_grp.iterrows():
            dt = pd.Timestamp(row['date']) + pd.offsets.MonthEnd(0)
            if dt in seen:
                continue
            seen.add(dt)
            dates.append(dt)
            totals.append(float(row['balance']))

        # Ensure current snapshot date is the last entry (use DB total if present)
        if is_rr:
            db_total = float(df[df['loan_pool'] == mapped]['current_balance'].sum())
        else:
            db_total = None
        if snap_ts not in seen:
            dates.append(snap_ts)
            totals.append(db_total if db_total is not None else 0.0)
        elif db_total is not None:
            idx = dates.index(snap_ts)
            totals[idx] = db_total

        if not dates:
            continue

        if is_rr and pcts:
            # Use exact grade balances for the snap month from the DB; allocate
            # earlier months proportionally with the same percentages.
            grade_vals = {g: [] for g in all_gl}
            pdf = df[df['loan_pool'] == mapped]
            for di, dt in enumerate(dates):
                pool_total = totals[di]
                if dt == snap_ts:
                    for g in all_gl:
                        grade_vals[g].append(
                            float(pdf[pdf['current_grade'] == g]['current_balance'].sum())
                        )
                else:
                    for g in all_gl:
                        grade_vals[g].append(pool_total * pcts.get(g, 0.0))
        else:
            grade_vals = {}

        hist_bal_data[mapped] = {
            'dates': dates,
            'grades': grade_vals,
            'total': totals,
        }
        pool_order.append(mapped)
        risk_rated[mapped] = bool(grade_vals)

    # Add pools that are in DB snapshot but not in monthly file (single-point entry)
    for pool in db_pools:
        if pool in hist_bal_data:
            continue
        pdf = df[df['loan_pool'] == pool]
        ptotal = float(pdf['current_balance'].sum())
        grade_vals = {}
        for g in all_gl:
            grade_vals[g] = [float(pdf[pdf['current_grade'] == g]['current_balance'].sum())]
        hist_bal_data[pool] = {
            'dates': [snap_ts],
            'grades': grade_vals,
            'total': [ptotal],
        }
        pool_order.append(pool)
        risk_rated[pool] = True

    if not hist_bal_data:
        return {}

    # Reorder pool_order to follow config['pool_order'] (with config-listed pools
    # first in their declared order, then any extra pools alphabetically). This
    # ensures every TCT sheet that consults impaired['pool_order'] uses the
    # same canonical order as the rest of the report.
    cfg_order = config.get('pool_order', []) or []
    nrr = set(config.get('not_risk_rated', []) or [])
    order_idx = {name: i for i, name in enumerate(cfg_order)}
    fallback = len(cfg_order)
    rr_pools = [p for p in pool_order if p not in nrr]
    nrr_pools = [p for p in pool_order if p in nrr]
    rr_pools.sort(key=lambda p: (order_idx.get(p, fallback), str(p)))
    nrr_pools.sort(key=lambda p: (order_idx.get(p, fallback), str(p)))
    pool_order = rr_pools + nrr_pools

    return {
        'hist_bal_data': hist_bal_data,
        'pool_order': pool_order,
        'risk_rated': risk_rated,
    }


def _grade_pct_from_last_month(pdata):
    """Return grade percentage distribution from the most recent month with data.

    Looks backwards through the pool's date list for the last month where
    grades sum to a nonzero value.  Returns {grade_label: fraction} or {}
    if the pool has no grade data at all.
    """
    grades = pdata.get('grades', {})
    if not grades:
        return {}
    n = len(pdata['dates'])
    for i in range(n - 1, -1, -1):
        total = sum(vals[i] for vals in grades.values())
        if total > 0:
            return {g: vals[i] / total for g, vals in grades.items()}
    return {}


def extend_hist_bal_with_monthly(hist_bal_data, monthly_balances):
    """Extend hist_bal_data with pool-level monthly balance records.

    Only adds months *after* the last date already in hist_bal_data.
    Monthly file dates are normalized to month-end for consistency.
    Grade values are distributed proportionally using the most recent month's
    grade percentages from the prior report.
    """
    if monthly_balances is None or monthly_balances.empty:
        return

    # Find the latest date already in hist_bal_data
    latest_existing = pd.Timestamp.min
    for pdata in hist_bal_data.values():
        for d in pdata.get('dates', []):
            ts = pd.Timestamp(d)
            if ts > latest_existing:
                latest_existing = ts

    if latest_existing == pd.Timestamp.min:
        return  # no existing data to extend from

    # Pre-compute grade percentage distributions per pool (before adding new months)
    grade_pcts = {}
    for pool, pdata in hist_bal_data.items():
        grade_pcts[pool] = _grade_pct_from_last_month(pdata)

    # Pool name normalization map (handle trailing spaces, parentheticals)
    pool_norm = {}
    for p in hist_bal_data:
        pool_norm[p.strip().lower()] = p
        # Also handle parenthetical variants e.g. "Re-write(RW TM)" → "Re-write"
        clean = re.sub(r'\s*\(.*\)\s*$', '', p).strip()
        if clean.lower() != p.strip().lower():
            pool_norm[clean.lower()] = p

    for pool_key, grp in monthly_balances.groupby('pool'):
        pk = str(pool_key).strip()
        norm_key = pk.lower()
        mapped = pool_norm.get(norm_key)
        if not mapped:
            clean = re.sub(r'\s*\(.*\)\s*$', '', pk).strip()
            mapped = pool_norm.get(clean.lower())
        if not mapped:
            continue

        pdata = hist_bal_data[mapped]
        pcts = grade_pcts.get(mapped, {})

        for _, row in grp.sort_values('date').iterrows():
            dt = pd.Timestamp(row['date'])
            # Normalize to month-end
            dt = dt + pd.offsets.MonthEnd(0)
            # Only add months after what the prior report already had
            if dt <= latest_existing:
                continue
            # Check not already present (e.g. from DB extension)
            if dt in set(pd.Timestamp(d) for d in pdata['dates']):
                continue
            pool_total = float(row['balance'])
            pdata['dates'].append(dt)
            pdata['total'].append(pool_total)
            # Distribute total across grades using prior month's percentages
            for g, vals in pdata.get('grades', {}).items():
                vals.append(pool_total * pcts.get(g, 0.0))


def extend_hist_bal_with_db(hist_bal_data, df, snap, grades, config):
    """Extend hist_bal_data with new months from the current DB snapshot.

    Computes grade-level balances for each pool from `df` and appends them
    as new monthly columns *after* whatever the prior report already contains.
    Only adds months not yet present.
    """
    no_score = config.get('no_score_label', 'Not Reported')
    snap_ts = pd.Timestamp(snap)

    for pool, pdata in hist_bal_data.items():
        existing_dates = [pd.Timestamp(d) for d in pdata.get('dates', [])]
        pdf = df[df['loan_pool'] == pool]
        pgrades = pdata.get('grades', {})

        if snap_ts in existing_dates:
            # Already present (e.g. from monthly file) — update grade values in-place
            idx = existing_dates.index(snap_ts)
            if pgrades:
                for g in list(pgrades.keys()):
                    bal = pdf[pdf['current_grade'] == g]['current_balance'].sum()
                    pgrades[g][idx] = bal
            pdata['total'][idx] = pdf['current_balance'].sum()
            continue

        if pgrades:
            # Risk-rated: compute per-grade balances
            for g in list(pgrades.keys()):
                bal = pdf[pdf['current_grade'] == g]['current_balance'].sum()
                pgrades[g].append(bal)
        pool_total = pdf['current_balance'].sum()
        pdata['total'].append(pool_total)
        pdata['dates'].append(snap_ts)

    # Handle pools in DB that aren't in hist_bal_data yet (new pools)
    all_gl = [g['label'] for g in grades]
    all_gl.append(no_score)
    for pool in df['loan_pool'].unique():
        if pool in hist_bal_data:
            continue
        pdf = df[df['loan_pool'] == pool]
        pool_grades = {}
        for g in all_gl:
            pool_grades[g] = [pdf[pdf['current_grade'] == g]['current_balance'].sum()]
        hist_bal_data[pool] = {
            'dates': [snap_ts],
            'grades': pool_grades,
            'total': [pdf['current_balance'].sum()],
        }


def load_historical_data(config):
    """Load all historical data for a client. Returns a dict with all historical DataFrames."""
    print("  Loading historical data...")
    co_rec = load_chargeoff_recovery_history(config)
    balances, alll_by_date = load_monthly_balances(config)
    dq = load_delinquency_history(config)

    # Compute annual average balances per pool from monthly data
    avg_balances = {}  # {year: {pool: avg_balance}}
    if not balances.empty:
        balances['year'] = balances['date'].dt.year
        for (year, pool), grp in balances.groupby(['year', 'pool']):
            avg_balances.setdefault(int(year), {})
            avg_balances[int(year)][pool] = grp['balance'].mean()

    # Compute delinquency % per pool per year
    dq_pct = {}  # {year: {pool: dq_pct}}
    for qlabel, pools in dq.items():
        year = int(qlabel[:4])
        for pool, dq_bal in pools.items():
            # Get total balance for that pool at that time
            total = avg_balances.get(year, {}).get(pool, 0)
            if total > 0:
                pct = dq_bal / total
                dq_pct.setdefault(year, {})
                # Average across quarters within a year
                if pool in dq_pct[year]:
                    dq_pct[year][pool] = (dq_pct[year][pool] + pct) / 2
                else:
                    dq_pct[year][pool] = pct

    hist = {
        'chargeoffs': co_rec['chargeoffs'],
        'recoveries': co_rec['recoveries'],
        'years': co_rec['years'],
        'co_monthly': co_rec.get('co_monthly', {}),
        'rc_monthly': co_rec.get('rc_monthly', {}),
        'monthly_balances': balances,
        'avg_balances': avg_balances,
        'delinquency': dq,
        'dq_pct': dq_pct,
        'alll_by_date': alll_by_date,
    }

    # Print summary
    if co_rec['years']:
        print(f"    Charge-off/recovery years: {co_rec['years'][0]}-{co_rec['years'][-1]}")
        total_co = sum(sum(p.values()) for p in co_rec['chargeoffs'].values())
        total_rc = sum(sum(p.values()) for p in co_rec['recoveries'].values())
        print(f"    Total charge-offs: ${total_co:,.2f}  Recoveries: ${total_rc:,.2f}")
    if not balances.empty:
        print(f"    Monthly balance records: {len(balances)} ({balances['date'].min().strftime('%Y-%m')} to {balances['date'].max().strftime('%Y-%m')})")
    if dq:
        print(f"    Delinquency quarters: {len(dq)}")

    return hist


def _find_prior_tct_report(config, snap):
    """Return path to the most recent prior TCT report (snap_date < snap), or None."""
    cu = config['credit_union']
    safe_cu = cu.replace(' ', '_').replace('/', '-')
    rpt_dir = os.path.join(BASE, 'Reports')
    if not os.path.isdir(rpt_dir):
        return None
    pattern = re.compile(
        rf'(\d{{4}}-\d{{2}}-\d{{2}})_CECL_Migration_{re.escape(safe_cu)}_TCT_Model\.xlsx$',
        re.IGNORECASE,
    )
    candidates = []
    for root, dirs, files in os.walk(rpt_dir):
        for f in files:
            if f.startswith('~$'):
                continue
            m = pattern.match(f)
            if not m:
                continue
            d = m.group(1)
            # Allow equal-date matches only when the file lives in the
            # dedicated WARM-baseline subdir; same-date top-level files
            # are this run's own output.
            in_warm_baselines = '_warm_baselines' in os.path.normpath(
                os.path.relpath(root, rpt_dir)
            ).split(os.sep)
            if d < snap or (d == snap and in_warm_baselines):
                candidates.append((d, in_warm_baselines,
                                   os.path.join(root, f)))
    if not candidates:
        return None
    candidates.sort(key=lambda c: (c[0], c[1]), reverse=True)
    return candidates[0][2]


def _load_acl_months_from_tct(filepath):
    """Read WARM Months column from a prior TCT report's '> Detail_HIst Balances' sheet."""
    from openpyxl import load_workbook as _load_wb
    try:
        wb = _load_wb(filepath, read_only=True, data_only=True)
    except Exception:
        return {}
    if '> Detail_HIst Balances' not in wb.sheetnames:
        wb.close()
        return {}
    ws = wb['> Detail_HIst Balances']
    rows = [list(r) for r in ws.iter_rows(min_row=1, max_row=ws.max_row,
                                          max_col=ws.max_column, values_only=True)]
    wb.close()

    acl_months = {}
    r = 4
    while r < len(rows):
        a = rows[r][0] if rows[r] else None
        if a is None or str(a).strip() == '':
            r += 1
            continue
        pool_name = str(a).strip()
        if pool_name in ('Current Grade', 'Total', 'Balance', '% of Loans',
                         'WARM\nMonths', 'Loss Factor Historical Detail'):
            r += 1
            continue
        if r + 1 >= len(rows):
            break
        next_a = rows[r + 1][0] if rows[r + 1] else None
        next_label = str(next_a).strip() if next_a else ''
        if next_label not in ('Current Grade', 'Balance'):
            r += 1
            continue
        # Find WARM column in header row r+1
        hdr = rows[r + 1]
        warm_ci = None
        for ci, v in enumerate(hdr):
            if isinstance(v, str) and 'WARM' in v:
                warm_ci = ci
                break
        if warm_ci is not None:
            for gr in range(r + 2, min(r + 20, len(rows))):
                v = rows[gr][warm_ci] if warm_ci < len(rows[gr]) else None
                if v is not None:
                    try:
                        acl_months[pool_name] = int(v)
                    except (ValueError, TypeError):
                        pass
                    break
        r += 2
    return acl_months


def _find_prior_warm_xlsx(config, snap):
    """Return path to the most recent prior CECL-Migration-WARM xlsx (snap_prefix < snap).

    Searches data_directory and fallback_report_folder. Skips ~$ temp files,
    DNU prefixes, and any non-.xlsx files (e.g. PDFs).
    """
    data_dir = config.get('data_directory', '')
    if not data_dir:
        return None
    if not os.path.isabs(data_dir):
        data_dir = os.path.join(BASE, data_dir)

    cu = config['credit_union']
    snap_prefix = snap[:7] if snap else ''

    search_dirs = [data_dir]
    fb_folder = config.get('credit_pull', {}).get('fallback_report_folder', '')
    if fb_folder and fb_folder != data_dir:
        if not os.path.isabs(fb_folder):
            fb_folder = os.path.join(BASE, fb_folder)
        search_dirs.append(fb_folder)

    pattern = re.compile(r'^(\d{4}-\d{2})(?:-\d{2})?\s+CECL-Migration-WARM.*\.xlsx$',
                         re.IGNORECASE)
    candidates = []
    for sdir in search_dirs:
        if not os.path.isdir(sdir):
            continue
        for root, dirs, files in os.walk(sdir):
            for f in files:
                if f.startswith('~$') or f.upper().startswith('DNU'):
                    continue
                m = pattern.match(f)
                if not m:
                    continue
                pfx = m.group(1)
                if pfx < snap_prefix and cu.lower().split()[0] in f.lower():
                    candidates.append((pfx, os.path.join(root, f)))
    if not candidates:
        return None
    candidates.sort(reverse=True)
    return candidates[0][1]


def _load_acl_months_from_warm_xlsx(filepath):
    """Read the BS CO DQ Data Enter tab and return {pool: acl_months_int}."""
    try:
        bs_df = pd.read_excel(filepath, sheet_name='BS CO DQ Data Enter', header=None)
    except (ValueError, KeyError, FileNotFoundError):
        return {}
    acl_months = {}
    if bs_df.shape[1] <= 6:
        return {}
    for idx in range(4, bs_df.shape[0]):
        pool_name = bs_df.iloc[idx, 0]
        months_val = bs_df.iloc[idx, 6]
        if pd.isna(pool_name) or str(pool_name).strip() == '':
            continue
        pn = str(pool_name).strip()
        if pn.upper().startswith(('HIDE', 'EXCLUDE', 'GRAND TOTAL')):
            continue
        if pd.isna(months_val):
            break
        try:
            acl_months[pn] = int(months_val)
        except (ValueError, TypeError):
            continue
    return acl_months


def _parse_display_co_recov_dq(found):
    """Parse the 'Display CO-Recov -DQ' tab and return a dict of warm_* keys.

    Reads per-year Charge offs / Recoveries / Net Charge offs / DQ % sections
    plus column-J totals (warm_co_totals, warm_rc_totals, warm_net_co).

    Returns dict with any of: warm_co, warm_rc, warm_net, warm_co_totals,
    warm_rc_totals, warm_dq_pct, warm_net_co. Returns {} if tab not present.

    Used by both load_impaired_data (legacy CECL-Migration-WARM file) and
    load_impaired_from_tct_baseline (Reports/_warm_baselines/*_TCT_Model.xlsx).
    """
    out = {}
    try:
        dq_display_df = pd.read_excel(found, sheet_name='Display CO-Recov -DQ',
                                      header=None)
    except (ValueError, KeyError, FileNotFoundError):
        return out

    # Row 3 (0-indexed 2) has year numbers in cols 1..N
    dq_years = []
    for c in range(1, dq_display_df.shape[1]):
        v = dq_display_df.iloc[2, c]
        if pd.notna(v):
            try:
                dq_years.append(int(v))
            except (ValueError, TypeError):
                break
        else:
            break

    # Find the DQ % section header row
    dq_start = None
    for idx in range(len(dq_display_df)):
        val = dq_display_df.iloc[idx, 0]
        if pd.notna(val) and str(val).strip().upper().startswith('DQ'):
            dq_start = idx + 1
            break

    def _parse_section(section_label, exact_start=False):
        """Return {year: {pool: value}} for a section identified by label."""
        sec_start = None
        for idx2 in range(len(dq_display_df)):
            val2 = dq_display_df.iloc[idx2, 0]
            if pd.notna(val2):
                cell_text = str(val2).strip().lower()
                if exact_start:
                    if cell_text.startswith(section_label.lower()):
                        sec_start = idx2 + 1
                        break
                else:
                    if section_label.lower() in cell_text:
                        sec_start = idx2 + 1
                        break
        if sec_start is None or not dq_years:
            return {}
        result = {}
        for idx2 in range(sec_start, min(sec_start + 30, len(dq_display_df))):
            pool_name = dq_display_df.iloc[idx2, 0]
            if pd.isna(pool_name) or str(pool_name).strip() == '':
                break
            pn = str(pool_name).strip()
            if pn.upper().startswith(('HIDE', 'EXCLUDE')):
                continue
            for ci, yr in enumerate(dq_years):
                col_idx = 1 + ci
                if col_idx >= dq_display_df.shape[1]:
                    break
                v = dq_display_df.iloc[idx2, col_idx]
                if pd.notna(v):
                    try:
                        result.setdefault(yr, {})[pn] = float(v)
                    except (ValueError, TypeError):
                        pass
        return result

    def _parse_section_totals(section_label, exact_start=False):
        """Return {pool: acl_total} for a section's total column (col J)."""
        sec_start = None
        for idx2 in range(len(dq_display_df)):
            val2 = dq_display_df.iloc[idx2, 0]
            if pd.notna(val2):
                cell_text = str(val2).strip().lower()
                if exact_start:
                    if cell_text.startswith(section_label.lower()):
                        sec_start = idx2 + 1
                        break
                else:
                    if section_label.lower() in cell_text:
                        sec_start = idx2 + 1
                        break
        if sec_start is None:
            return {}
        total_col = 9
        result = {}
        for idx2 in range(sec_start, min(sec_start + 30, len(dq_display_df))):
            pool_name = dq_display_df.iloc[idx2, 0]
            if pd.isna(pool_name) or str(pool_name).strip() == '':
                break
            pn = str(pool_name).strip()
            if pn.upper().startswith(('HIDE', 'EXCLUDE')):
                continue
            v = (dq_display_df.iloc[idx2, total_col]
                 if dq_display_df.shape[1] > total_col else None)
            if pd.notna(v):
                try:
                    result[pn] = float(v)
                except (ValueError, TypeError):
                    pass
        return result

    warm_co = _parse_section('charge offs', exact_start=True)
    warm_rc = _parse_section('recoveries', exact_start=True)
    warm_net = _parse_section('net charge offs')
    warm_co_totals = _parse_section_totals('charge offs', exact_start=True)
    warm_rc_totals = _parse_section_totals('recoveries', exact_start=True)

    if warm_co:
        out['warm_co'] = warm_co
        out['warm_rc'] = warm_rc
        out['warm_net'] = warm_net
        out['warm_co_totals'] = warm_co_totals
        out['warm_rc_totals'] = warm_rc_totals
        print(f"    WARM CO/RC data: {len(warm_co)} years, "
              f"CO pools: {sum(len(v) for v in warm_co.values())}, "
              f"RC pools: {sum(len(v) for v in warm_rc.values())}")

    if dq_start and dq_years:
        warm_dq_pct = {}
        for idx in range(dq_start, min(dq_start + 30, len(dq_display_df))):
            pool_name = dq_display_df.iloc[idx, 0]
            if pd.isna(pool_name) or str(pool_name).strip() == '':
                break
            pn = str(pool_name).strip()
            if pn.upper().startswith(('HIDE', 'EXCLUDE')):
                continue
            for ci, yr in enumerate(dq_years):
                col_idx = 1 + ci
                if col_idx >= dq_display_df.shape[1]:
                    break
                v = dq_display_df.iloc[idx, col_idx]
                if pd.notna(v):
                    try:
                        warm_dq_pct.setdefault(yr, {})[pn] = float(v)
                    except (ValueError, TypeError):
                        pass
        if warm_dq_pct:
            out['warm_dq_pct'] = warm_dq_pct
            print(f"    WARM DQ% data: {len(warm_dq_pct)} years, "
                  f"{sum(len(v) for v in warm_dq_pct.values())} pool-year entries")

    # Net Chargeoff totals per pool from CO-Recov-DQ (col J total)
    net_co_start = None
    for idx in range(len(dq_display_df)):
        val = dq_display_df.iloc[idx, 0]
        if (pd.notna(val) and 'net' in str(val).strip().lower()
                and 'charge' in str(val).strip().lower()):
            net_co_start = idx + 1
            break
    if net_co_start:
        warm_net_co = {}
        total_col = 9
        for idx in range(net_co_start, min(net_co_start + 30, len(dq_display_df))):
            pool_name = dq_display_df.iloc[idx, 0]
            if pd.isna(pool_name) or str(pool_name).strip() == '':
                break
            pn = str(pool_name).strip()
            if pn.upper().startswith(('HIDE', 'EXCLUDE')):
                continue
            v = (dq_display_df.iloc[idx, total_col]
                 if dq_display_df.shape[1] > total_col else None)
            if pd.notna(v):
                try:
                    warm_net_co[pn] = float(v)
                except (ValueError, TypeError):
                    pass
        if warm_net_co:
            out['warm_net_co'] = warm_net_co
            print(f"    WARM Net CO data: {len(warm_net_co)} pools")

    return out


def load_impaired_data(config, snap):
    """Load impaired-loan summary from the existing CECL-Migration-WARM working file.

    Reads the 'Impaired Loans' tab, column L (category) and P (Sum of Provision Amount)
    from the summary pivot at rows 5-10 (skipping rows labelled 'HIDE').

    Returns dict with:
      'items': {category: provision_amount, ...},
      'total_spec_id': float  (sum of all provision amounts)
    or empty dict if file/tab not found.
    """
    data_dir = config.get('data_directory', '')
    if not data_dir:
        return {}
    # Resolve data_dir (may be absolute or relative)
    if not os.path.isabs(data_dir):
        data_dir = os.path.join(BASE, data_dir)

    cu = config['credit_union']
    safe_cu = cu.replace(' ', '_').replace('/', '-')
    # Build expected filename:  e.g. "2025-12 CECL-Migration-WARM - Franklin Trust FCU.xlsx"
    # snap is like "2025-12-31" — extract YYYY-MM
    snap_prefix = snap[:7] if snap else ''

    # Search for the file in data_dir, then fallback_report_folder
    target_name = f"{snap_prefix} CECL-Migration-WARM - {cu}.xlsx"
    search_dirs = [data_dir]
    fb_folder = config.get('credit_pull', {}).get('fallback_report_folder', '')
    if fb_folder and fb_folder != data_dir:
        if not os.path.isabs(fb_folder):
            fb_folder = os.path.join(BASE, fb_folder)
        search_dirs.append(fb_folder)

    found = None
    for sdir in search_dirs:
        if not os.path.isdir(sdir):
            continue
        for root, dirs, files in os.walk(sdir):
            for f in files:
                if f.startswith('~$') or f.upper().startswith('DNU'):
                    continue
                if f == target_name:
                    found = os.path.join(root, f)
                    break
            if found:
                break
        if found:
            break

    # Fallback: search by pattern
    if not found:
        pattern = re.compile(rf'^{re.escape(snap_prefix)}.*CECL-Migration-WARM.*\.xlsx$', re.IGNORECASE)
        for sdir in search_dirs:
            if not os.path.isdir(sdir):
                continue
            for root, dirs, files in os.walk(sdir):
                for f in files:
                    if f.startswith('~$') or f.upper().startswith('DNU'):
                        continue
                    if pattern.match(f):
                        found = os.path.join(root, f)
                        break
                if found:
                    break
            if found:
                break

    if not found:
        print(f"    No CECL-Migration-WARM file found for {snap_prefix}")
        return {}

    print(f"    Loading impaired loan data from: {os.path.basename(found)}")
    try:
        imp_df = pd.read_excel(found, sheet_name='Impaired Loans', header=None)
    except (ValueError, KeyError):
        print(f"    'Impaired Loans' tab not found in {os.path.basename(found)}")
        return {}

    # The summary pivot is at rows 4-10 (0-indexed: 3-9)
    # Column L = index 11, Column P = index 15
    # Row 20 has the "Total" row; stop before it
    items = {}
    total = 0.0
    for idx in range(4, min(30, len(imp_df))):  # rows 5-30 (1-indexed)
        cat = imp_df.iloc[idx, 11] if imp_df.shape[1] > 11 else None
        prov = imp_df.iloc[idx, 15] if imp_df.shape[1] > 15 else 0
        if cat is None or pd.isna(cat) or str(cat).strip() == '':
            continue
        cat_str = str(cat).strip()
        if cat_str.upper() in ('HIDE', 'TOTAL', 'CALCULATION'):
            continue
        prov_val = 0.0
        try:
            prov_val = float(prov) if pd.notna(prov) else 0.0
        except (ValueError, TypeError):
            continue  # skip non-numeric rows
        items[cat_str] = prov_val
        total += prov_val

    print(f"    Impaired loan categories: {len(items)}, Total: ${total:,.2f}")
    result = {'items': items, 'total_spec_id': total}

    # ── Extract per-pool per-grade "Balance Removed" from detail rows ──
    # Detail rows start after the "Data Entry" marker row.
    # Col Q (16) = Balance Removed, Col R (17) = Loan Pool, Col S (18) = Credit Grade
    spec_id_by_pool = {}   # {pool: {grade: balance_removed, ...}, ...}
    for idx in range(30, len(imp_df)):
        a_val = imp_df.iloc[idx, 0] if imp_df.shape[1] > 0 else None
        if a_val is not None and not pd.isna(a_val):
            lbl = str(a_val).strip()
            if lbl in ('Data Entry', 'Impairment Type'):
                continue
            # This should be a detail row with impairment type in col A
            bal_removed = imp_df.iloc[idx, 16] if imp_df.shape[1] > 16 else 0
            pool_name = imp_df.iloc[idx, 17] if imp_df.shape[1] > 17 else None
            grade = imp_df.iloc[idx, 18] if imp_df.shape[1] > 18 else None
            if pool_name is None or pd.isna(pool_name):
                continue
            pool_str = str(pool_name).strip()
            grade_str = str(grade).strip() if grade is not None and not pd.isna(grade) else ''
            try:
                bal_val = float(bal_removed) if pd.notna(bal_removed) else 0.0
            except (ValueError, TypeError):
                bal_val = 0.0
            if bal_val > 0 and pool_str:
                if pool_str not in spec_id_by_pool:
                    spec_id_by_pool[pool_str] = {}
                spec_id_by_pool[pool_str][grade_str] = (
                    spec_id_by_pool[pool_str].get(grade_str, 0) + bal_val
                )
    if spec_id_by_pool:
        result['spec_id_by_pool'] = spec_id_by_pool
        n_pools = len(spec_id_by_pool)
        total_removed = sum(sum(g.values()) for g in spec_id_by_pool.values())
        print(f"    Specific ID by pool: {n_pools} pools, "
              f"Total removed: ${total_removed:,.2f}")

    # Load Improved/Deteriorated by grade from "Executive Summary (3)" tab
    try:
        es3_df = pd.read_excel(found, sheet_name='Executive Summary (3)', header=None)
        imp_grades = {}   # {grade: balance}
        det_grades = {}   # {grade: balance}

        def _find_section_grades(df, section_keyword):
            """Dynamically find grade/balance rows after a section header."""
            grades = {}
            # Search col C (idx 2) for section header containing the keyword
            start_idx = None
            for idx in range(len(df)):
                val = df.iloc[idx, 2] if df.shape[1] > 2 else None
                if val is None or pd.isna(val):
                    # Also check col B (idx 1) for merged header cells
                    val = df.iloc[idx, 1] if df.shape[1] > 1 else None
                if val is not None and not pd.isna(val):
                    s = str(val).strip().lower()
                    if section_keyword in s and 'summary' in s:
                        start_idx = idx
                        break
            if start_idx is None:
                return grades
            # Find the "Grade" header row after the section header
            for idx in range(start_idx, min(start_idx + 5, len(df))):
                val = df.iloc[idx, 2] if df.shape[1] > 2 else None
                if val is not None and not pd.isna(val) and str(val).strip().lower() == 'grade':
                    # Read grade rows starting from next row
                    for gi in range(idx + 1, min(idx + 15, len(df))):
                        g_val = df.iloc[gi, 2] if df.shape[1] > 2 else None
                        b_val = df.iloc[gi, 3] if df.shape[1] > 3 else 0
                        if g_val is None or pd.isna(g_val):
                            continue
                        g = str(g_val).strip()
                        if g.lower().startswith('total'):
                            break
                        if g.lower().startswith('hide'):
                            continue
                        try:
                            grades[g] = float(b_val) if pd.notna(b_val) else 0.0
                        except (ValueError, TypeError):
                            grades[g] = 0.0
                    break
            return grades

        imp_grades = _find_section_grades(es3_df, 'improved')
        det_grades = _find_section_grades(es3_df, 'deteriorated')
        result['exec_summary_3'] = {'improved': imp_grades, 'deteriorated': det_grades}
        print(f"    Executive Summary (3): {len(imp_grades)} improved grades, {len(det_grades)} deteriorated grades")
    except (ValueError, KeyError):
        print(f"    'Executive Summary (3)' tab not found in {os.path.basename(found)}")

    # Also read Pooled Totals and ACL Balance from "ACL Env by Pool Mgmt Adj" tab
    try:
        acl_df = pd.read_excel(found, sheet_name='ACL Env by Pool Mgmt Adj', header=None)
    except (ValueError, KeyError):
        print(f"    'ACL Env by Pool Mgmt Adj' tab not found")
        return result

    # Search column A for key labels, read value from column K (index 10)
    for idx in range(len(acl_df)):
        label = acl_df.iloc[idx, 0]  # column A
        if pd.isna(label):
            continue
        lbl = str(label).strip()
        k_val = acl_df.iloc[idx, 10] if acl_df.shape[1] > 10 else None  # column K
        if lbl.startswith('Pooled Totals'):
            try:
                result['pooled_total_allowance'] = float(k_val) if pd.notna(k_val) else 0.0
                print(f"    Pooled Total Allowance: ${result['pooled_total_allowance']:,.2f}")
            except (ValueError, TypeError):
                pass
        elif lbl.startswith('Allowance for Credit Loss Balance'):
            try:
                result['acl_balance'] = float(k_val) if pd.notna(k_val) else 0.0
                print(f"    ACL Balance: ${result['acl_balance']:,.2f}")
            except (ValueError, TypeError):
                pass

    # ── Parse the full per-pool per-grade ACL data from the same tab ──
    acl_pools = {}   # {pool_name: {'grades': {grade: {balance, spec_id, calc_bal, base_rate, mgmt_adj, factor, allow_before}}, 'total': {...}}}
    acl_impaired = {}  # {label: allowance}
    acl_summary = {}   # pooled_total_spec_id, total_spec_allow, total_allow_needed, acl_bal, adjustment
    current_pool = None
    current_grades = {}
    for idx in range(len(acl_df)):
        a_val = acl_df.iloc[idx, 0]
        if pd.isna(a_val):
            continue
        label = str(a_val).strip()

        # Pool header row: next row has "Current Grade" header
        if idx + 1 < len(acl_df):
            next_a = acl_df.iloc[idx + 1, 0]
            if pd.notna(next_a) and str(next_a).strip() == 'Current Grade':
                # Save previous pool
                if current_pool and current_grades:
                    acl_pools[current_pool]['grades'] = current_grades
                current_pool = label
                acl_pools[current_pool] = {'grades': {}, 'total': {}}
                current_grades = {}
                continue

        # Grade data row (inside a pool block): A=grade, B=balance, ...
        if current_pool and label not in ('Current Grade', 'Total'):
            b = acl_df.iloc[idx, 1] if acl_df.shape[1] > 1 else 0
            c = acl_df.iloc[idx, 2] if acl_df.shape[1] > 2 else 0
            d = acl_df.iloc[idx, 3] if acl_df.shape[1] > 3 else 0
            e = acl_df.iloc[idx, 4] if acl_df.shape[1] > 4 else 0
            f = acl_df.iloc[idx, 5] if acl_df.shape[1] > 5 else 0
            g = acl_df.iloc[idx, 6] if acl_df.shape[1] > 6 else 0
            h = acl_df.iloc[idx, 7] if acl_df.shape[1] > 7 else 0
            current_grades[label] = {
                'balance': float(b) if pd.notna(b) else 0.0,
                'spec_id': float(c) if pd.notna(c) else 0.0,
                'calc_bal': float(d) if pd.notna(d) else 0.0,
                'base_rate': float(e) if pd.notna(e) else 0.0,
                'mgmt_adj': float(f) if pd.notna(f) else 0.0,
                'factor': float(g) if pd.notna(g) else 0.0,
                'allow_before': float(h) if pd.notna(h) else 0.0,
            }

        # Total row for current pool
        if current_pool and label == 'Total':
            e = acl_df.iloc[idx, 4] if acl_df.shape[1] > 4 else 0
            f_val = acl_df.iloc[idx, 5] if acl_df.shape[1] > 5 else 0
            g_val = acl_df.iloc[idx, 6] if acl_df.shape[1] > 6 else 0
            h = acl_df.iloc[idx, 7] if acl_df.shape[1] > 7 else 0
            i_val = acl_df.iloc[idx, 8] if acl_df.shape[1] > 8 else 0
            j = acl_df.iloc[idx, 9] if acl_df.shape[1] > 9 else 0
            k = acl_df.iloc[idx, 10] if acl_df.shape[1] > 10 else 0
            acl_pools[current_pool]['total'] = {
                'balance': float(acl_df.iloc[idx, 1]) if pd.notna(acl_df.iloc[idx, 1]) else 0.0,
                'spec_id': float(acl_df.iloc[idx, 2]) if pd.notna(acl_df.iloc[idx, 2]) else 0.0,
                'base_rate': float(e) if pd.notna(e) else 0.0,
                'mgmt_adj': float(f_val) if pd.notna(f_val) else 0.0,
                'factor': float(g_val) if pd.notna(g_val) else 0.0,
                'allow_before': float(h) if pd.notna(h) else 0.0,
                'env_factor': float(i_val) if pd.notna(i_val) else 0.0,
                'env_allow': float(j) if pd.notna(j) else 0.0,
                'total_allow': float(k) if pd.notna(k) else 0.0,
            }
            if current_grades:
                acl_pools[current_pool]['grades'] = current_grades
            current_grades = {}
            current_pool = None

        # Pooled Totals row
        if label.startswith('Pooled Totals'):
            acl_summary['pooled_balance'] = float(acl_df.iloc[idx, 1]) if pd.notna(acl_df.iloc[idx, 1]) else 0.0
            acl_summary['pooled_spec_id'] = float(acl_df.iloc[idx, 2]) if pd.notna(acl_df.iloc[idx, 2]) else 0.0
            acl_summary['pooled_allow_before'] = float(acl_df.iloc[idx, 7]) if pd.notna(acl_df.iloc[idx, 7]) else 0.0
            acl_summary['pooled_env_allow'] = float(acl_df.iloc[idx, 9]) if pd.notna(acl_df.iloc[idx, 9]) else 0.0
            acl_summary['pooled_total_allow'] = float(acl_df.iloc[idx, 10]) if pd.notna(acl_df.iloc[idx, 10]) else 0.0

        # Impaired Loans section
        if label in ('Delinquent Loans', 'Known Losses', 'Repossessions',
                      'Foreclosed Real Estate', 'Deceased', 'Bankruptcy'):
            k = acl_df.iloc[idx, 10] if acl_df.shape[1] > 10 else 0
            acl_impaired[label] = float(k) if pd.notna(k) else 0.0

        if label == 'Total Specifically Identified Allowance':
            k = acl_df.iloc[idx, 10] if acl_df.shape[1] > 10 else 0
            acl_summary['total_spec_allow'] = float(k) if pd.notna(k) else 0.0
        if label == 'Total Allowance Needed':
            k = acl_df.iloc[idx, 10] if acl_df.shape[1] > 10 else 0
            acl_summary['total_allow_needed'] = float(k) if pd.notna(k) else 0.0
        if label.startswith('Allowance for Credit Loss Balance'):
            k = acl_df.iloc[idx, 10] if acl_df.shape[1] > 10 else 0
            acl_summary['acl_balance'] = float(k) if pd.notna(k) else 0.0
        if label.startswith('Adjustment'):
            k = acl_df.iloc[idx, 10] if acl_df.shape[1] > 10 else 0
            acl_summary['adjustment'] = float(k) if pd.notna(k) else 0.0

    # Filter out HIDE pools and Exclude
    acl_pools = {k: v for k, v in acl_pools.items()
                 if not k.upper().startswith('HIDE') and k != 'Exclude'}
    result['acl_pools'] = acl_pools
    result['pool_order'] = list(acl_pools.keys())
    result['acl_impaired'] = acl_impaired
    result['acl_summary'] = acl_summary
    print(f"    ACL per-pool data: {len(acl_pools)} pools, "
          f"{len(acl_impaired)} impaired categories")

    # Helper: read ALL pool blocks from a DQ/CO Data Entry tab.
    # Each block has a "Loan Status" header in col P with pool name in col A.
    # Returns (grand_total_dict, per_pool_dict).
    def _read_migration_blocks(sheet_df):
        by_pool = {}
        grand_total = {}
        for idx in range(len(sheet_df)):
            cell_p = sheet_df.iloc[idx, 15] if sheet_df.shape[1] > 15 else None
            if not (pd.notna(cell_p) and str(cell_p).strip() == 'Loan Status'):
                continue
            pool_raw = sheet_df.iloc[idx, 0] if pd.notna(sheet_df.iloc[idx, 0]) else ''
            pool_key = str(pool_raw).strip()
            block = {}
            for di in range(1, 5):
                ri = idx + di
                if ri >= len(sheet_df):
                    break
                status = sheet_df.iloc[ri, 15] if sheet_df.shape[1] > 15 else None
                balance = sheet_df.iloc[ri, 16] if sheet_df.shape[1] > 16 else 0
                pct = sheet_df.iloc[ri, 17] if sheet_df.shape[1] > 17 else 0
                if pd.notna(status):
                    s = str(status).strip()
                    try:
                        block[s] = {
                            'balance': float(balance) if pd.notna(balance) else 0.0,
                            'pct': float(pct) if pd.notna(pct) else 0.0,
                        }
                    except (ValueError, TypeError):
                        pass
            if not block:
                continue
            if pool_key.lower().startswith('grand total'):
                grand_total = block
            elif pool_key.lower().startswith(('hide', 'exclude', 'credit grade', 'risk rated')):
                continue
            else:
                by_pool[pool_key] = block
        return grand_total, by_pool

    # Load DQ by migration status from "DQ Data Entry" tab (all pools + grand total)
    try:
        dq_df = pd.read_excel(found, sheet_name='DQ Data Entry', header=None)
        dq_grand, dq_by_pool = _read_migration_blocks(dq_df)
        if dq_grand:
            result['dq_by_status'] = dq_grand
            total_dq = sum(v['balance'] for v in dq_grand.values())
            print(f"    DQ by migration status: {len(dq_grand)} categories, Total DQ: ${total_dq:,.2f}")
        if dq_by_pool:
            result['dq_by_pool'] = dq_by_pool
            print(f"    DQ per-pool data: {len(dq_by_pool)} pools")
    except (ValueError, KeyError):
        print(f"    'DQ Data Entry' tab not found in {os.path.basename(found)}")

    # Load CO by migration status from "CO Data Entry" tab (all pools + grand total)
    try:
        co_df = pd.read_excel(found, sheet_name='CO Data Entry', header=None)
        co_grand, co_by_pool = _read_migration_blocks(co_df)
        if co_grand:
            result['co_by_status'] = co_grand
            total_co = sum(v['balance'] for v in co_grand.values())
            print(f"    CO by migration status: {len(co_grand)} categories, Total CO: ${total_co:,.2f}")
        if co_by_pool:
            result['co_by_pool'] = co_by_pool
            print(f"    CO per-pool data: {len(co_by_pool)} pools")
    except (ValueError, KeyError):
        print(f"    'CO Data Entry' tab not found in {os.path.basename(found)}")

    # ── Economic Stress Data from "BS CO DQ Data Enter" tab ──
    try:
        bs_df = pd.read_excel(found, sheet_name='BS CO DQ Data Enter', header=None)
        # Row 5 (0-indexed) has headers: L=State, M=County, N=Unemp%, O=FC/Person, P=BK, Q=Population
        # Row 6 (0-indexed) has values
        if bs_df.shape[0] > 6 and bs_df.shape[1] > 16:
            state = bs_df.iloc[6, 11] if pd.notna(bs_df.iloc[6, 11]) else ''
            county = bs_df.iloc[6, 12] if pd.notna(bs_df.iloc[6, 12]) else ''
            unemp = float(bs_df.iloc[6, 13]) if pd.notna(bs_df.iloc[6, 13]) else 0.0
            fc = bs_df.iloc[6, 14] if pd.notna(bs_df.iloc[6, 14]) else 0
            bk = bs_df.iloc[6, 15] if pd.notna(bs_df.iloc[6, 15]) else 0
            pop = bs_df.iloc[6, 16] if pd.notna(bs_df.iloc[6, 16]) else 1
            try:
                fc = int(fc)
            except (ValueError, TypeError):
                fc = 0
            try:
                bk = int(bk)
            except (ValueError, TypeError):
                bk = 0
            try:
                pop = int(pop)
            except (ValueError, TypeError):
                pop = 1
            result['economic_data'] = {
                'state': str(state).strip(),
                'county': str(county).strip(),
                'unemployment_rate': unemp,
                'foreclosures': fc,
                'bankruptcies': bk,
                'population': pop,
            }
            print(f"    Economic stress data: {result['economic_data']['state']}, "
                  f"{result['economic_data']['county']}")

        # ── Risk Rated flag per pool (col B = "Risk Rated Yes/No") ──
        risk_rated = {}
        for idx in range(4, bs_df.shape[0]):
            pool_name = bs_df.iloc[idx, 0]
            rr_val = bs_df.iloc[idx, 1] if bs_df.shape[1] > 1 else None
            if pd.isna(pool_name) or str(pool_name).strip() == '':
                continue
            pn = str(pool_name).strip()
            if pn.upper().startswith(('HIDE', 'EXCLUDE', 'GRAND TOTAL')):
                continue
            risk_rated[pn] = str(rr_val).strip().lower() == 'yes' if pd.notna(rr_val) else True
        if risk_rated:
            result['risk_rated'] = risk_rated
            nr_count = sum(1 for v in risk_rated.values() if not v)
            print(f"    Risk rated flags: {len(risk_rated)} pools ({nr_count} not risk-rated)")

        # ── ACL Months (Life of Loan) per pool ──
        # Row 3 (0-indexed) has header with "ACL Months" at col 6
        # Rows 4+ have pool name (col 0) and ACL months (col 6)
        # Stop at 'Exclude' or 'Grand Total' sentinel rows
        acl_months = {}
        if bs_df.shape[1] > 6:
            for idx in range(4, bs_df.shape[0]):
                pool_name = bs_df.iloc[idx, 0]
                months_val = bs_df.iloc[idx, 6]
                if pd.isna(pool_name) or str(pool_name).strip() == '':
                    continue
                pn = str(pool_name).strip()
                if pn.upper().startswith(('HIDE', 'EXCLUDE', 'GRAND TOTAL')):
                    continue
                if pd.isna(months_val):
                    break  # reached end of ACL months section
                try:
                    acl_months[pn] = int(months_val)
                except (ValueError, TypeError):
                    acl_months[pn] = 36
        if acl_months:
            result['acl_months'] = acl_months
            print(f"    ACL months (life of loan): {len(acl_months)} pools")
    except (ValueError, KeyError):
        print(f"    'BS CO DQ Data Enter' tab not found in {os.path.basename(found)}")

    # ── Environmental Factor Ranges from "Envir Fact Ranges" tab ──
    try:
        ef_df = pd.read_excel(found, sheet_name='Envir Fact Ranges', header=None)
        # Row 6 (0-indexed) is header row: col 1=Range, 2=Score, 3=Range, 4=Score, 5=Range, 6=Score
        # NCC ranges: cols 1-2 starting row 7, DQ: cols 3-4, ES: cols 5-6

        def _parse_range_label(lbl):
            """Parse range labels like '>5.00%', '4.00% to 4.99%', '<-5.00%'."""
            lbl = lbl.replace('%', '').strip()
            if lbl.startswith('>') or lbl.startswith('>='):
                val = float(lbl.lstrip('>= '))
                return (val, 999)
            elif lbl.startswith('<') or lbl.startswith('<='):
                val = float(lbl.lstrip('<= '))
                return (-999, val)
            elif ' to ' in lbl:
                parts = lbl.split(' to ')
                return (float(parts[0].strip()), float(parts[1].strip()) + 0.01)
            return (0, 0)

        def _read_range_col(col_lbl, col_sc):
            """Read label+score columns and return list of (lo, hi, score) + labels."""
            rows = []
            labels = []
            for i in range(7, ef_df.shape[0]):
                sc = ef_df.iloc[i, col_sc] if pd.notna(ef_df.iloc[i, col_sc]) else None
                lbl = str(ef_df.iloc[i, col_lbl]).strip() if pd.notna(ef_df.iloc[i, col_lbl]) else ''
                if sc is not None and lbl:
                    lo, hi = _parse_range_label(lbl)
                    rows.append((lo, hi, round(float(sc) * 100, 2)))
                    labels.append(lbl)
            return rows, labels

        parsed_ncc, lbl_ncc = _read_range_col(1, 2)   # Net Credit Change
        parsed_dq,  lbl_dq  = _read_range_col(3, 4)   # Delinquency
        parsed_es,  lbl_es  = _read_range_col(5, 6)   # Economic Stress

        if parsed_ncc and parsed_dq and parsed_es:
            result['env_ranges'] = {
                'ncc': parsed_ncc,
                'dq': parsed_dq,
                'es': parsed_es,
                'ncc_labels': lbl_ncc,
                'dq_labels': lbl_dq,
                'es_labels': lbl_es,
            }
            print(f"    Env factor ranges: NCC={len(parsed_ncc)}, "
                  f"DQ={len(parsed_dq)}, ES={len(parsed_es)}")
    except (ValueError, KeyError):
        print(f"    'Envir Fact Ranges' tab not found in {os.path.basename(found)}")

    # ── Balance Adjustments per pool from "Risk Change Data Entry" tab ──
    # Col A (0) = Grade, Col M (12) = Loan Report Balance, Col O (14) = Balance Sheet Total,
    # Col P (15) = Bal Adjustment, Col Q (16) = Specific Identification
    # Pool name in col N (13); Total rows have col A = "Total"
    try:
        rc_df = pd.read_excel(found, sheet_name='Risk Change Data Entry', header=None)
        bal_adj = {}       # {pool_name: adjustment_amount}
        pool_bal_detail = {}  # {pool_name: {grade: {loan_report_bal, bal_adj, balance_sheet_total, specific_id}}}
        total_bal_adj = 0.0
        total_in_portfolio = 0.0
        current_pool = None
        current_grades = {}  # grade detail accumulator for current pool
        skip_labels = {'', 'Loan Pool', '% of Loan Balance', 'Grand Total ', 'Grand Total'}

        def _safe_float(val):
            try:
                return float(val) if pd.notna(val) else 0.0
            except (ValueError, TypeError):
                return 0.0

        for idx in range(len(rc_df)):
            n_val = rc_df.iloc[idx, 13] if rc_df.shape[1] > 13 else None
            a_val = rc_df.iloc[idx, 0] if pd.notna(rc_df.iloc[idx, 0]) else None

            # Pool name row: col N has a string that matches pool name (not header labels)
            if pd.notna(n_val) and isinstance(n_val, str):
                nstr = n_val.strip()
                if nstr not in skip_labels:
                    current_pool = nstr
                    current_grades = {}

            if a_val is None or not current_pool:
                continue
            a_str = str(a_val).strip()

            # Skip header row and hidden grades
            if a_str in ('Current Grade', '') or a_str.upper().startswith('HIDE'):
                continue

            # Read per-row values: M=12, O=14, P=15, Q=16
            m_val = _safe_float(rc_df.iloc[idx, 12]) if rc_df.shape[1] > 12 else 0.0
            o_val = _safe_float(rc_df.iloc[idx, 14]) if rc_df.shape[1] > 14 else 0.0
            p_val = _safe_float(rc_df.iloc[idx, 15]) if rc_df.shape[1] > 15 else 0.0
            q_val = _safe_float(rc_df.iloc[idx, 16]) if rc_df.shape[1] > 16 else 0.0

            if a_str == 'Total':
                if not current_pool.upper().startswith('HIDE'):
                    bal_adj[current_pool] = p_val
                    total_bal_adj += p_val
                    total_in_portfolio += o_val
                    # Store grade detail plus total row
                    current_grades['Total'] = {
                        'loan_report_bal': m_val, 'bal_adj': p_val,
                        'balance_sheet_total': o_val, 'specific_id': q_val,
                    }
                    pool_bal_detail[current_pool] = current_grades
                current_pool = None
                current_grades = {}
            else:
                # Regular grade row
                current_grades[a_str] = {
                    'loan_report_bal': m_val, 'bal_adj': p_val,
                    'balance_sheet_total': o_val, 'specific_id': q_val,
                }

        result['balance_adjustments'] = bal_adj
        result['pool_bal_detail'] = pool_bal_detail
        result['total_balance_adjustment'] = round(total_bal_adj, 2)
        result['total_in_portfolio'] = round(total_in_portfolio, 2)
        if abs(total_bal_adj) > 0.01:
            print(f"    Balance adjustments: {sum(1 for v in bal_adj.values() if abs(v) > 0.01)} pools, "
                  f"Total: ${total_bal_adj:,.2f}")
    except (ValueError, KeyError):
        print(f"    'Risk Change Data Entry' tab not found in {os.path.basename(found)}")

    # ── Historical grade-level balances from "HIst Bal Data" tab ──
    try:
        hb_df = pd.read_excel(found, sheet_name='HIst Bal Data', header=None)
        # Row layout per pool block (15 rows):
        #   pool name | blank | ...
        #   "Current Grade" | blank | date1 | date2 | ...
        #   grade_label | blank | val1 | val2 | ...  (11 grades)
        #   "Total" | blank | val1 | val2 | ...
        #   blank row
        # Header rows 1-5 have metadata; dates are in row 5 (idx 4), col C onwards

        # Read dates from row 5 (index 4)
        hist_dates = []
        for c in range(2, hb_df.shape[1]):
            v = hb_df.iloc[4, c] if 4 < len(hb_df) else None
            if pd.notna(v):
                try:
                    hist_dates.append(pd.Timestamp(v))
                except Exception:
                    pass

        hist_bal_data = {}  # {pool: {dates: [...], grades: {grade: [vals]}, total: [vals]}}
        idx = 5  # start scanning after header rows
        while idx < len(hb_df):
            # Look for pool name row: col A has text, next row has "Current Grade"
            a_val = hb_df.iloc[idx, 0] if pd.notna(hb_df.iloc[idx, 0]) else None
            if a_val is None:
                idx += 1
                continue
            pool_name = str(a_val).strip()
            if pool_name in ('', 'Current Grade', 'Total'):
                idx += 1
                continue
            # Check next row is "Current Grade"
            if idx + 1 < len(hb_df):
                next_a = hb_df.iloc[idx + 1, 0]
                if pd.notna(next_a) and str(next_a).strip() == 'Current Grade':
                    # This is a pool header; read grade rows
                    pool_grades = {}
                    pool_total = []
                    gr_idx = idx + 2  # first grade row
                    while gr_idx < len(hb_df):
                        ga = hb_df.iloc[gr_idx, 0]
                        if pd.isna(ga) or str(ga).strip() == '':
                            break
                        glabel = str(ga).strip()
                        vals = []
                        for c in range(2, 2 + len(hist_dates)):
                            v = hb_df.iloc[gr_idx, c] if c < hb_df.shape[1] else 0
                            try:
                                vals.append(float(v) if pd.notna(v) else 0.0)
                            except (ValueError, TypeError):
                                vals.append(0.0)
                        if glabel == 'Total':
                            pool_total = vals
                            gr_idx += 1
                            break
                        else:
                            pool_grades[glabel] = vals
                        gr_idx += 1
                    if not pool_name.upper().startswith('HIDE'):
                        hist_bal_data[pool_name] = {
                            'dates': hist_dates,
                            'grades': pool_grades,
                            'total': pool_total,
                        }
                    idx = gr_idx
                    continue
            idx += 1

        result['hist_bal_data'] = hist_bal_data
        if hist_bal_data:
            print(f"    HIst Bal Data: {len(hist_bal_data)} pools, {len(hist_dates)} months")
    except (ValueError, KeyError):
        print(f"    'HIst Bal Data' tab not found in {os.path.basename(found)}")

    # ── DQ % / CO / RC per year from "Display CO-Recov -DQ" tab ──
    result.update(_parse_display_co_recov_dq(found))

    # ── Monthly CO/RC from WARM "Charge off History" tab ──
    try:
        co_hist_df = pd.read_excel(found, sheet_name='Charge off History', header=None)
        # Row 8 (0-indexed): section header "Charge offs" with dates in cols 2+
        # Rows 9-19: pool CO values (may include HIDE/Exclude rows after visible pools)
        # Row 33: section header "Recoveries" with dates in cols 2+
        # Rows 34-44: pool RC values (negative in WARM)
        co_dates = []
        for c in range(2, co_hist_df.shape[1]):
            v = co_hist_df.iloc[8, c]
            if pd.notna(v):
                try:
                    co_dates.append((c, pd.Timestamp(v)))
                except Exception:
                    pass

        def _parse_co_hist_section(start_idx):
            """Parse pool rows below a section header, returning {(yr,mo): {pool: val}}."""
            out = {}
            for ri in range(start_idx + 1, min(start_idx + 25, len(co_hist_df))):
                pn_raw = co_hist_df.iloc[ri, 0]
                if pd.isna(pn_raw):
                    continue
                pn = str(pn_raw).strip()
                if pn.upper().startswith(('HIDE', 'EXCLUDE', 'TOTAL')):
                    continue
                if pn == '':
                    break
                for ci, dt in co_dates:
                    v = co_hist_df.iloc[ri, ci]
                    val = float(v) if pd.notna(v) else 0.0
                    if val != 0:
                        ym = (dt.year, dt.month)
                        out.setdefault(ym, {})[pn] = val
            return out

        # Find section start rows
        co_start = None
        rc_start = None
        for idx in range(len(co_hist_df)):
            val = co_hist_df.iloc[idx, 0]
            if pd.notna(val):
                txt = str(val).strip().lower()
                if txt == 'charge offs':
                    co_start = idx
                elif txt == 'recoveries':
                    rc_start = idx

        if co_start is not None:
            result['warm_co_monthly'] = _parse_co_hist_section(co_start)
        if rc_start is not None:
            result['warm_rc_monthly'] = _parse_co_hist_section(rc_start)

        n_co = sum(len(v) for v in result.get('warm_co_monthly', {}).values())
        n_rc = sum(len(v) for v in result.get('warm_rc_monthly', {}).values())
        if n_co or n_rc:
            print(f"    WARM Charge off History: "
                  f"{len(result.get('warm_co_monthly', {}))} CO months ({n_co} entries), "
                  f"{len(result.get('warm_rc_monthly', {}))} RC months ({n_rc} entries)")
    except (ValueError, KeyError):
        pass  # tab not found, silently skip

    return result


def load_impaired_from_tct_baseline(config, snap):
    """Load impaired-loan data from the previously-generated TCT model baseline.

    Used when the source CECL-Migration-WARM file has no 'Impaired Loans' tab
    (because the TCT model is replacing it). Reads from
    Reports/_warm_baselines/<snap>_CECL_Migration_<CU>_TCT_Model.xlsx:

      - 'Impaired Loans' tab: pivot summary at rows 4-9 with
          col A=Impairment Type, col B=Provision Percentage,
          col L=Impairment Type, col N=Sum of Loss Given Default
        Allowance per category = Provision Pct * LGD.
      - 'Impaired Loans Pivot' tab: pool x grade pivot of Balance Removed.

    Returns dict with 'acl_impaired', 'spec_id_by_pool', 'total_spec_id'
    or {} if file/tabs not found.
    """
    cu = config['credit_union']
    safe_cu = cu.replace(' ', '_').replace('/', '-')
    baseline_dir = os.path.join(BASE, 'Reports', '_warm_baselines')
    if not os.path.isdir(baseline_dir):
        return {}
    target = f"{snap}_CECL_Migration_{safe_cu}_TCT_Model.xlsx"
    found = None
    for f in os.listdir(baseline_dir):
        if f == target and not f.startswith('~$'):
            found = os.path.join(baseline_dir, f)
            break
    if not found:
        # Fallback: pattern match for any TCT_Model baseline for this CU,
        # preferring the closest-dated baseline (snap exact, then most recent
        # prior, then earliest later if none prior exists). Provides historical
        # Display CO-Recov -DQ / Impaired tabs when no exact-snap baseline
        # exists.
        any_pat = re.compile(
            r'^(\d{4}-\d{2}-\d{2})_CECL_Migration_.*_TCT_Model\.xlsx$',
            re.IGNORECASE)
        prior_candidates = []  # (date, path) where date <= snap
        later_candidates = []  # (date, path) where date > snap
        for f in os.listdir(baseline_dir):
            if f.startswith('~$'):
                continue
            m = any_pat.match(f)
            if not m or safe_cu.lower() not in f.lower():
                continue
            d = m.group(1)
            path = os.path.join(baseline_dir, f)
            if d <= snap:
                prior_candidates.append((d, path))
            else:
                later_candidates.append((d, path))
        if prior_candidates:
            prior_candidates.sort(reverse=True)
            found = prior_candidates[0][1]
        elif later_candidates:
            later_candidates.sort()
            found = later_candidates[0][1]
    if not found:
        return {}

    print(f"    Loading impaired data from TCT baseline: {os.path.basename(found)}")
    try:
        from openpyxl import load_workbook as _lw
        wb = _lw(found, data_only=True, read_only=True)
    except Exception as e:
        print(f"    Warning: could not open baseline: {e}")
        return {}

    result = {}

    # ── Parse 'Impaired Loans' summary pivot for allowance per category ──
    acl_impaired = {}
    if 'Impaired Loans' in wb.sheetnames:
        ws = wb['Impaired Loans']
        # Iterate first ~25 rows; collect category->prov pct from cols A/B,
        # category->LGD from cols L/N.
        prov_pct = {}   # {category: pct}
        lgd_by_cat = {} # {category: lgd}
        for row in ws.iter_rows(min_row=1, max_row=25, max_col=14, values_only=True):
            if not row:
                continue
            a = row[0] if len(row) > 0 else None
            b = row[1] if len(row) > 1 else None
            l = row[11] if len(row) > 11 else None
            n = row[13] if len(row) > 13 else None
            if a and isinstance(a, str):
                a_str = a.strip()
                if a_str and a_str.upper() not in ('HIDE', 'IMPAIRMENT TYPE',
                                                   'TOTAL', 'CALCULATION',
                                                   'DATA ENTRY'):
                    try:
                        if b is not None:
                            prov_pct[a_str] = float(b)
                    except (ValueError, TypeError):
                        pass
            if l and isinstance(l, str):
                l_str = l.strip()
                if l_str and l_str.upper() not in ('HIDE', 'IMPAIRMENT TYPE',
                                                   'TOTAL'):
                    try:
                        if n is not None:
                            lgd_by_cat[l_str] = float(n)
                    except (ValueError, TypeError):
                        pass
        for cat, pct in prov_pct.items():
            lgd = lgd_by_cat.get(cat, 0.0)
            acl_impaired[cat] = pct * lgd
        if acl_impaired:
            result['acl_impaired'] = acl_impaired
            total_allow = sum(acl_impaired.values())
            print(f"    Impaired allowance: {len(acl_impaired)} categories, "
                  f"Total allow: ${total_allow:,.2f}")

    # ── Parse 'Impaired Loans Pivot' for spec_id_by_pool ──
    spec_id_by_pool = {}
    if 'Impaired Loans Pivot' in wb.sheetnames:
        ws = wb['Impaired Loans Pivot']
        # Header row at row 4: col A "Row Labels", cols B..(N-1) pool names,
        # last col "Grand Total".
        header_row = None
        for ridx, row in enumerate(ws.iter_rows(min_row=1, max_row=10,
                                                values_only=True), start=1):
            if not row:
                continue
            a = row[0]
            if isinstance(a, str) and a.strip().lower() == 'row labels':
                header_row = ridx
                pool_cols = []
                for ci, val in enumerate(row[1:], start=2):
                    if not val:
                        continue
                    s = str(val).strip()
                    if s.lower() in ('grand total', 'column labels'):
                        continue
                    pool_cols.append((ci, s))
                break

        if header_row is not None:
            for row in ws.iter_rows(min_row=header_row + 1,
                                    max_row=ws.max_row, values_only=True):
                if not row:
                    continue
                grade = row[0]
                if not grade:
                    continue
                g_str = str(grade).strip()
                if not g_str or g_str.lower() in ('grand total',):
                    continue
                if g_str.lower().startswith('hide'):
                    continue
                for ci, pname in pool_cols:
                    val = row[ci - 1] if ci - 1 < len(row) else None
                    if val is None:
                        continue
                    try:
                        v = float(val)
                    except (ValueError, TypeError):
                        continue
                    if v <= 0:
                        continue
                    if pname.lower() == 'exclude':
                        continue
                    spec_id_by_pool.setdefault(pname, {})
                    spec_id_by_pool[pname][g_str] = (
                        spec_id_by_pool[pname].get(g_str, 0.0) + v
                    )
        if spec_id_by_pool:
            result['spec_id_by_pool'] = spec_id_by_pool
            n = len(spec_id_by_pool)
            tot = sum(sum(g.values()) for g in spec_id_by_pool.values())
            print(f"    Spec ID by pool: {n} pools, "
                  f"Total removed: ${tot:,.2f}")

    if 'acl_impaired' in result or 'spec_id_by_pool' in result:
        result['total_spec_id'] = sum(result.get('acl_impaired', {}).values())

    # ── DQ % / CO / RC per year from "Display CO-Recov -DQ" tab ──
    # When the source CECL-Migration-WARM file is absent (modern setups have
    # only a TCT_Model baseline in Reports/_warm_baselines/), this is the only
    # path that populates warm_dq_pct, warm_co, warm_rc, warm_net, etc.
    result.update(_parse_display_co_recov_dq(found))

    try:
        wb.close()
    except Exception:
        pass
    return result


def load_wizard_impaired(config):
    """Load impaired-loan data persisted by the setup wizard.

    Reads ``cfg["impaired_loans"]["data_rows"]`` (written by
    ``cecl_ui.services.config_service.build_yaml_from_wizard``) where
    each row carries the wizard-resolved ``loan_pool``, ``credit_grade``,
    ``balance_removed``, ``provision_amount``, and ``impairment_type``.

    Returns ``None`` when no wizard rows are present, otherwise a dict
    with the same keys as :func:`load_standalone_impaired`:
      'acl_impaired'     : {impairment_type: total_provision_amount}
      'spec_id_by_pool'  : {pool: {grade: balance_removed, ...}}
      'total_spec_id'    : sum of all provision_amounts
    """
    imp_cfg = config.get('impaired_loans') or {}
    rows = imp_cfg.get('data_rows') or []
    if not rows:
        return None

    acl_impaired: dict[str, float] = {}
    spec_id_by_pool: dict[str, dict[str, float]] = {}
    n_rows = 0
    for r in rows:
        if not isinstance(r, dict):
            continue
        try:
            bal = float(r.get('balance_removed') or 0.0)
        except (TypeError, ValueError):
            bal = 0.0
        if bal <= 0:
            continue
        pool = str(r.get('loan_pool') or '').strip()
        if not pool:
            continue
        grade = str(r.get('credit_grade') or '').strip()
        spec_id_by_pool.setdefault(pool, {})
        spec_id_by_pool[pool][grade] = (
            spec_id_by_pool[pool].get(grade, 0.0) + bal
        )
        imp_type = str(r.get('impairment_type') or '').strip() or 'Other'
        prov = r.get('provision_amount')
        try:
            prov_val = float(prov) if prov is not None else 0.0
        except (TypeError, ValueError):
            prov_val = 0.0
        acl_impaired[imp_type] = acl_impaired.get(imp_type, 0.0) + prov_val
        n_rows += 1

    if not spec_id_by_pool:
        return None

    result = {
        'acl_impaired': acl_impaired,
        'spec_id_by_pool': spec_id_by_pool,
        'total_spec_id': sum(acl_impaired.values()),
    }
    n_pools = len(spec_id_by_pool)
    tot_removed = sum(sum(g.values()) for g in spec_id_by_pool.values())
    print(f"    Wizard impaired loans: {n_rows} row(s), "
          f"{n_pools} pool(s), ${tot_removed:,.2f} balance removed, "
          f"${result['total_spec_id']:,.2f} provision")
    return result


def _load_dq_history_from_db(config):
    """Aggregate ``loan_code_delinquency_history`` into per-year/per-pool DQ%.

    Returns ``{year: {pool: dq_pct}}`` suitable for overlaying onto
    ``hist['impaired']['warm_dq_pct']``. Rows are bucketed to the most
    recent date in their calendar year (DQ is point-in-time, so the
    year-end / latest available quarter for that year is the right
    proxy for the WARM-style annual DQ%).

    Resolution rules per (year, pool):
      * Sum ``dq_amount`` and ``total_balance`` across all loan_codes
        mapping to that pool for the chosen as_of_date.
      * If ``total_balance`` totals to > 0, return
        ``dq_amount / total_balance``.
      * Otherwise, if ANY row carries a non-null ``dq_pct``, return the
        balance-weighted average (or simple average when all balances
        are null).

    Returns an empty dict on any error (missing table, no rows, etc.).
    """
    cu = (config.get('credit_union') or '').strip()
    if not cu:
        return {}
    pool_map = config.get('pool_map') or {}
    default_pool = config.get('default_pool') or ''
    try:
        from cecl_credentials import get_database_url
        from sqlalchemy import create_engine, text as _sql_text
    except Exception:  # noqa: BLE001
        return {}
    try:
        eng = create_engine(get_database_url())
        with eng.begin() as conn:
            rows = conn.execute(
                _sql_text(
                    "SELECT as_of_date, loan_code, dq_amount, "
                    "       total_balance, dq_pct "
                    "FROM loan_code_delinquency_history "
                    "WHERE cu = :cu"
                ),
                {"cu": cu},
            ).fetchall()
    except Exception as exc:  # noqa: BLE001
        print(f"    DQ history DB read skipped: {type(exc).__name__}: {exc}")
        return {}
    if not rows:
        return {}

    # year -> latest as_of_date in that year
    latest_in_year: dict[int, str] = {}
    for r in rows:
        d = r[0].isoformat() if hasattr(r[0], 'isoformat') else str(r[0])
        try:
            yr = int(d[:4])
        except (TypeError, ValueError):
            continue
        if yr not in latest_in_year or d > latest_in_year[yr]:
            latest_in_year[yr] = d

    # Collect contributions per (year, pool).
    by_yp: dict[tuple[int, str], dict[str, float]] = {}
    for r in rows:
        d = r[0].isoformat() if hasattr(r[0], 'isoformat') else str(r[0])
        try:
            yr = int(d[:4])
        except (TypeError, ValueError):
            continue
        if latest_in_year.get(yr) != d:
            continue  # only use the latest as_of_date per calendar year
        code = str(r[1]).strip()
        pool = pool_map.get(code) or pool_map.get(code.upper()) or default_pool
        pool = (pool or '').strip()
        if not pool or pool.lower() in ('ignore', 'exclude'):
            continue
        amt = float(r[2] or 0.0)
        tot = float(r[3]) if r[3] is not None else None
        pct = float(r[4]) if r[4] is not None else None
        agg = by_yp.setdefault((yr, pool), {
            'amount': 0.0, 'total': 0.0, 'pct_sum': 0.0,
            'pct_weight': 0.0, 'pct_count': 0,
        })
        agg['amount'] += amt
        if tot is not None:
            agg['total'] += tot
        if pct is not None:
            w = tot if tot is not None else 1.0
            agg['pct_sum'] += pct * w
            agg['pct_weight'] += w
            agg['pct_count'] += 1

    out: dict[int, dict[str, float]] = {}
    for (yr, pool), agg in by_yp.items():
        pct = None
        if agg['total'] > 0:
            pct = agg['amount'] / agg['total']
        elif agg['pct_count'] > 0:
            pct = (agg['pct_sum'] / agg['pct_weight']
                   if agg['pct_weight'] else None)
        if pct is None:
            continue
        out.setdefault(yr, {})[pool] = round(pct, 6)
    return out


def load_standalone_impaired(config, snap, df=None):
    """Load impaired-loan data from the standalone Impaired Loans file.

    Searches for files matching patterns like:
      "2026-03 Impaired Loans - Franklin Trust FCU.xlsx"
      "2026- 03 Impaired Loans - Franklin Trust FCU.xlsx"

    Returns dict with:
      'acl_impaired': {category: provision_amount, ...}
      'spec_id_by_pool': {pool: {grade: balance_removed, ...}, ...}
      'total_spec_id': float
    or empty dict if file/tab not found.
    """
    data_dir = config.get('data_directory', '')
    if not data_dir:
        return {}
    if not os.path.isabs(data_dir):
        data_dir = os.path.join(BASE, data_dir)

    cu = config['credit_union']
    snap_prefix = snap[:7] if snap else ''  # e.g. "2026-03"
    pool_map = config.get('pool_map', {})
    default_pool = config.get('default_pool', 'Other/Uncategorized')

    # Search for the standalone impaired loans file. Filenames vary:
    #   "2026-03 Impaired Loans - Franklin Trust FCU.xlsx"
    #   "2025-06_CECL-Migration-Impaired_Loans_-_Honolulu_FD_FCU.xlsx"
    # The separator between "Impaired" and "Loans" can be whitespace,
    # underscore, or hyphen. The CU name in the filename is often an
    # abbreviation (e.g. "Honolulu_FD_FCU" for "Honolulu Fire Department
    # FCU"), so we don't require it in the filename — we're already
    # restricted to the CU-specific data_directory.
    pattern = re.compile(
        rf'^{re.escape(snap_prefix[:4])}\s*[-_]?\s*{re.escape(snap_prefix[5:7])}'
        rf'.*Impaired[\s_-]+Loans.*\.xlsx$',
        re.IGNORECASE)

    search_dirs = [data_dir]
    fb_folder = config.get('credit_pull', {}).get('fallback_report_folder', '')
    if fb_folder and fb_folder != data_dir:
        if not os.path.isabs(fb_folder):
            fb_folder = os.path.join(BASE, fb_folder)
        search_dirs.append(fb_folder)

    found = None
    for sdir in search_dirs:
        if not os.path.isdir(sdir):
            continue
        for root, dirs, files in os.walk(sdir):
            for f in files:
                if f.startswith('~$') or f.upper().startswith('DNU'):
                    continue
                # Skip WARM files — those are handled by load_impaired_data.
                # Note: 'CECL-Migration-Impaired_Loans*.xlsx' (a *standalone*
                # impaired-loan workbook from the CECL Migration suite) must
                # NOT be excluded here, so check 'WARM' alone.
                if 'WARM' in f.upper():
                    continue
                if pattern.match(f):
                    found = os.path.join(root, f)
                    break
            if found:
                break
        if found:
            break

    if not found:
        return {}

    print(f"    Loading standalone impaired loans from: {os.path.basename(found)}")
    try:
        from openpyxl import load_workbook
        wb = load_workbook(found, data_only=True)
        # Tab name varies — may have leading/trailing whitespace
        # (e.g. ' Impaired Loans' in CECL-Migration workbooks).
        ws = None
        for sn in wb.sheetnames:
            if sn.strip().lower() == 'impaired loans':
                ws = wb[sn]
                break
        if ws is None:
            raise KeyError("No 'Impaired Loans' sheet (got: "
                           f"{wb.sheetnames!r})")
    except (KeyError, Exception) as exc:
        print(f"    Error reading impaired loans file: {exc}")
        return {}

    # ── Parse summary section ──
    # Layout varies: locate the header row by scanning the first ~22
    # rows for a cell whose value contains "Impairment Type" in the
    # right-hand columns (the left side may have a separate
    # "Impairment Type/Provision Percentage" definitions block).
    # Then find the columns labelled "Impairment Type" and "Sum of
    # Provision Amount" (or just "Provision Amount").
    acl_impaired = {}
    hdr_row = None
    type_col = prov_col = None
    for r in range(1, 23):
        for c in range(1, min(ws.max_column, 30) + 1):
            v = ws.cell(row=r, column=c).value
            if not v:
                continue
            s = str(v).strip().lower()
            if s == 'impaired type' or s == 'impairment type':
                # Prefer the right-hand "Impairment Type" header (the
                # one that has Provision Amount in the same row).
                # Verify by scanning that row for a Provision header.
                row_vals = [ws.cell(row=r, column=cc).value
                            for cc in range(1, ws.max_column + 1)]
                for cc, vv in enumerate(row_vals, start=1):
                    if vv and 'provision amount' in str(vv).lower():
                        hdr_row = r
                        type_col = c
                        prov_col = cc
                        break
                if hdr_row is not None:
                    break
        if hdr_row is not None:
            break

    if hdr_row is None:
        # Fall back to the original Franklin-Trust layout
        hdr_row = 4
        type_col = 12
        prov_col = 16

    for row in range(hdr_row + 1, hdr_row + 17):
        cat = ws.cell(row=row, column=type_col).value
        prov = ws.cell(row=row, column=prov_col).value
        if not cat or str(cat).strip().upper() in ('', 'HIDE', 'TOTAL',
                                                    'CALCULATION', 'IMPAIRMENT TYPE'):
            continue
        cat_str = str(cat).strip()
        try:
            prov_val = float(prov) if prov else 0.0
        except (ValueError, TypeError):
            prov_val = 0.0
        acl_impaired[cat_str] = prov_val

    # ── Build member→grade lookup from df ──
    grade_lookup = {}  # {member_suffix_str: grade}
    if df is not None and 'member_number' in df.columns and 'current_grade' in df.columns:
        for _, r in df.iterrows():
            mem = str(r['member_number']).strip()
            grade_lookup[mem] = r['current_grade']

    # ── Parse detail rows (row 24+) ──
    spec_id_by_pool = {}  # {pool: {grade: balance_removed}}
    total_removed = 0.0
    suffix_len = config.get('account_suffix_length', 3)
    for row in range(24, ws.max_row + 1):
        imp_type = ws.cell(row=row, column=1).value
        if not imp_type or str(imp_type).strip() in ('', 'HIDE'):
            continue
        balance = ws.cell(row=row, column=5).value
        if not balance:
            continue

        member = ws.cell(row=row, column=2).value
        suffix = ws.cell(row=row, column=3).value
        loan_type = ws.cell(row=row, column=4).value
        removed = ws.cell(row=row, column=17).value  # Balance Removed

        try:
            removed_val = float(removed) if removed else 0.0
        except (ValueError, TypeError):
            removed_val = 0.0
        if removed_val <= 0:
            continue

        # Map loan type to pool
        lt = str(loan_type).strip() if loan_type else ''
        pool = pool_map.get(lt, default_pool)

        # Look up grade from df using member+suffix
        grade = ''
        if member is not None and suffix is not None:
            try:
                mem_int = int(float(member))
                suf_int = int(float(suffix))
            except (TypeError, ValueError):
                # Spreadsheet placeholders like 'xxxx' or blanks — skip
                # grade lookup for this row rather than crashing the
                # whole report. The balance still aggregates into the
                # pool below.
                mem_int = suf_int = None
            if mem_int is not None:
                mem_key = f"{mem_int}{suf_int:0{suffix_len}d}"
                grade = grade_lookup.get(mem_key, '')
                if not grade:
                    # Try with string suffix as-is
                    mem_key2 = f"{mem_int}{str(suf_int).zfill(suffix_len)}"
                    grade = grade_lookup.get(mem_key2, '')

        spec_id_by_pool.setdefault(pool, {})
        spec_id_by_pool[pool][grade] = (
            spec_id_by_pool[pool].get(grade, 0) + removed_val
        )
        total_removed += removed_val

    wb.close()

    result = {
        'acl_impaired': acl_impaired,
        'spec_id_by_pool': spec_id_by_pool,
        'total_spec_id': sum(acl_impaired.values()),
    }

    imp_count = sum(1 for _ in acl_impaired.values() if _ > 0)
    n_pools = len(spec_id_by_pool)
    print(f"    Impaired categories: {len(acl_impaired)} "
          f"({imp_count} with provision), "
          f"Spec ID: {n_pools} pools, "
          f"Total removed: ${total_removed:,.2f}")

    return result


# ── Styling Helpers ────────────────────────────────────────────────
def hdr_row(ws, row, ncol, fill=None):
    for c in range(1, ncol + 1):
        cell = ws.cell(row=row, column=c)
        cell.font = HDR_FONT
        cell.fill = fill or HDR_FILL
        cell.alignment = Alignment(horizontal='center', wrap_text=True)
        cell.border = THIN

def style_rows(ws, r1, r2, ncol, mcols=(), pcols=()):
    for r in range(r1, r2 + 1):
        for c in range(1, ncol + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = NORM
            cell.border = THIN
            if (r - r1) % 2 == 1:
                cell.fill = ALT_FILL
            if c in mcols:
                cell.number_format = MONEY
            elif c in pcols:
                cell.number_format = PCT

def auto_w(ws, ncol, mn=10, mx=25):
    for c in range(1, ncol + 1):
        lt = get_column_letter(c)
        best = mn
        for row in ws.iter_rows(min_col=c, max_col=c):
            for cell in row:
                if cell.value:
                    best = max(best, min(len(str(cell.value)) + 2, mx))
        ws.column_dimensions[lt].width = best

def write_title(ws, row, text_val, col=1):
    ws.cell(row=row, column=col, value=text_val).font = TITLE_FONT

def write_sub(ws, row, text_val, col=1):
    ws.cell(row=row, column=col, value=text_val).font = SUB_FONT


# ── Calculation Helpers ────────────────────────────────────────────
def calc_net_credit_change(df):
    """Return improved%, deteriorated%, net_change% for a DataFrame."""
    total = df['current_balance'].sum()
    if total == 0:
        return 0, 0, 0
    imp = df[df['migration_status'] == 'Improved']['current_balance'].sum() / total
    det = df[df['migration_status'] == 'Deteriorated']['current_balance'].sum() / total
    return imp, det, imp - det

def calc_economic_stress(config):
    """Calculate Economic Stress Index from config economic data."""
    ed = config.get('economic_data', {})
    unemp = ed.get('unemployment_rate', 0) * 100
    pop = ed.get('population', 1)
    bk_pct = (ed.get('bankruptcies', 0) / pop) * 100 if pop else 0
    fc_pct = (ed.get('foreclosures', 0) / pop) * 100 if pop else 0
    return unemp + bk_pct + fc_pct

def calc_env_factor_pool(ncc_pct, dq_variance, econ_stress):
    """Calculate environmental factor for a pool."""
    ncc_score = score_from_ranges(ncc_pct * 100, NCC_RANGES)
    dq_score = score_from_ranges(dq_variance * 100, DQ_RANGES)
    es_score = score_from_ranges(econ_stress, ES_RANGES)
    return ncc_score + dq_score + es_score

def get_acl_base_rates(grades, config):
    """Get ACL base loss rates per grade. Uses config reserve_rates as fallback."""
    return {g['label']: g['reserve_rate'] for g in grades}

def get_dist_factor(grade_idx, num_grades):
    """Get distribution factor for a grade position."""
    if grade_idx < len(DIST_FACTORS):
        return DIST_FACTORS[grade_idx] / 100.0
    return DIST_FACTORS[-1] / 100.0


# ── Admin-default + per-pool management-adjustment resolver ──────────
def _load_admin_default_mgmt_adj():
    """Read firm-wide default from admin_defaults.yaml. 0.0 on error."""
    try:
        import yaml as _yaml
        from pathlib import Path as _Path
        p = _Path(__file__).resolve().parent / 'admin_defaults.yaml'
        if not p.exists():
            return 0.0
        data = _yaml.safe_load(p.read_text(encoding='utf-8')) or {}
        return float(data.get('default_mgmt_adj', 0.0) or 0.0)
    except Exception:
        return 0.0


def _build_pool_use_default_map(config):
    """Return ``{pool_name: bool}`` from ``config['pools']``."""
    out = {}
    for p in (config.get('pools') or []):
        name = (p.get('name') or '').strip()
        if name:
            out[name] = bool(p.get('use_default_mgmt_adj'))
    return out


def _resolve_mgmt_adj_grade(pool, grade_label, grade_idx, num_grades,
                             pool_use_default, mgmt_adj_by_pool,
                             admin_default, prior_mgmt_adj_map,
                             base_rate=None):
    """Mirror of report_tct's resolver for the legacy generate_report
    path. Precedence: prior > manual×dist > admin×dist (only when
    use_default AND no manual AND base_rate==0) > 0.
    """
    pm = prior_mgmt_adj_map.get(pool, {}) if prior_mgmt_adj_map else {}
    if grade_label in pm:
        return pm[grade_label]
    dist = get_dist_factor(grade_idx, num_grades)
    manual = mgmt_adj_by_pool.get(pool, 0) or 0
    if manual:
        return float(manual) * dist
    if (pool_use_default.get(pool, False)
            and admin_default
            and (base_rate is None or float(base_rate or 0) == 0)):
        return float(admin_default) * dist
    return 0.0


# ══════════════════════════════════════════════════════════════════
# SHEET BUILDERS
# ══════════════════════════════════════════════════════════════════

def sheet_cover_tct(wb, cu, snap):
    """Cover sheet - TCT/Franklin Trust style."""
    ws = wb.active
    ws.title = "Cover"
    ws['B4'] = "RISK BASED PRICING"
    ws['B4'].font = Font(name='Calibri', bold=True, size=22)
    ws['B6'] = "ACL/Credit Migration Report"
    ws['B6'].font = Font(name='Calibri', bold=True, size=16)
    items = ["CECL Compliant", "Risk Change by Type",
             "Improved/Deteriorated Loan Analysis", "Environmental Factor",
             "Allowance for Credit Loss (ACL)", "Summary of Deteriorated Loans"]
    for i, item in enumerate(items):
        ws.cell(row=8 + i, column=2, value=item).font = Font(name='Calibri', size=12)
    ws['B16'] = "Prepared For:"
    ws['B16'].font = Font(name='Calibri', size=12, italic=True)
    ws['B18'] = cu
    ws['B18'].font = Font(name='Calibri', bold=True, size=20)
    ws['B21'] = f"For Period Ending"
    ws['B22'] = snap
    ws['B22'].font = Font(name='Calibri', bold=True, size=14)
    ws['B25'] = "Presented by:"
    ws['B27'] = "TCT Risk Solutions"
    ws['B27'].font = Font(name='Calibri', bold=True, size=16, color='C0392B')
    ws['B28'] = "Take Charge Today"
    ws['B28'].font = Font(name='Calibri', italic=True, size=12, color='2E86C1')
    dt = datetime.now().strftime('%B %d, %Y')
    ws['B30'] = f"Report Generated: {dt}"
    ws['B30'].font = Font(name='Calibri', size=10, color='888888')


def sheet_cover_vizo(wb, cu, snap, supplemental=False):
    """Cover sheet - Vizo/Credit Union B style."""
    ws = wb.active
    ws.title = "Cover"
    ws['B6'] = "CECL Credit Migration Report"
    ws['B6'].font = Font(name='Calibri', bold=True, size=20)
    if supplemental:
        ws['B7'] = "Supplemental Reports"
        ws['B7'].font = Font(name='Calibri', bold=True, size=16)
    ws['B10'] = cu
    ws['B10'].font = Font(name='Calibri', bold=True, size=18)
    ws['B12'] = snap
    ws['B12'].font = Font(name='Calibri', bold=True, size=14)
    ws['B16'] = "TCT Risk Solutions"
    ws['B16'].font = Font(name='Calibri', bold=True, size=14, color='C0392B')
    ws['B20'] = "All reports are confidential."
    ws['B20'].font = Font(name='Calibri', size=10, italic=True, color='888888')


def sheet_report_overview(wb, report_type="main"):
    """Report Overview / Index page (Vizo style)."""
    ws = wb.create_sheet("Report Overview")
    ws['A2'] = "Report Overview"
    ws['A2'].font = Font(name='Calibri', bold=True, size=16)
    ws['A4'] = ("The CECL Credit Migration Reports from TCT, Inc. presents a comprehensive picture "
                "of the changing nature of risk in the credit union's loan portfolio.")
    ws['A4'].font = NORM
    ws['A4'].alignment = Alignment(wrap_text=True)
    ws.merge_cells('A4:H6')

    ws['A8'] = "Report Index:"
    ws['A8'].font = Font(name='Calibri', bold=True, size=14, color='1B4F72')

    if report_type == "main":
        sections = [
            ("Executive Summary", ["CECL Adjustment & Improved/Deteriorated",
                                    "Improved & Deteriorated Loans Risk Change By Credit Score"]),
            ("Detailed Reporting", ["Allowance & Provision for Credit Loss Reserve Analysis",
                                    "Risk Change by Credit Score - Total Loans",
                                    "Risk Change by Credit Score - Loan Pools",
                                    "Environmental Factor Provision for Loan Loss",
                                    "Loss Factor Calculation", "Delinquency Calculation"]),
        ]
    else:
        sections = [
            ("Supplemental Reporting Package", [
                "Historical Loan Balances by Credit Score",
                "Loss Factor Historical Detail",
                "Charge off and Recoveries Historical Detail",
                "Balance Adjustment Detail"]),
        ]

    r = 10
    for title, items in sections:
        ws.cell(row=r, column=1, value=title).font = Font(bold=True, size=12, color='1B4F72')
        for item in items:
            r += 1
            ws.cell(row=r, column=2, value=item).font = NORM
        r += 2


def sheet_exec_summary(wb, cu, snap, df, grades, config):
    """Executive Summary – works for both TCT and Vizo formats."""
    ws = wb.create_sheet("Executive Summary")
    ws['A1'] = cu
    ws['A1'].font = TITLE_FONT
    ws['A2'] = "Executive Summary"
    ws['A2'].font = Font(bold=True, size=14)
    ws['A3'] = f"For Period Ending {snap}"
    ws['A3'].font = Font(size=12, color='555555')

    no_score = config.get('no_score_label', 'Not Reported')
    total = df['current_balance'].sum()
    imp_bal = df[df['migration_status'] == 'Improved']['current_balance'].sum()
    det_bal = df[df['migration_status'] == 'Deteriorated']['current_balance'].sum()
    unc_bal = df[df['migration_status'] == 'Unchanged']['current_balance'].sum()
    ncc = (imp_bal - det_bal) / total * 100 if total else 0
    total_reserve = df['expected_loss_amount'].sum()

    # CECL Adjustment box
    r = 5
    ws.cell(row=r, column=1, value="CECL Adjustment").font = SUB_FONT
    for label, val in [("Total Specifically Identified Allowance", 0),
                       ("Total Allowance Needed", total_reserve),
                       ("Allowance for Credit Loss Balance", 0),
                       ("Adjustment (Overfunded)", total_reserve)]:
        r += 1
        ws.cell(row=r, column=1, value=label).font = NORM
        ws.cell(row=r, column=3, value=val).number_format = MONEY

    # Improved/Deteriorated summary by grade
    r += 2
    ws.cell(row=r, column=1, value="Improved Loans Summary").font = SUB_FONT
    r += 1
    ws.cell(row=r, column=1, value="Grade"); ws.cell(row=r, column=2, value="Balance")
    hdr_row(ws, r, 2)
    grade_labels = [g['label'] for g in grades]
    imp_df = df[df['migration_status'] == 'Improved']
    for gl in grade_labels:
        r += 1
        g_bal = imp_df[imp_df['current_grade'] == gl]['current_balance'].sum()
        ws.cell(row=r, column=1, value=gl)
        ws.cell(row=r, column=2, value=g_bal).number_format = MONEY
    r += 1
    ws.cell(row=r, column=1, value="Total Improved").font = Font(bold=True)
    ws.cell(row=r, column=2, value=imp_bal).number_format = MONEY

    r += 2
    ws.cell(row=r, column=1, value="Deteriorated Loans Summary").font = SUB_FONT
    r += 1
    ws.cell(row=r, column=1, value="Grade"); ws.cell(row=r, column=2, value="Balance")
    hdr_row(ws, r, 2)
    det_df = df[df['migration_status'] == 'Deteriorated']
    for gl in grade_labels:
        r += 1
        g_bal = det_df[det_df['current_grade'] == gl]['current_balance'].sum()
        ws.cell(row=r, column=1, value=gl)
        ws.cell(row=r, column=2, value=g_bal).number_format = MONEY
    r += 1
    ws.cell(row=r, column=1, value="Total Impaired").font = Font(bold=True)
    ws.cell(row=r, column=2, value=det_bal).number_format = MONEY

    # Net Credit Change per pool
    r += 2
    ws.cell(row=r, column=1, value="Improved/Deteriorated by Pool").font = SUB_FONT
    r += 1
    for hdr_i, h in enumerate(["Pool", "Improved %", "Deteriorated %", "Net Change %"]):
        ws.cell(row=r, column=1 + hdr_i, value=h)
    hdr_row(ws, r, 4)
    pools = sorted(df['loan_pool'].unique())
    for pool in pools:
        r += 1
        pdf = df[df['loan_pool'] == pool]
        imp_p, det_p, net_p = calc_net_credit_change(pdf)
        ws.cell(row=r, column=1, value=pool)
        ws.cell(row=r, column=2, value=imp_p).number_format = PCT
        ws.cell(row=r, column=3, value=det_p).number_format = PCT
        ws.cell(row=r, column=4, value=net_p).number_format = PCT
    r += 1
    ws.cell(row=r, column=1, value="Grand Total").font = Font(bold=True)
    imp_t, det_t, net_t = calc_net_credit_change(df)
    ws.cell(row=r, column=2, value=imp_t).number_format = PCT
    ws.cell(row=r, column=3, value=det_t).number_format = PCT
    ws.cell(row=r, column=4, value=net_t).number_format = PCT

    auto_w(ws, 4)


def sheet_risk_change_all(wb, cu, snap, df, grades, config, hist=None):
    """Risk Change By Credit Score - Grand Total (dollar + percent matrices)."""
    ws = wb.create_sheet("Risk Change-All Loans")
    no_score = config.get('no_score_label', 'Not Reported')
    gl = [g['label'] for g in grades] + [no_score]
    matrix = risk_change_matrix(df, grades, no_score)

    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Risk Change By Credit Score"
    ws['A2'].font = Font(bold=True, size=12)
    ws['A3'] = f"For Period Ending {snap}"

    # Score range map (supports optional display overrides per client)
    range_overrides = config.get('risk_change_range_labels', {})
    rng = {g['label']: range_overrides.get(g['label'], f"{g['min_score']}-{g['max_score']}") for g in grades}
    rng[no_score] = ""

    gray_fill = PatternFill('solid', fgColor='D9D9D9')
    gray_hdr_font = Font(name='Calibri', bold=True, size=10, color='000000')

    def apply_gray_header(row_num):
        for c in range(1, ncol + 1):
            cell = ws.cell(row=row_num, column=c)
            cell.font = gray_hdr_font
            cell.fill = gray_fill
            cell.alignment = Alignment(horizontal='center', wrap_text=True)
            cell.border = THIN

    def apply_plain_rows(r1, r2):
        for rr in range(r1, r2 + 1):
            for cc in range(1, ncol + 1):
                cell = ws.cell(row=rr, column=cc)
                if not cell.font or not cell.font.bold:
                    cell.font = NORM
                cell.border = THIN

    # ─── Dollar Matrix ───
    r = 5
    ws.cell(row=r, column=1, value="Dollar")
    ws.cell(row=r, column=4, value="Original Credit Grade")
    r += 1
    ws.cell(row=r, column=1, value="Current Credit Grade")
    ws.cell(row=r, column=2, value="")
    for j, g in enumerate(gl):
        ws.cell(row=r, column=3 + j, value=g)
    ws.cell(row=r, column=ncol, value="Grand Total")
    ncol = 3 + len(gl)
    apply_gray_header(r)

    start = r + 1
    for i, g in enumerate(gl):
        r += 1
        ws.cell(row=r, column=1, value=g)
        ws.cell(row=r, column=2, value=rng.get(g, ''))
        rtotal = 0
        for j, og in enumerate(gl):
            v = matrix.loc[g, og] if g in matrix.index and og in matrix.columns else 0
            ws.cell(row=r, column=3 + j, value=v).number_format = MONEY
            rtotal += v
            if j < i:
                ws.cell(row=r, column=3 + j).fill = DET_FILL
            elif j > i:
                ws.cell(row=r, column=3 + j).fill = IMP_FILL
        ws.cell(row=r, column=ncol, value=rtotal).number_format = MONEY
    # Total row
    r += 1
    ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
    for j, og in enumerate(gl):
        col_total = sum(matrix.loc[g2, og] for g2 in gl if g2 in matrix.index and og in matrix.columns)
        ws.cell(row=r, column=3 + j, value=col_total).number_format = MONEY
    ws.cell(row=r, column=ncol, value=df['current_balance'].sum()).number_format = MONEY
    apply_plain_rows(start, r)

    # Balance adjustments row
    r += 1
    ws.cell(row=r, column=1, value="Loans Not Risk Rated and Adjustments")
    _imp = hist.get('impaired', {}) if hist else {}
    _tba = _imp.get('total_balance_adjustment', 0.0)
    ws.cell(row=r, column=ncol, value=_tba).number_format = MONEY
    r += 1
    ws.cell(row=r, column=1, value="Total in Portfolio").font = Font(bold=True)
    _tip = _imp.get('total_in_portfolio', df['current_balance'].sum() + _tba)
    if ncol > 3:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=ncol - 1)
        ws.cell(row=r, column=2, value="Section 3 Allowance for Credit Loss Calculation")
        ws.cell(row=r, column=2).font = Font(bold=True)
        ws.cell(row=r, column=2).alignment = Alignment(horizontal='center')
    ws.cell(row=r, column=ncol, value=_tip).number_format = MONEY
    apply_plain_rows(r - 1, r)

    # ─── Percent Matrix ───
    r += 2
    ws.cell(row=r, column=1, value="Percent")
    ws.cell(row=r, column=4, value="Original Credit Grade")
    r += 1
    ws.cell(row=r, column=1, value="Current Credit Grade")
    for j, g in enumerate(gl):
        ws.cell(row=r, column=3 + j, value=g)
    ws.cell(row=r, column=ncol, value="Grand Total")
    hdr_row(ws, r, ncol)

    total = df['current_balance'].sum()
    start2 = r + 1
    for i, g in enumerate(gl):
        r += 1
        ws.cell(row=r, column=1, value=g)
        ws.cell(row=r, column=2, value=rng.get(g, ''))
        rtotal = 0
        for j, og in enumerate(gl):
            v = matrix.loc[g, og] if g in matrix.index and og in matrix.columns else 0
            col_total = sum(matrix.loc[g2, og] for g2 in gl if g2 in matrix.index and og in matrix.columns)
            pct = v / col_total if col_total else 0
            ws.cell(row=r, column=3 + j, value=pct).number_format = PCT
            rtotal += v
        ws.cell(row=r, column=ncol, value=rtotal / total if total else 0).number_format = PCT
    r += 1
    ws.cell(row=r, column=1, value="Grand Total").font = Font(bold=True)
    for j in range(len(gl)):
        ws.cell(row=r, column=3 + j, value=1.0).number_format = PCT
    ws.cell(row=r, column=ncol, value=1.0).number_format = PCT
    style_rows(ws, start2, r, ncol, pcols=set(range(3, ncol + 1)))

    # ─── Net Credit Change box ───
    r += 2
    imp_bal = df[df['migration_status'] == 'Improved']['current_balance'].sum()
    det_bal = df[df['migration_status'] == 'Deteriorated']['current_balance'].sum()
    unc_bal = df[df['migration_status'] == 'Unchanged']['current_balance'].sum()
    for lbl, vd, vp in [
        ("Total-Improved", imp_bal, imp_bal / total if total else 0),
        ("Total-Deteriorated", det_bal, det_bal / total if total else 0),
        ("Total Unchanged", unc_bal, unc_bal / total if total else 0),
        ("Total In Portfolio", total, 1.0),
    ]:
        r += 1
        ws.cell(row=r, column=1, value=label)
        ws.cell(row=r, column=2, value=vd).number_format = MONEY
        ws.cell(row=r, column=3, value=vp).number_format = PCT
        r += 1
    ws.cell(row=r, column=1, value="Net Change").font = Font(bold=True, size=12)
    ws.cell(row=r, column=2, value=imp_bal - det_bal).number_format = MONEY
    ws.cell(row=r, column=3, value=(imp_bal - det_bal) / total if total else 0).number_format = PCT

    auto_w(ws, ncol)


def sheet_impdet_summary(wb, cu, snap, df):
    """Analysis of Improved/Deteriorated Summary - all pools."""
    ws = wb.create_sheet("Improved-Deteriorated Summary")
    ws['A1'] = cu
    ws['A1'].font = TITLE_FONT
    ws['A2'] = "Analysis of Improved/Deteriorated Summary"
    ws['A3'] = f"For Period Ending {snap}"

    pools = sorted(df['loan_pool'].unique())
    r = 5
    for pool in pools:
        pdf = df[df['loan_pool'] == pool]
        ptotal = pdf['current_balance'].sum()
        ws.cell(row=r, column=1, value=pool).font = Font(bold=True, size=11)
        r += 1
        for h_i, h in enumerate(["", "$", "%"]):
            ws.cell(row=r, column=1 + h_i, value=h)
        hdr_row(ws, r, 3)
        r += 1
        for status in ['Improved', 'Deteriorated', 'Unchanged']:
            bal = pdf[pdf['migration_status'] == status]['current_balance'].sum()
            pct = bal / ptotal if ptotal else 0
            lbl = f"Total-{status}" if status != 'Unchanged' else "Total Unchanged"
            ws.cell(row=r, column=1, value=lbl)
            ws.cell(row=r, column=2, value=bal).number_format = MONEY
            ws.cell(row=r, column=3, value=pct).number_format = PCT
            if status == 'Improved':
                ws.cell(row=r, column=1).fill = IMP_FILL
            elif status == 'Deteriorated':
                ws.cell(row=r, column=1).fill = DET_FILL
            r += 1
        ws.cell(row=r, column=1, value="Total In Pool").font = Font(bold=True)
        ws.cell(row=r, column=2, value=ptotal).number_format = MONEY
        r += 1
        net = pdf[pdf['migration_status'] == 'Improved']['current_balance'].sum() - \
              pdf[pdf['migration_status'] == 'Deteriorated']['current_balance'].sum()
        ws.cell(row=r, column=1, value="Net Change").font = Font(bold=True, size=12)
        ws.cell(row=r, column=2, value=net).number_format = MONEY
        ws.cell(row=r, column=3, value=net / ptotal if ptotal else 0).number_format = PCT
        r += 2

    # Grand Total
    total = df['current_balance'].sum()
    ws.cell(row=r, column=1, value="Grand Total").font = Font(bold=True, size=12)
    r += 1
    for status in ['Improved', 'Deteriorated', 'Unchanged']:
        bal = df[df['migration_status'] == status]['current_balance'].sum()
        lbl = f"Total-{status}" if status != 'Unchanged' else "Total Unchanged"
        ws.cell(row=r, column=1, value=lbl)
        ws.cell(row=r, column=2, value=bal).number_format = MONEY
        ws.cell(row=r, column=3, value=bal / total if total else 0).number_format = PCT
        r += 1
    ws.cell(row=r, column=1, value="Total In Portfolio").font = Font(bold=True)
    ws.cell(row=r, column=2, value=total).number_format = MONEY
    r += 1
    net = df[df['migration_status'] == 'Improved']['current_balance'].sum() - \
          df[df['migration_status'] == 'Deteriorated']['current_balance'].sum()
    ws.cell(row=r, column=1, value="Net Change").font = Font(bold=True, size=14)
    ws.cell(row=r, column=2, value=net).number_format = MONEY
    ws.cell(row=r, column=3, value=net / total if total else 0).number_format = PCT
    auto_w(ws, 3)


def _pool_risk_sheet(wb, cu, snap, pool_df, pool_name, grades, config):
    """Risk Change per pool with matrix, net credit change, delinquency/chargeoff stubs."""
    no_score = config.get('no_score_label', 'Not Reported')
    gl = [g['label'] for g in grades] + [no_score]
    matrix = risk_change_matrix(pool_df, grades, no_score)
    safe = re.sub(r'[^\w\s-]', '', pool_name)[:25]
    ws = wb.create_sheet(safe)

    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=12)
    ws['A2'] = "Risk Change By Credit Score"
    ws['A3'] = f"For Period Ending {snap}"
    ws['A4'] = pool_name
    ws['A4'].font = Font(bold=True, size=14)

    range_overrides = config.get('risk_change_range_labels', {})
    rng = {g['label']: range_overrides.get(g['label'], f"{g['min_score']}-{g['max_score']}") for g in grades}
    rng[no_score] = ""
    ncol = 3 + len(gl)

    gray_fill = PatternFill('solid', fgColor='D9D9D9')
    gray_hdr_font = Font(name='Calibri', bold=True, size=10, color='000000')

    def apply_gray_header(row_num):
        for c in range(1, ncol + 1):
            cell = ws.cell(row=row_num, column=c)
            cell.font = gray_hdr_font
            cell.fill = gray_fill
            cell.alignment = Alignment(horizontal='center', wrap_text=True)
            cell.border = THIN

    def apply_plain_rows(r1, r2):
        for rr in range(r1, r2 + 1):
            for cc in range(1, ncol + 1):
                cell = ws.cell(row=rr, column=cc)
                if not cell.font or not cell.font.bold:
                    cell.font = NORM
                cell.border = THIN

    # Dollar matrix
    r = 6
    ws.cell(row=r, column=1, value="Dollar")
    ws.cell(row=r, column=4, value="Original Credit Grade")
    r += 1
    ws.cell(row=r, column=1, value="Current Credit Grade")
    ws.cell(row=r, column=2, value="")
    for j, g in enumerate(gl):
        ws.cell(row=r, column=3 + j, value=g)
    ws.cell(row=r, column=ncol, value="Grand Total")
    apply_gray_header(r)
    start = r + 1
    for i, g in enumerate(gl):
        r += 1
        ws.cell(row=r, column=1, value=g)
        ws.cell(row=r, column=2, value=rng.get(g, ''))
        rt = 0
        for j, og in enumerate(gl):
            v = matrix.loc[g, og] if g in matrix.index and og in matrix.columns else 0
            ws.cell(row=r, column=3 + j, value=v).number_format = MONEY
            rt += v
            if j < i: ws.cell(row=r, column=3+j).fill = DET_FILL
            elif j > i: ws.cell(row=r, column=3+j).fill = IMP_FILL
        ws.cell(row=r, column=ncol, value=rt).number_format = MONEY
    r += 1
    ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
    pool_total = pool_df['current_balance'].sum()
    for j, og in enumerate(gl):
        ct = sum(matrix.loc[g2, og] for g2 in gl if g2 in matrix.index and og in matrix.columns)
        ws.cell(row=r, column=3+j, value=ct).number_format = MONEY
    ws.cell(row=r, column=ncol, value=pool_total).number_format = MONEY
    apply_plain_rows(start, r)
    r += 1
    ws.cell(row=r, column=1, value="Loans Not Risk Rated and Adjustments")
    ws.cell(row=r, column=ncol, value=0).number_format = MONEY
    r += 1
    ws.cell(row=r, column=1, value="Total in Pool").font = Font(bold=True)
    if ncol > 3:
        ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=ncol - 1)
        ws.cell(row=r, column=2, value="Section 3 Allowance for Credit Loss Calculation")
        ws.cell(row=r, column=2).font = Font(bold=True)
        ws.cell(row=r, column=2).alignment = Alignment(horizontal='center')
    ws.cell(row=r, column=ncol, value=pool_total).number_format = MONEY
    apply_plain_rows(r - 1, r)

    # Percent matrix
    r += 2
    ws.cell(row=r, column=1, value="Percent")
    ws.cell(row=r, column=4, value="Original Credit Grade")
    r += 1
    ws.cell(row=r, column=1, value="Current Credit Grade")
    for j, g in enumerate(gl):
        ws.cell(row=r, column=3+j, value=g)
    ws.cell(row=r, column=ncol, value="Grand Total")
    apply_gray_header(r)
    start2 = r + 1
    for i, g in enumerate(gl):
        r += 1
        ws.cell(row=r, column=1, value=g)
        ws.cell(row=r, column=2, value=rng.get(g, ''))
        rt = 0
        for j, og in enumerate(gl):
            v = matrix.loc[g, og] if g in matrix.index and og in matrix.columns else 0
            ct = sum(matrix.loc[g2, og] for g2 in gl if g2 in matrix.index and og in matrix.columns)
            ws.cell(row=r, column=3+j, value=v/ct if ct else 0).number_format = PCT
            rt += v
        ws.cell(row=r, column=ncol, value=rt/pool_total if pool_total else 0).number_format = PCT
    r += 1
    ws.cell(row=r, column=1, value="Grand Total").font = Font(bold=True)
    for j in range(len(gl)):
        ws.cell(row=r, column=3+j, value=1.0).number_format = PCT
    ws.cell(row=r, column=ncol, value=1.0).number_format = PCT
    style_rows(ws, start2, r, ncol, pcols=set(range(3, ncol+1)))

    # Net Credit Change box
    r += 2
    ws.cell(row=r, column=1, value="Net Credit Change").font = SUB_FONT
    imp = pool_df[pool_df['migration_status']=='Improved']['current_balance'].sum()
    det = pool_df[pool_df['migration_status']=='Deteriorated']['current_balance'].sum()
    unc = pool_df[pool_df['migration_status']=='Unchanged']['current_balance'].sum()
    for lbl, vd, vp in [("Total-Improved", imp, imp/pool_total if pool_total else 0),
                        ("Total-Deteriorated", det, det/pool_total if pool_total else 0),
                        ("Total Unchanged", unc, unc/pool_total if pool_total else 0),
                        ("Total In Portfolio", pool_total, 1.0)]:
        r += 1
        ws.cell(row=r, column=1, value=lbl)
        ws.cell(row=r, column=2, value=vd).number_format = MONEY
        ws.cell(row=r, column=3, value=vp).number_format = PCT
        r += 1
    ws.cell(row=r, column=1, value="Net Change").font = Font(bold=True, size=12)
    ws.cell(row=r, column=2, value=imp - det).number_format = MONEY
    ws.cell(row=r, column=3, value=(imp-det)/pool_total if pool_total else 0).number_format = PCT

    auto_w(ws, ncol)


def sheet_pool_risk_changes(wb, cu, snap, df, grades, config):
    """Create one Risk Change sheet per pool."""
    pools = sorted(df['loan_pool'].unique())
    for pool in pools:
        _pool_risk_sheet(wb, cu, snap, df, pool, grades, config)


def sheet_acl_reserve(wb, cu, snap, df, grades, config, hist=None):
    """Allowance & Provision for Credit Loss Reserve Analysis."""
    ws = wb.create_sheet("ACL Reserve Analysis")
    no_score = config.get('no_score_label', 'Not Reported')
    gl = [g['label'] for g in grades] + [no_score]
    econ_stress = calc_economic_stress(config)

    # Compute life loss rates from historical data
    co_data = hist.get('chargeoffs', {}) if hist else {}
    rc_data = hist.get('recoveries', {}) if hist else {}
    avg_bals = hist.get('avg_balances', {}) if hist else {}
    hist_years = hist.get('years', []) if hist else []
    dq_pct = hist.get('dq_pct', {}) if hist else {}

    pool_life_loss = {}
    for pool in sorted(df['loan_pool'].unique()):
        rates = []
        for y in hist_years:
            net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
            avg = avg_bals.get(y, {}).get(pool, 0)
            if avg > 0:
                rates.append(net / avg)
        pool_life_loss[pool] = sum(rates) / len(rates) if rates else 0

    # Compute delinquency variance per pool
    pool_dq_var = {}
    for pool in sorted(df['loan_pool'].unique()):
        dq_rates = [dq_pct.get(y, {}).get(pool, 0) for y in sorted(dq_pct.keys())]
        if len(dq_rates) >= 2:
            avg_dq = sum(dq_rates) / len(dq_rates)
            pool_dq_var[pool] = dq_rates[-1] - avg_dq
        else:
            pool_dq_var[pool] = 0

    mgmt_adj_by_pool = config.get('mgmt_adj_by_pool', {})
    pool_use_default = _build_pool_use_default_map(config)
    admin_default_mgmt_adj = _load_admin_default_mgmt_adj()
    _imp = hist.get('impaired', {}) if hist else {}
    prior_mgmt_adj = _imp.get('prior_mgmt_adj', {})
    prior_env_factor = _imp.get('prior_env_factor', {})
    spec_id_by_pool = _imp.get('spec_id_by_pool', {})
    acl_impaired = _imp.get('acl_impaired', {})
    acl_summary = _imp.get('acl_summary', {})
    spec_id_by_pool = _imp.get('spec_id_by_pool', {})
    acl_impaired = _imp.get('acl_impaired', {})
    acl_summary = _imp.get('acl_summary', {})

    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Allowance & Provision for Credit Loss Reserve Analysis"
    ws['A3'] = f"For Period Ending {snap}"

    headers = ["Current Grade", "Balance", "Specific\nIdentification",
               "Loan Loss Calc.\nBalance", "ACL Base\nLoss Rate",
               "Management\nAdjustment", "Allowance\nFactor",
               "Allowance before\nEnvironmental", "Environmental\nFactor",
               "Environmental\n Allowance", "Total Allowance"]
    pools = sorted(df['loan_pool'].unique())
    r = 5
    grand_allowance = 0

    for pool in pools:
        pdf = df[df['loan_pool'] == pool]
        pool_total = pdf['current_balance'].sum()
        # Compute env factor for this pool
        imp_p, det_p, ncc = calc_net_credit_change(pdf)
        dq_var = pool_dq_var.get(pool, 0)
        env_factor_calc = calc_env_factor_pool(ncc, dq_var, econ_stress) / 100.0  # as decimal
        env_factor = prior_env_factor.get(pool, env_factor_calc)

        ws.cell(row=r, column=1, value=pool).font = Font(bold=True, size=12)
        r += 1
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers))
        start = r + 1
        pool_allowance_before = 0

        for gi, g in enumerate(gl):
            r += 1
            g_df = pdf[pdf['current_grade'] == g]
            balance = g_df['current_balance'].sum()
            specific_id = spec_id_by_pool.get(pool, {}).get(g, 0)
            calc_bal = balance - specific_id
            # ACL base rate: life loss rate × distribution factor
            life_loss = pool_life_loss.get(pool, 0)
            dist = get_dist_factor(gi, len(gl))
            base_rate = max(0, life_loss * dist)
            if base_rate == 0 and life_loss == 0:
                # Fallback to config reserve rate only when no data exists
                base_rate = next((gr['reserve_rate'] for gr in grades if gr['label'] == g), 0.005)
                if g == no_score:
                    base_rate = np.median([gr['reserve_rate'] for gr in grades])
            mgmt_adj = _resolve_mgmt_adj_grade(
                pool, g, gi, len(gl),
                pool_use_default, mgmt_adj_by_pool,
                admin_default_mgmt_adj, prior_mgmt_adj,
                base_rate=base_rate,
            )
            factor = base_rate + mgmt_adj
            allowance_before = calc_bal * factor
            pool_allowance_before += allowance_before

            ws.cell(row=r, column=1, value=g)
            ws.cell(row=r, column=2, value=balance).number_format = MONEY
            ws.cell(row=r, column=3, value=specific_id).number_format = MONEY
            ws.cell(row=r, column=4, value=calc_bal).number_format = MONEY
            ws.cell(row=r, column=5, value=base_rate).number_format = PCT4
            ws.cell(row=r, column=6, value=mgmt_adj).number_format = PCT4
            ws.cell(row=r, column=7, value=factor).number_format = PCT4
            ws.cell(row=r, column=8, value=allowance_before).number_format = MONEY
            ws.cell(row=r, column=9, value="")
            ws.cell(row=r, column=10, value="")
            ws.cell(row=r, column=11, value="")

        # Pool total row
        r += 1
        env_allowance = pool_allowance_before * env_factor
        total_allowance = pool_allowance_before + env_allowance
        grand_allowance += total_allowance
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        ws.cell(row=r, column=2, value=pool_total).number_format = MONEY
        ws.cell(row=r, column=8, value=pool_allowance_before).number_format = MONEY
        ws.cell(row=r, column=9, value=env_factor).number_format = PCT
        ws.cell(row=r, column=10, value=env_allowance).number_format = MONEY
        ws.cell(row=r, column=11, value=total_allowance).number_format = MONEY
        style_rows(ws, start, r, len(headers), mcols={2,3,4,8,10,11}, pcols={5,6,7,9})
    pooled_balance = df['current_balance'].sum()
    pooled_spec_id = sum(sum(g.values()) for g in spec_id_by_pool.values()) if spec_id_by_pool else 0
    pooled_calc_bal = pooled_balance - pooled_spec_id
    ws.cell(row=r, column=1, value="Pooled Totals").font = Font(bold=True, size=12)
    ws.cell(row=r, column=2, value=pooled_balance).number_format = MONEY
    ws.cell(row=r, column=3, value=pooled_spec_id).number_format = MONEY
    ws.cell(row=r, column=4, value=pooled_calc_bal).number_format = MONEY
    ws.cell(row=r, column=11, value=grand_allowance).number_format = MONEY

    # Impaired loans section
    r += 2
    ws.cell(row=r, column=1, value="Impaired Loans").font = Font(bold=True)
    ws.cell(row=r, column=10, value="Allowance").font = Font(bold=True)
    total_spec_allow = 0
    for lbl in ["Delinquent Loans", "Known Losses", "Repossessions",
                "Foreclosed Real Estate", "Deceased", "Bankruptcy"]:
        imp_val = acl_impaired.get(lbl, 0)
        r += 1
        ws.cell(row=r, column=1, value=lbl)
        ws.cell(row=r, column=11, value=imp_val).number_format = MONEY
        total_spec_allow += imp_val
    r += 1
    ws.cell(row=r, column=1, value="Total Specifically Identified Allowance").font = Font(bold=True)
    ws.cell(row=r, column=11, value=total_spec_allow).number_format = MONEY
    total_allow_needed = grand_allowance + total_spec_allow
    r += 1
    ws.cell(row=r, column=1, value="Total Allowance Needed").font = Font(bold=True)
    ws.cell(row=r, column=11, value=total_allow_needed).number_format = MONEY
    acl_bal = acl_summary.get('acl_balance', config.get('acl_balance', 0))
    r += 1
    ws.cell(row=r, column=1, value=f"Allowance for Credit Loss Balance as of {snap}")
    ws.cell(row=r, column=11, value=acl_bal).number_format = MONEY
    adjustment = total_allow_needed - acl_bal
    r += 1
    adj_label = "Adjustment (Underfunded)" if adjustment >= 0 else "Adjustment (Overfunded)"
    ws.cell(row=r, column=1, value=adj_label).font = Font(bold=True)
    ws.cell(row=r, column=11, value=adjustment).number_format = MONEY

    auto_w(ws, len(headers))


def sheet_env_factor(wb, cu, snap, df, grades, config, hist=None):
    """Environmental Factor for PLL."""
    ws = wb.create_sheet("Environmental Factor")
    ed = config.get('economic_data', {})
    econ_stress = calc_economic_stress(config)
    no_score = config.get('no_score_label', 'Not Reported')
    dq_pct = hist.get('dq_pct', {}) if hist else {}

    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Environmental Factor for PLL"
    ws['A3'] = f"For Period Ending {snap}"

    # Economic Stress Index
    r = 5
    ws.cell(row=r, column=1, value="Economic Stress Index Calculation").font = SUB_FONT
    r += 1
    ws.cell(row=r, column=1, value="State")
    ws.cell(row=r, column=2, value="Unemployment Rate")
    ws.cell(row=r, column=3, value="Foreclosures")
    ws.cell(row=r, column=4, value="Bankruptcies")
    ws.cell(row=r, column=5, value="Population")
    hdr_row(ws, r, 5)
    r += 1
    ws.cell(row=r, column=1, value=ed.get('state', ''))
    ws.cell(row=r, column=2, value=ed.get('unemployment_rate', 0)).number_format = PCT
    ws.cell(row=r, column=3, value=ed.get('foreclosures', 0))
    ws.cell(row=r, column=4, value=ed.get('bankruptcies', 0))
    ws.cell(row=r, column=5, value=ed.get('population', 0)).number_format = '#,##0'
    r += 1
    ws.cell(row=r, column=1, value="County")
    ws.cell(row=r, column=2, value="Unemployment Rate")
    ws.cell(row=r, column=3, value="Bankruptcy %")
    ws.cell(row=r, column=4, value="Foreclosure %")
    ws.cell(row=r, column=5, value="Economic Stress Index")
    hdr_row(ws, r, 5)
    r += 1
    pop = ed.get('population', 1)
    ws.cell(row=r, column=1, value=ed.get('county', ''))
    ws.cell(row=r, column=2, value=ed.get('unemployment_rate', 0)).number_format = PCT
    ws.cell(row=r, column=3, value=ed.get('bankruptcies', 0) / pop if pop else 0).number_format = PCT
    ws.cell(row=r, column=4, value=ed.get('foreclosures', 0) / pop if pop else 0).number_format = PCT
    ws.cell(row=r, column=5, value=econ_stress / 100).number_format = PCT

    # Per-pool environmental factors with real delinquency variance
    r += 2
    headers = ["Portfolio Segment", "Net Credit\nChange", "Net Credit\nScore",
               "Delinquency\nVariance", "Delinquency\nScore",
               "Economic Stress\nActual", "Economic Stress\nScore",
               "Environmental\nFactor"]
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))

    pools = sorted(df['loan_pool'].unique())
    start = r + 1
    for pool in pools:
        r += 1
        pdf = df[df['loan_pool'] == pool]
        _, _, ncc = calc_net_credit_change(pdf)
        ncc_score = score_from_ranges(ncc * 100, NCC_RANGES)

        # Compute delinquency variance from historical data
        dq_rates = [dq_pct.get(y, {}).get(pool, 0) for y in sorted(dq_pct.keys())]
        if len(dq_rates) >= 2:
            avg_dq = sum(dq_rates) / len(dq_rates)
            dq_var = dq_rates[-1] - avg_dq
        else:
            dq_var = 0

        dq_score = score_from_ranges(dq_var * 100, DQ_RANGES)
        es_score = score_from_ranges(econ_stress, ES_RANGES)
        env_f = ncc_score + dq_score + es_score

        ws.cell(row=r, column=1, value=pool)
        ws.cell(row=r, column=2, value=ncc).number_format = PCT
        ws.cell(row=r, column=3, value=ncc_score / 100).number_format = PCT
        ws.cell(row=r, column=4, value=dq_var).number_format = PCT
        ws.cell(row=r, column=5, value=dq_score / 100).number_format = PCT
        ws.cell(row=r, column=6, value=econ_stress / 100).number_format = PCT
        ws.cell(row=r, column=7, value=es_score / 100).number_format = PCT
        ws.cell(row=r, column=8, value=env_f / 100).number_format = PCT
    style_rows(ws, start, r, len(headers), pcols=set(range(2, 9)))
    auto_w(ws, len(headers))


def sheet_loss_factor(wb, cu, snap, df, grades, config, hist=None):
    """Loss Factor Calculation summary."""
    ws = wb.create_sheet("Loss Factor Calculation")
    no_score = config.get('no_score_label', 'Not Reported')
    gl = [g['label'] for g in grades] + [no_score]

    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Loss Factor Calculation"
    ws['A3'] = f"For Period Ending {snap}"

    headers = ["Current Grade", "Average Balance", "Life Loss Rate",
               "Distribution Factor", "ACL Base Loss Rate", "% of Loans"]
    pools = sorted(df['loan_pool'].unique())

    # Compute life loss rate per pool from historical data
    co_data = hist.get('chargeoffs', {}) if hist else {}
    rc_data = hist.get('recoveries', {}) if hist else {}
    avg_bals = hist.get('avg_balances', {}) if hist else {}
    years = hist.get('years', []) if hist else []
    pool_life_loss = {}
    for pool in pools:
        rates = []
        for y in years:
            net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
            avg = avg_bals.get(y, {}).get(pool, 0)
            rate = net / avg if avg > 0 else 0
            ws.cell(row=r, column=2 + yi, value=rate).number_format = PCT4
            rates.append(rate)
        avg_rate = sum(rates) / len(rates) if rates else 0
        ws.cell(row=r, column=ncol, value=avg_rate).number_format = PCT4
    style_rows(ws, start, r, ncol, pcols=set(range(2, ncol + 1)))

    # Average balances section
    r += 3
    ws.cell(row=r, column=1, value="Average Balances by Pool").font = SUB_FONT
    r += 1
    headers2 = ["Pool"] + year_strs
    ncol2 = len(headers2)
    for hi, h in enumerate(headers2):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, ncol2)
    start2 = r + 1
    for pool in pools:
        r += 1
        ws.cell(row=r, column=1, value=pool)
        for yi, y in enumerate(years):
            avg = avg_bals.get(y, {}).get(pool, 0)
            ws.cell(row=r, column=2 + yi, value=avg).number_format = MONEY
        style_rows(ws, start2, r, ncol2, mcols=set(range(2, ncol2 + 1)))
        auto_w(ws, max(ncol, ncol2))
    else:
        ws['A5'] = "No historical data available."
        ws['A5'].font = Font(italic=True, color='888888')


def sheet_chargeoff_recovery(wb, cu, snap, config, hist=None):
    """Charge off and Recoveries summary."""
    ws = wb.create_sheet("Charge offs & Recoveries")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Charge off and Recoveries"
    ws['A3'] = f"For Period Ending {snap}"

    pools = config.get('pool_map', {})
    pool_names = sorted(set(pools.values()))

    co_data = hist.get('chargeoffs', {}) if hist else {}
    rc_data = hist.get('recoveries', {}) if hist else {}
    years = hist.get('years', []) if hist else []
    if not years:
        years = list(range(2019, int(snap[:4]) + 1))
    year_strs = [str(y) for y in years]

    # ─── Charge offs ───
    r = 5
    headers = ["Charge offs"] + year_strs + ["ACL Charge offs"]
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))
    start = r + 1
    grand_co = {y: 0 for y in years}
    for pool in pool_names:
        r += 1
        ws.cell(row=r, column=1, value=pool)
        row_total = 0
        for yi, y in enumerate(years):
            val = co_data.get(y, {}).get(pool, 0)
            ws.cell(row=r, column=2 + yi, value=val).number_format = MONEY
            row_total += val
        ws.cell(row=r, column=ncol, value=row_total).number_format = MONEY
    r += 1
    ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
    acl_co_total = 0
    for yi, y in enumerate(years):
        ws.cell(row=r, column=2 + yi, value=grand_co[y]).number_format = MONEY
        acl_co_total += grand_co[y]
    ws.cell(row=r, column=ncol, value=acl_co_total).number_format = MONEY
    style_rows(ws, start, r, len(headers), mcols=set(range(2, len(headers) + 1)))

    # ─── Recoveries ───
    r += 2
    headers2 = ["Recoveries"] + year_strs + ["ACL Recoveries"]
    for hi, h in enumerate(headers2):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers2))
    start2 = r + 1
    grand_rc = {y: 0 for y in years}
    for pool in pool_names:
        r += 1
        ws.cell(row=r, column=1, value=pool)
        row_total = 0
        for yi, y in enumerate(years):
            val = rc_data.get(y, {}).get(pool, 0)
            ws.cell(row=r, column=2 + yi, value=val).number_format = MONEY
            row_total += val
        ws.cell(row=r, column=ncol, value=row_total).number_format = MONEY
    r += 1
    ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
    acl_rc_total = 0
    for yi, y in enumerate(years):
        ws.cell(row=r, column=2 + yi, value=grand_rc[y]).number_format = MONEY
        acl_rc_total += grand_rc[y]
    ws.cell(row=r, column=ncol, value=acl_rc_total).number_format = MONEY
    style_rows(ws, start2, r, len(headers2), mcols=set(range(2, len(headers2) + 1)))

    # ─── Net Charge offs ───
    r += 2
    headers3 = ["Net Charge offs"] + year_strs + ["Net Charge offs"]
    for hi, h in enumerate(headers3):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers3))
    start3 = r + 1
    grand_net = {y: 0 for y in years}
    for pool in pool_names:
        r += 1
        ws.cell(row=r, column=1, value=pool)
        row_total = 0
        for yi, y in enumerate(years):
            net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
            ws.cell(row=r, column=2 + yi, value=net).number_format = MONEY
            grand_net[y] += net
            row_total += net
        ws.cell(row=r, column=ncol, value=row_total).number_format = MONEY
    r += 1
    ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
    net_total = 0
    for yi, y in enumerate(years):
        ws.cell(row=r, column=2 + yi, value=grand_net[y]).number_format = MONEY
        net_total += grand_net[y]
    ws.cell(row=r, column=ncol, value=net_total).number_format = MONEY
    style_rows(ws, start3, r, len(headers3), mcols=set(range(2, len(headers3) + 1)))

    # ─── Life of Loan Loss Rate ───
    avg_bals = hist.get('avg_balances', {}) if hist else {}
    r += 2
    headers4 = ["Life Loss Rate"] + year_strs + ["Average"]
    for hi, h in enumerate(headers4):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers4))
    start4 = r + 1
    for pool in pool_names:
        r += 1
        ws.cell(row=r, column=1, value=pool)
        rates = []
        for yi, y in enumerate(years):
            net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
            avg = avg_bals.get(y, {}).get(pool, 0)
            rate = net / avg if avg > 0 else 0
            ws.cell(row=r, column=2 + yi, value=rate).number_format = PCT4
            rates.append(rate)
        avg_rate = sum(rates) / len(rates) if rates else 0
        ws.cell(row=r, column=len(headers4), value=avg_rate).number_format = PCT4
    style_rows(ws, start4, r, len(headers4), pcols=set(range(2, len(headers4) + 1)))

    auto_w(ws, len(headers))


def sheet_delinquency(wb, cu, snap, config, hist=None):
    """Delinquency Calculation with historical data."""
    ws = wb.create_sheet("Delinquency Calculation")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Delinquency Calculation"
    ws['A3'] = f"For Period Ending {snap}"

    pools = sorted(set(config.get('pool_map', {}).values()))
    dq_pct = hist.get('dq_pct', {}) if hist else {}
    years = sorted(dq_pct.keys()) if dq_pct else list(range(2019, int(snap[:4]) + 1))
    year_strs = [str(y) for y in years]

    r = 5
    headers = ["DQ %"] + year_strs + ["Average", "Variance from Avg"]
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))
    start = r + 1
    for pool in pools:
        r += 1
        ws.cell(row=r, column=1, value=pool)
        rates = []
        for yi, y in enumerate(years):
            val = dq_pct.get(y, {}).get(pool, 0)
            ws.cell(row=r, column=2 + yi, value=val).number_format = PCT
            rates.append(val)
        avg = sum(rates) / len(rates) if rates else 0
        ws.cell(row=r, column=len(headers) - 1, value=avg).number_format = PCT
        # Variance = most recent - average
        current = rates[-1] if rates else 0
        ws.cell(row=r, column=len(headers), value=current - avg).number_format = PCT
    style_rows(ws, start, r, len(headers), pcols=set(range(2, len(headers) + 1)))
    auto_w(ws, len(headers))


def sheet_balance_adj(wb, cu, snap, df, grades, config):
    """FAS 114 / Balance Adjustment sheet."""
    no_score = config.get('no_score_label', 'Not Reported')
    gl = [g['label'] for g in grades] + [no_score]

    ws = wb.create_sheet("Balance Adjustment")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Balance Adjustment"
    ws['A3'] = f"For Period Ending {snap}"

    pools = sorted(df['loan_pool'].unique())
    r = 5
    headers = ["Current Grade", "Loan Report Balance", "Bal Adjustment", "Balance Sheet Total"]

    for pool in pools:
        pdf = df[df['loan_pool'] == pool]
        ws.cell(row=r, column=1, value=pool).font = Font(bold=True, size=12)
        r += 1
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers))
        start = r + 1
        pool_total = 0
        for g in gl:
            r += 1
            bal = pdf[pdf['current_grade'] == g]['current_balance'].sum()
            pool_total += bal
            ws.cell(row=r, column=1, value=g)
            ws.cell(row=r, column=2, value=bal).number_format = MONEY
            ws.cell(row=r, column=3, value=0).number_format = MONEY
            ws.cell(row=r, column=4, value=bal).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        ws.cell(row=r, column=2, value=pool_total).number_format = MONEY
        ws.cell(row=r, column=3, value=0).number_format = MONEY
        ws.cell(row=r, column=4, value=pool_total).number_format = MONEY
        style_rows(ws, start, r, len(headers), mcols={2,3,4})
        r += 2

    # Grand total
    ws.cell(row=r, column=1, value="Grand Total").font = Font(bold=True, size=12)
    total = df['current_balance'].sum()
    ws.cell(row=r, column=2, value=total).number_format = MONEY
    ws.cell(row=r, column=3, value=0).number_format = MONEY
    ws.cell(row=r, column=4, value=total).number_format = MONEY
    auto_w(ws, len(headers))


def sheet_env_ranges(wb):
    """Environmental Factor Ranges reference table."""
    ws = wb.create_sheet("Env Factor Ranges")
    ws['A1'] = "Environmental Factor Ranges"
    ws['A1'].font = Font(bold=True, size=14)

    # Net Credit Change
    r = 3
    ws.cell(row=r, column=1, value="Net Credit Change").font = SUB_FONT
    ws.cell(row=r, column=1+0, value="Range"); ws.cell(row=r, column=2, value="Score")
    ws.cell(row=r, column=4, value="Delinquency").font = SUB_FONT
    ws.cell(row=r, column=4, value="Range"); ws.cell(row=r, column=5, value="Score")
    ws.cell(row=r, column=7, value="Economic Stress Score").font = SUB_FONT
    ws.cell(row=r, column=7, value="Range"); ws.cell(row=r, column=8, value="Score")
    hdr_row(ws, r, 8)

    ncc_rows = [
        ("<-18.00%", "7.00%"), ("-17.99% to -16.00%", "6.00%"),
        ("-15.99% to -14.00%", "5.00%"), ("-13.99% to -11.00%", "4.00%"),
        ("-10.99% to -8.00%", "3.00%"), ("-7.99% to -6.00%", "2.00%"),
        ("-5.99% to -4.00%", "1.00%"), ("-3.99% to 3.99%", "0.00%"),
        ("4.00% to 5.99%", "-1.00%"), ("6.00% to 7.99%", "-2.00%"),
        ("8.00% to 8.99%", "-3.00%"), ("9.00% to 10.99%", "-4.00%"),
        ("11.00% to 12.99%", "-5.00%"), ("13.00% to 14.99%", "-6.00%"),
        (">15.00%", "-7.00%"),
    ]
    dq_rows = [
        (">5.00%", "20.00%"), ("4.00% to 4.99%", "17.00%"),
        ("3.00% to 3.99%", "12.00%"), ("2.50% to 2.99%", "8.00%"),
        ("2.00% to 2.49%", "4.00%"), ("1.50% to 1.99%", "2.50%"),
        ("1.00% to 1.49%", "1.50%"), (".50% to .99%", "0.75%"),
        ("-.49% to .49%", "0.00%"), ("-.99% to -.50%", "-0.75%"),
        ("-1.49% to -1.00%", "-1.50%"), ("-1.99% to -1.50%", "-2.50%"),
        ("-2.49% to -2.00%", "-4.00%"), ("-2.99% to -2.50%", "-8.00%"),
        ("-3.99% to -3.00%", "-12.00%"), ("-4.99% to -4.00%", "-17.00%"),
        ("<-5.00%", "-20.00%"),
    ]
    es_rows = [
        (">25.00%", "10.00%"), ("24.00% to 24.99%", "8.00%"),
        ("22.00% to 23.99%", "7.00%"), ("20.00% to 21.99%", "6.00%"),
        ("18.00% to 19.99%", "5.00%"), ("16.00% to 17.99%", "4.00%"),
        ("14.00% to 15.99%", "3.50%"), ("12.00% to 13.99%", "3.00%"),
        ("10.00% to 11.99%", "2.00%"), ("8.00% to 9.99%", "1.00%"),
        ("6.00% to 7.99%", "0.00%"), ("4.00% to 5.99%", "0.00%"),
        ("2.00% to 3.99%", "-1.00%"), (".00% to 1.99%", "-2.00%"),
    ]

    for i, (rng, sc) in enumerate(ncc_rows):
        ws.cell(row=r + 1 + i, column=1, value=rng)
        ws.cell(row=r + 1 + i, column=2, value=sc)
    for i, (rng, sc) in enumerate(dq_rows):
        ws.cell(row=r + 1 + i, column=4, value=rng)
        ws.cell(row=r + 1 + i, column=5, value=sc)
    for i, (rng, sc) in enumerate(es_rows):
        ws.cell(row=r + 1 + i, column=7, value=rng)
        ws.cell(row=r + 1 + i, column=8, value=sc)
    auto_w(ws, 8)


def sheet_grade_config(wb, grades, config):
    """Grade ranges & loan code reference."""
    ws = wb.create_sheet("Grade Ranges & Loan Codes")
    ws['A1'] = "Credit Grade Configuration"
    ws['A1'].font = Font(bold=True, size=14)

    headers = ["Grade", "Score Range", "Reserve Rate"]
    r = 3
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))
    for g in grades:
        r += 1
        ws.cell(row=r, column=1, value=g['label'])
        ws.cell(row=r, column=2, value=f"{g['min_score']}-{g['max_score']}")
        ws.cell(row=r, column=3, value=g['reserve_rate']).number_format = PCT

    r += 3
    ws.cell(row=r, column=1, value="Loan Type Codes").font = Font(bold=True, size=14)
    r += 1
    ws.cell(row=r, column=1, value="Code"); ws.cell(row=r, column=2, value="Loan Pool")
    hdr_row(ws, r, 2)
    for code, pool in sorted(config.get('pool_map', {}).items(), key=lambda x: x[1]):
        r += 1
        ws.cell(row=r, column=1, value=str(code))
        ws.cell(row=r, column=2, value=pool)
    auto_w(ws, 3)


def sheet_all_loans(wb, cu, snap, df, grades, config):
    """All Loans detail listing."""
    ws = wb.create_sheet("All Loans")
    no_score = config.get('no_score_label', 'Not Reported')
    ws['A1'] = "Credit Grade Analysis - All Loans"
    ws['A1'].font = Font(bold=True, size=14)
    ws['F1'] = snap

    headers = ["Member #", "Loan Pool", "Current Balance",
               "Original Score", "Original Grade",
               "Current Score", "Current Grade",
               "Migration Status", "Reserve Rate", "Expected Loss"]
    r = 2
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))

    start = r + 1
    for _, loan in df.iterrows():
        r += 1
        ws.cell(row=r, column=1, value=str(loan.get('member_number', '')))
        ws.cell(row=r, column=2, value=loan.get('loan_pool', ''))
        ws.cell(row=r, column=3, value=loan.get('current_balance', 0))
        ws.cell(row=r, column=4, value=loan.get('original_fico_score', 0))
        ws.cell(row=r, column=5, value=loan.get('original_grade', no_score))
        ws.cell(row=r, column=6, value=loan.get('current_fico_score', 0))
        ws.cell(row=r, column=7, value=loan.get('current_grade', no_score))
        ws.cell(row=r, column=8, value=loan.get('migration_status', 'Unchanged'))
        ws.cell(row=r, column=9, value=loan.get('reserve_rate', 0))
        ws.cell(row=r, column=10, value=loan.get('expected_loss_amount', 0))
        status = loan.get('migration_status', 'Unchanged')
        if status == 'Improved':
            ws.cell(row=r, column=8).fill = IMP_FILL
        elif status == 'Deteriorated':
            ws.cell(row=r, column=8).fill = DET_FILL
    style_rows(ws, start, r, len(headers), mcols={3, 10}, pcols={9})
    auto_w(ws, len(headers), mx=18)


def sheet_hist_balances(wb, cu, snap, df, grades, config, hist=None):
    """Historical Loan Balances by pool with monthly balance data."""
    ws = wb.create_sheet("Historical Balances")
    no_score = config.get('no_score_label', 'Not Reported')
    gl = [g['label'] for g in grades] + [no_score]

    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Historical Loan Balances by Pool"
    ws['A3'] = f"For Period Ending {snap}"

    monthly = hist.get('monthly_balances', pd.DataFrame()) if hist else pd.DataFrame()

    if not monthly.empty:
        # Use monthly balance data - show quarterly snapshots per pool
        # Get quarter-end dates (month-end for Mar, Jun, Sep, Dec)
        monthly['quarter'] = monthly['date'].dt.to_period('Q')
        # Get last date per quarter per pool
        qtr_data = monthly.groupby(['pool', 'quarter']).last().reset_index()
        quarters = sorted(qtr_data['quarter'].unique())
        # Limit to last 20 quarters to keep sheet manageable
        if len(quarters) > 20:
            quarters = quarters[-20:]
        qtr_strs = [str(q) for q in quarters]

        pool_names = sorted(qtr_data['pool'].unique())
        r = 5
        headers = ["Pool"] + qtr_strs
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pool_names:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            for qi, q in enumerate(quarters):
                val = qtr_data[(qtr_data['pool'] == pool) & (qtr_data['quarter'] == q)]
                bal = val['balance'].values[0] if len(val) > 0 else 0
                ws.cell(row=r, column=2 + qi, value=bal).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        for qi, q in enumerate(quarters):
            total = qtr_data[qtr_data['quarter'] == q]['balance'].sum()
            ws.cell(row=r, column=2 + qi, value=total).number_format = MONEY
        style_rows(ws, start, r, ncol, mcols=set(range(2, ncol + 1)))
        auto_w(ws, ncol)
    else:
        # Fallback: just show current data by grade
        pools = sorted(df['loan_pool'].unique())
        r = 6
        for pool in pools:
            ws.cell(row=r, column=1, value=pool).font = Font(bold=True, size=12)
            r += 1
            ws.cell(row=r, column=1, value="Current Grade")
            ws.cell(row=r, column=2, value=snap)
            hdr_row(ws, r, 2)
            pdf = df[df['loan_pool'] == pool]
            for g in gl:
                r += 1
                ws.cell(row=r, column=1, value=g)
                ws.cell(row=r, column=2, value=pdf[pdf['current_grade'] == g]['current_balance'].sum()).number_format = MONEY
            r += 1
            ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
            ws.cell(row=r, column=2, value=pdf['current_balance'].sum()).number_format = MONEY
            r += 2
        auto_w(ws, 2)


def sheet_loss_factor_hist(wb, cu, snap, df, grades, config, hist=None):
    """Loss Factor Historical Detail with charge-off/recovery data."""
    ws = wb.create_sheet("Loss Factor Historical")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Loss Factor Historical Detail"
    ws['A3'] = f"For Period Ending {snap}"

    co_data = hist.get('chargeoffs', {}) if hist else {}
    rc_data = hist.get('recoveries', {}) if hist else {}
    avg_bals = hist.get('avg_balances', {}) if hist else {}
    years = hist.get('years', []) if hist else []

    pools = sorted(set(config.get('pool_map', {}).values()))

    if years:
        year_strs = [str(y) for y in years]
        r = 5
        # Net charge-off rates per pool per year
        headers = ["Pool"] + year_strs + ["Average Life\nLoss Rate"]
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            rates = []
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                avg = avg_bals.get(y, {}).get(pool, 0)
                rate = net / avg if avg > 0 else 0
                ws.cell(row=r, column=2 + yi, value=rate).number_format = PCT4
                rates.append(rate)
            avg_rate = sum(rates) / len(rates) if rates else 0
            ws.cell(row=r, column=ncol, value=avg_rate).number_format = PCT4

        style_rows(ws, start, r, ncol, pcols=set(range(2, ncol + 1)))

        # Average balances section
        r += 3
        ws.cell(row=r, column=1, value="Average Balances by Pool").font = SUB_FONT
        r += 1
        headers2 = ["Pool"] + year_strs
        ncol2 = len(headers2)
        for hi, h in enumerate(headers2):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol2)
        start2 = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            for yi, y in enumerate(years):
                avg = avg_bals.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=avg).number_format = MONEY
        style_rows(ws, start2, r, ncol2, mcols=set(range(2, ncol2 + 1)))
        auto_w(ws, max(ncol, ncol2))
    else:
        ws['A5'] = "No historical data available."
        ws['A5'].font = Font(italic=True, color='888888')


def sheet_chargeoff_hist(wb, cu, snap, config, hist=None):
    """Charge off / Recoveries Historical Detail."""
    ws = wb.create_sheet("Chargeoff Historical")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Charge off and Recoveries Historical Detail"
    ws['A3'] = f"For Period Ending {snap}"

    co_data = hist.get('chargeoffs', {}) if hist else {}
    rc_data = hist.get('recoveries', {}) if hist else {}
    years = hist.get('years', []) if hist else []
    pools = sorted(set(config.get('pool_map', {}).values()))

    if years:
        year_strs = [str(y) for y in years]

        # Charge-offs by pool by year
        r = 5
        ws.cell(row=r, column=1, value="Charge offs by Year").font = SUB_FONT
        r += 1
        headers = ["Pool"] + year_strs + ["Total"]
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_tot = 0
            for yi, y in enumerate(years):
                val = co_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=val).number_format = MONEY
                row_tot += val
            ws.cell(row=r, column=ncol, value=row_tot).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        grand = 0
        for yi, y in enumerate(years):
            yt = sum(co_data.get(y, {}).get(p, 0) for p in pools)
            ws.cell(row=r, column=2 + yi, value=yt).number_format = MONEY
            grand += yt
        ws.cell(row=r, column=ncol, value=grand).number_format = MONEY
        style_rows(ws, start, r, ncol, mcols=set(range(2, ncol + 1)))

        # Recoveries by pool by year
        r += 3
        ws.cell(row=r, column=1, value="Recoveries by Year").font = SUB_FONT
        r += 1
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start2 = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_tot = 0
            for yi, y in enumerate(years):
                val = rc_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=val).number_format = MONEY
                row_tot += val
            ws.cell(row=r, column=ncol, value=row_tot).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        grand = 0
        for yi, y in enumerate(years):
            yt = sum(rc_data.get(y, {}).get(p, 0) for p in pools)
            ws.cell(row=r, column=2 + yi, value=yt).number_format = MONEY
            grand += yt
        ws.cell(row=r, column=ncol, value=grand).number_format = MONEY
        style_rows(ws, start2, r, ncol, mcols=set(range(2, ncol + 1)))

        # Net Charge offs
        r += 2
        headers3 = ["Net Charge offs"] + year_strs + ["Net Charge offs"]
        for hi, h in enumerate(headers3):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers3))
        start3 = r + 1
        grand_net = {y: 0 for y in years}
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_total = 0
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=net).number_format = MONEY
                grand_net[y] += net
                row_total += net
            ws.cell(row=r, column=len(headers3), value=row_total).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        net_total = 0
        for yi, y in enumerate(years):
            ws.cell(row=r, column=2 + yi, value=grand_net[y]).number_format = MONEY
            net_total += grand_net[y]
        ws.cell(row=r, column=len(headers3), value=net_total).number_format = MONEY
        style_rows(ws, start3, r, len(headers3), mcols=set(range(2, len(headers3) + 1)))

        # Life of Loan Loss Rate
        r += 2
        headers4 = ["Life Loss Rate"] + year_strs + ["Average"]
        for hi, h in enumerate(headers4):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers4))
        start4 = r + 1
        for pool in pool_names:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            rates = []
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                avg = avg_bals.get(y, {}).get(pool, 0)
                rate = net / avg if avg > 0 else 0
                ws.cell(row=r, column=2 + yi, value=rate).number_format = PCT4
                rates.append(rate)
            avg_rate = sum(rates) / len(rates) if rates else 0
            ws.cell(row=r, column=len(headers4), value=avg_rate).number_format = PCT4
        style_rows(ws, start4, r, len(headers4), pcols=set(range(2, len(headers4) + 1)))

        auto_w(ws, len(headers))


def sheet_delinquency(wb, cu, snap, config, hist=None):
    """Delinquency Calculation with historical data."""
    ws = wb.create_sheet("Delinquency Calculation")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Delinquency Calculation"
    ws['A3'] = f"For Period Ending {snap}"

    pools = sorted(set(config.get('pool_map', {}).values()))
    dq_pct = hist.get('dq_pct', {}) if hist else {}
    years = sorted(dq_pct.keys()) if dq_pct else list(range(2019, int(snap[:4]) + 1))
    year_strs = [str(y) for y in years]

    r = 5
    headers = ["DQ %"] + year_strs + ["Average", "Variance from Avg"]
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))
    start = r + 1
    for pool in pools:
        r += 1
        ws.cell(row=r, column=1, value=pool)
        rates = []
        for yi, y in enumerate(years):
            val = dq_pct.get(y, {}).get(pool, 0)
            ws.cell(row=r, column=2 + yi, value=val).number_format = PCT
            rates.append(val)
        avg = sum(rates) / len(rates) if rates else 0
        ws.cell(row=r, column=len(headers) - 1, value=avg).number_format = PCT
        # Variance = most recent - average
        current = rates[-1] if rates else 0
        ws.cell(row=r, column=len(headers), value=current - avg).number_format = PCT
    style_rows(ws, start, r, len(headers), pcols=set(range(2, len(headers) + 1)))
    auto_w(ws, len(headers))


def sheet_balance_adj(wb, cu, snap, df, grades, config):
    """FAS 114 / Balance Adjustment sheet."""
    no_score = config.get('no_score_label', 'Not Reported')
    gl = [g['label'] for g in grades] + [no_score]

    ws = wb.create_sheet("Balance Adjustment")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Balance Adjustment"
    ws['A3'] = f"For Period Ending {snap}"

    pools = sorted(df['loan_pool'].unique())
    r = 5
    headers = ["Current Grade", "Loan Report Balance", "Bal Adjustment", "Balance Sheet Total"]

    for pool in pools:
        pdf = df[df['loan_pool'] == pool]
        ws.cell(row=r, column=1, value=pool).font = Font(bold=True, size=12)
        r += 1
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers))
        start = r + 1
        pool_total = 0
        for g in gl:
            r += 1
            bal = pdf[pdf['current_grade'] == g]['current_balance'].sum()
            pool_total += bal
            ws.cell(row=r, column=1, value=g)
            ws.cell(row=r, column=2, value=bal).number_format = MONEY
            ws.cell(row=r, column=3, value=0).number_format = MONEY
            ws.cell(row=r, column=4, value=bal).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        ws.cell(row=r, column=2, value=pool_total).number_format = MONEY
        ws.cell(row=r, column=3, value=0).number_format = MONEY
        ws.cell(row=r, column=4, value=pool_total).number_format = MONEY
        style_rows(ws, start, r, len(headers), mcols={2,3,4})
        r += 2

    # Grand total
    ws.cell(row=r, column=1, value="Grand Total").font = Font(bold=True, size=12)
    total = df['current_balance'].sum()
    ws.cell(row=r, column=2, value=total).number_format = MONEY
    ws.cell(row=r, column=3, value=0).number_format = MONEY
    ws.cell(row=r, column=4, value=total).number_format = MONEY
    auto_w(ws, len(headers))


def sheet_env_ranges(wb):
    """Environmental Factor Ranges reference table."""
    ws = wb.create_sheet("Env Factor Ranges")
    ws['A1'] = "Environmental Factor Ranges"
    ws['A1'].font = Font(bold=True, size=14)

    # Net Credit Change
    r = 3
    ws.cell(row=r, column=1, value="Net Credit Change").font = SUB_FONT
    ws.cell(row=r, column=1+0, value="Range"); ws.cell(row=r, column=2, value="Score")
    ws.cell(row=r, column=4, value="Delinquency").font = SUB_FONT
    ws.cell(row=r, column=4, value="Range"); ws.cell(row=r, column=5, value="Score")
    ws.cell(row=r, column=7, value="Economic Stress Score").font = SUB_FONT
    ws.cell(row=r, column=7, value="Range"); ws.cell(row=r, column=8, value="Score")
    hdr_row(ws, r, 8)

    ncc_rows = [
        ("<-18.00%", "7.00%"), ("-17.99% to -16.00%", "6.00%"),
        ("-15.99% to -14.00%", "5.00%"), ("-13.99% to -11.00%", "4.00%"),
        ("-10.99% to -8.00%", "3.00%"), ("-7.99% to -6.00%", "2.00%"),
        ("-5.99% to -4.00%", "1.00%"), ("-3.99% to 3.99%", "0.00%"),
        ("4.00% to 5.99%", "-1.00%"), ("6.00% to 7.99%", "-2.00%"),
        ("8.00% to 8.99%", "-3.00%"), ("9.00% to 10.99%", "-4.00%"),
        ("11.00% to 12.99%", "-5.00%"), ("13.00% to 14.99%", "-6.00%"),
        (">15.00%", "-7.00%"),
    ]
    dq_rows = [
        (">5.00%", "20.00%"), ("4.00% to 4.99%", "17.00%"),
        ("3.00% to 3.99%", "12.00%"), ("2.50% to 2.99%", "8.00%"),
        ("2.00% to 2.49%", "4.00%"), ("1.50% to 1.99%", "2.50%"),
        ("1.00% to 1.49%", "1.50%"), (".50% to .99%", "0.75%"),
        ("-.49% to .49%", "0.00%"), ("-.99% to -.50%", "-0.75%"),
        ("-1.49% to -1.00%", "-1.50%"), ("-1.99% to -1.50%", "-2.50%"),
        ("-2.49% to -2.00%", "-4.00%"), ("-2.99% to -2.50%", "-8.00%"),
        ("-3.99% to -3.00%", "-12.00%"), ("-4.99% to -4.00%", "-17.00%"),
        ("<-5.00%", "-20.00%"),
    ]
    es_rows = [
        (">25.00%", "10.00%"), ("24.00% to 24.99%", "8.00%"),
        ("22.00% to 23.99%", "7.00%"), ("20.00% to 21.99%", "6.00%"),
        ("18.00% to 19.99%", "5.00%"), ("16.00% to 17.99%", "4.00%"),
        ("14.00% to 15.99%", "3.50%"), ("12.00% to 13.99%", "3.00%"),
        ("10.00% to 11.99%", "2.00%"), ("8.00% to 9.99%", "1.00%"),
        ("6.00% to 7.99%", "0.00%"), ("4.00% to 5.99%", "0.00%"),
        ("2.00% to 3.99%", "-1.00%"), (".00% to 1.99%", "-2.00%"),
    ]

    for i, (rng, sc) in enumerate(ncc_rows):
        ws.cell(row=r + 1 + i, column=1, value=rng)
        ws.cell(row=r + 1 + i, column=2, value=sc)
    for i, (rng, sc) in enumerate(dq_rows):
        ws.cell(row=r + 1 + i, column=4, value=rng)
        ws.cell(row=r + 1 + i, column=5, value=sc)
    for i, (rng, sc) in enumerate(es_rows):
        ws.cell(row=r + 1 + i, column=7, value=rng)
        ws.cell(row=r + 1 + i, column=8, value=sc)
    auto_w(ws, 8)


def sheet_grade_config(wb, grades, config):
    """Grade ranges & loan code reference."""
    ws = wb.create_sheet("Grade Ranges & Loan Codes")
    ws['A1'] = "Credit Grade Configuration"
    ws['A1'].font = Font(bold=True, size=14)

    headers = ["Grade", "Score Range", "Reserve Rate"]
    r = 3
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))
    for g in grades:
        r += 1
        ws.cell(row=r, column=1, value=g['label'])
        ws.cell(row=r, column=2, value=f"{g['min_score']}-{g['max_score']}")
        ws.cell(row=r, column=3, value=g['reserve_rate']).number_format = PCT

    r += 3
    ws.cell(row=r, column=1, value="Loan Type Codes").font = Font(bold=True, size=14)
    r += 1
    ws.cell(row=r, column=1, value="Code"); ws.cell(row=r, column=2, value="Loan Pool")
    hdr_row(ws, r, 2)
    for code, pool in sorted(config.get('pool_map', {}).items(), key=lambda x: x[1]):
        r += 1
        ws.cell(row=r, column=1, value=str(code))
        ws.cell(row=r, column=2, value=pool)
    auto_w(ws, 3)


def sheet_all_loans(wb, cu, snap, df, grades, config):
    """All Loans detail listing."""
    ws = wb.create_sheet("All Loans")
    no_score = config.get('no_score_label', 'Not Reported')
    ws['A1'] = "Credit Grade Analysis - All Loans"
    ws['A1'].font = Font(bold=True, size=14)
    ws['F1'] = snap

    headers = ["Member #", "Loan Pool", "Current Balance",
               "Original Score", "Original Grade",
               "Current Score", "Current Grade",
               "Migration Status", "Reserve Rate", "Expected Loss"]
    r = 2
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))

    start = r + 1
    for _, loan in df.iterrows():
        r += 1
        ws.cell(row=r, column=1, value=str(loan.get('member_number', '')))
        ws.cell(row=r, column=2, value=loan.get('loan_pool', ''))
        ws.cell(row=r, column=3, value=loan.get('current_balance', 0))
        ws.cell(row=r, column=4, value=loan.get('original_fico_score', 0))
        ws.cell(row=r, column=5, value=loan.get('original_grade', no_score))
        ws.cell(row=r, column=6, value=loan.get('current_fico_score', 0))
        ws.cell(row=r, column=7, value=loan.get('current_grade', no_score))
        ws.cell(row=r, column=8, value=loan.get('migration_status', 'Unchanged'))
        ws.cell(row=r, column=9, value=loan.get('reserve_rate', 0))
        ws.cell(row=r, column=10, value=loan.get('expected_loss_amount', 0))
        status = loan.get('migration_status', 'Unchanged')
        if status == 'Improved':
            ws.cell(row=r, column=8).fill = IMP_FILL
        elif status == 'Deteriorated':
            ws.cell(row=r, column=8).fill = DET_FILL
    style_rows(ws, start, r, len(headers), mcols={3, 10}, pcols={9})
    auto_w(ws, len(headers), mx=18)


def sheet_hist_balances(wb, cu, snap, df, grades, config, hist=None):
    """Historical Loan Balances by pool with monthly balance data."""
    ws = wb.create_sheet("Historical Balances")
    no_score = config.get('no_score_label', 'Not Reported')
    gl = [g['label'] for g in grades] + [no_score]

    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Historical Loan Balances by Pool"
    ws['A3'] = f"For Period Ending {snap}"

    monthly = hist.get('monthly_balances', pd.DataFrame()) if hist else pd.DataFrame()

    if not monthly.empty:
        # Use monthly balance data - show quarterly snapshots per pool
        # Get quarter-end dates (month-end for Mar, Jun, Sep, Dec)
        monthly['quarter'] = monthly['date'].dt.to_period('Q')
        # Get last date per quarter per pool
        qtr_data = monthly.groupby(['pool', 'quarter']).last().reset_index()
        quarters = sorted(qtr_data['quarter'].unique())
        # Limit to last 20 quarters to keep sheet manageable
        if len(quarters) > 20:
            quarters = quarters[-20:]
        qtr_strs = [str(q) for q in quarters]

        pool_names = sorted(qtr_data['pool'].unique())
        r = 5
        headers = ["Pool"] + qtr_strs
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pool_names:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            for qi, q in enumerate(quarters):
                val = qtr_data[(qtr_data['pool'] == pool) & (qtr_data['quarter'] == q)]
                bal = val['balance'].values[0] if len(val) > 0 else 0
                ws.cell(row=r, column=2 + qi, value=bal).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        for qi, q in enumerate(quarters):
            total = qtr_data[qtr_data['quarter'] == q]['balance'].sum()
            ws.cell(row=r, column=2 + qi, value=total).number_format = MONEY
        style_rows(ws, start, r, ncol, mcols=set(range(2, ncol + 1)))
        auto_w(ws, ncol)
    else:
        # Fallback: just show current data by grade
        pools = sorted(df['loan_pool'].unique())
        r = 6
        for pool in pools:
            ws.cell(row=r, column=1, value=pool).font = Font(bold=True, size=12)
            r += 1
            ws.cell(row=r, column=1, value="Current Grade")
            ws.cell(row=r, column=2, value=snap)
            hdr_row(ws, r, 2)
            pdf = df[df['loan_pool'] == pool]
            for g in gl:
                r += 1
                ws.cell(row=r, column=1, value=g)
                ws.cell(row=r, column=2, value=pdf[pdf['current_grade'] == g]['current_balance'].sum()).number_format = MONEY
            r += 1
            ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
            ws.cell(row=r, column=2, value=pdf['current_balance'].sum()).number_format = MONEY
            r += 2
        auto_w(ws, 2)


def sheet_loss_factor_hist(wb, cu, snap, df, grades, config, hist=None):
    """Loss Factor Historical Detail with charge-off/recovery data."""
    ws = wb.create_sheet("Loss Factor Historical")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Loss Factor Historical Detail"
    ws['A3'] = f"For Period Ending {snap}"

    co_data = hist.get('chargeoffs', {}) if hist else {}
    rc_data = hist.get('recoveries', {}) if hist else {}
    avg_bals = hist.get('avg_balances', {}) if hist else {}
    years = hist.get('years', []) if hist else []

    pools = sorted(set(config.get('pool_map', {}).values()))

    if years:
        year_strs = [str(y) for y in years]
        r = 5
        # Net charge-off rates per pool per year
        headers = ["Pool"] + year_strs + ["Average Life\nLoss Rate"]
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            rates = []
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                avg = avg_bals.get(y, {}).get(pool, 0)
                rate = net / avg if avg > 0 else 0
                ws.cell(row=r, column=2 + yi, value=rate).number_format = PCT4
                rates.append(rate)
            avg_rate = sum(rates) / len(rates) if rates else 0
            ws.cell(row=r, column=ncol, value=avg_rate).number_format = PCT4

        style_rows(ws, start, r, ncol, pcols=set(range(2, ncol + 1)))

        # Average balances section
        r += 3
        ws.cell(row=r, column=1, value="Average Balances by Pool").font = SUB_FONT
        r += 1
        headers2 = ["Pool"] + year_strs
        ncol2 = len(headers2)
        for hi, h in enumerate(headers2):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol2)
        start2 = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            for yi, y in enumerate(years):
                avg = avg_bals.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=avg).number_format = MONEY
        style_rows(ws, start2, r, ncol2, mcols=set(range(2, ncol2 + 1)))
        auto_w(ws, max(ncol, ncol2))
    else:
        ws['A5'] = "No historical data available."
        ws['A5'].font = Font(italic=True, color='888888')


def sheet_chargeoff_hist(wb, cu, snap, config, hist=None):
    """Charge off / Recoveries Historical Detail."""
    ws = wb.create_sheet("Chargeoff Historical")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Charge off and Recoveries Historical Detail"
    ws['A3'] = f"For Period Ending {snap}"

    co_data = hist.get('chargeoffs', {}) if hist else {}
    rc_data = hist.get('recoveries', {}) if hist else {}
    years = hist.get('years', []) if hist else []
    pools = sorted(set(config.get('pool_map', {}).values()))

    if years:
        year_strs = [str(y) for y in years]

        # Charge-offs by pool by year
        r = 5
        ws.cell(row=r, column=1, value="Charge offs by Year").font = SUB_FONT
        r += 1
        headers = ["Pool"] + year_strs + ["Total"]
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_tot = 0
            for yi, y in enumerate(years):
                val = co_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=val).number_format = MONEY
                row_tot += val
            ws.cell(row=r, column=ncol, value=row_tot).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        grand = 0
        for yi, y in enumerate(years):
            yt = sum(co_data.get(y, {}).get(p, 0) for p in pools)
            ws.cell(row=r, column=2 + yi, value=yt).number_format = MONEY
            grand += yt
        ws.cell(row=r, column=ncol, value=grand).number_format = MONEY
        style_rows(ws, start, r, ncol, mcols=set(range(2, ncol + 1)))

        # Recoveries by pool by year
        r += 3
        ws.cell(row=r, column=1, value="Recoveries by Year").font = SUB_FONT
        r += 1
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start2 = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_tot = 0
            for yi, y in enumerate(years):
                val = rc_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=val).number_format = MONEY
                row_tot += val
            ws.cell(row=r, column=ncol, value=row_tot).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        grand = 0
        for yi, y in enumerate(years):
            yt = sum(rc_data.get(y, {}).get(p, 0) for p in pools)
            ws.cell(row=r, column=2 + yi, value=yt).number_format = MONEY
            grand += yt
        ws.cell(row=r, column=ncol, value=grand).number_format = MONEY
        style_rows(ws, start2, r, ncol, mcols=set(range(2, ncol + 1)))

        # Net Charge offs
        r += 2
        headers3 = ["Net Charge offs"] + year_strs + ["Net Charge offs"]
        for hi, h in enumerate(headers3):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers3))
        start3 = r + 1
        grand_net = {y: 0 for y in years}
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_total = 0
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=net).number_format = MONEY
                grand_net[y] += net
                row_total += net
            ws.cell(row=r, column=len(headers3), value=row_total).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        net_total = 0
        for yi, y in enumerate(years):
            ws.cell(row=r, column=2 + yi, value=grand_net[y]).number_format = MONEY
            net_total += grand_net[y]
        ws.cell(row=r, column=len(headers3), value=net_total).number_format = MONEY
        style_rows(ws, start3, r, len(headers3), mcols=set(range(2, len(headers3) + 1)))

        # Life of Loan Loss Rate
        r += 2
        headers4 = ["Life Loss Rate"] + year_strs + ["Average"]
        for hi, h in enumerate(headers4):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers4))
        start4 = r + 1
        for pool in pool_names:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            rates = []
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                avg = avg_bals.get(y, {}).get(pool, 0)
                rate = net / avg if avg > 0 else 0
                ws.cell(row=r, column=2 + yi, value=rate).number_format = PCT4
                rates.append(rate)
            avg_rate = sum(rates) / len(rates) if rates else 0
            ws.cell(row=r, column=len(headers4), value=avg_rate).number_format = PCT4
        style_rows(ws, start4, r, len(headers4), pcols=set(range(2, len(headers4) + 1)))

        auto_w(ws, len(headers))


def sheet_delinquency(wb, cu, snap, config, hist=None):
    """Delinquency Calculation with historical data."""
    ws = wb.create_sheet("Delinquency Calculation")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Delinquency Calculation"
    ws['A3'] = f"For Period Ending {snap}"

    pools = sorted(set(config.get('pool_map', {}).values()))
    dq_pct = hist.get('dq_pct', {}) if hist else {}
    years = sorted(dq_pct.keys()) if dq_pct else list(range(2019, int(snap[:4]) + 1))
    year_strs = [str(y) for y in years]

    r = 5
    headers = ["DQ %"] + year_strs + ["Average", "Variance from Avg"]
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))
    start = r + 1
    for pool in pools:
        r += 1
        ws.cell(row=r, column=1, value=pool)
        rates = []
        for yi, y in enumerate(years):
            val = dq_pct.get(y, {}).get(pool, 0)
            ws.cell(row=r, column=2 + yi, value=val).number_format = PCT
            rates.append(val)
        avg = sum(rates) / len(rates) if rates else 0
        ws.cell(row=r, column=len(headers) - 1, value=avg).number_format = PCT
        # Variance = most recent - average
        current = rates[-1] if rates else 0
        ws.cell(row=r, column=len(headers), value=current - avg).number_format = PCT
    style_rows(ws, start, r, len(headers), pcols=set(range(2, len(headers) + 1)))
    auto_w(ws, len(headers))


def sheet_balance_adj(wb, cu, snap, df, grades, config):
    """FAS 114 / Balance Adjustment sheet."""
    no_score = config.get('no_score_label', 'Not Reported')
    gl = [g['label'] for g in grades] + [no_score]

    ws = wb.create_sheet("Balance Adjustment")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Balance Adjustment"
    ws['A3'] = f"For Period Ending {snap}"

    pools = sorted(df['loan_pool'].unique())
    r = 5
    headers = ["Current Grade", "Loan Report Balance", "Bal Adjustment", "Balance Sheet Total"]

    for pool in pools:
        pdf = df[df['loan_pool'] == pool]
        ws.cell(row=r, column=1, value=pool).font = Font(bold=True, size=12)
        r += 1
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers))
        start = r + 1
        pool_total = 0
        for g in gl:
            r += 1
            bal = pdf[pdf['current_grade'] == g]['current_balance'].sum()
            pool_total += bal
            ws.cell(row=r, column=1, value=g)
            ws.cell(row=r, column=2, value=bal).number_format = MONEY
            ws.cell(row=r, column=3, value=0).number_format = MONEY
            ws.cell(row=r, column=4, value=bal).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        ws.cell(row=r, column=2, value=pool_total).number_format = MONEY
        ws.cell(row=r, column=3, value=0).number_format = MONEY
        ws.cell(row=r, column=4, value=pool_total).number_format = MONEY
        style_rows(ws, start, r, len(headers), mcols={2,3,4})
        r += 2

    # Grand total
    ws.cell(row=r, column=1, value="Grand Total").font = Font(bold=True, size=12)
    total = df['current_balance'].sum()
    ws.cell(row=r, column=2, value=total).number_format = MONEY
    ws.cell(row=r, column=3, value=0).number_format = MONEY
    ws.cell(row=r, column=4, value=total).number_format = MONEY
    auto_w(ws, len(headers))


def sheet_env_ranges(wb):
    """Environmental Factor Ranges reference table."""
    ws = wb.create_sheet("Env Factor Ranges")
    ws['A1'] = "Environmental Factor Ranges"
    ws['A1'].font = Font(bold=True, size=14)

    # Net Credit Change
    r = 3
    ws.cell(row=r, column=1, value="Net Credit Change").font = SUB_FONT
    ws.cell(row=r, column=1+0, value="Range"); ws.cell(row=r, column=2, value="Score")
    ws.cell(row=r, column=4, value="Delinquency").font = SUB_FONT
    ws.cell(row=r, column=4, value="Range"); ws.cell(row=r, column=5, value="Score")
    ws.cell(row=r, column=7, value="Economic Stress Score").font = SUB_FONT
    ws.cell(row=r, column=7, value="Range"); ws.cell(row=r, column=8, value="Score")
    hdr_row(ws, r, 8)

    ncc_rows = [
        ("<-18.00%", "7.00%"), ("-17.99% to -16.00%", "6.00%"),
        ("-15.99% to -14.00%", "5.00%"), ("-13.99% to -11.00%", "4.00%"),
        ("-10.99% to -8.00%", "3.00%"), ("-7.99% to -6.00%", "2.00%"),
        ("-5.99% to -4.00%", "1.00%"), ("-3.99% to 3.99%", "0.00%"),
        ("4.00% to 5.99%", "-1.00%"), ("6.00% to 7.99%", "-2.00%"),
        ("8.00% to 8.99%", "-3.00%"), ("9.00% to 10.99%", "-4.00%"),
        ("11.00% to 12.99%", "-5.00%"), ("13.00% to 14.99%", "-6.00%"),
        (">15.00%", "-7.00%"),
    ]
    dq_rows = [
        (">5.00%", "20.00%"), ("4.00% to 4.99%", "17.00%"),
        ("3.00% to 3.99%", "12.00%"), ("2.50% to 2.99%", "8.00%"),
        ("2.00% to 2.49%", "4.00%"), ("1.50% to 1.99%", "2.50%"),
        ("1.00% to 1.49%", "1.50%"), (".50% to .99%", "0.75%"),
        ("-.49% to .49%", "0.00%"), ("-.99% to -.50%", "-0.75%"),
        ("-1.49% to -1.00%", "-1.50%"), ("-1.99% to -1.50%", "-2.50%"),
        ("-2.49% to -2.00%", "-4.00%"), ("-2.99% to -2.50%", "-8.00%"),
        ("-3.99% to -3.00%", "-12.00%"), ("-4.99% to -4.00%", "-17.00%"),
        ("<-5.00%", "-20.00%"),
    ]
    es_rows = [
        (">25.00%", "10.00%"), ("24.00% to 24.99%", "8.00%"),
        ("22.00% to 23.99%", "7.00%"), ("20.00% to 21.99%", "6.00%"),
        ("18.00% to 19.99%", "5.00%"), ("16.00% to 17.99%", "4.00%"),
        ("14.00% to 15.99%", "3.50%"), ("12.00% to 13.99%", "3.00%"),
        ("10.00% to 11.99%", "2.00%"), ("8.00% to 9.99%", "1.00%"),
        ("6.00% to 7.99%", "0.00%"), ("4.00% to 5.99%", "0.00%"),
        ("2.00% to 3.99%", "-1.00%"), (".00% to 1.99%", "-2.00%"),
    ]

    for i, (rng, sc) in enumerate(ncc_rows):
        ws.cell(row=r + 1 + i, column=1, value=rng)
        ws.cell(row=r + 1 + i, column=2, value=sc)
    for i, (rng, sc) in enumerate(dq_rows):
        ws.cell(row=r + 1 + i, column=4, value=rng)
        ws.cell(row=r + 1 + i, column=5, value=sc)
    for i, (rng, sc) in enumerate(es_rows):
        ws.cell(row=r + 1 + i, column=7, value=rng)
        ws.cell(row=r + 1 + i, column=8, value=sc)
    auto_w(ws, 8)


def sheet_grade_config(wb, grades, config):
    """Grade ranges & loan code reference."""
    ws = wb.create_sheet("Grade Ranges & Loan Codes")
    ws['A1'] = "Credit Grade Configuration"
    ws['A1'].font = Font(bold=True, size=14)

    headers = ["Grade", "Score Range", "Reserve Rate"]
    r = 3
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))
    for g in grades:
        r += 1
        ws.cell(row=r, column=1, value=g['label'])
        ws.cell(row=r, column=2, value=f"{g['min_score']}-{g['max_score']}")
        ws.cell(row=r, column=3, value=g['reserve_rate']).number_format = PCT

    r += 3
    ws.cell(row=r, column=1, value="Loan Type Codes").font = Font(bold=True, size=14)
    r += 1
    ws.cell(row=r, column=1, value="Code"); ws.cell(row=r, column=2, value="Loan Pool")
    hdr_row(ws, r, 2)
    for code, pool in sorted(config.get('pool_map', {}).items(), key=lambda x: x[1]):
        r += 1
        ws.cell(row=r, column=1, value=str(code))
        ws.cell(row=r, column=2, value=pool)
    auto_w(ws, 3)


def sheet_all_loans(wb, cu, snap, df, grades, config):
    """All Loans detail listing."""
    ws = wb.create_sheet("All Loans")
    no_score = config.get('no_score_label', 'Not Reported')
    ws['A1'] = "Credit Grade Analysis - All Loans"
    ws['A1'].font = Font(bold=True, size=14)
    ws['F1'] = snap

    headers = ["Member #", "Loan Pool", "Current Balance",
               "Original Score", "Original Grade",
               "Current Score", "Current Grade",
               "Migration Status", "Reserve Rate", "Expected Loss"]
    r = 2
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))

    start = r + 1
    for _, loan in df.iterrows():
        r += 1
        ws.cell(row=r, column=1, value=str(loan.get('member_number', '')))
        ws.cell(row=r, column=2, value=loan.get('loan_pool', ''))
        ws.cell(row=r, column=3, value=loan.get('current_balance', 0))
        ws.cell(row=r, column=4, value=loan.get('original_fico_score', 0))
        ws.cell(row=r, column=5, value=loan.get('original_grade', no_score))
        ws.cell(row=r, column=6, value=loan.get('current_fico_score', 0))
        ws.cell(row=r, column=7, value=loan.get('current_grade', no_score))
        ws.cell(row=r, column=8, value=loan.get('migration_status', 'Unchanged'))
        ws.cell(row=r, column=9, value=loan.get('reserve_rate', 0))
        ws.cell(row=r, column=10, value=loan.get('expected_loss_amount', 0))
        status = loan.get('migration_status', 'Unchanged')
        if status == 'Improved':
            ws.cell(row=r, column=8).fill = IMP_FILL
        elif status == 'Deteriorated':
            ws.cell(row=r, column=8).fill = DET_FILL
    style_rows(ws, start, r, len(headers), mcols={3, 10}, pcols={9})
    auto_w(ws, len(headers), mx=18)


def sheet_hist_balances(wb, cu, snap, df, grades, config, hist=None):
    """Historical Loan Balances by pool with monthly balance data."""
    ws = wb.create_sheet("Historical Balances")
    no_score = config.get('no_score_label', 'Not Reported')
    gl = [g['label'] for g in grades] + [no_score]

    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Historical Loan Balances by Pool"
    ws['A3'] = f"For Period Ending {snap}"

    monthly = hist.get('monthly_balances', pd.DataFrame()) if hist else pd.DataFrame()

    if not monthly.empty:
        # Use monthly balance data - show quarterly snapshots per pool
        # Get quarter-end dates (month-end for Mar, Jun, Sep, Dec)
        monthly['quarter'] = monthly['date'].dt.to_period('Q')
        # Get last date per quarter per pool
        qtr_data = monthly.groupby(['pool', 'quarter']).last().reset_index()
        quarters = sorted(qtr_data['quarter'].unique())
        # Limit to last 20 quarters to keep sheet manageable
        if len(quarters) > 20:
            quarters = quarters[-20:]
        qtr_strs = [str(q) for q in quarters]

        pool_names = sorted(qtr_data['pool'].unique())
        r = 5
        headers = ["Pool"] + qtr_strs
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pool_names:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            for qi, q in enumerate(quarters):
                val = qtr_data[(qtr_data['pool'] == pool) & (qtr_data['quarter'] == q)]
                bal = val['balance'].values[0] if len(val) > 0 else 0
                ws.cell(row=r, column=2 + qi, value=bal).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        for qi, q in enumerate(quarters):
            total = qtr_data[qtr_data['quarter'] == q]['balance'].sum()
            ws.cell(row=r, column=2 + qi, value=total).number_format = MONEY
        style_rows(ws, start, r, ncol, mcols=set(range(2, ncol + 1)))
        auto_w(ws, ncol)
    else:
        # Fallback: just show current data by grade
        pools = sorted(df['loan_pool'].unique())
        r = 6
        for pool in pools:
            ws.cell(row=r, column=1, value=pool).font = Font(bold=True, size=12)
            r += 1
            ws.cell(row=r, column=1, value="Current Grade")
            ws.cell(row=r, column=2, value=snap)
            hdr_row(ws, r, 2)
            pdf = df[df['loan_pool'] == pool]
            for g in gl:
                r += 1
                ws.cell(row=r, column=1, value=g)
                ws.cell(row=r, column=2, value=pdf[pdf['current_grade'] == g]['current_balance'].sum()).number_format = MONEY
            r += 1
            ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
            ws.cell(row=r, column=2, value=pdf['current_balance'].sum()).number_format = MONEY
            r += 2
        auto_w(ws, 2)


def sheet_loss_factor_hist(wb, cu, snap, df, grades, config, hist=None):
    """Loss Factor Historical Detail with charge-off/recovery data."""
    ws = wb.create_sheet("Loss Factor Historical")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Loss Factor Historical Detail"
    ws['A3'] = f"For Period Ending {snap}"

    co_data = hist.get('chargeoffs', {}) if hist else {}
    rc_data = hist.get('recoveries', {}) if hist else {}
    avg_bals = hist.get('avg_balances', {}) if hist else {}
    years = hist.get('years', []) if hist else []

    pools = sorted(set(config.get('pool_map', {}).values()))

    if years:
        year_strs = [str(y) for y in years]
        r = 5
        # Net charge-off rates per pool per year
        headers = ["Pool"] + year_strs + ["Average Life\nLoss Rate"]
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            rates = []
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                avg = avg_bals.get(y, {}).get(pool, 0)
                rate = net / avg if avg > 0 else 0
                ws.cell(row=r, column=2 + yi, value=rate).number_format = PCT4
                rates.append(rate)
            avg_rate = sum(rates) / len(rates) if rates else 0
            ws.cell(row=r, column=ncol, value=avg_rate).number_format = PCT4

        style_rows(ws, start, r, ncol, pcols=set(range(2, ncol + 1)))

        # Average balances section
        r += 3
        ws.cell(row=r, column=1, value="Average Balances by Pool").font = SUB_FONT
        r += 1
        headers2 = ["Pool"] + year_strs
        ncol2 = len(headers2)
        for hi, h in enumerate(headers2):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol2)
        start2 = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            for yi, y in enumerate(years):
                avg = avg_bals.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=avg).number_format = MONEY
        style_rows(ws, start2, r, ncol2, mcols=set(range(2, ncol2 + 1)))
        auto_w(ws, max(ncol, ncol2))
    else:
        ws['A5'] = "No historical data available."
        ws['A5'].font = Font(italic=True, color='888888')


def sheet_chargeoff_hist(wb, cu, snap, config, hist=None):
    """Charge off / Recoveries Historical Detail."""
    ws = wb.create_sheet("Chargeoff Historical")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Charge off and Recoveries Historical Detail"
    ws['A3'] = f"For Period Ending {snap}"

    co_data = hist.get('chargeoffs', {}) if hist else {}
    rc_data = hist.get('recoveries', {}) if hist else {}
    years = hist.get('years', []) if hist else []
    pools = sorted(set(config.get('pool_map', {}).values()))

    if years:
        year_strs = [str(y) for y in years]

        # Charge-offs by pool by year
        r = 5
        ws.cell(row=r, column=1, value="Charge offs by Year").font = SUB_FONT
        r += 1
        headers = ["Pool"] + year_strs + ["Total"]
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_tot = 0
            for yi, y in enumerate(years):
                val = co_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=val).number_format = MONEY
                row_tot += val
            ws.cell(row=r, column=ncol, value=row_tot).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        grand = 0
        for yi, y in enumerate(years):
            yt = sum(co_data.get(y, {}).get(p, 0) for p in pools)
            ws.cell(row=r, column=2 + yi, value=yt).number_format = MONEY
            grand += yt
        ws.cell(row=r, column=ncol, value=grand).number_format = MONEY
        style_rows(ws, start, r, ncol, mcols=set(range(2, ncol + 1)))

        # Recoveries by pool by year
        r += 3
        ws.cell(row=r, column=1, value="Recoveries by Year").font = SUB_FONT
        r += 1
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start2 = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_tot = 0
            for yi, y in enumerate(years):
                val = rc_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=val).number_format = MONEY
                row_tot += val
            ws.cell(row=r, column=ncol, value=row_tot).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        grand = 0
        for yi, y in enumerate(years):
            yt = sum(rc_data.get(y, {}).get(p, 0) for p in pools)
            ws.cell(row=r, column=2 + yi, value=yt).number_format = MONEY
            grand += yt
        ws.cell(row=r, column=ncol, value=grand).number_format = MONEY
        style_rows(ws, start2, r, ncol, mcols=set(range(2, ncol + 1)))

        # Net Charge offs
        r += 2
        headers3 = ["Net Charge offs"] + year_strs + ["Net Charge offs"]
        for hi, h in enumerate(headers3):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers3))
        start3 = r + 1
        grand_net = {y: 0 for y in years}
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_total = 0
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=net).number_format = MONEY
                grand_net[y] += net
                row_total += net
            ws.cell(row=r, column=len(headers3), value=row_total).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        net_total = 0
        for yi, y in enumerate(years):
            ws.cell(row=r, column=2 + yi, value=grand_net[y]).number_format = MONEY
            net_total += grand_net[y]
        ws.cell(row=r, column=len(headers3), value=net_total).number_format = MONEY
        style_rows(ws, start3, r, len(headers3), mcols=set(range(2, len(headers3) + 1)))

        # Life of Loan Loss Rate
        r += 2
        headers4 = ["Life Loss Rate"] + year_strs + ["Average"]
        for hi, h in enumerate(headers4):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers4))
        start4 = r + 1
        for pool in pool_names:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            rates = []
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                avg = avg_bals.get(y, {}).get(pool, 0)
                rate = net / avg if avg > 0 else 0
                ws.cell(row=r, column=2 + yi, value=rate).number_format = PCT4
                rates.append(rate)
            avg_rate = sum(rates) / len(rates) if rates else 0
            ws.cell(row=r, column=len(headers4), value=avg_rate).number_format = PCT4
        style_rows(ws, start4, r, len(headers4), pcols=set(range(2, len(headers4) + 1)))

        auto_w(ws, len(headers))


def sheet_delinquency(wb, cu, snap, config, hist=None):
    """Delinquency Calculation with historical data."""
    ws = wb.create_sheet("Delinquency Calculation")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Delinquency Calculation"
    ws['A3'] = f"For Period Ending {snap}"

    pools = sorted(set(config.get('pool_map', {}).values()))
    dq_pct = hist.get('dq_pct', {}) if hist else {}
    years = sorted(dq_pct.keys()) if dq_pct else list(range(2019, int(snap[:4]) + 1))
    year_strs = [str(y) for y in years]

    r = 5
    headers = ["DQ %"] + year_strs + ["Average", "Variance from Avg"]
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))
    start = r + 1
    for pool in pools:
        r += 1
        ws.cell(row=r, column=1, value=pool)
        rates = []
        for yi, y in enumerate(years):
            val = dq_pct.get(y, {}).get(pool, 0)
            ws.cell(row=r, column=2 + yi, value=val).number_format = PCT
            rates.append(val)
        avg = sum(rates) / len(rates) if rates else 0
        ws.cell(row=r, column=len(headers) - 1, value=avg).number_format = PCT
        # Variance = most recent - average
        current = rates[-1] if rates else 0
        ws.cell(row=r, column=len(headers), value=current - avg).number_format = PCT
    style_rows(ws, start, r, len(headers), pcols=set(range(2, len(headers) + 1)))
    auto_w(ws, len(headers))


def sheet_balance_adj(wb, cu, snap, df, grades, config):
    """FAS 114 / Balance Adjustment sheet."""
    no_score = config.get('no_score_label', 'Not Reported')
    gl = [g['label'] for g in grades] + [no_score]

    ws = wb.create_sheet("Balance Adjustment")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Balance Adjustment"
    ws['A3'] = f"For Period Ending {snap}"

    pools = sorted(df['loan_pool'].unique())
    r = 5
    headers = ["Current Grade", "Loan Report Balance", "Bal Adjustment", "Balance Sheet Total"]

    for pool in pools:
        pdf = df[df['loan_pool'] == pool]
        ws.cell(row=r, column=1, value=pool).font = Font(bold=True, size=12)
        r += 1
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers))
        start = r + 1
        pool_total = 0
        for g in gl:
            r += 1
            bal = pdf[pdf['current_grade'] == g]['current_balance'].sum()
            pool_total += bal
            ws.cell(row=r, column=1, value=g)
            ws.cell(row=r, column=2, value=bal).number_format = MONEY
            ws.cell(row=r, column=3, value=0).number_format = MONEY
            ws.cell(row=r, column=4, value=bal).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        ws.cell(row=r, column=2, value=pool_total).number_format = MONEY
        ws.cell(row=r, column=3, value=0).number_format = MONEY
        ws.cell(row=r, column=4, value=pool_total).number_format = MONEY
        style_rows(ws, start, r, len(headers), mcols={2,3,4})
        r += 2

    # Grand total
    ws.cell(row=r, column=1, value="Grand Total").font = Font(bold=True, size=12)
    total = df['current_balance'].sum()
    ws.cell(row=r, column=2, value=total).number_format = MONEY
    ws.cell(row=r, column=3, value=0).number_format = MONEY
    ws.cell(row=r, column=4, value=total).number_format = MONEY
    auto_w(ws, len(headers))


def sheet_env_ranges(wb):
    """Environmental Factor Ranges reference table."""
    ws = wb.create_sheet("Env Factor Ranges")
    ws['A1'] = "Environmental Factor Ranges"
    ws['A1'].font = Font(bold=True, size=14)

    # Net Credit Change
    r = 3
    ws.cell(row=r, column=1, value="Net Credit Change").font = SUB_FONT
    ws.cell(row=r, column=1+0, value="Range"); ws.cell(row=r, column=2, value="Score")
    ws.cell(row=r, column=4, value="Delinquency").font = SUB_FONT
    ws.cell(row=r, column=4, value="Range"); ws.cell(row=r, column=5, value="Score")
    ws.cell(row=r, column=7, value="Economic Stress Score").font = SUB_FONT
    ws.cell(row=r, column=7, value="Range"); ws.cell(row=r, column=8, value="Score")
    hdr_row(ws, r, 8)

    ncc_rows = [
        ("<-18.00%", "7.00%"), ("-17.99% to -16.00%", "6.00%"),
        ("-15.99% to -14.00%", "5.00%"), ("-13.99% to -11.00%", "4.00%"),
        ("-10.99% to -8.00%", "3.00%"), ("-7.99% to -6.00%", "2.00%"),
        ("-5.99% to -4.00%", "1.00%"), ("-3.99% to 3.99%", "0.00%"),
        ("4.00% to 5.99%", "-1.00%"), ("6.00% to 7.99%", "-2.00%"),
        ("8.00% to 8.99%", "-3.00%"), ("9.00% to 10.99%", "-4.00%"),
        ("11.00% to 12.99%", "-5.00%"), ("13.00% to 14.99%", "-6.00%"),
        (">15.00%", "-7.00%"),
    ]
    dq_rows = [
        (">5.00%", "20.00%"), ("4.00% to 4.99%", "17.00%"),
        ("3.00% to 3.99%", "12.00%"), ("2.50% to 2.99%", "8.00%"),
        ("2.00% to 2.49%", "4.00%"), ("1.50% to 1.99%", "2.50%"),
        ("1.00% to 1.49%", "1.50%"), (".50% to .99%", "0.75%"),
        ("-.49% to .49%", "0.00%"), ("-.99% to -.50%", "-0.75%"),
        ("-1.49% to -1.00%", "-1.50%"), ("-1.99% to -1.50%", "-2.50%"),
        ("-2.49% to -2.00%", "-4.00%"), ("-2.99% to -2.50%", "-8.00%"),
        ("-3.99% to -3.00%", "-12.00%"), ("-4.99% to -4.00%", "-17.00%"),
        ("<-5.00%", "-20.00%"),
    ]
    es_rows = [
        (">25.00%", "10.00%"), ("24.00% to 24.99%", "8.00%"),
        ("22.00% to 23.99%", "7.00%"), ("20.00% to 21.99%", "6.00%"),
        ("18.00% to 19.99%", "5.00%"), ("16.00% to 17.99%", "4.00%"),
        ("14.00% to 15.99%", "3.50%"), ("12.00% to 13.99%", "3.00%"),
        ("10.00% to 11.99%", "2.00%"), ("8.00% to 9.99%", "1.00%"),
        ("6.00% to 7.99%", "0.00%"), ("4.00% to 5.99%", "0.00%"),
        ("2.00% to 3.99%", "-1.00%"), (".00% to 1.99%", "-2.00%"),
    ]

    for i, (rng, sc) in enumerate(ncc_rows):
        ws.cell(row=r + 1 + i, column=1, value=rng)
        ws.cell(row=r + 1 + i, column=2, value=sc)
    for i, (rng, sc) in enumerate(dq_rows):
        ws.cell(row=r + 1 + i, column=4, value=rng)
        ws.cell(row=r + 1 + i, column=5, value=sc)
    for i, (rng, sc) in enumerate(es_rows):
        ws.cell(row=r + 1 + i, column=7, value=rng)
        ws.cell(row=r + 1 + i, column=8, value=sc)
    auto_w(ws, 8)


def sheet_grade_config(wb, grades, config):
    """Grade ranges & loan code reference."""
    ws = wb.create_sheet("Grade Ranges & Loan Codes")
    ws['A1'] = "Credit Grade Configuration"
    ws['A1'].font = Font(bold=True, size=14)

    headers = ["Grade", "Score Range", "Reserve Rate"]
    r = 3
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))
    for g in grades:
        r += 1
        ws.cell(row=r, column=1, value=g['label'])
        ws.cell(row=r, column=2, value=f"{g['min_score']}-{g['max_score']}")
        ws.cell(row=r, column=3, value=g['reserve_rate']).number_format = PCT

    r += 3
    ws.cell(row=r, column=1, value="Loan Type Codes").font = Font(bold=True, size=14)
    r += 1
    ws.cell(row=r, column=1, value="Code"); ws.cell(row=r, column=2, value="Loan Pool")
    hdr_row(ws, r, 2)
    for code, pool in sorted(config.get('pool_map', {}).items(), key=lambda x: x[1]):
        r += 1
        ws.cell(row=r, column=1, value=str(code))
        ws.cell(row=r, column=2, value=pool)
    auto_w(ws, 3)


def sheet_all_loans(wb, cu, snap, df, grades, config):
    """All Loans detail listing."""
    ws = wb.create_sheet("All Loans")
    no_score = config.get('no_score_label', 'Not Reported')
    ws['A1'] = "Credit Grade Analysis - All Loans"
    ws['A1'].font = Font(bold=True, size=14)
    ws['F1'] = snap

    headers = ["Member #", "Loan Pool", "Current Balance",
               "Original Score", "Original Grade",
               "Current Score", "Current Grade",
               "Migration Status", "Reserve Rate", "Expected Loss"]
    r = 2
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))

    start = r + 1
    for _, loan in df.iterrows():
        r += 1
        ws.cell(row=r, column=1, value=str(loan.get('member_number', '')))
        ws.cell(row=r, column=2, value=loan.get('loan_pool', ''))
        ws.cell(row=r, column=3, value=loan.get('current_balance', 0))
        ws.cell(row=r, column=4, value=loan.get('original_fico_score', 0))
        ws.cell(row=r, column=5, value=loan.get('original_grade', no_score))
        ws.cell(row=r, column=6, value=loan.get('current_fico_score', 0))
        ws.cell(row=r, column=7, value=loan.get('current_grade', no_score))
        ws.cell(row=r, column=8, value=loan.get('migration_status', 'Unchanged'))
        ws.cell(row=r, column=9, value=loan.get('reserve_rate', 0))
        ws.cell(row=r, column=10, value=loan.get('expected_loss_amount', 0))
        status = loan.get('migration_status', 'Unchanged')
        if status == 'Improved':
            ws.cell(row=r, column=8).fill = IMP_FILL
        elif status == 'Deteriorated':
            ws.cell(row=r, column=8).fill = DET_FILL
    style_rows(ws, start, r, len(headers), mcols={3, 10}, pcols={9})
    auto_w(ws, len(headers), mx=18)


def sheet_hist_balances(wb, cu, snap, df, grades, config, hist=None):
    """Historical Loan Balances by pool with monthly balance data."""
    ws = wb.create_sheet("Historical Balances")
    no_score = config.get('no_score_label', 'Not Reported')
    gl = [g['label'] for g in grades] + [no_score]

    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Historical Loan Balances by Pool"
    ws['A3'] = f"For Period Ending {snap}"

    monthly = hist.get('monthly_balances', pd.DataFrame()) if hist else pd.DataFrame()

    if not monthly.empty:
        # Use monthly balance data - show quarterly snapshots per pool
        # Get quarter-end dates (month-end for Mar, Jun, Sep, Dec)
        monthly['quarter'] = monthly['date'].dt.to_period('Q')
        # Get last date per quarter per pool
        qtr_data = monthly.groupby(['pool', 'quarter']).last().reset_index()
        quarters = sorted(qtr_data['quarter'].unique())
        # Limit to last 20 quarters to keep sheet manageable
        if len(quarters) > 20:
            quarters = quarters[-20:]
        qtr_strs = [str(q) for q in quarters]

        pool_names = sorted(qtr_data['pool'].unique())
        r = 5
        headers = ["Pool"] + qtr_strs
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pool_names:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            for qi, q in enumerate(quarters):
                val = qtr_data[(qtr_data['pool'] == pool) & (qtr_data['quarter'] == q)]
                bal = val['balance'].values[0] if len(val) > 0 else 0
                ws.cell(row=r, column=2 + qi, value=bal).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        for qi, q in enumerate(quarters):
            total = qtr_data[qtr_data['quarter'] == q]['balance'].sum()
            ws.cell(row=r, column=2 + qi, value=total).number_format = MONEY
        style_rows(ws, start, r, ncol, mcols=set(range(2, ncol + 1)))
        auto_w(ws, ncol)
    else:
        # Fallback: just show current data by grade
        pools = sorted(df['loan_pool'].unique())
        r = 6
        for pool in pools:
            ws.cell(row=r, column=1, value=pool).font = Font(bold=True, size=12)
            r += 1
            ws.cell(row=r, column=1, value="Current Grade")
            ws.cell(row=r, column=2, value=snap)
            hdr_row(ws, r, 2)
            pdf = df[df['loan_pool'] == pool]
            for g in gl:
                r += 1
                ws.cell(row=r, column=1, value=g)
                ws.cell(row=r, column=2, value=pdf[pdf['current_grade'] == g]['current_balance'].sum()).number_format = MONEY
            r += 1
            ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
            ws.cell(row=r, column=2, value=pdf['current_balance'].sum()).number_format = MONEY
            r += 2
        auto_w(ws, 2)


def sheet_loss_factor_hist(wb, cu, snap, df, grades, config, hist=None):
    """Loss Factor Historical Detail with charge-off/recovery data."""
    ws = wb.create_sheet("Loss Factor Historical")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Loss Factor Historical Detail"
    ws['A3'] = f"For Period Ending {snap}"

    co_data = hist.get('chargeoffs', {}) if hist else {}
    rc_data = hist.get('recoveries', {}) if hist else {}
    avg_bals = hist.get('avg_balances', {}) if hist else {}
    years = hist.get('years', []) if hist else []

    pools = sorted(set(config.get('pool_map', {}).values()))

    if years:
        year_strs = [str(y) for y in years]
        r = 5
        # Net charge-off rates per pool per year
        headers = ["Pool"] + year_strs + ["Average Life\nLoss Rate"]
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            rates = []
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                avg = avg_bals.get(y, {}).get(pool, 0)
                rate = net / avg if avg > 0 else 0
                ws.cell(row=r, column=2 + yi, value=rate).number_format = PCT4
                rates.append(rate)
            avg_rate = sum(rates) / len(rates) if rates else 0
            ws.cell(row=r, column=ncol, value=avg_rate).number_format = PCT4

        style_rows(ws, start, r, ncol, pcols=set(range(2, ncol + 1)))

        # Average balances section
        r += 3
        ws.cell(row=r, column=1, value="Average Balances by Pool").font = SUB_FONT
        r += 1
        headers2 = ["Pool"] + year_strs
        ncol2 = len(headers2)
        for hi, h in enumerate(headers2):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol2)
        start2 = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            for yi, y in enumerate(years):
                avg = avg_bals.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=avg).number_format = MONEY
        style_rows(ws, start2, r, ncol2, mcols=set(range(2, ncol2 + 1)))
        auto_w(ws, max(ncol, ncol2))
    else:
        ws['A5'] = "No historical data available."
        ws['A5'].font = Font(italic=True, color='888888')


def sheet_chargeoff_hist(wb, cu, snap, config, hist=None):
    """Charge off / Recoveries Historical Detail."""
    ws = wb.create_sheet("Chargeoff Historical")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Charge off and Recoveries Historical Detail"
    ws['A3'] = f"For Period Ending {snap}"

    co_data = hist.get('chargeoffs', {}) if hist else {}
    rc_data = hist.get('recoveries', {}) if hist else {}
    years = hist.get('years', []) if hist else []
    pools = sorted(set(config.get('pool_map', {}).values()))

    if years:
        year_strs = [str(y) for y in years]

        # Charge-offs by pool by year
        r = 5
        ws.cell(row=r, column=1, value="Charge offs by Year").font = SUB_FONT
        r += 1
        headers = ["Pool"] + year_strs + ["Total"]
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_tot = 0
            for yi, y in enumerate(years):
                val = co_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=val).number_format = MONEY
                row_tot += val
            ws.cell(row=r, column=ncol, value=row_tot).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        grand = 0
        for yi, y in enumerate(years):
            yt = sum(co_data.get(y, {}).get(p, 0) for p in pools)
            ws.cell(row=r, column=2 + yi, value=yt).number_format = MONEY
            grand += yt
        ws.cell(row=r, column=ncol, value=grand).number_format = MONEY
        style_rows(ws, start, r, ncol, mcols=set(range(2, ncol + 1)))

        # Recoveries by pool by year
        r += 3
        ws.cell(row=r, column=1, value="Recoveries by Year").font = SUB_FONT
        r += 1
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start2 = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_tot = 0
            for yi, y in enumerate(years):
                val = rc_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=val).number_format = MONEY
                row_tot += val
            ws.cell(row=r, column=ncol, value=row_tot).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        grand = 0
        for yi, y in enumerate(years):
            yt = sum(rc_data.get(y, {}).get(p, 0) for p in pools)
            ws.cell(row=r, column=2 + yi, value=yt).number_format = MONEY
            grand += yt
        ws.cell(row=r, column=ncol, value=grand).number_format = MONEY
        style_rows(ws, start2, r, ncol, mcols=set(range(2, ncol + 1)))

        # Net Charge offs
        r += 2
        headers3 = ["Net Charge offs"] + year_strs + ["Net Charge offs"]
        for hi, h in enumerate(headers3):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers3))
        start3 = r + 1
        grand_net = {y: 0 for y in years}
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_total = 0
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=net).number_format = MONEY
                grand_net[y] += net
                row_total += net
            ws.cell(row=r, column=len(headers3), value=row_total).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        net_total = 0
        for yi, y in enumerate(years):
            ws.cell(row=r, column=2 + yi, value=grand_net[y]).number_format = MONEY
            net_total += grand_net[y]
        ws.cell(row=r, column=len(headers3), value=net_total).number_format = MONEY
        style_rows(ws, start3, r, len(headers3), mcols=set(range(2, len(headers3) + 1)))

        # Life of Loan Loss Rate
        r += 2
        headers4 = ["Life Loss Rate"] + year_strs + ["Average"]
        for hi, h in enumerate(headers4):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers4))
        start4 = r + 1
        for pool in pool_names:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            rates = []
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                avg = avg_bals.get(y, {}).get(pool, 0)
                rate = net / avg if avg > 0 else 0
                ws.cell(row=r, column=2 + yi, value=rate).number_format = PCT4
                rates.append(rate)
            avg_rate = sum(rates) / len(rates) if rates else 0
            ws.cell(row=r, column=len(headers4), value=avg_rate).number_format = PCT4
        style_rows(ws, start4, r, len(headers4), pcols=set(range(2, len(headers4) + 1)))

        auto_w(ws, len(headers))


def sheet_delinquency(wb, cu, snap, config, hist=None):
    """Delinquency Calculation with historical data."""
    ws = wb.create_sheet("Delinquency Calculation")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Delinquency Calculation"
    ws['A3'] = f"For Period Ending {snap}"

    pools = sorted(set(config.get('pool_map', {}).values()))
    dq_pct = hist.get('dq_pct', {}) if hist else {}
    years = sorted(dq_pct.keys()) if dq_pct else list(range(2019, int(snap[:4]) + 1))
    year_strs = [str(y) for y in years]

    r = 5
    headers = ["DQ %"] + year_strs + ["Average", "Variance from Avg"]
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))
    start = r + 1
    for pool in pools:
        r += 1
        ws.cell(row=r, column=1, value=pool)
        rates = []
        for yi, y in enumerate(years):
            val = dq_pct.get(y, {}).get(pool, 0)
            ws.cell(row=r, column=2 + yi, value=val).number_format = PCT
            rates.append(val)
        avg = sum(rates) / len(rates) if rates else 0
        ws.cell(row=r, column=len(headers) - 1, value=avg).number_format = PCT
        # Variance = most recent - average
        current = rates[-1] if rates else 0
        ws.cell(row=r, column=len(headers), value=current - avg).number_format = PCT
    style_rows(ws, start, r, len(headers), pcols=set(range(2, len(headers) + 1)))
    auto_w(ws, len(headers))


def sheet_balance_adj(wb, cu, snap, df, grades, config):
    """FAS 114 / Balance Adjustment sheet."""
    no_score = config.get('no_score_label', 'Not Reported')
    gl = [g['label'] for g in grades] + [no_score]

    ws = wb.create_sheet("Balance Adjustment")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Balance Adjustment"
    ws['A3'] = f"For Period Ending {snap}"

    pools = sorted(df['loan_pool'].unique())
    r = 5
    headers = ["Current Grade", "Loan Report Balance", "Bal Adjustment", "Balance Sheet Total"]

    for pool in pools:
        pdf = df[df['loan_pool'] == pool]
        ws.cell(row=r, column=1, value=pool).font = Font(bold=True, size=12)
        r += 1
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers))
        start = r + 1
        pool_total = 0
        for g in gl:
            r += 1
            bal = pdf[pdf['current_grade'] == g]['current_balance'].sum()
            pool_total += bal
            ws.cell(row=r, column=1, value=g)
            ws.cell(row=r, column=2, value=bal).number_format = MONEY
            ws.cell(row=r, column=3, value=0).number_format = MONEY
            ws.cell(row=r, column=4, value=bal).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        ws.cell(row=r, column=2, value=pool_total).number_format = MONEY
        ws.cell(row=r, column=3, value=0).number_format = MONEY
        ws.cell(row=r, column=4, value=pool_total).number_format = MONEY
        style_rows(ws, start, r, len(headers), mcols={2,3,4})
        r += 2

    # Grand total
    ws.cell(row=r, column=1, value="Grand Total").font = Font(bold=True, size=12)
    total = df['current_balance'].sum()
    ws.cell(row=r, column=2, value=total).number_format = MONEY
    ws.cell(row=r, column=3, value=0).number_format = MONEY
    ws.cell(row=r, column=4, value=total).number_format = MONEY
    auto_w(ws, len(headers))


def sheet_env_ranges(wb):
    """Environmental Factor Ranges reference table."""
    ws = wb.create_sheet("Env Factor Ranges")
    ws['A1'] = "Environmental Factor Ranges"
    ws['A1'].font = Font(bold=True, size=14)

    # Net Credit Change
    r = 3
    ws.cell(row=r, column=1, value="Net Credit Change").font = SUB_FONT
    ws.cell(row=r, column=1+0, value="Range"); ws.cell(row=r, column=2, value="Score")
    ws.cell(row=r, column=4, value="Delinquency").font = SUB_FONT
    ws.cell(row=r, column=4, value="Range"); ws.cell(row=r, column=5, value="Score")
    ws.cell(row=r, column=7, value="Economic Stress Score").font = SUB_FONT
    ws.cell(row=r, column=7, value="Range"); ws.cell(row=r, column=8, value="Score")
    hdr_row(ws, r, 8)

    ncc_rows = [
        ("<-18.00%", "7.00%"), ("-17.99% to -16.00%", "6.00%"),
        ("-15.99% to -14.00%", "5.00%"), ("-13.99% to -11.00%", "4.00%"),
        ("-10.99% to -8.00%", "3.00%"), ("-7.99% to -6.00%", "2.00%"),
        ("-5.99% to -4.00%", "1.00%"), ("-3.99% to 3.99%", "0.00%"),
        ("4.00% to 5.99%", "-1.00%"), ("6.00% to 7.99%", "-2.00%"),
        ("8.00% to 8.99%", "-3.00%"), ("9.00% to 10.99%", "-4.00%"),
        ("11.00% to 12.99%", "-5.00%"), ("13.00% to 14.99%", "-6.00%"),
        (">15.00%", "-7.00%"),
    ]
    dq_rows = [
        (">5.00%", "20.00%"), ("4.00% to 4.99%", "17.00%"),
        ("3.00% to 3.99%", "12.00%"), ("2.50% to 2.99%", "8.00%"),
        ("2.00% to 2.49%", "4.00%"), ("1.50% to 1.99%", "2.50%"),
        ("1.00% to 1.49%", "1.50%"), (".50% to .99%", "0.75%"),
        ("-.49% to .49%", "0.00%"), ("-.99% to -.50%", "-0.75%"),
        ("-1.49% to -1.00%", "-1.50%"), ("-1.99% to -1.50%", "-2.50%"),
        ("-2.49% to -2.00%", "-4.00%"), ("-2.99% to -2.50%", "-8.00%"),
        ("-3.99% to -3.00%", "-12.00%"), ("-4.99% to -4.00%", "-17.00%"),
        ("<-5.00%", "-20.00%"),
    ]
    es_rows = [
        (">25.00%", "10.00%"), ("24.00% to 24.99%", "8.00%"),
        ("22.00% to 23.99%", "7.00%"), ("20.00% to 21.99%", "6.00%"),
        ("18.00% to 19.99%", "5.00%"), ("16.00% to 17.99%", "4.00%"),
        ("14.00% to 15.99%", "3.50%"), ("12.00% to 13.99%", "3.00%"),
        ("10.00% to 11.99%", "2.00%"), ("8.00% to 9.99%", "1.00%"),
        ("6.00% to 7.99%", "0.00%"), ("4.00% to 5.99%", "0.00%"),
        ("2.00% to 3.99%", "-1.00%"), (".00% to 1.99%", "-2.00%"),
    ]

    for i, (rng, sc) in enumerate(ncc_rows):
        ws.cell(row=r + 1 + i, column=1, value=rng)
        ws.cell(row=r + 1 + i, column=2, value=sc)
    for i, (rng, sc) in enumerate(dq_rows):
        ws.cell(row=r + 1 + i, column=4, value=rng)
        ws.cell(row=r + 1 + i, column=5, value=sc)
    for i, (rng, sc) in enumerate(es_rows):
        ws.cell(row=r + 1 + i, column=7, value=rng)
        ws.cell(row=r + 1 + i, column=8, value=sc)
    auto_w(ws, 8)


def sheet_grade_config(wb, grades, config):
    """Grade ranges & loan code reference."""
    ws = wb.create_sheet("Grade Ranges & Loan Codes")
    ws['A1'] = "Credit Grade Configuration"
    ws['A1'].font = Font(bold=True, size=14)

    headers = ["Grade", "Score Range", "Reserve Rate"]
    r = 3
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))
    for g in grades:
        r += 1
        ws.cell(row=r, column=1, value=g['label'])
        ws.cell(row=r, column=2, value=f"{g['min_score']}-{g['max_score']}")
        ws.cell(row=r, column=3, value=g['reserve_rate']).number_format = PCT

    r += 3
    ws.cell(row=r, column=1, value="Loan Type Codes").font = Font(bold=True, size=14)
    r += 1
    ws.cell(row=r, column=1, value="Code"); ws.cell(row=r, column=2, value="Loan Pool")
    hdr_row(ws, r, 2)
    for code, pool in sorted(config.get('pool_map', {}).items(), key=lambda x: x[1]):
        r += 1
        ws.cell(row=r, column=1, value=str(code))
        ws.cell(row=r, column=2, value=pool)
    auto_w(ws, 3)


def sheet_all_loans(wb, cu, snap, df, grades, config):
    """All Loans detail listing."""
    ws = wb.create_sheet("All Loans")
    no_score = config.get('no_score_label', 'Not Reported')
    ws['A1'] = "Credit Grade Analysis - All Loans"
    ws['A1'].font = Font(bold=True, size=14)
    ws['F1'] = snap

    headers = ["Member #", "Loan Pool", "Current Balance",
               "Original Score", "Original Grade",
               "Current Score", "Current Grade",
               "Migration Status", "Reserve Rate", "Expected Loss"]
    r = 2
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))

    start = r + 1
    for _, loan in df.iterrows():
        r += 1
        ws.cell(row=r, column=1, value=str(loan.get('member_number', '')))
        ws.cell(row=r, column=2, value=loan.get('loan_pool', ''))
        ws.cell(row=r, column=3, value=loan.get('current_balance', 0))
        ws.cell(row=r, column=4, value=loan.get('original_fico_score', 0))
        ws.cell(row=r, column=5, value=loan.get('original_grade', no_score))
        ws.cell(row=r, column=6, value=loan.get('current_fico_score', 0))
        ws.cell(row=r, column=7, value=loan.get('current_grade', no_score))
        ws.cell(row=r, column=8, value=loan.get('migration_status', 'Unchanged'))
        ws.cell(row=r, column=9, value=loan.get('reserve_rate', 0))
        ws.cell(row=r, column=10, value=loan.get('expected_loss_amount', 0))
        status = loan.get('migration_status', 'Unchanged')
        if status == 'Improved':
            ws.cell(row=r, column=8).fill = IMP_FILL
        elif status == 'Deteriorated':
            ws.cell(row=r, column=8).fill = DET_FILL
    style_rows(ws, start, r, len(headers), mcols={3, 10}, pcols={9})
    auto_w(ws, len(headers), mx=18)


def sheet_hist_balances(wb, cu, snap, df, grades, config, hist=None):
    """Historical Loan Balances by pool with monthly balance data."""
    ws = wb.create_sheet("Historical Balances")
    no_score = config.get('no_score_label', 'Not Reported')
    gl = [g['label'] for g in grades] + [no_score]

    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Historical Loan Balances by Pool"
    ws['A3'] = f"For Period Ending {snap}"

    monthly = hist.get('monthly_balances', pd.DataFrame()) if hist else pd.DataFrame()

    if not monthly.empty:
        # Use monthly balance data - show quarterly snapshots per pool
        # Get quarter-end dates (month-end for Mar, Jun, Sep, Dec)
        monthly['quarter'] = monthly['date'].dt.to_period('Q')
        # Get last date per quarter per pool
        qtr_data = monthly.groupby(['pool', 'quarter']).last().reset_index()
        quarters = sorted(qtr_data['quarter'].unique())
        # Limit to last 20 quarters to keep sheet manageable
        if len(quarters) > 20:
            quarters = quarters[-20:]
        qtr_strs = [str(q) for q in quarters]

        pool_names = sorted(qtr_data['pool'].unique())
        r = 5
        headers = ["Pool"] + qtr_strs
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pool_names:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            for qi, q in enumerate(quarters):
                val = qtr_data[(qtr_data['pool'] == pool) & (qtr_data['quarter'] == q)]
                bal = val['balance'].values[0] if len(val) > 0 else 0
                ws.cell(row=r, column=2 + qi, value=bal).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        for qi, q in enumerate(quarters):
            total = qtr_data[qtr_data['quarter'] == q]['balance'].sum()
            ws.cell(row=r, column=2 + qi, value=total).number_format = MONEY
        style_rows(ws, start, r, ncol, mcols=set(range(2, ncol + 1)))
        auto_w(ws, ncol)
    else:
        # Fallback: just show current data by grade
        pools = sorted(df['loan_pool'].unique())
        r = 6
        for pool in pools:
            ws.cell(row=r, column=1, value=pool).font = Font(bold=True, size=12)
            r += 1
            ws.cell(row=r, column=1, value="Current Grade")
            ws.cell(row=r, column=2, value=snap)
            hdr_row(ws, r, 2)
            pdf = df[df['loan_pool'] == pool]
            for g in gl:
                r += 1
                ws.cell(row=r, column=1, value=g)
                ws.cell(row=r, column=2, value=pdf[pdf['current_grade'] == g]['current_balance'].sum()).number_format = MONEY
            r += 1
            ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
            ws.cell(row=r, column=2, value=pdf['current_balance'].sum()).number_format = MONEY
            r += 2
        auto_w(ws, 2)


def sheet_loss_factor_hist(wb, cu, snap, df, grades, config, hist=None):
    """Loss Factor Historical Detail with charge-off/recovery data."""
    ws = wb.create_sheet("Loss Factor Historical")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Loss Factor Historical Detail"
    ws['A3'] = f"For Period Ending {snap}"

    co_data = hist.get('chargeoffs', {}) if hist else {}
    rc_data = hist.get('recoveries', {}) if hist else {}
    avg_bals = hist.get('avg_balances', {}) if hist else {}
    years = hist.get('years', []) if hist else []

    pools = sorted(set(config.get('pool_map', {}).values()))

    if years:
        year_strs = [str(y) for y in years]
        r = 5
        # Net charge-off rates per pool per year
        headers = ["Pool"] + year_strs + ["Average Life\nLoss Rate"]
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            rates = []
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                avg = avg_bals.get(y, {}).get(pool, 0)
                rate = net / avg if avg > 0 else 0
                ws.cell(row=r, column=2 + yi, value=rate).number_format = PCT4
                rates.append(rate)
            avg_rate = sum(rates) / len(rates) if rates else 0
            ws.cell(row=r, column=ncol, value=avg_rate).number_format = PCT4

        style_rows(ws, start, r, ncol, pcols=set(range(2, ncol + 1)))

        # Average balances section
        r += 3
        ws.cell(row=r, column=1, value="Average Balances by Pool").font = SUB_FONT
        r += 1
        headers2 = ["Pool"] + year_strs
        ncol2 = len(headers2)
        for hi, h in enumerate(headers2):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol2)
        start2 = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            for yi, y in enumerate(years):
                avg = avg_bals.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=avg).number_format = MONEY
        style_rows(ws, start2, r, ncol2, mcols=set(range(2, ncol2 + 1)))
        auto_w(ws, max(ncol, ncol2))
    else:
        ws['A5'] = "No historical data available."
        ws['A5'].font = Font(italic=True, color='888888')


def sheet_chargeoff_hist(wb, cu, snap, config, hist=None):
    """Charge off / Recoveries Historical Detail."""
    ws = wb.create_sheet("Chargeoff Historical")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Charge off and Recoveries Historical Detail"
    ws['A3'] = f"For Period Ending {snap}"

    co_data = hist.get('chargeoffs', {}) if hist else {}
    rc_data = hist.get('recoveries', {}) if hist else {}
    years = hist.get('years', []) if hist else []
    pools = sorted(set(config.get('pool_map', {}).values()))

    if years:
        year_strs = [str(y) for y in years]

        # Charge-offs by pool by year
        r = 5
        ws.cell(row=r, column=1, value="Charge offs by Year").font = SUB_FONT
        r += 1
        headers = ["Pool"] + year_strs + ["Total"]
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_tot = 0
            for yi, y in enumerate(years):
                val = co_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=val).number_format = MONEY
                row_tot += val
            ws.cell(row=r, column=ncol, value=row_tot).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        grand = 0
        for yi, y in enumerate(years):
            yt = sum(co_data.get(y, {}).get(p, 0) for p in pools)
            ws.cell(row=r, column=2 + yi, value=yt).number_format = MONEY
            grand += yt
        ws.cell(row=r, column=ncol, value=grand).number_format = MONEY
        style_rows(ws, start, r, ncol, mcols=set(range(2, ncol + 1)))

        # Recoveries by pool by year
        r += 3
        ws.cell(row=r, column=1, value="Recoveries by Year").font = SUB_FONT
        r += 1
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start2 = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_tot = 0
            for yi, y in enumerate(years):
                val = rc_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=val).number_format = MONEY
                row_tot += val
            ws.cell(row=r, column=ncol, value=row_tot).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        grand = 0
        for yi, y in enumerate(years):
            yt = sum(rc_data.get(y, {}).get(p, 0) for p in pools)
            ws.cell(row=r, column=2 + yi, value=yt).number_format = MONEY
            grand += yt
        ws.cell(row=r, column=ncol, value=grand).number_format = MONEY
        style_rows(ws, start2, r, ncol, mcols=set(range(2, ncol + 1)))

        # Net Charge offs
        r += 2
        headers3 = ["Net Charge offs"] + year_strs + ["Net Charge offs"]
        for hi, h in enumerate(headers3):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers3))
        start3 = r + 1
        grand_net = {y: 0 for y in years}
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_total = 0
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=net).number_format = MONEY
                grand_net[y] += net
                row_total += net
            ws.cell(row=r, column=len(headers3), value=row_total).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        net_total = 0
        for yi, y in enumerate(years):
            ws.cell(row=r, column=2 + yi, value=grand_net[y]).number_format = MONEY
            net_total += grand_net[y]
        ws.cell(row=r, column=len(headers3), value=net_total).number_format = MONEY
        style_rows(ws, start3, r, len(headers3), mcols=set(range(2, len(headers3) + 1)))

        # Life of Loan Loss Rate
        r += 2
        headers4 = ["Life Loss Rate"] + year_strs + ["Average"]
        for hi, h in enumerate(headers4):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers4))
        start4 = r + 1
        for pool in pool_names:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            rates = []
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                avg = avg_bals.get(y, {}).get(pool, 0)
                rate = net / avg if avg > 0 else 0
                ws.cell(row=r, column=2 + yi, value=rate).number_format = PCT4
                rates.append(rate)
            avg_rate = sum(rates) / len(rates) if rates else 0
            ws.cell(row=r, column=len(headers4), value=avg_rate).number_format = PCT4
        style_rows(ws, start4, r, len(headers4), pcols=set(range(2, len(headers4) + 1)))

        auto_w(ws, len(headers))


def sheet_delinquency(wb, cu, snap, config, hist=None):
    """Delinquency Calculation with historical data."""
    ws = wb.create_sheet("Delinquency Calculation")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Delinquency Calculation"
    ws['A3'] = f"For Period Ending {snap}"

    pools = sorted(set(config.get('pool_map', {}).values()))
    dq_pct = hist.get('dq_pct', {}) if hist else {}
    years = sorted(dq_pct.keys()) if dq_pct else list(range(2019, int(snap[:4]) + 1))
    year_strs = [str(y) for y in years]

    r = 5
    headers = ["DQ %"] + year_strs + ["Average", "Variance from Avg"]
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))
    start = r + 1
    for pool in pools:
        r += 1
        ws.cell(row=r, column=1, value=pool)
        rates = []
        for yi, y in enumerate(years):
            val = dq_pct.get(y, {}).get(pool, 0)
            ws.cell(row=r, column=2 + yi, value=val).number_format = PCT
            rates.append(val)
        avg = sum(rates) / len(rates) if rates else 0
        ws.cell(row=r, column=len(headers) - 1, value=avg).number_format = PCT
        # Variance = most recent - average
        current = rates[-1] if rates else 0
        ws.cell(row=r, column=len(headers), value=current - avg).number_format = PCT
    style_rows(ws, start, r, len(headers), pcols=set(range(2, len(headers) + 1)))
    auto_w(ws, len(headers))


def sheet_balance_adj(wb, cu, snap, df, grades, config):
    """FAS 114 / Balance Adjustment sheet."""
    no_score = config.get('no_score_label', 'Not Reported')
    gl = [g['label'] for g in grades] + [no_score]

    ws = wb.create_sheet("Balance Adjustment")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Balance Adjustment"
    ws['A3'] = f"For Period Ending {snap}"

    pools = sorted(df['loan_pool'].unique())
    r = 5
    headers = ["Current Grade", "Loan Report Balance", "Bal Adjustment", "Balance Sheet Total"]

    for pool in pools:
        pdf = df[df['loan_pool'] == pool]
        ws.cell(row=r, column=1, value=pool).font = Font(bold=True, size=12)
        r += 1
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers))
        start = r + 1
        pool_total = 0
        for g in gl:
            r += 1
            bal = pdf[pdf['current_grade'] == g]['current_balance'].sum()
            pool_total += bal
            ws.cell(row=r, column=1, value=g)
            ws.cell(row=r, column=2, value=bal).number_format = MONEY
            ws.cell(row=r, column=3, value=0).number_format = MONEY
            ws.cell(row=r, column=4, value=bal).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        ws.cell(row=r, column=2, value=pool_total).number_format = MONEY
        ws.cell(row=r, column=3, value=0).number_format = MONEY
        ws.cell(row=r, column=4, value=pool_total).number_format = MONEY
        style_rows(ws, start, r, len(headers), mcols={2,3,4})
        r += 2

    # Grand total
    ws.cell(row=r, column=1, value="Grand Total").font = Font(bold=True, size=12)
    total = df['current_balance'].sum()
    ws.cell(row=r, column=2, value=total).number_format = MONEY
    ws.cell(row=r, column=3, value=0).number_format = MONEY
    ws.cell(row=r, column=4, value=total).number_format = MONEY
    auto_w(ws, len(headers))


def sheet_env_ranges(wb):
    """Environmental Factor Ranges reference table."""
    ws = wb.create_sheet("Env Factor Ranges")
    ws['A1'] = "Environmental Factor Ranges"
    ws['A1'].font = Font(bold=True, size=14)

    # Net Credit Change
    r = 3
    ws.cell(row=r, column=1, value="Net Credit Change").font = SUB_FONT
    ws.cell(row=r, column=1+0, value="Range"); ws.cell(row=r, column=2, value="Score")
    ws.cell(row=r, column=4, value="Delinquency").font = SUB_FONT
    ws.cell(row=r, column=4, value="Range"); ws.cell(row=r, column=5, value="Score")
    ws.cell(row=r, column=7, value="Economic Stress Score").font = SUB_FONT
    ws.cell(row=r, column=7, value="Range"); ws.cell(row=r, column=8, value="Score")
    hdr_row(ws, r, 8)

    ncc_rows = [
        ("<-18.00%", "7.00%"), ("-17.99% to -16.00%", "6.00%"),
        ("-15.99% to -14.00%", "5.00%"), ("-13.99% to -11.00%", "4.00%"),
        ("-10.99% to -8.00%", "3.00%"), ("-7.99% to -6.00%", "2.00%"),
        ("-5.99% to -4.00%", "1.00%"), ("-3.99% to 3.99%", "0.00%"),
        ("4.00% to 5.99%", "-1.00%"), ("6.00% to 7.99%", "-2.00%"),
        ("8.00% to 8.99%", "-3.00%"), ("9.00% to 10.99%", "-4.00%"),
        ("11.00% to 12.99%", "-5.00%"), ("13.00% to 14.99%", "-6.00%"),
        (">15.00%", "-7.00%"),
    ]
    dq_rows = [
        (">5.00%", "20.00%"), ("4.00% to 4.99%", "17.00%"),
        ("3.00% to 3.99%", "12.00%"), ("2.50% to 2.99%", "8.00%"),
        ("2.00% to 2.49%", "4.00%"), ("1.50% to 1.99%", "2.50%"),
        ("1.00% to 1.49%", "1.50%"), (".50% to .99%", "0.75%"),
        ("-.49% to .49%", "0.00%"), ("-.99% to -.50%", "-0.75%"),
        ("-1.49% to -1.00%", "-1.50%"), ("-1.99% to -1.50%", "-2.50%"),
        ("-2.49% to -2.00%", "-4.00%"), ("-2.99% to -2.50%", "-8.00%"),
        ("-3.99% to -3.00%", "-12.00%"), ("-4.99% to -4.00%", "-17.00%"),
        ("<-5.00%", "-20.00%"),
    ]
    es_rows = [
        (">25.00%", "10.00%"), ("24.00% to 24.99%", "8.00%"),
        ("22.00% to 23.99%", "7.00%"), ("20.00% to 21.99%", "6.00%"),
        ("18.00% to 19.99%", "5.00%"), ("16.00% to 17.99%", "4.00%"),
        ("14.00% to 15.99%", "3.50%"), ("12.00% to 13.99%", "3.00%"),
        ("10.00% to 11.99%", "2.00%"), ("8.00% to 9.99%", "1.00%"),
        ("6.00% to 7.99%", "0.00%"), ("4.00% to 5.99%", "0.00%"),
        ("2.00% to 3.99%", "-1.00%"), (".00% to 1.99%", "-2.00%"),
    ]

    for i, (rng, sc) in enumerate(ncc_rows):
        ws.cell(row=r + 1 + i, column=1, value=rng)
        ws.cell(row=r + 1 + i, column=2, value=sc)
    for i, (rng, sc) in enumerate(dq_rows):
        ws.cell(row=r + 1 + i, column=4, value=rng)
        ws.cell(row=r + 1 + i, column=5, value=sc)
    for i, (rng, sc) in enumerate(es_rows):
        ws.cell(row=r + 1 + i, column=7, value=rng)
        ws.cell(row=r + 1 + i, column=8, value=sc)
    auto_w(ws, 8)


def sheet_grade_config(wb, grades, config):
    """Grade ranges & loan code reference."""
    ws = wb.create_sheet("Grade Ranges & Loan Codes")
    ws['A1'] = "Credit Grade Configuration"
    ws['A1'].font = Font(bold=True, size=14)

    headers = ["Grade", "Score Range", "Reserve Rate"]
    r = 3
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))
    for g in grades:
        r += 1
        ws.cell(row=r, column=1, value=g['label'])
        ws.cell(row=r, column=2, value=f"{g['min_score']}-{g['max_score']}")
        ws.cell(row=r, column=3, value=g['reserve_rate']).number_format = PCT

    r += 3
    ws.cell(row=r, column=1, value="Loan Type Codes").font = Font(bold=True, size=14)
    r += 1
    ws.cell(row=r, column=1, value="Code"); ws.cell(row=r, column=2, value="Loan Pool")
    hdr_row(ws, r, 2)
    for code, pool in sorted(config.get('pool_map', {}).items(), key=lambda x: x[1]):
        r += 1
        ws.cell(row=r, column=1, value=str(code))
        ws.cell(row=r, column=2, value=pool)
    auto_w(ws, 3)


def sheet_all_loans(wb, cu, snap, df, grades, config):
    """All Loans detail listing."""
    ws = wb.create_sheet("All Loans")
    no_score = config.get('no_score_label', 'Not Reported')
    ws['A1'] = "Credit Grade Analysis - All Loans"
    ws['A1'].font = Font(bold=True, size=14)
    ws['F1'] = snap

    headers = ["Member #", "Loan Pool", "Current Balance",
               "Original Score", "Original Grade",
               "Current Score", "Current Grade",
               "Migration Status", "Reserve Rate", "Expected Loss"]
    r = 2
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))

    start = r + 1
    for _, loan in df.iterrows():
        r += 1
        ws.cell(row=r, column=1, value=str(loan.get('member_number', '')))
        ws.cell(row=r, column=2, value=loan.get('loan_pool', ''))
        ws.cell(row=r, column=3, value=loan.get('current_balance', 0))
        ws.cell(row=r, column=4, value=loan.get('original_fico_score', 0))
        ws.cell(row=r, column=5, value=loan.get('original_grade', no_score))
        ws.cell(row=r, column=6, value=loan.get('current_fico_score', 0))
        ws.cell(row=r, column=7, value=loan.get('current_grade', no_score))
        ws.cell(row=r, column=8, value=loan.get('migration_status', 'Unchanged'))
        ws.cell(row=r, column=9, value=loan.get('reserve_rate', 0))
        ws.cell(row=r, column=10, value=loan.get('expected_loss_amount', 0))
        status = loan.get('migration_status', 'Unchanged')
        if status == 'Improved':
            ws.cell(row=r, column=8).fill = IMP_FILL
        elif status == 'Deteriorated':
            ws.cell(row=r, column=8).fill = DET_FILL
    style_rows(ws, start, r, len(headers), mcols={3, 10}, pcols={9})
    auto_w(ws, len(headers), mx=18)


def sheet_hist_balances(wb, cu, snap, df, grades, config, hist=None):
    """Historical Loan Balances by pool with monthly balance data."""
    ws = wb.create_sheet("Historical Balances")
    no_score = config.get('no_score_label', 'Not Reported')
    gl = [g['label'] for g in grades] + [no_score]

    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Historical Loan Balances by Pool"
    ws['A3'] = f"For Period Ending {snap}"

    monthly = hist.get('monthly_balances', pd.DataFrame()) if hist else pd.DataFrame()

    if not monthly.empty:
        # Use monthly balance data - show quarterly snapshots per pool
        # Get quarter-end dates (month-end for Mar, Jun, Sep, Dec)
        monthly['quarter'] = monthly['date'].dt.to_period('Q')
        # Get last date per quarter per pool
        qtr_data = monthly.groupby(['pool', 'quarter']).last().reset_index()
        quarters = sorted(qtr_data['quarter'].unique())
        # Limit to last 20 quarters to keep sheet manageable
        if len(quarters) > 20:
            quarters = quarters[-20:]
        qtr_strs = [str(q) for q in quarters]

        pool_names = sorted(qtr_data['pool'].unique())
        r = 5
        headers = ["Pool"] + qtr_strs
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pool_names:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            for qi, q in enumerate(quarters):
                val = qtr_data[(qtr_data['pool'] == pool) & (qtr_data['quarter'] == q)]
                bal = val['balance'].values[0] if len(val) > 0 else 0
                ws.cell(row=r, column=2 + qi, value=bal).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        for qi, q in enumerate(quarters):
            total = qtr_data[qtr_data['quarter'] == q]['balance'].sum()
            ws.cell(row=r, column=2 + qi, value=total).number_format = MONEY
        style_rows(ws, start, r, ncol, mcols=set(range(2, ncol + 1)))
        auto_w(ws, ncol)
    else:
        # Fallback: just show current data by grade
        pools = sorted(df['loan_pool'].unique())
        r = 6
        for pool in pools:
            ws.cell(row=r, column=1, value=pool).font = Font(bold=True, size=12)
            r += 1
            ws.cell(row=r, column=1, value="Current Grade")
            ws.cell(row=r, column=2, value=snap)
            hdr_row(ws, r, 2)
            pdf = df[df['loan_pool'] == pool]
            for g in gl:
                r += 1
                ws.cell(row=r, column=1, value=g)
                ws.cell(row=r, column=2, value=pdf[pdf['current_grade'] == g]['current_balance'].sum()).number_format = MONEY
            r += 1
            ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
            ws.cell(row=r, column=2, value=pdf['current_balance'].sum()).number_format = MONEY
            r += 2
        auto_w(ws, 2)


def sheet_loss_factor_hist(wb, cu, snap, df, grades, config, hist=None):
    """Loss Factor Historical Detail with charge-off/recovery data."""
    ws = wb.create_sheet("Loss Factor Historical")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Loss Factor Historical Detail"
    ws['A3'] = f"For Period Ending {snap}"

    co_data = hist.get('chargeoffs', {}) if hist else {}
    rc_data = hist.get('recoveries', {}) if hist else {}
    avg_bals = hist.get('avg_balances', {}) if hist else {}
    years = hist.get('years', []) if hist else []

    pools = sorted(set(config.get('pool_map', {}).values()))

    if years:
        year_strs = [str(y) for y in years]
        r = 5
        # Net charge-off rates per pool per year
        headers = ["Pool"] + year_strs + ["Average Life\nLoss Rate"]
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            rates = []
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                avg = avg_bals.get(y, {}).get(pool, 0)
                rate = net / avg if avg > 0 else 0
                ws.cell(row=r, column=2 + yi, value=rate).number_format = PCT4
                rates.append(rate)
            avg_rate = sum(rates) / len(rates) if rates else 0
            ws.cell(row=r, column=ncol, value=avg_rate).number_format = PCT4

        style_rows(ws, start, r, ncol, pcols=set(range(2, ncol + 1)))

        # Average balances section
        r += 3
        ws.cell(row=r, column=1, value="Average Balances by Pool").font = SUB_FONT
        r += 1
        headers2 = ["Pool"] + year_strs
        ncol2 = len(headers2)
        for hi, h in enumerate(headers2):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol2)
        start2 = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            for yi, y in enumerate(years):
                avg = avg_bals.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=avg).number_format = MONEY
        style_rows(ws, start2, r, ncol2, mcols=set(range(2, ncol2 + 1)))
        auto_w(ws, max(ncol, ncol2))
    else:
        ws['A5'] = "No historical data available."
        ws['A5'].font = Font(italic=True, color='888888')


def sheet_chargeoff_hist(wb, cu, snap, config, hist=None):
    """Charge off / Recoveries Historical Detail."""
    ws = wb.create_sheet("Chargeoff Historical")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Charge off and Recoveries Historical Detail"
    ws['A3'] = f"For Period Ending {snap}"

    co_data = hist.get('chargeoffs', {}) if hist else {}
    rc_data = hist.get('recoveries', {}) if hist else {}
    years = hist.get('years', []) if hist else []
    pools = sorted(set(config.get('pool_map', {}).values()))

    if years:
        year_strs = [str(y) for y in years]

        # Charge-offs by pool by year
        r = 5
        ws.cell(row=r, column=1, value="Charge offs by Year").font = SUB_FONT
        r += 1
        headers = ["Pool"] + year_strs + ["Total"]
        ncol = len(headers)
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_tot = 0
            for yi, y in enumerate(years):
                val = co_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=val).number_format = MONEY
                row_tot += val
            ws.cell(row=r, column=ncol, value=row_tot).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        grand = 0
        for yi, y in enumerate(years):
            yt = sum(co_data.get(y, {}).get(p, 0) for p in pools)
            ws.cell(row=r, column=2 + yi, value=yt).number_format = MONEY
            grand += yt
        ws.cell(row=r, column=ncol, value=grand).number_format = MONEY
        style_rows(ws, start, r, ncol, mcols=set(range(2, ncol + 1)))

        # Recoveries by pool by year
        r += 3
        ws.cell(row=r, column=1, value="Recoveries by Year").font = SUB_FONT
        r += 1
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, ncol)
        start2 = r + 1
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_tot = 0
            for yi, y in enumerate(years):
                val = rc_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=val).number_format = MONEY
                row_tot += val
            ws.cell(row=r, column=ncol, value=row_tot).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        grand = 0
        for yi, y in enumerate(years):
            yt = sum(rc_data.get(y, {}).get(p, 0) for p in pools)
            ws.cell(row=r, column=2 + yi, value=yt).number_format = MONEY
            grand += yt
        ws.cell(row=r, column=ncol, value=grand).number_format = MONEY
        style_rows(ws, start2, r, ncol, mcols=set(range(2, ncol + 1)))

        # Net Charge offs
        r += 2
        headers3 = ["Net Charge offs"] + year_strs + ["Net Charge offs"]
        for hi, h in enumerate(headers3):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers3))
        start3 = r + 1
        grand_net = {y: 0 for y in years}
        for pool in pools:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            row_total = 0
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                ws.cell(row=r, column=2 + yi, value=net).number_format = MONEY
                grand_net[y] += net
                row_total += net
            ws.cell(row=r, column=len(headers3), value=row_total).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        net_total = 0
        for yi, y in enumerate(years):
            ws.cell(row=r, column=2 + yi, value=grand_net[y]).number_format = MONEY
            net_total += grand_net[y]
        ws.cell(row=r, column=len(headers3), value=net_total).number_format = MONEY
        style_rows(ws, start3, r, len(headers3), mcols=set(range(2, len(headers3) + 1)))

        # Life of Loan Loss Rate
        r += 2
        headers4 = ["Life Loss Rate"] + year_strs + ["Average"]
        for hi, h in enumerate(headers4):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers4))
        start4 = r + 1
        for pool in pool_names:
            r += 1
            ws.cell(row=r, column=1, value=pool)
            rates = []
            for yi, y in enumerate(years):
                net = co_data.get(y, {}).get(pool, 0) - rc_data.get(y, {}).get(pool, 0)
                avg = avg_bals.get(y, {}).get(pool, 0)
                rate = net / avg if avg > 0 else 0
                ws.cell(row=r, column=2 + yi, value=rate).number_format = PCT4
                rates.append(rate)
            avg_rate = sum(rates) / len(rates) if rates else 0
            ws.cell(row=r, column=len(headers4), value=avg_rate).number_format = PCT4
        style_rows(ws, start4, r, len(headers4), pcols=set(range(2, len(headers4) + 1)))

        auto_w(ws, len(headers))


def sheet_delinquency(wb, cu, snap, config, hist=None):
    """Delinquency Calculation with historical data."""
    ws = wb.create_sheet("Delinquency Calculation")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Delinquency Calculation"
    ws['A3'] = f"For Period Ending {snap}"

    pools = sorted(set(config.get('pool_map', {}).values()))
    dq_pct = hist.get('dq_pct', {}) if hist else {}
    years = sorted(dq_pct.keys()) if dq_pct else list(range(2019, int(snap[:4]) + 1))
    year_strs = [str(y) for y in years]

    r = 5
    headers = ["DQ %"] + year_strs + ["Average", "Variance from Avg"]
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))
    start = r + 1
    for pool in pools:
        r += 1
        ws.cell(row=r, column=1, value=pool)
        rates = []
        for yi, y in enumerate(years):
            val = dq_pct.get(y, {}).get(pool, 0)
            ws.cell(row=r, column=2 + yi, value=val).number_format = PCT
            rates.append(val)
        avg = sum(rates) / len(rates) if rates else 0
        ws.cell(row=r, column=len(headers) - 1, value=avg).number_format = PCT
        # Variance = most recent - average
        current = rates[-1] if rates else 0
        ws.cell(row=r, column=len(headers), value=current - avg).number_format = PCT
    style_rows(ws, start, r, len(headers), pcols=set(range(2, len(headers) + 1)))
    auto_w(ws, len(headers))


def sheet_balance_adj(wb, cu, snap, df, grades, config):
    """FAS 114 / Balance Adjustment sheet."""
    no_score = config.get('no_score_label', 'Not Reported')
    gl = [g['label'] for g in grades] + [no_score]

    ws = wb.create_sheet("Balance Adjustment")
    ws['A1'] = cu
    ws['A1'].font = Font(bold=True, size=14)
    ws['A2'] = "Balance Adjustment"
    ws['A3'] = f"For Period Ending {snap}"

    pools = sorted(df['loan_pool'].unique())
    r = 5
    headers = ["Current Grade", "Loan Report Balance", "Bal Adjustment", "Balance Sheet Total"]

    for pool in pools:
        pdf = df[df['loan_pool'] == pool]
        ws.cell(row=r, column=1, value=pool).font = Font(bold=True, size=12)
        r += 1
        for hi, h in enumerate(headers):
            ws.cell(row=r, column=1 + hi, value=h)
        hdr_row(ws, r, len(headers))
        start = r + 1
        pool_total = 0
        for g in gl:
            r += 1
            bal = pdf[pdf['current_grade'] == g]['current_balance'].sum()
            pool_total += bal
            ws.cell(row=r, column=1, value=g)
            ws.cell(row=r, column=2, value=bal).number_format = MONEY
            ws.cell(row=r, column=3, value=0).number_format = MONEY
            ws.cell(row=r, column=4, value=bal).number_format = MONEY
        r += 1
        ws.cell(row=r, column=1, value="Total").font = Font(bold=True)
        ws.cell(row=r, column=2, value=pool_total).number_format = MONEY
        ws.cell(row=r, column=3, value=0).number_format = MONEY
        ws.cell(row=r, column=4, value=pool_total).number_format = MONEY
        style_rows(ws, start, r, len(headers), mcols={2,3,4})
        r += 2

    # Grand total
    ws.cell(row=r, column=1, value="Grand Total").font = Font(bold=True, size=12)
    total = df['current_balance'].sum()
    ws.cell(row=r, column=2, value=total).number_format = MONEY
    ws.cell(row=r, column=3, value=0).number_format = MONEY
    ws.cell(row=r, column=4, value=total).number_format = MONEY
    auto_w(ws, len(headers))


def sheet_env_ranges(wb):
    """Environmental Factor Ranges reference table."""
    ws = wb.create_sheet("Env Factor Ranges")
    ws['A1'] = "Environmental Factor Ranges"
    ws['A1'].font = Font(bold=True, size=14)

    # Net Credit Change
    r = 3
    ws.cell(row=r, column=1, value="Net Credit Change").font = SUB_FONT
    ws.cell(row=r, column=1+0, value="Range"); ws.cell(row=r, column=2, value="Score")
    ws.cell(row=r, column=4, value="Delinquency").font = SUB_FONT
    ws.cell(row=r, column=4, value="Range"); ws.cell(row=r, column=5, value="Score")
    ws.cell(row=r, column=7, value="Economic Stress Score").font = SUB_FONT
    ws.cell(row=r, column=7, value="Range"); ws.cell(row=r, column=8, value="Score")
    hdr_row(ws, r, 8)

    ncc_rows = [
        ("<-18.00%", "7.00%"), ("-17.99% to -16.00%", "6.00%"),
        ("-15.99% to -14.00%", "5.00%"), ("-13.99% to -11.00%", "4.00%"),
        ("-10.99% to -8.00%", "3.00%"), ("-7.99% to -6.00%", "2.00%"),
        ("-5.99% to -4.00%", "1.00%"), ("-3.99% to 3.99%", "0.00%"),
        ("4.00% to 5.99%", "-1.00%"), ("6.00% to 7.99%", "-2.00%"),
        ("8.00% to 8.99%", "-3.00%"), ("9.00% to 10.99%", "-4.00%"),
        ("11.00% to 12.99%", "-5.00%"), ("13.00% to 14.99%", "-6.00%"),
        (">15.00%", "-7.00%"),
    ]
    dq_rows = [
        (">5.00%", "20.00%"), ("4.00% to 4.99%", "17.00%"),
        ("3.00% to 3.99%", "12.00%"), ("2.50% to 2.99%", "8.00%"),
        ("2.00% to 2.49%", "4.00%"), ("1.50% to 1.99%", "2.50%"),
        ("1.00% to 1.49%", "1.50%"), (".50% to .99%", "0.75%"),
        ("-.49% to .49%", "0.00%"), ("-.99% to -.50%", "-0.75%"),
        ("-1.49% to -1.00%", "-1.50%"), ("-1.99% to -1.50%", "-2.50%"),
        ("-2.49% to -2.00%", "-4.00%"), ("-2.99% to -2.50%", "-8.00%"),
        ("-3.99% to -3.00%", "-12.00%"), ("-4.99% to -4.00%", "-17.00%"),
        ("<-5.00%", "-20.00%"),
    ]
    es_rows = [
        (">25.00%", "10.00%"), ("24.00% to 24.99%", "8.00%"),
        ("22.00% to 23.99%", "7.00%"), ("20.00% to 21.99%", "6.00%"),
        ("18.00% to 19.99%", "5.00%"), ("16.00% to 17.99%", "4.00%"),
        ("14.00% to 15.99%", "3.50%"), ("12.00% to 13.99%", "3.00%"),
        ("10.00% to 11.99%", "2.00%"), ("8.00% to 9.99%", "1.00%"),
        ("6.00% to 7.99%", "0.00%"), ("4.00% to 5.99%", "0.00%"),
        ("2.00% to 3.99%", "-1.00%"), (".00% to 1.99%", "-2.00%"),
    ]

    for i, (rng, sc) in enumerate(ncc_rows):
        ws.cell(row=r + 1 + i, column=1, value=rng)
        ws.cell(row=r + 1 + i, column=2, value=sc)
    for i, (rng, sc) in enumerate(dq_rows):
        ws.cell(row=r + 1 + i, column=4, value=rng)
        ws.cell(row=r + 1 + i, column=5, value=sc)
    for i, (rng, sc) in enumerate(es_rows):
        ws.cell(row=r + 1 + i, column=7, value=rng)
        ws.cell(row=r + 1 + i, column=8, value=sc)
    auto_w(ws, 8)


def sheet_grade_config(wb, grades, config):
    """Grade ranges & loan code reference."""
    ws = wb.create_sheet("Grade Ranges & Loan Codes")
    ws['A1'] = "Credit Grade Configuration"
    ws['A1'].font = Font(bold=True, size=14)

    headers = ["Grade", "Score Range", "Reserve Rate"]
    r = 3
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))
    for g in grades:
        r += 1
        ws.cell(row=r, column=1, value=g['label'])
        ws.cell(row=r, column=2, value=f"{g['min_score']}-{g['max_score']}")
        ws.cell(row=r, column=3, value=g['reserve_rate']).number_format = PCT

    r += 3
    ws.cell(row=r, column=1, value="Loan Type Codes").font = Font(bold=True, size=14)
    r += 1
    ws.cell(row=r, column=1, value="Code"); ws.cell(row=r, column=2, value="Loan Pool")
    hdr_row(ws, r, 2)
    for code, pool in sorted(config.get('pool_map', {}).items(), key=lambda x: x[1]):
        r += 1
        ws.cell(row=r, column=1, value=str(code))
        ws.cell(row=r, column=2, value=pool)
    auto_w(ws, 3)


def sheet_all_loans(wb, cu, snap, df, grades, config):
    """All Loans detail listing."""
    ws = wb.create_sheet("All Loans")
    no_score = config.get('no_score_label', 'Not Reported')
    ws['A1'] = "Credit Grade Analysis - All Loans"
    ws['A1'].font = Font(bold=True, size=14)
    ws['F1'] = snap

    headers = ["Member #", "Loan Pool", "Current Balance",
               "Original Score", "Original Grade",
               "Current Score", "Current Grade",
               "Migration Status", "Reserve Rate", "Expected Loss"]
    r = 2
    for hi, h in enumerate(headers):
        ws.cell(row=r, column=1 + hi, value=h)
    hdr_row(ws, r, len(headers))

    start = r + 1
    for _, loan in df.iterrows():
        r += 1
        ws.cell(row=r, column=1, value=str(loan.get('member_number', '')))
        ws.cell(row=r, column=2, value=loan.get('loan_pool', ''))
        ws.cell(row=r, column=3, value=loan.get('current_balance', 0))
        ws.cell(row=r, column=4, value=loan.get('original_fico_score', 0))
        ws.cell(row=r, column=5, value=loan.get('original_grade', no_score))
        ws.cell(row=r, column=6, value=loan.get('current_fico_score', 0))
        ws.cell(row=r, column=7, value=loan.get('current_grade', no_score))
        ws.cell(row=r, column=8, value=loan.get('migration_status', 'Unchanged'))
        ws.cell(row=r, column=9, value=loan.get('reserve_rate', 0))
        ws.cell(row=r, column=10, value=loan.get('expected_loss_amount', 0))


# ── Main Entry Point ──────────────────────────────────────────────

def generate_report(client_name, snapshot_date=None, reports=None):
    """Generate CECL reports for a client.

    Args:
        client_name: Config name (e.g. 'franklin', 'ontario', 'maple')
        snapshot_date: Optional YYYY-MM-DD string; defaults to latest in DB
        reports: Optional list of report types to generate (e.g. ['tct', 'vizo', 'vizo_supp']).
                 If None, uses the client config 'reports' section.

    Returns:
        List of output file paths that were saved.
    """
    config = load_config(client_name)
    cu = config['credit_union']
    grades = config['credit_grades']
    no_score = config.get('no_score_label', 'Not Reported')

    # Drop any grade in credit_grades that duplicates the no_score_label so
    # downstream report tabs don't render two "Not Reported" rows/columns.
    if grades:
        grades = [g for g in grades if (g.get('label') or '').strip() != no_score]
        config = dict(config)
        config['credit_grades'] = grades

    # Determine snapshot date
    if not snapshot_date:
        snapshot_date = latest_date(cu)
        if not snapshot_date:
            print(f"  No data found for {cu}")
            return []

    print(f"\n{'='*60}")
    print(f"  Generating reports for {cu} - {snapshot_date}")
    print(f"{'='*60}")

    audit = get_audit_logger()
    audit.info("BEGIN report generation | client=%s | cu=%s | date=%s | types=%s",
               client_name, cu, snapshot_date, reports or "config-default")

    # Load loan data
    df = load_loans(cu, snapshot_date, config)
    if df.empty:
        print(f"  No loan data for {snapshot_date}")
        return []

    df = calculate_cecl(df, grades, no_score)

    # Load historical data
    hist = load_historical_data(config)

    # Load impaired data from WARM working file
    impaired = load_impaired_data(config, snapshot_date)
    if impaired:
        hist['impaired'] = impaired
    else:
        # No WARM file — try loading hist_bal_data from the prior TCT report
        prior = load_prior_tct_hist_bal(config, snapshot_date)
        if prior:
            # Extend with intermediate months from the monthly balances file
            hbd = prior.get('hist_bal_data', {})
            if hbd:
                # Snapshot WARM-reported pool totals for the snapshot month
                # BEFORE extension overwrites them with loan-extract sums.
                # These represent the authoritative monthly balance per pool
                # (used for Risk Change "Total in Pool" / adjustments).
                snap_ts = pd.Timestamp(snapshot_date)
                snap_ym = snap_ts.to_period('M')
                warm_snap = {}
                for pool, pdata in hbd.items():
                    dates = pdata.get('dates') or []
                    tots = pdata.get('total') or []
                    best_idx = None
                    for i, d in enumerate(dates):
                        try:
                            d_ym = pd.Timestamp(d).to_period('M')
                        except Exception:
                            continue
                        if d_ym <= snap_ym:
                            best_idx = i
                        if d_ym == snap_ym:
                            break
                    if best_idx is not None and best_idx < len(tots):
                        try:
                            warm_snap[str(pool).strip()] = float(tots[best_idx])
                        except (TypeError, ValueError):
                            continue
                prior['warm_snapshot_balances'] = warm_snap

                extend_hist_bal_with_monthly(hbd, hist.get('monthly_balances'))
                extend_hist_bal_with_db(hbd, df, snapshot_date, grades, config)
                n_dates = max(len(d.get('dates', [])) for d in hbd.values())
                print(f"    Extended hist bal: {len(hbd)} pools, {n_dates} months")

            # Merge current-year CO/recovery from file-parsed data into
            # the prior report's warm_co/warm_rc so we get the authoritative
            # prior-year values plus fresh current-year data.
            snap_year = int(snapshot_date[:4])
            if prior.get('warm_co') is not None:
                cur_co = hist.get('chargeoffs', {}).get(snap_year, {})
                cur_rc = hist.get('recoveries', {}).get(snap_year, {})
                if cur_co:
                    prior['warm_co'][snap_year] = cur_co
                if cur_rc:
                    prior['warm_rc'][snap_year] = cur_rc
                # Compute net for current year
                if cur_co or cur_rc:
                    net_yr = {}
                    all_pools = set(list(cur_co.keys()) + list(cur_rc.keys()))
                    for p in all_pools:
                        net_yr[p] = cur_co.get(p, 0) - cur_rc.get(p, 0)
                    prior['warm_net'][snap_year] = net_yr

                # Merge current-year monthly CO/RC data before recalculating
                # ACL totals so windowing can use monthly granularity
                co_monthly_file = hist.get('co_monthly', {})
                rc_monthly_file = hist.get('rc_monthly', {})
                wco_m = prior.get('warm_co_monthly', {})
                wrc_m = prior.get('warm_rc_monthly', {})
                if co_monthly_file or rc_monthly_file:
                    for ym, pools_d in co_monthly_file.items():
                        if ym[0] >= snap_year:
                            wco_m[ym] = pools_d
                    for ym, pools_d in rc_monthly_file.items():
                        if ym[0] >= snap_year:
                            wrc_m[ym] = pools_d
                    if wco_m:
                        prior['warm_co_monthly'] = wco_m
                    if wrc_m:
                        prior['warm_rc_monthly'] = wrc_m

                # Recalculate ACL totals using acl_months window with
                # monthly-level precision for the earliest (partial) year
                acl_map = prior.get('acl_months', {})
                snap_month = int(snapshot_date[5:7])
                for pool in set(p for yr in prior['warm_co'].values()
                                for p in yr):
                    pool_acl = acl_map.get(pool, 36)
                    earliest_abs = (snap_year * 12 + snap_month) - pool_acl + 1
                    earliest_yr = (earliest_abs - 1) // 12
                    earliest_mo = earliest_abs - earliest_yr * 12

                    co_tot = 0
                    rc_tot = 0
                    for y in prior['warm_co']:
                        if y < earliest_yr:
                            continue
                        if y == earliest_yr:
                            # Partial year: sum only months in the window
                            partial = 0
                            has_monthly = False
                            for m in range(earliest_mo, 13):
                                v = wco_m.get((y, m), {}).get(pool, 0)
                                if v:
                                    has_monthly = True
                                partial += v
                            if has_monthly:
                                co_tot += partial
                            else:
                                full = prior['warm_co'].get(y, {}).get(pool, 0)
                                months_in = 12 - earliest_mo + 1
                                co_tot += full * months_in / 12 if full else 0
                        else:
                            co_tot += prior['warm_co'].get(y, {}).get(pool, 0)

                    for y in prior['warm_rc']:
                        if y < earliest_yr:
                            continue
                        if y == earliest_yr:
                            partial = 0
                            has_monthly = False
                            for m in range(earliest_mo, 13):
                                v = wrc_m.get((y, m), {}).get(pool, 0)
                                if v:
                                    has_monthly = True
                                partial += v
                            if has_monthly:
                                # Monthly RC may be stored negative; align sign
                                full_year = prior['warm_rc'].get(y, {}).get(pool, 0)
                                if full_year and (full_year > 0) != (partial > 0):
                                    partial = -partial
                                rc_tot += partial
                            else:
                                full = prior['warm_rc'].get(y, {}).get(pool, 0)
                                months_in = 12 - earliest_mo + 1
                                rc_tot += full * months_in / 12 if full else 0
                        else:
                            rc_tot += prior['warm_rc'].get(y, {}).get(pool, 0)

                    prior['warm_co_totals'][pool] = co_tot
                    prior['warm_rc_totals'][pool] = rc_tot
                    prior['warm_net_co'][pool] = co_tot - rc_tot

                n_co_yr = len(prior['warm_co'])
                print(f"    Merged CO/RC: {n_co_yr} years "
                      f"(added {snap_year} from file-parsed data)")

                # ── Overlay prior-report CO/Rc onto top-level hist keys ──
                # Going forward, historical (prior-year) data should come from
                # the prior TCT report, not from re-parsing source files. Only
                # the current snapshot year is taken from raw-file parsing.
                # This makes prior-year totals stable across runs and aligned
                # with the previously-validated report.
                #
                # User can force a full re-parse of historicals by deleting
                # the prior report from the Reports/ folder.
                hist['chargeoffs'] = {
                    y: dict(pools) for y, pools in prior['warm_co'].items()
                }
                hist['recoveries'] = {
                    y: dict(pools) for y, pools in prior['warm_rc'].items()
                }
                hist['co_monthly'] = {
                    ym: dict(pools)
                    for ym, pools in prior.get('warm_co_monthly', {}).items()
                }
                hist['rc_monthly'] = {
                    ym: dict(pools)
                    for ym, pools in prior.get('warm_rc_monthly', {}).items()
                }
                hist['years'] = sorted(set(hist['chargeoffs'])
                                       | set(hist['recoveries']))
                tot_co = sum(sum(p.values())
                             for p in hist['chargeoffs'].values())
                tot_rc = sum(sum(p.values())
                             for p in hist['recoveries'].values())
                print(f"    Historical CO/Rc sourced from prior TCT report: "
                      f"${tot_co:,.2f} CO / ${tot_rc:,.2f} Rc "
                      f"(years {hist['years'][0]}-{hist['years'][-1]})")

            # Merge current-year DQ% from file-parsed data
            file_dq = hist.get('dq_pct', {})
            if file_dq and prior.get('warm_dq_pct') is not None:
                cur_dq = file_dq.get(snap_year, {})
                if cur_dq:
                    prior['warm_dq_pct'][snap_year] = cur_dq

            hist['impaired'] = prior
        else:
            # No WARM file and no prior TCT report — build a fresh hist_bal_data
            # from the monthly balances workbook + current snapshot grades so the
            # Historical Trends Balance tab still has data to display.
            fresh = build_hist_bal_from_monthly(
                hist.get('monthly_balances'), df, snapshot_date, grades, config,
            )
            if fresh:
                hbd = fresh.get('hist_bal_data', {})
                n_dates = max((len(d.get('dates', [])) for d in hbd.values()), default=0)
                print(f"    Built fresh hist bal from monthly file: "
                      f"{len(hbd)} pools, {n_dates} months")
                hist['impaired'] = fresh

    # ── Load standalone Impaired Loans file (if available) ──
    standalone_imp = load_standalone_impaired(config, snapshot_date, df)
    if standalone_imp:
        imp = hist.get('impaired', {})
        # Overlay fresh impaired data onto whatever was loaded
        imp['acl_impaired'] = standalone_imp['acl_impaired']
        imp['spec_id_by_pool'] = standalone_imp['spec_id_by_pool']
        imp['total_spec_id'] = standalone_imp['total_spec_id']
        hist['impaired'] = imp

    # ── Load wizard-entered impaired loans (highest precedence) ──
    # When the user has entered/uploaded impaired loans in the setup
    # wizard, those rows are the most current source of truth and
    # override anything loaded from WARM / standalone file / prior TCT
    # baseline. The wizard's impaired_parser has already resolved
    # loan_pool + credit_grade via lookup against the loan-data extract.
    wizard_imp = load_wizard_impaired(config)
    if wizard_imp:
        imp = hist.get('impaired', {})
        imp['acl_impaired'] = wizard_imp['acl_impaired']
        imp['spec_id_by_pool'] = wizard_imp['spec_id_by_pool']
        imp['total_spec_id'] = wizard_imp['total_spec_id']
        hist['impaired'] = imp

    # ── Overlay DQ% from loan_code_delinquency_history table ──
    # The wizard's "Historical DQ" step writes rows here from three
    # sources (loan-extract derivation, 5300 backfill, manual entry).
    # DB rows take precedence over the WARM-derived warm_dq_pct: any
    # (year, pool) cell present in the DB overwrites the WARM value;
    # other cells are left alone.
    db_dq = _load_dq_history_from_db(config)
    if db_dq:
        imp = hist.get('impaired') or {}
        existing = imp.get('warm_dq_pct') or {}
        for yr, by_pool in db_dq.items():
            existing.setdefault(yr, {}).update(by_pool)
        imp['warm_dq_pct'] = existing
        hist['impaired'] = imp
        n_cells = sum(len(v) for v in db_dq.values())
        print(f"    Overlaid DQ% from loan_code_delinquency_history: "
              f"{len(db_dq)} year(s), {n_cells} pool-year cell(s).")

    # ── Fallback: load impaired data from prior TCT baseline ──
    # When the source WARM has no 'Impaired Loans' tab and there's no
    # standalone Impaired Loans file, pull from the previously-generated
    # TCT model (which carries 'Impaired Loans' + 'Impaired Loans Pivot' tabs).
    # Even when impaired data is already present from a more current source
    # (wizard / standalone file), still call the baseline loader so that its
    # historical warm_dq_pct / warm_co / warm_rc data from the
    # 'Display CO-Recov -DQ' tab gets merged in (impaired keys are skipped
    # to preserve precedence).
    imp_now = hist.get('impaired', {})
    has_imp = bool(imp_now.get('acl_impaired')) or bool(imp_now.get('spec_id_by_pool'))
    baseline_imp = load_impaired_from_tct_baseline(config, snapshot_date)
    if baseline_imp:
        if not has_imp:
            imp_now.update(baseline_imp)
        else:
            # Only merge historical CO/RC/DQ keys; preserve current impaired data.
            for k, v in baseline_imp.items():
                if k.startswith('warm_') and k not in imp_now:
                    imp_now[k] = v
        hist['impaired'] = imp_now

    # ── Carry-forward WARM Months from prior reports (per-pool ACL months) ──
    # Long-term plan: TCT reports are replacing the legacy CECL-Migration-WARM
    # xlsx files. Going forward, each new quarter inherits its per-pool WARM
    # Months from the most recent prior TCT report; the legacy WARM xlsx is
    # only consulted as a fallback during phase-out (or for the very first TCT
    # generation when no prior TCT exists yet).
    #
    # Override priority (highest wins):
    #   1. Current quarter's CECL-Migration-WARM xlsx (if user maintains one)
    #   2. Most recent prior TCT report's "> Detail_HIst Balances" WARM column
    #   3. Most recent prior CECL-Migration-WARM xlsx (legacy fallback)
    imp_now = hist.get('impaired', {})
    cur_acl = dict(imp_now.get('acl_months', {}) or {})
    pool_order = imp_now.get('pool_order', []) or list(
        imp_now.get('hist_bal_data', {}).keys()
    )
    pre_count = len(cur_acl)

    needs_fill = [p for p in pool_order if p not in cur_acl]
    if needs_fill or not cur_acl:
        # 2. Prior TCT report
        prior_tct = _find_prior_tct_report(config, snapshot_date)
        added_tct = 0
        if prior_tct:
            tct_acl = _load_acl_months_from_tct(prior_tct)
            for pool, months in tct_acl.items():
                if pool not in cur_acl:
                    cur_acl[pool] = months
                    added_tct += 1
            if added_tct:
                print(f"    Carried forward WARM Months from prior TCT "
                      f"({os.path.basename(prior_tct)}): {added_tct} pools")

        # 3. Legacy WARM xlsx (phase-out fallback)
        still_missing = [p for p in pool_order if p not in cur_acl]
        if still_missing or not cur_acl:
            prior_warm = _find_prior_warm_xlsx(config, snapshot_date)
            if prior_warm:
                prior_acl = _load_acl_months_from_warm_xlsx(prior_warm)
                added_warm = 0
                for pool, months in prior_acl.items():
                    if pool not in cur_acl:
                        cur_acl[pool] = months
                        added_warm += 1
                if added_warm:
                    print(f"    Carried forward WARM Months from legacy WARM xlsx "
                          f"({os.path.basename(prior_warm)}): {added_warm} pools")

        if len(cur_acl) > pre_count:
            imp_now['acl_months'] = cur_acl
            hist['impaired'] = imp_now

    # ── YAML overrides for per-pool ACL months and risk_rated flag ──
    # Lets users edit pool settings post-setup without touching WARM files.
    cfg_acl_overrides = config.get('acl_months_by_pool') or {}
    if cfg_acl_overrides:
        imp_now = hist.get('impaired', {})
        cur_acl = dict(imp_now.get('acl_months', {}) or {})
        applied = 0
        for pool, months in cfg_acl_overrides.items():
            try:
                m = int(months)
            except (TypeError, ValueError):
                continue
            if m > 0:
                cur_acl[pool] = m
                applied += 1
        if applied:
            imp_now['acl_months'] = cur_acl
            hist['impaired'] = imp_now
            print(f"    YAML acl_months_by_pool override: {applied} pool(s)")

    cfg_nrr = set(config.get('not_risk_rated', []) or [])
    if cfg_nrr:
        imp_now = hist.get('impaired', {})
        rr_map = dict(imp_now.get('risk_rated', {}) or {})
        for pool in cfg_nrr:
            rr_map[pool] = False
        imp_now['risk_rated'] = rr_map
        hist['impaired'] = imp_now

    # ── Set ACL Balance from Monthly loan balances file (ALLL Balance row) ──
    alll_by_date = hist.get('alll_by_date', {})
    if alll_by_date:
        snap_dt = pd.Timestamp(snapshot_date)
        snap_ym = snap_dt.to_period('M')
        acl_bal = None
        for dt, val in alll_by_date.items():
            if pd.Timestamp(dt).to_period('M') == snap_ym:
                acl_bal = val
                break
        if acl_bal is not None:
            imp = hist.get('impaired', {})
            acl_sum = imp.get('acl_summary', {})
            acl_sum['acl_balance'] = acl_bal
            imp['acl_summary'] = acl_sum
            imp['acl_balance'] = acl_bal  # also at top level for report_vizo
            hist['impaired'] = imp
            print(f"    ACL Balance from monthly file: ${acl_bal:,.2f}")

    # ── Compute balance adjustments from monthly file vs loan file ──
    _compute_balance_adjustments(df, hist, config, snapshot_date)

    # Determine which reports to generate
    if reports is None:
        rpt_cfg = config.get('reports', {})
        reports = [k for k, v in rpt_cfg.items() if v]
    if not reports:
        reports = ['tct']  # default fallback

    os.makedirs(RPT_DIR, exist_ok=True)
    saved = []

    for rpt_type in reports:
        try:
            if rpt_type == 'tct':
                wb, fname = compose_tct_new(client_name, snapshot_date, df, config, grades, hist)
            elif rpt_type == 'vizo':
                wb, fname = compose_vizo_main_new(client_name, snapshot_date, df, config, grades, hist)
            elif rpt_type == 'vizo_supp':
                wb, fname = compose_vizo_supp_new(client_name, snapshot_date, df, config, grades, hist)
            else:
                print(f"  Unknown report type: {rpt_type}")
                continue

            output_path = os.path.join(RPT_DIR, fname)

            # "All Loans" and "Risk Change-All Loans" tabs are intentionally
            # left unlocked so users can sort/filter. (Previously protected
            # with a password; removed per user request.)

            wb.save(output_path)
            print(f"  Saved {rpt_type}: {output_path}")
            saved.append(output_path)
            log_report_generation(client_name, cu, snapshot_date, rpt_type, output_path, success=True)

            # Post-processing: patch charts
            if rpt_type == 'vizo':
                try:
                    patch_dq_pie_zero_labels(output_path)
                    patch_impdet_charts(output_path)
                    patch_drawing_onecell_to_twocell(output_path)
                    patch_remove_chart_borders_and_axis_lines(output_path)
                    print(f"  Patched charts in {fname}")
                except Exception as e:
                    print(f"  Warning: Chart patching failed: {e}")

        except Exception as e:
            print(f"  ERROR generating {rpt_type}: {e}")
            log_report_generation(client_name, cu, snapshot_date, rpt_type, None, success=False)
            import traceback
            traceback.print_exc()

    if saved:
        print(f"\n  {len(saved)} report(s) saved to {RPT_DIR}")
    else:
        print(f"\n  No reports were generated.")

    return saved


if __name__ == '__main__':
    import sys
    parser = argparse.ArgumentParser(
        description="Generate CECL reports (TCT, Vizo, Supplemental)",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python generate_report.py --client franklin --date 2025-12-31
  python generate_report.py --client franklin --reports tct vizo vizo_supp
  python generate_report.py --all --date 2025-12-31
  python generate_report.py --list
        """,
    )
    parser.add_argument('--client', help='Client config name (e.g., "franklin")')
    parser.add_argument('--date', help='Snapshot date (YYYY-MM-DD), defaults to latest')
    parser.add_argument('--reports', nargs='+', choices=['tct', 'vizo', 'vizo_supp'],
                        help='Report types to generate (overrides config)')
    parser.add_argument('--all', action='store_true', help='Generate for all clients')
    parser.add_argument('--list', action='store_true', help='List available clients')
    args = parser.parse_args()

    if args.list:
        print("Available Clients:")
        print(f"  {'Config':20s}  {'Credit Union':40s}  {'Reports'}")
        print(f"  {'-'*20}  {'-'*40}  {'-'*20}")
        for c in list_clients():
            cfg = load_config(c)
            rpts = [k for k, v in cfg.get('reports', {}).items() if v]
            print(f"  {c:20s}  {cfg['credit_union']:40s}  {', '.join(rpts)}")
        sys.exit(0)

    if args.all:
        log_session_start('generate_report.py', f'--all --date={args.date}')
        clients = list_clients()
        print(f"Processing {len(clients)} client(s): {', '.join(clients)}")
        for client_name in clients:
            generate_report(client_name, args.date, args.reports)
        log_session_end('generate_report.py')
    elif args.client:
        log_session_start('generate_report.py', f'--client {args.client} --date={args.date} --reports={args.reports}')
        generate_report(args.client, args.date, args.reports)
        log_session_end('generate_report.py')
    else:
        parser.print_help()
