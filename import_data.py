"""
Config-driven CECL data import pipeline.
Reads client YAML configs and imports loan data from any CU's file format.

Usage:
    python import_data.py --client ontario
    python import_data.py --client ontario --file "AIRESLOANS 2025-12.xlsx"
    python import_data.py --all
    python import_data.py --list
"""
import os
import re
import sys
import shutil
import argparse
import calendar
from datetime import date

import pandas as pd
import yaml
from sqlalchemy import create_engine, text
from dotenv import load_dotenv
from cecl_audit_log import get_audit_logger, log_data_import, log_session_start, log_session_end
from cecl_credentials import get_database_url

load_dotenv()

BASE_FOLDER = os.path.dirname(os.path.abspath(__file__))
CONFIG_FOLDER = os.path.join(BASE_FOLDER, 'client_configs')
UPLOAD_FOLDER = os.path.join(BASE_FOLDER, 'Raw_Uploads')
ARCHIVE_FOLDER = os.path.join(BASE_FOLDER, 'Archive')

db_url = get_database_url()
engine = create_engine(db_url)


def resolve_path(path_value, base=BASE_FOLDER):
    """Resolve configured paths: keep absolute paths, join relative paths to base."""
    if not path_value:
        return ''
    return path_value if os.path.isabs(path_value) else os.path.join(base, path_value)


def load_client_config(client_name):
    """Load a client YAML config file."""
    config_path = os.path.join(CONFIG_FOLDER, f'{client_name}.yaml')
    if not os.path.exists(config_path):
        raise FileNotFoundError(f"Config not found: {config_path}")
    with open(config_path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)


def list_clients():
    """List available client configs (excluding template)."""
    clients = []
    for f in os.listdir(CONFIG_FOLDER):
        if f.endswith('.yaml') and not f.startswith('_'):
            clients.append(os.path.splitext(f)[0])
    return sorted(clients)


def extract_snapshot_date(source_text, config):
    """Extract snapshot date using the client's date regex pattern."""
    date_pattern = config['date_pattern']
    date_fmt = config.get('date_format', 'YYYY-MM')
    text = str(source_text)
    match = re.search(date_pattern, text)

    def _from_month_name(year_token, mon_token):
        month = _MONTH_MAP.get(str(mon_token)[:3].lower())
        if not month:
            return None
        try:
            year = int(year_token)
        except (TypeError, ValueError):
            return None
        last_day = calendar.monthrange(year, month)[1]
        return date(year, month, last_day).isoformat()

    if not match:
        # Fallback: handle month-name formats like "Mar_2026", "March-2026",
        # or "2026_Mar" that the configured numeric regex won't catch.
        mon_first = re.search(
            r"(?i)(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*"
            r"[-_ \.]+(20\d{2})",
            text,
        )
        if mon_first:
            return _from_month_name(mon_first.group(2), mon_first.group(1))
        year_first = re.search(
            r"(?i)(20\d{2})[-_ \.]+"
            r"(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*",
            text,
        )
        if year_first:
            return _from_month_name(year_first.group(1), year_first.group(2))
        # Fallback: "DDMonYY" / "DDMonYYYY" smushed-together filenames like
        # "30JUN25" or "01Jan2026" (no separators between day/month/year).
        # The day is captured but ignored — we still snap to month-end below.
        ddmonyy = re.search(
            r"(?i)(?<!\d)(\d{1,2})(jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)[a-z]*"
            r"[-_ \.]?(\d{2}|\d{4})(?!\d)",
            text,
        )
        if ddmonyy:
            yr_token = ddmonyy.group(3)
            if len(yr_token) == 2:
                yr_token = "20" + yr_token
            return _from_month_name(yr_token, ddmonyy.group(2))
        return None

    # If either captured group is a month name, route through the name parser
    # rather than treating it as a number.
    g1 = match.group(1) if match.lastindex and match.lastindex >= 1 else None
    g2 = match.group(2) if match.lastindex and match.lastindex >= 2 else None
    if g1 and g1[:3].lower() in _MONTH_MAP:
        return _from_month_name(g2, g1)
    if g2 and g2[:3].lower() in _MONTH_MAP:
        return _from_month_name(g1, g2)

    if date_fmt == 'MMDDYY':
        month, day, year = int(match.group(1)), int(match.group(2)), int(match.group(3))
        year += 2000 if year < 100 else 0
        return date(year, month, day).isoformat()
    elif date_fmt == 'MMYY':
        month, year = int(match.group(1)), int(match.group(2))
        year += 2000 if year < 100 else 0
        last_day = calendar.monthrange(year, month)[1]
        return date(year, month, last_day).isoformat()
    else:
        year, month = int(match.group(1)), int(match.group(2))
        last_day = calendar.monthrange(year, month)[1]
        return date(year, month, last_day).isoformat()


def clean_balance(series, balance_format):
    """Clean a balance column based on config rules."""
    s = series.astype(str)
    for char in balance_format.get('remove_chars', []):
        s = s.str.replace(char, '', regex=False)
    if balance_format.get('accounting_negatives', False):
        s = s.str.replace('(', '-', regex=False).str.replace(')', '', regex=False)
    return pd.to_numeric(s, errors='coerce')


def map_pool_codes(series, config):
    """Map raw loan pool codes to pool names using config."""
    split_char = config.get('pool_code_split')
    if split_char:
        raw = series.astype(str).str.split(split_char).str[0].str.strip()
    else:
        raw = series.astype(str).str.strip()
    # Normalize float strings like "85.0" to "85" for numeric codes
    raw = raw.apply(lambda x: str(int(float(x))) if x.replace('.', '', 1).isdigit() else x)
    pool_map = {str(k): v for k, v in config['pool_map'].items()}
    default = config.get('default_pool', 'Other/Uncategorized')
    return raw.map(pool_map).fillna(default)


# Month name -> number for sorting credit pull sheet names
_MONTH_MAP = {
    'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6,
    'jul': 7, 'aug': 8, 'sep': 9, 'oct': 10, 'nov': 11, 'dec': 12,
}


def _sort_credit_pull_sheets(sheet_names):
    """Sort credit pull sheet names most-recent-first.

    Handles names like 'Dec-25 Credit Pull', 'Dec-24 Credit Pull', etc.
    """
    def _sort_key(name):
        m = re.search(r'([A-Za-z]{3})-(\d{2})', name)
        if m:
            month = _MONTH_MAP.get(m.group(1).lower(), 0)
            year = int(m.group(2))
            return (-year, -month)
        return (0, 0)
    return sorted(sheet_names, key=_sort_key)


def load_credit_pull_scores(config):
    """
    Load current credit scores from a credit pull source.
    Priority:
      1. Standalone credit pull file (latest bureau scores)
      2. Credit pull tabs in existing CECL/WARM report (older bureau scores)
      3. WARM "All Loans" current score column (per-loan carry-forward scores)
    Returns a tuple of:
      - member_scores: dict {member_number (int): fico_score (int)} — member-level
      - acct_scores: dict {account_number (str): fico_score (int)} — per-loan level
      - pull_as_of: pd.Timestamp or None — best-effort date of the latest credit
        pull source (config override > standalone file mtime > fallback report mtime)
    """
    cp_config = config.get('credit_pull', {})
    if not cp_config:
        return {}, {}, None

    scores = {}
    warm_acct_scores = {}  # per-loan scores from WARM All Loans
    pull_as_of = None      # tracked across the source files we read

    # Optional explicit override from the wizard.
    pull_override = (cp_config.get('pull_as_of_date') or '').strip()
    if pull_override:
        try:
            pull_as_of = pd.Timestamp(pull_override)
        except (ValueError, TypeError):
            print(f"    WARNING: Could not parse pull_as_of_date '{pull_override}'")
            pull_as_of = None

    # 1. Try standalone credit pull file
    file_pattern = cp_config.get('file_pattern')
    if file_pattern:
        source_folder = resolve_path(cp_config.get('source_folder', ''))
        if os.path.isdir(source_folder):
            pattern_re = re.compile(file_pattern, re.IGNORECASE)
            matching_files = []
            for root, _dirs, files in os.walk(source_folder):
                for fname in files:
                    if fname.startswith('~$'):
                        continue
                    if pattern_re.search(fname):
                        matching_files.append(os.path.join(root, fname))

            for fpath in sorted(matching_files, key=os.path.getmtime, reverse=True):
                    fname = os.path.basename(fpath)
                    member_col = cp_config.get('member_column', 'Member Number')
                    score_col = cp_config.get('score_column', 'FICO')
                    df = pd.read_excel(fpath)
                    df.columns = [str(c).strip() for c in df.columns]
                    if member_col in df.columns and score_col in df.columns:
                        # Strip the same suffix the Aires file uses so the
                        # member-number key matches what import_file builds.
                        suffix_length = config.get('account_suffix_length', 0) or 0
                        for _, row in df.iterrows():
                            m = row[member_col]
                            s = row[score_col]
                            if pd.notna(m) and pd.notna(s):
                                try:
                                    score_val = int(float(s))
                                    if score_val <= 0:
                                        continue  # skip zero = no data
                                    m_str = str(m).strip()
                                    if m_str.endswith('.0'):
                                        m_str = m_str[:-2]
                                    if suffix_length and len(m_str) > suffix_length:
                                        m_key = int(m_str[:-suffix_length])
                                    else:
                                        m_key = int(float(m))
                                    scores[m_key] = score_val
                                except (ValueError, TypeError):
                                    pass
                        print(f"    Credit pull file: {fname} ({len(scores)} scores loaded)")
                    if pull_as_of is None:
                        try:
                            pull_as_of = pd.Timestamp(os.path.getmtime(fpath), unit='s')
                        except (OSError, ValueError):
                            pass
                    break  # use first matching file

    # 2. Also check credit pull tabs in existing CECL report to fill gaps
    # (older credit pulls may cover members not in the latest standalone file)
    warm_file_path = None  # track the WARM file for step 3
    report_pattern = cp_config.get('fallback_report_pattern')
    report_folder = cp_config.get('fallback_report_folder')
    if report_pattern and report_folder:
        folder_path = resolve_path(report_folder)
        if os.path.isdir(folder_path):
            pattern_re = re.compile(report_pattern, re.IGNORECASE)
            sheet_pattern = cp_config.get('fallback_sheet_pattern', 'Credit Pull')
            member_idx = cp_config.get('fallback_member_col', 0)
            score_idx = cp_config.get('fallback_score_col', 1)
            matching_reports = []
            for root, _dirs, files in os.walk(folder_path):
                for fname in files:
                    if fname.startswith('~$'):
                        continue
                    if pattern_re.search(fname):
                        matching_reports.append(os.path.join(root, fname))

            for fpath in sorted(matching_reports, key=os.path.getmtime, reverse=True):
                    warm_file_path = fpath
                    if pull_as_of is None:
                        try:
                            pull_as_of = pd.Timestamp(os.path.getmtime(fpath), unit='s')
                        except (OSError, ValueError):
                            pass
                    fname = os.path.basename(fpath)
                    xl = pd.ExcelFile(fpath)
                    cp_sheets = [s for s in xl.sheet_names
                                 if sheet_pattern.lower() in s.lower()]
                    if not cp_sheets:
                        continue
                    # Sort credit pull sheets: most recent first (by date in name)
                    # so we fill newest scores first and older sheets fill gaps
                    cp_sheets = _sort_credit_pull_sheets(cp_sheets)
                    for sheet_name in cp_sheets:
                        sheet_count = 0
                        df = pd.read_excel(xl, sheet_name, header=None, skiprows=1)
                        for _, row in df.iterrows():
                            m = row.iloc[member_idx]
                            s = row.iloc[score_idx]
                            if pd.notna(m) and pd.notna(s):
                                try:
                                    score_val = int(float(s))
                                    mem_id = int(float(m))
                                    if score_val > 0 and mem_id not in scores:
                                        scores[mem_id] = score_val
                                        sheet_count += 1
                                except (ValueError, TypeError):
                                    pass
                        print(f"    Credit pull tab: '{sheet_name}' from {fname} (+{sheet_count} scores, total {len(scores)})")
                    break  # use first matching report

    # 3. Fill remaining gaps from WARM "All Loans" current score column.
    # The WARM workbook carries forward scores from prior periods per loan.
    # This is useful on first import when there's no previous snapshot in the DB.
    # Scores are keyed per-loan (member-suffix) since different loans for the
    # same member may carry different scores.
    warm_scores_cfg = cp_config.get('warm_scores', {})
    if warm_scores_cfg:
        warm_sheet = warm_scores_cfg.get('sheet', 'All Loans')
        warm_member_col = warm_scores_cfg.get('member_col', 0)
        warm_suffix_col = warm_scores_cfg.get('suffix_col', 1)
        warm_score_col = warm_scores_cfg.get('score_col', 8)
        warm_skip_rows = warm_scores_cfg.get('skip_rows', 2)  # header + column labels
        suffix_length = config.get('account_suffix_length', 0)

        # Use the same WARM file found in step 2, or search for it
        warm_path = warm_file_path
        if not warm_path:
            warm_pattern = warm_scores_cfg.get('file_pattern', report_pattern)
            warm_folder = warm_scores_cfg.get('folder', report_folder)
            if warm_pattern and warm_folder:
                folder_path = resolve_path(warm_folder)
                if os.path.isdir(folder_path):
                    pattern_re = re.compile(warm_pattern, re.IGNORECASE)
                    for root, _dirs, files in os.walk(folder_path):
                        for fname in sorted(files, reverse=True):
                            if fname.startswith('~$'):
                                continue
                            if pattern_re.search(fname):
                                warm_path = os.path.join(root, fname)
                                break
                        if warm_path:
                            break

        if warm_path:
            try:
                import openpyxl
                wb = openpyxl.load_workbook(warm_path, data_only=True, read_only=True)
                if warm_sheet in wb.sheetnames:
                    ws = wb[warm_sheet]
                    warm_count = 0
                    warm_member_count = 0
                    for row_idx, row in enumerate(ws.iter_rows(
                            min_row=warm_skip_rows + 1, values_only=True)):
                        if len(row) <= max(warm_member_col, warm_suffix_col, warm_score_col):
                            continue
                        m = row[warm_member_col]
                        suffix = row[warm_suffix_col]
                        s = row[warm_score_col]
                        if m is not None and s is not None:
                            try:
                                mem_id = int(float(m))
                                score_val = int(float(s))
                                suffix_str = str(int(float(suffix))).zfill(suffix_length) if suffix is not None and suffix_length > 0 else str(int(float(suffix))) if suffix is not None else ''
                                # Build account-level key (member+suffix) matching loan file formatformat
                                acct_key = f"{mem_id}{suffix_str}"
                                if score_val > 0:
                                    if acct_key not in warm_acct_scores:
                                        warm_acct_scores[acct_key] = score_val
                                        warm_count += 1
                                    # Also add member-level for credit pull gap-fill
                                    if mem_id not in scores:
                                        scores[mem_id] = score_val
                                        warm_member_count += 1
                            except (ValueError, TypeError):
                                pass
                    if warm_count > 0:
                        print(f"    WARM '{warm_sheet}' scores: +{warm_member_count} members, +{warm_count} accounts (total {len(scores)} members)")
                wb.close()
            except Exception as e:
                print(f"    WARNING: Could not read WARM scores: {e}")

    return scores, warm_acct_scores, pull_as_of


def extract_member_number(account_series, suffix_length):
    """Strip the loan suffix from account numbers to get the member number."""
    acct_str = account_series.astype(str).str.strip()
    # Drop trailing ".0" produced when pandas reads numeric account columns as float
    acct_str = acct_str.str.replace(r'\.0+$', '', regex=True)
    if suffix_length and suffix_length > 0:
        return acct_str.str[:-suffix_length].astype(int)
    return pd.to_numeric(acct_str, errors='coerce').fillna(0).astype(int)


def _clean_id_series(series):
    """Trim whitespace and strip trailing '.0' that pandas leaves on
    numeric-looking ID columns. Returns a string series."""
    s = series.astype(str).str.strip()
    return s.str.replace(r'\.0+$', '', regex=True)


def derive_member_account(df, config, has_header):
    """Return (member_only_str, full_account_str) honoring the three input
    modes captured by the wizard:

      * fixed_suffix : single column, last N chars are the account/suffix.
      * delimiter    : single column, member & account split by a delimiter.
      * split        : two columns; member col + account/suffix col.

    Falls back to the legacy `account_suffix_length` behavior when the
    `member_account` block is absent (older configs).
    """
    col_map = config['column_mappings']
    ma = config.get('member_account') or {}
    mode = ma.get('mode') or 'fixed_suffix'

    def _col(field):
        ref = col_map[field]
        return df[ref] if has_header else df.iloc[:, ref]

    member_raw = _clean_id_series(_col('member_number'))

    if mode == 'split' and col_map.get('loan_suffix'):
        suffix_raw = _clean_id_series(_col('loan_suffix'))
        # Pad suffix to 3 chars by default (match historical convention).
        pad_len = int(ma.get('suffix_length') or 3)
        suffix_padded = suffix_raw.str.zfill(pad_len) if pad_len > 0 else suffix_raw
        full = member_raw + suffix_padded
        return member_raw, full

    if mode == 'delimiter':
        delim = ma.get('delimiter') or '-'
        # Split once: left=member, right=account
        parts = member_raw.str.split(delim, n=1, expand=True)
        member_only = parts[0].fillna(member_raw)
        # Reconstruct the "full account" identifier without the delimiter so
        # the DB key matches the credit-pull / WARM convention of one
        # contiguous string.
        if parts.shape[1] > 1:
            account_part = parts[1].fillna('')
            full = member_only + account_part
        else:
            full = member_only
        return member_only, full

    # mode == 'fixed_suffix' (or unknown): use legacy suffix-length logic.
    suffix_length = int(
        ma.get('suffix_length')
        if ma.get('suffix_length') is not None
        else config.get('account_suffix_length', 0) or 0
    )
    if suffix_length > 0:
        member_only = member_raw.str[:-suffix_length]
    else:
        member_only = member_raw
    return member_only, member_raw


def _load_previous_fico(cu_name, current_snapshot):
    """Load current_fico_score from the most recent previous snapshot.

    Returns a dict mapping raw account number (str) -> previous current_fico_score,
    keyed at the loan level (not member level) so that each loan can carry forward
    its own score from the prior period.
    """
    with engine.connect() as conn:
        row = conn.execute(
            text("SELECT MAX(snapshot_date) FROM monthly_loan_data "
                 "WHERE credit_union = :cu AND snapshot_date < :snap"),
            {"cu": cu_name, "snap": current_snapshot}
        ).fetchone()
    prev_date = row[0] if row and row[0] else None
    if not prev_date:
        return {}

    prev_df = pd.read_sql(
        text("SELECT member_number, current_fico_score FROM monthly_loan_data "
             "WHERE credit_union = :cu AND snapshot_date = :snap"),
        engine, params={"cu": cu_name, "snap": str(prev_date)}
    )
    if prev_df.empty:
        return {}

    # Build lookup keyed by raw account number (includes suffix)
    scores = {}
    for _, r in prev_df.iterrows():
        acct = str(r['member_number']).strip()
        fico = int(r['current_fico_score'])
        if fico > 0:
            scores[acct] = fico
    return scores


def _excel_letter_to_index(letter: str) -> int:
    """'A' -> 0, 'Z' -> 25, 'AA' -> 26, etc. Raises ValueError on bad input."""
    s = str(letter).strip().upper()
    if not s or not s.isalpha():
        raise ValueError(f"Not a column letter: {letter!r}")
    n = 0
    for ch in s:
        n = n * 26 + (ord(ch) - ord('A') + 1)
    return n - 1


def _normalize_col_map_for_no_header(col_map):
    """When the loan extracts have no header row, the wizard stores mapping
    values like ``"col_A"`` (or already-integer positions on legacy configs).
    Return a copy with every value coerced to a 0-based integer position."""
    out = {}
    for field, val in (col_map or {}).items():
        if val is None or val == "":
            continue
        if isinstance(val, int):
            out[field] = val
            continue
        s = str(val).strip()
        if s.lower().startswith("col_"):
            try:
                out[field] = _excel_letter_to_index(s[4:])
                continue
            except ValueError:
                pass
        try:
            out[field] = int(s)
        except ValueError:
            try:
                out[field] = _excel_letter_to_index(s)
            except ValueError:
                # Leave as-is; will fail loudly downstream with a clear msg.
                out[field] = val
    return out


def import_file(file_path, config, snapshot_date, credit_pull_scores=None,
                warm_acct_scores=None, pull_as_of=None):
    """Import a single file using the client config. Returns count of imported rows."""
    cu_name = config['credit_union']
    col_map = config['column_mappings']
    has_header = config.get('has_header', True)
    suffix_length = config.get('account_suffix_length', 0)

    # When the wizard set a "fixed loan pool code for every row" on this
    # file, the loan_pool_code column is optional — we synthesize a
    # constant series for it below.
    static_pool_code = (col_map.get('loan_pool_code_static') or '').strip()

    ext = os.path.splitext(file_path)[1].lower()
    if has_header:
        # header_row is 1-indexed; pandas wants 0-indexed. 0/1/missing
        # → default header=0. Required for AIRES-style extracts whose
        # real headers live on row 2 (row 1 = column position numbers).
        try:
            hr_cfg = int(config.get('header_row') or 0)
        except (TypeError, ValueError):
            hr_cfg = 0
        pd_header = hr_cfg - 1 if hr_cfg > 1 else 0
        if ext == '.csv':
            df = pd.read_csv(file_path, header=pd_header)
        else:
            df = pd.read_excel(file_path, header=pd_header)
        df.columns = [str(c).strip() for c in df.columns]

        required = ['member_number', 'current_balance']
        if not static_pool_code:
            required.append('loan_pool_code')
        for field in required:
            src_col = col_map.get(field)
            if not src_col or src_col not in df.columns:
                raise ValueError(f"Required column '{field}' mapped to '{src_col}' not found. "
                                 f"Available: {list(df.columns)[:20]}")
    else:
        if ext == '.csv':
            df = pd.read_csv(file_path, header=None)
        else:
            df = pd.read_excel(file_path, header=None)

        # Wizard stores mapping values as "col_A"/"col_B" strings; the
        # importer needs 0-based integer positions when reading without
        # headers. Translate once up-front and rebind the local for the rest
        # of the function (and for derive_member_account, which reads
        # col_map via config).
        col_map = _normalize_col_map_for_no_header(col_map)
        config = dict(config)
        config['column_mappings'] = col_map

        required = ['member_number', 'current_balance']
        if not static_pool_code:
            required.append('loan_pool_code')
        for field in required:
            pos = col_map.get(field)
            if not isinstance(pos, int) or pos < 0 or pos >= len(df.columns):
                raise ValueError(f"Required column '{field}' mapped to position {pos!r} "
                                 f"but file has only {len(df.columns)} columns")

    # Build the output DataFrame — access by name or position index
    def col(field):
        if field == 'loan_pool_code' and static_pool_code:
            return pd.Series([static_pool_code] * len(df), index=df.index)
        return df[col_map[field]] if has_header else df.iloc[:, col_map[field]]

    # AIRES file provides the original FICO score (optional)
    if col_map.get('original_fico_score') or 'original_fico_score' in col_map:
        try:
            original_fico = pd.to_numeric(col('original_fico_score'), errors='coerce').fillna(0).astype(int)
        except KeyError:
            original_fico = pd.Series([0] * len(df), index=df.index, dtype=int)
    else:
        original_fico = pd.Series([0] * len(df), index=df.index, dtype=int)

    # Extract member-only & full-account identifiers honoring the wizard's
    # member/account format selection (fixed-suffix / delimiter / split).
    member_only_str, full_account_str = derive_member_account(df, config, has_header)
    raw_account = full_account_str           # full account string (DB key)
    raw_acct_str = full_account_str          # alias used downstream
    # member_numbers is used as a join key against credit_pull_scores (int keys).
    member_numbers = pd.to_numeric(member_only_str, errors='coerce').fillna(0).astype(int)

    # ----------------------------------------------------------------------
    # Original-FICO baseline fallback (wizard's "Original Score Baseline"
    # step). Fills in original_fico for loans whose extract doesn't carry
    # the original score (e.g. VISA / credit-card files). Two lookup keys
    # are tried, in order of specificity:
    #   1. full account string (member + suffix) — exact loan match.
    #   2. member-only string  — member-level fallback when no suffix was
    #      mapped on the baseline file.
    # Only loans with original_fico == 0 are touched; existing scores from
    # the extract are preserved.
    # ----------------------------------------------------------------------
    baseline_cfg = config.get('original_fico_baseline') or {}
    baseline_rows = baseline_cfg.get('rows') or []
    if baseline_rows:
        full_lookup: dict[str, int] = {}
        member_lookup: dict[str, int] = {}
        for r in baseline_rows:
            mem = str(r.get('member') or '').strip()
            if not mem:
                continue
            try:
                score = int(r.get('score') or 0)
            except (TypeError, ValueError):
                continue
            if score <= 0:
                continue
            suf = str(r.get('suffix') or '').strip()
            if suf:
                # Match the importer's convention of zero-padding suffixes.
                ma_b = config.get('member_account') or {}
                pad_len = int(ma_b.get('suffix_length') or 0) or len(suf)
                if pad_len > 0 and suf.isdigit():
                    suf_padded = suf.zfill(pad_len)
                else:
                    suf_padded = suf
                full_lookup[mem + suf_padded] = score
                full_lookup[mem + suf] = score  # also key as-given
            else:
                member_lookup[mem] = score
        if full_lookup or member_lookup:
            missing_mask = (original_fico == 0)
            # Optional pool scope: only fill within the listed loan_pool
            # value(s). Empty/missing list = apply globally.
            scoped_pools = [
                str(p).strip() for p in (baseline_cfg.get('pools') or [])
                if str(p).strip()
            ]
            if scoped_pools:
                # Derive the mapped pool name the same way clean_data does
                # later on (map raw loan_pool_code -> pool_map names). Use
                # the static value when the file carries no code column
                # (e.g. credit-card files with loan_pool_code_static set).
                try:
                    if static_pool_code:
                        raw_code_series = pd.Series(
                            [static_pool_code] * len(df), index=df.index
                        )
                    else:
                        raw_code_series = col('loan_pool_code')
                    pool_series = map_pool_codes(
                        raw_code_series, config
                    ).astype(str).str.strip()
                except KeyError:
                    pool_series = pd.Series([''] * len(df), index=df.index)
                pool_mask = pool_series.isin(scoped_pools)
                missing_mask = missing_mask & pool_mask
            if missing_mask.any():
                # Try full-account match first.
                filled = pd.Series(
                    [0] * len(df), index=df.index, dtype='Int64'
                )
                if full_lookup:
                    filled = raw_acct_str.map(full_lookup).astype('Int64')
                # Member-only fallback for still-unmatched rows.
                if member_lookup:
                    member_filled = member_only_str.map(member_lookup).astype('Int64')
                    filled = filled.fillna(member_filled)
                apply_mask = missing_mask & filled.notna()
                if apply_mask.any():
                    original_fico = original_fico.where(
                        ~apply_mask, filled
                    ).fillna(0).astype(int)
                    scope_msg = (
                        f" (pool scope: {', '.join(scoped_pools)})"
                        if scoped_pools else ""
                    )
                    print(
                        f"    Original-FICO baseline: filled {int(apply_mask.sum())} "
                        f"loan(s) from {len(baseline_rows)} baseline row(s)"
                        f"{scope_msg}"
                    )

    # Current FICO priority:
    #   1. WARM "All Loans" per-loan scores (authoritative source matching WARM's final scores)
    #   2. Credit pull (member-level) for loans not in WARM
    #   3. Previous snapshot per-loan carry-forward
    #   4. Original FICO
    # raw_acct_str is set above (alias of full_account_str)

    if warm_acct_scores:
        # Primary: WARM per-loan scores (exactly matches WARM's computed current scores)
        current_fico = raw_acct_str.map(warm_acct_scores)
        warm_matched = current_fico.notna().sum()
        unmatched = current_fico.isna().sum()

        # Secondary: credit pull (member-level) for loans not in WARM
        cp_filled = 0
        if unmatched > 0 and credit_pull_scores:
            cp_mapped = member_numbers.map(credit_pull_scores)
            current_fico = current_fico.fillna(cp_mapped)
            cp_filled = current_fico.notna().sum() - warm_matched
            unmatched = current_fico.isna().sum()

        parts = [f"WARM per-loan matched: {warm_matched}"]
        if cp_filled > 0:
            parts.append(f"credit pull: {cp_filled}")
        if unmatched > 0:
            parts.append(f"fallback to original: {unmatched}")
        print(f"    {', '.join(parts)}")

        # Final fallback to original score
        current_fico = current_fico.fillna(original_fico).astype(int)

    elif credit_pull_scores:
        current_fico = member_numbers.map(credit_pull_scores)
        matched = current_fico.notna().sum()
        unmatched = current_fico.isna().sum()

        if unmatched > 0:
            prev_scores = _load_previous_fico(cu_name, snapshot_date)
            if prev_scores:
                prev_mapped = raw_acct_str.map(prev_scores)
                prev_filled_count = prev_mapped.notna().sum()
                current_fico = current_fico.fillna(prev_mapped)
                unmatched = current_fico.isna().sum()
            else:
                prev_filled_count = 0

            parts = [f"Credit pull matched: {matched}"]
            if prev_filled_count > 0:
                parts.append(f"previous snapshot: {prev_filled_count}")
            if unmatched > 0:
                parts.append(f"fallback to original: {unmatched}")
            print(f"    {', '.join(parts)}")
        else:
            print(f"    Credit pull matched: {matched}")

        # Final fallback to original score
        current_fico = current_fico.fillna(original_fico).astype(int)
    else:
        current_fico = original_fico.copy()
        print(f"    WARNING: No credit pull data — current score = original score")

    # ----------------------------------------------------------------------
    # Original-score fallback for loans newer than the credit pull.
    # When the wizard's `prefer_original_for_new_loans` flag is set, any loan
    # whose open_date is after the credit-pull as-of date gets its current
    # score replaced with the loan-file's original score, on the theory that
    # the original was pulled at origination — i.e. more recent than the
    # bureau-wide pull. Without an open_date column or a known pull date this
    # block is a no-op (the bureau scores already loaded above stand).
    # ----------------------------------------------------------------------
    cp_cfg = config.get('credit_pull') or {}
    if cp_cfg.get('prefer_original_for_new_loans') and pull_as_of is not None:
        open_date_col = (config.get('column_mappings') or {}).get('open_date')
        if open_date_col is not None and (
            (has_header and open_date_col in df.columns)
            or (not has_header and isinstance(open_date_col, int)
                and open_date_col < len(df.columns))
        ):
            open_date_series = pd.to_datetime(
                df[open_date_col] if has_header else df.iloc[:, open_date_col],
                errors='coerce',
            )
            newer_mask = open_date_series.notna() & (open_date_series > pd.Timestamp(pull_as_of))
            if newer_mask.any():
                # Per-member: pick the original score from the most-recently-opened
                # loan and apply it to every loan that member has. This matches
                # the user's intent that all of a member's loans share one
                # current score derived from their freshest origination data.
                tmp = pd.DataFrame({
                    'member': member_numbers,
                    'open_date': open_date_series,
                    'orig_fico': original_fico,
                })
                # Only consider loans with a positive original score; zeros
                # represent missing data and shouldn't override real scores.
                tmp_valid = tmp[tmp['orig_fico'] > 0]
                if not tmp_valid.empty:
                    idx_max = tmp_valid.groupby('member')['open_date'].idxmax()
                    member_to_latest = (
                        tmp_valid.loc[idx_max].set_index('member')['orig_fico']
                    )
                    replacement = member_numbers.map(member_to_latest)
                    apply_mask = newer_mask & replacement.notna()
                    if apply_mask.any():
                        current_fico = current_fico.where(
                            ~apply_mask, replacement.astype('Int64')
                        )
                        # Coerce back to int after the masked update.
                        current_fico = current_fico.fillna(0).astype(int)
                        print(
                            f"    Original-score fallback: {int(apply_mask.sum())} "
                            f"loans newer than {pd.Timestamp(pull_as_of).date()} "
                            f"now use member's most-recent original score"
                        )

    clean_data = pd.DataFrame({
        'credit_union': cu_name,
        'snapshot_date': snapshot_date,
        'member_number': raw_account,
        'current_balance': clean_balance(
            col('current_balance'), config.get('balance_format', {})
        ),
        'current_fico_score': current_fico,
        'original_fico_score': original_fico,
        'loan_pool': map_pool_codes(col('loan_pool_code'), config),
    })

    # When original FICO is 0 but current is known, treat as unchanged (WARM convention)
    mask = (clean_data['original_fico_score'] == 0) & (clean_data['current_fico_score'] > 0)
    if mask.any():
        clean_data.loc[mask, 'original_fico_score'] = clean_data.loc[mask, 'current_fico_score']
        print(f"    Original FICO gap-fill: {mask.sum()} loans set original = current")

    clean_data = clean_data.dropna(subset=['current_balance'])
    clean_data = clean_data[clean_data['current_balance'] > 0]

    if len(clean_data) == 0:
        return 0

    with engine.begin() as conn:
        # Scope the pre-insert delete to only the pools represented in this
        # file. With multi-file imports (e.g. AIRES extract + a separate
        # VISA/credit-card extract for the same snapshot), an unconditional
        # delete by (cu, snapshot_date) would have the second file wipe out
        # the first file's rows. Deleting only the pools we're about to
        # re-insert keeps each file's pools refreshable independently.
        pools_in_file = sorted({str(p) for p in clean_data['loan_pool'].dropna().unique()})
        if pools_in_file:
            conn.execute(
                text(
                    "DELETE FROM monthly_loan_data "
                    "WHERE credit_union = :cu "
                    "AND snapshot_date = :sd "
                    "AND loan_pool = ANY(:pools)"
                ),
                {"cu": cu_name, "sd": snapshot_date, "pools": pools_in_file},
            )
        clean_data.to_sql('monthly_loan_data', conn, if_exists='append', index=False)

    return len(clean_data)


def process_client(client_name, specific_file=None):
    """Process all matching files for a client."""
    config = load_client_config(client_name)
    cu_name = config['credit_union']
    file_pattern_re = re.compile(config['file_pattern'], re.IGNORECASE)

    # Per-file loan-data extracts (wizard "Column Mappings" step). When
    # present, each entry overrides ``column_mappings`` / ``member_account``
    # / ``has_header`` / ``account_suffix_length`` for files whose name
    # matches the extract's own ``file_pattern``. Falls back to the
    # top-level mapping when no per-file pattern matches.
    extracts_raw = config.get('loan_data_extracts') or []
    extracts: list[tuple[re.Pattern, dict]] = []
    for e in extracts_raw:
        pat = (e or {}).get('file_pattern') or ''
        if not pat:
            continue
        try:
            extracts.append((re.compile(pat, re.IGNORECASE), e))
        except re.error as exc:
            print(f"  WARNING: loan_data_extracts entry {e.get('label','?')!r} "
                  f"has invalid file_pattern {pat!r}: {exc}. Skipping.")

    # Optional custom loan source folder (absolute or relative), useful for external client folders.
    configured_loan_folder = config.get('loan_file_folder')
    if configured_loan_folder:
        scan_folder = resolve_path(configured_loan_folder)
        if not os.path.isdir(scan_folder):
            raise FileNotFoundError(f"Loan file folder not found: {scan_folder}")
    else:
        # Look in per-client subfolder first, then fallback to main Raw_Uploads
        client_upload = os.path.join(UPLOAD_FOLDER, client_name)
        if os.path.isdir(client_upload):
            scan_folder = client_upload
        else:
            scan_folder = UPLOAD_FOLDER

    recursive_scan = bool(config.get('loan_file_recursive', False))
    archive_imported = bool(config.get('archive_imported_files', True))

    client_archive = None
    if archive_imported:
        archive_dir = config.get('archive_directory')
        client_archive = resolve_path(archive_dir) if archive_dir else os.path.join(ARCHIVE_FOLDER, client_name)
        os.makedirs(client_archive, exist_ok=True)

    print(f"\n{'='*60}")
    print(f"Processing: {cu_name}")
    print(f"Scanning: {scan_folder}")
    print(f"Recursive scan: {'Yes' if recursive_scan else 'No'}")
    print(f"Archive imported files: {'Yes' if archive_imported else 'No'}")
    print(f"{'='*60}")

    # Load credit pull scores (current FICO) before processing loan files
    credit_pull_scores, warm_acct_scores, pull_as_of = load_credit_pull_scores(config)
    if pull_as_of is not None:
        print(f"  Credit-pull as-of date: {pd.Timestamp(pull_as_of).date()}")

    files_to_process = []
    if recursive_scan:
        for root, _, files in os.walk(scan_folder):
            for filename in files:
                files_to_process.append((root, filename))
    else:
        for filename in os.listdir(scan_folder):
            files_to_process.append((scan_folder, filename))

    files_to_process.sort(key=lambda x: os.path.relpath(os.path.join(x[0], x[1]), scan_folder).lower())

    files_processed = 0
    for root, filename in files_to_process:
        file_path = os.path.join(root, filename)
        if not os.path.isfile(file_path):
            continue
        if filename.startswith("~$"):
            continue

        relative_file = os.path.relpath(file_path, scan_folder)
        if specific_file and filename != specific_file and relative_file != specific_file:
            continue

        # Route this file to a per-file extract (column_mappings,
        # member_account, has_header overrides) when one matches. Falls
        # back to the top-level mapping for back-compat.
        per_file_cfg = config
        matched_extract = None
        for pat, extract in extracts:
            if pat.search(filename) or pat.search(relative_file):
                matched_extract = extract
                break
        if matched_extract is not None:
            per_file_cfg = dict(config)
            per_file_cfg['column_mappings'] = dict(
                matched_extract.get('column_mappings') or {}
            )
            ma_over = matched_extract.get('member_account')
            if ma_over:
                per_file_cfg['member_account'] = dict(ma_over)
                if (ma_over.get('mode') == 'fixed_suffix'
                        and ma_over.get('suffix_length') is not None):
                    per_file_cfg['account_suffix_length'] = int(
                        ma_over.get('suffix_length') or 0
                    )
            if 'has_header' in matched_extract:
                per_file_cfg['has_header'] = bool(
                    matched_extract.get('has_header')
                )
            # Per-file header_row override (1-indexed). Lets AIRES-style
            # extracts (row 1 = position numbers, row 2 = real headers)
            # coexist with conventional extracts in the same client.
            if 'header_row' in matched_extract:
                try:
                    per_file_cfg['header_row'] = int(
                        matched_extract.get('header_row') or 0
                    )
                except (TypeError, ValueError):
                    per_file_cfg['header_row'] = 0
            label_txt = matched_extract.get('label') or '(unlabeled)'
            print(f"    Using extract mapping: {label_txt}")
        elif extracts:
            # We have per-file extracts configured but none match this
            # file. Top-level pattern is the global catch-all; if the
            # file matched that we'd still be here. Skip with a warning
            # rather than risk mis-mapping with the first extract's
            # columns.
            if not file_pattern_re.search(filename):
                continue
            print(f"    WARNING: {filename} matched top-level file_pattern but "
                  "no loan_data_extracts entry. Using top-level mapping as "
                  "fallback.")

        if not file_pattern_re.search(filename) and matched_extract is None:
            continue

        print(f"\n  File: {relative_file}")

        date_source = relative_file if str(config.get('date_source', 'filename')).lower() == 'path' else filename
        snapshot_date = extract_snapshot_date(date_source, config)
        if not snapshot_date and date_source != relative_file:
            # Fallback: allow a filename regex to match a dated folder segment in recursive paths.
            snapshot_date = extract_snapshot_date(relative_file, config)
        if not snapshot_date:
            print(f"    SKIPPED: Could not extract date from filename")
            continue

        try:
            count = import_file(file_path, per_file_cfg, snapshot_date, credit_pull_scores, warm_acct_scores, pull_as_of)
            if count > 0:
                if archive_imported and client_archive:
                    archive_target = os.path.join(client_archive, relative_file)
                    os.makedirs(os.path.dirname(archive_target), exist_ok=True)
                    shutil.move(file_path, archive_target)
                print(f"    SUCCESS: Imported {count} loans for {snapshot_date}")
                log_data_import(client_name, cu_name, file_path, count, success=True)
                files_processed += 1
            else:
                print(f"    WARNING: No valid loan records found")
                log_data_import(client_name, cu_name, file_path, 0, success=False)
        except Exception as e:
            print(f"    ERROR: {e}")
            log_data_import(client_name, cu_name, file_path, 0, success=False)

    if files_processed == 0:
        print(f"\n  No new files found to import.")
    else:
        print(f"\n  Imported {files_processed} file(s).")

    return files_processed


def main():
    parser = argparse.ArgumentParser(description="Import loan data for CECL analysis")
    parser.add_argument('--client', help='Client config name (e.g., "ontario")')
    parser.add_argument('--file', help='Specific filename to import')
    parser.add_argument('--all', action='store_true', help='Process all configured clients')
    parser.add_argument('--list', action='store_true', help='List available client configs')
    args = parser.parse_args()

    if args.list:
        print("Available clients:")
        for c in list_clients():
            cfg = load_client_config(c)
            print(f"  {c:20s} -> {cfg['credit_union']}")
        return

    if args.all:
        log_session_start('import_data.py', '--all')
        for client_name in list_clients():
            process_client(client_name)
        log_session_end('import_data.py')
    elif args.client:
        log_session_start('import_data.py', f'--client {args.client} --file={args.file}')
        process_client(args.client, args.file)
        log_session_end('import_data.py')
    else:
        parser.print_help()
        print("\nAvailable clients:", ', '.join(list_clients()))


if __name__ == '__main__':
    main()