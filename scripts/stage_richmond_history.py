"""Stage Richmond history: content-routed loan files (2024-2025) for import +
normalized CO/recovery files (2023-2025) for the report to read at run time.
Dates come from the FOLDER name (most reliable). CO files are rewritten to a
single canonical 6-column layout matching historical_file_formats in the config.
"""
import os, re, shutil, calendar
import pandas as pd

BASE = r'Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access\Clients\Credit Union of Richmond 66929\Portfolio Management\CECL Migration IDLR'
DST = r'Z:\Shared\TCT Files\CECL - CM Files\Raw_Uploads\credit_union_of_richmond'
os.makedirs(DST, exist_ok=True)

MONTHS = {'jan': 1, 'feb': 2, 'mar': 3, 'apr': 4, 'may': 5, 'jun': 6, 'june': 6,
          'jul': 7, 'july': 7, 'aug': 8, 'sep': 9, 'sept': 9, 'oct': 10, 'nov': 11, 'dec': 12}

def folder_date(name):
    m = re.match(r'([A-Za-z]+)\s+(20\d{2})', name.strip())
    if not m:
        return None
    mo = MONTHS.get(m.group(1)[:4].lower()) or MONTHS.get(m.group(1)[:3].lower())
    if not mo:
        return None
    yr = int(m.group(2))
    return yr, mo, calendar.monthrange(yr, mo)[1]

def read_head(path, n=3):
    try:
        return pd.read_excel(path, header=0, nrows=n)
    except Exception:
        return None

def classify(path):
    df = read_head(path)
    if df is None:
        return None, None
    cols = [str(c).strip().lower() for c in df.columns]
    cset = set(cols)
    def has(*names):
        return any(any(nm in c for c in cols) for nm in names)
    # CO/recovery: has charge off amount + recovery amount
    if has('charge off amount', 'charge-off amount', 'chargeoff amount') and has('recovery amount'):
        return 'CO', df
    # Credit cards: 81-col with Account Identifier -PK
    if has('account identifier') and has('account balance total amount'):
        return 'CC', df
    # AIRES v2 (Feb 2026+): named cols 'Current Balance' + 'FICO', no 'LOAN BALANCE'
    if has('account number') and has('current balance') and has('fico') and not has('loan balance'):
        return 'AIRESV2', df
    # AIRES loan (old): has Account Number + LOAN BALANCE/LOAN TYPE
    if has('account number') and (has('loan balance') or has('loan type')):
        return 'AIRES', df
    # Negative shares: has acct_type / current balance and small
    if has('acct_type') or (('current balance' in cset or has('current balance')) and df.shape[1] <= 12 and has('loan type code')):
        return 'NEGSH', df
    return None, df

def co_col(df, *names):
    for i, c in enumerate(df.columns):
        cl = str(c).strip().lower()
        if any(nm in cl for nm in names):
            return i
    return None

def normalize_co(src, dst):
    """Rewrite CO file to canonical 6 cols: acct,code,acct2,recovery,chargeoff,date."""
    df = pd.read_excel(src, header=0)
    ai = co_col(df, 'account #', 'account#', 'account number')
    ci = co_col(df, 'loan type code', 'loan type')
    ri = co_col(df, 'recovery amount')
    coi = co_col(df, 'charge off amount', 'charge-off amount', 'chargeoff amount')
    di = co_col(df, 'charge off date', 'charge-off date', 'chargeoff date')
    # leading unnamed member col (2025 layout): if col0 is unnamed numeric, prefer it as account
    acct = df.iloc[:, ai] if ai is not None else df.iloc[:, 0]
    if ai is not None and ai > 0 and str(df.columns[0]).lower().startswith('unnamed'):
        acct = df.iloc[:, 0]
    out = pd.DataFrame({
        'Account #': acct.values,
        'Loan Type Code': df.iloc[:, ci].values if ci is not None else '',
        'Account #2': acct.values,
        'Recovery Amount': df.iloc[:, ri].values if ri is not None else 0,
        'Charge off Amount': df.iloc[:, coi].values if coi is not None else 0,
        'Charge off Date': df.iloc[:, di].values if di is not None else pd.NaT,
    })
    out.to_excel(dst, index=False)

def stage():
    loan_years = {'2024', '2025', '2026'}
    co_years = {'2023', '2024', '2025', '2026'}
    skip_folders = {'June 2026'}  # incomplete data per user
    summary = []
    for yr in sorted(set(loan_years) | set(co_years)):
        yp = os.path.join(BASE, yr)
        if not os.path.isdir(yp):
            continue
        for mo in sorted(os.listdir(yp)):
            mp = os.path.join(yp, mo)
            if not os.path.isdir(mp):
                continue
            if mo in skip_folders:
                print('  SKIP folder (incomplete):', yr, mo); continue
            fd = folder_date(mo)
            if not fd:
                print('  SKIP folder (no date):', yr, mo); continue
            Y, M, D = fd
            tag = '%02d-%02d-%04d' % (M, D, Y)
            picks = {'AIRES': [], 'AIRESV2': [], 'CC': [], 'NEGSH': [], 'CO': []}
            for f in os.listdir(mp):
                if not f.lower().endswith(('.xlsx', '.xls')):
                    continue
                if 'investment' in f.lower() or 'cross reference' in f.lower():
                    continue
                kind, _ = classify(os.path.join(mp, f))
                if kind:
                    picks[kind].append(f)
            got = []
            # loan files -> only for loan_years, canonical copy
            if yr in loan_years:
                for kind, canon in [('AIRES', 'AIRES'), ('AIRESV2', 'AIRESV2'), ('CC', 'Credit Cards'), ('NEGSH', 'Negative Shares')]:
                    if picks[kind]:
                        src = os.path.join(mp, sorted(picks[kind], key=lambda x: os.path.getsize(os.path.join(mp, x)))[-1])
                        ext = os.path.splitext(src)[1].lower()
                        dst = os.path.join(DST, '%s %s%s' % (canon, tag, ext))
                        shutil.copy2(src, dst)
                        got.append(kind)
            # CO files -> for co_years, normalized
            if yr in co_years and picks['CO']:
                src = os.path.join(mp, sorted(picks['CO'], key=lambda x: os.path.getsize(os.path.join(mp, x)))[-1])
                dst = os.path.join(DST, 'Charge-Off and Recoveries %s.xlsx' % tag)
                try:
                    normalize_co(src, dst)
                    got.append('CO')
                except Exception as e:
                    print('  CO normalize FAIL', yr, mo, e)
            summary.append((yr, mo, tag, got))
    print('\n=== staging summary ===')
    for yr, mo, tag, got in summary:
        print('  %-4s %-10s %s  %s' % (yr, mo, tag, got))

stage()
