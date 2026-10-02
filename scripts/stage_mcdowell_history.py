"""Stage McDowell Cornerstone history for import.

Filenames drift month-to-month, so files are routed by HEADER CONTENT (not name)
and copied to canonical names with the FOLDER-derived month-end date:
  Aires MM-DD-YYYY.xlsx / Buy Side MM-DD-YYYY.xlsx / Negative Shares MM-DD-YYYY.xlsx
  Charge-Off and Recoveries MM-DD-YYYY.xlsx  (already the canonical 5-col layout)
Loan files staged for 2025-2026; CO/recovery files staged for 2024-2026.
"""
import os, re, shutil, calendar
import pandas as pd

BASE = (r'Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access\Clients'
        r'\McDowell Cornerstone CU 60149\Portfolio Management\CECL Migration IDLR')
DST = r'Z:\Shared\TCT Files\CECL - CM Files\Raw_Uploads\mcdowell_cornerstone_cu'
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
        return None
    cols = [str(c).strip().lower() for c in df.columns]
    def has(*names):
        return any(any(nm in c for c in cols) for nm in names)
    # CO/recovery: charge off amount + recovery amount columns
    if has('charge off amount', 'charge-off amount', 'chargeoff amount') and has('recovery amount'):
        return 'CO'
    # Buy Side (LoanStreet): 'LoanStreet ID' + 'Principal Owned at Period End'
    if has('loanstreet id') and has('principal owned'):
        return 'BUYSIDE'
    # AIRES core loan: Account + Loan Type Code + Current Balance + FICO
    if has('account') and has('loan type code') and has('current balance') and has('fico'):
        return 'AIRES'
    # Negative shares: ACCOUNT + BALANCE + Loan Type Code, small
    if has('account') and has('balance') and has('loan type code') and df.shape[1] <= 8:
        return 'NEGSH'
    return None

def stage():
    loan_years = {'2025', '2026'}
    co_years = {'2024', '2025', '2026'}
    summary = []
    for yr in sorted(co_years | loan_years):
        yp = os.path.join(BASE, yr)
        if not os.path.isdir(yp):
            continue
        for mo in sorted(os.listdir(yp)):
            mp = os.path.join(yp, mo)
            if not os.path.isdir(mp):
                continue
            fd = folder_date(mo)
            if not fd:
                print('  SKIP folder (no date):', yr, mo); continue
            Y, M, D = fd
            tag = '%02d-%02d-%04d' % (M, D, Y)
            picks = {'AIRES': [], 'BUYSIDE': [], 'NEGSH': [], 'CO': []}
            for f in os.listdir(mp):
                if not f.lower().endswith(('.xlsx', '.xls')):
                    continue
                kind = classify(os.path.join(mp, f))
                if kind:
                    picks[kind].append(f)
            got = []
            if yr in loan_years:
                for kind, canon in [('AIRES', 'Aires'), ('BUYSIDE', 'Buy Side'),
                                    ('NEGSH', 'Negative Shares')]:
                    if picks[kind]:
                        src = os.path.join(mp, sorted(picks[kind],
                              key=lambda x: os.path.getsize(os.path.join(mp, x)))[-1])
                        ext = os.path.splitext(src)[1].lower()
                        dst = os.path.join(DST, '%s %s%s' % (canon, tag, ext))
                        shutil.copy2(src, dst)
                        got.append(kind)
            if yr in co_years:
                for f in picks['CO']:
                    src = os.path.join(mp, f)
                    # unique-ish canonical name; keep both CU + LoanStreet CO files
                    lbl = 'LS ' if 'loan street' in f.lower() or 'loanstreet' in f.lower() else ''
                    dst = os.path.join(DST, 'Charge-Off and Recoveries %s%s.xlsx' % (lbl, tag))
                    shutil.copy2(src, dst)
                    got.append('CO' + (':LS' if lbl else ''))
            summary.append((yr, mo, tag, got))
    print('=== staging summary ===')
    for yr, mo, tag, got in summary:
        print('  %-4s %-10s %s  %s' % (yr, mo, tag, got))

if __name__ == '__main__':
    stage()
    print('\nfiles in DST:', len(os.listdir(DST)))
