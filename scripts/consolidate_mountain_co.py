"""Consolidate Mountain CU's monthly 'Loan Charge Offs - Recoveries' files into a
single worksheet (one header row) plus pool/loan-type breakdowns so the CU can
validate the charge-off / recovery data that feeds the report's CO-Recov-DQ tab."""
import os, re
import yaml
import pandas as pd

DD = r'\\EgnyteDrive\tctconsult\Shared\TCT Files\CECL - CM Files\Raw_Uploads\mountain_cu'
CFG = r'\\EgnyteDrive\tctconsult\Shared\TCT Files\CECL - CM Files\client_configs\mountain_cu.yaml'
OUT = r'\\EgnyteDrive\tctconsult\Shared\TCT Files\CECL - CM Files\Reports\Mountain_CU_ChargeOff_Recovery_Consolidated.xlsx'

MON = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,'jul':7,'aug':8,
       'sep':9,'oct':10,'nov':11,'dec':12}
KEEP = ['Charge Off Date', 'Loan Suffix', 'Account', 'Account Number',
        'Loan Type', 'Charge Off Amount', 'Recovery Amount']

cfg = yaml.safe_load(open(CFG, encoding='utf-8'))
POOL_MAP = cfg.get('pool_map', {}) or {}
DEFAULT_POOL = cfg.get('default_pool', 'Ignore')

def code_to_pool(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return DEFAULT_POOL
    s = str(v).strip()
    cands = [s]
    try:
        i = int(float(s)); cands += [str(i), str(i).zfill(4)]
    except (ValueError, TypeError):
        pass
    for c in cands:
        if c in POOL_MAP:
            return POOL_MAP[c]
    return DEFAULT_POOL

def period_from_name(fn):
    m = re.search(r'([A-Za-z]{3})\s*[_\-]?\s*(\d{2})\b', fn)
    if not m:
        return ''
    mo = MON.get(m.group(1).lower())
    return '20%s-%02d' % (m.group(2), mo) if mo else ''

files = sorted(f for f in os.listdir(DD)
               if f.lower().startswith('loan') and 'charge' in f.lower()
               and f.lower().endswith(('.xlsx', '.xls')))

frames = []
for f in files:
    df = pd.read_excel(os.path.join(DD, f), sheet_name=0)
    df.columns = [str(c).strip() for c in df.columns]          # 'Account ' -> 'Account'
    df = df[[c for c in KEEP if c in df.columns]].copy()
    df = df[~(df.get('Charge Off Amount').isna()
              & df.get('Recovery Amount').isna()
              & df.get('Account Number').isna())]
    df.insert(0, 'Source File', f)
    df.insert(1, 'Period', period_from_name(f))
    frames.append(df)

allrows = pd.concat(frames, ignore_index=True)
allrows['Pool'] = allrows['Loan Type'].map(code_to_pool)
allrows['CO'] = pd.to_numeric(allrows['Charge Off Amount'], errors='coerce').fillna(0)
allrows['RC'] = pd.to_numeric(allrows['Recovery Amount'], errors='coerce').fillna(0)
allrows['Year'] = allrows['Period'].str[:4]
# Rows with no Loan Type code are per-file TOTALS/summary rows (verified: each
# equals its file's detail sum). The report excludes them (blank code -> Ignore)
# so they are NOT double-counted. Flag them explicitly for the CU.
is_subtotal = allrows['Loan Type'].isna()
allrows['Row Type'] = ['FILE SUBTOTAL (excluded)' if b else 'DETAIL' for b in is_subtotal]
allrows.loc[is_subtotal, 'Pool'] = 'FILE SUBTOTAL (excluded)'

main_cols = ['Source File', 'Period', 'Row Type', 'Charge Off Date', 'Loan Suffix', 'Account',
             'Account Number', 'Loan Type', 'Pool', 'Charge Off Amount', 'Recovery Amount']
main = allrows[main_cols].sort_values(['Period', 'Row Type', 'Pool', 'Source File']).reset_index(drop=True)

def pivot(val):
    p = allrows.pivot_table(index='Pool', columns='Year', values=val,
                            aggfunc='sum', fill_value=0.0, margins=True, margins_name='Total')
    return p.round(2)

co_pool = pivot('CO')
rc_pool = pivot('RC')

by_type = (allrows.assign(**{'Loan Type': allrows['Loan Type'].where(
               allrows['Loan Type'].notna(), '(blank / no code)')})
           .groupby(['Loan Type', 'Pool'], dropna=False)
           .agg(Rows=('CO', 'size'), Charge_Off=('CO', 'sum'), Recovery=('RC', 'sum'))
           .reset_index().sort_values(['Pool', 'Loan Type']))
by_type['Charge_Off'] = by_type['Charge_Off'].round(2)
by_type['Recovery'] = by_type['Recovery'].round(2)

# Reconciliation: naive sum (incl. per-file subtotal rows) vs the report's
# detail-only figure. The subtotal rows would DOUBLE-COUNT if included.
excl = allrows[is_subtotal]
recon = pd.DataFrame({
    'Measure': ['Charge Offs', 'Recoveries'],
    'Naive sum (incl. file subtotal rows)': [allrows['CO'].sum(), allrows['RC'].sum()],
    'File subtotal rows (double-count, excluded)': [excl['CO'].sum(), excl['RC'].sum()],
    'Detail only = report figure': [allrows['CO'].sum() - excl['CO'].sum(),
                                    allrows['RC'].sum() - excl['RC'].sum()],
}).round(2)

blank_ct = int(is_subtotal.sum())

with pd.ExcelWriter(OUT, engine='openpyxl') as xw:
    main.to_excel(xw, sheet_name='CO_Recovery', index=False)
    recon.to_excel(xw, sheet_name='Reconciliation', index=False)
    co_pool.to_excel(xw, sheet_name='Charge Offs by Pool-Year')
    rc_pool.to_excel(xw, sheet_name='Recoveries by Pool-Year')
    by_type.to_excel(xw, sheet_name='By Loan Type', index=False)

print('WROTE', OUT)
print('files:', len(files), '| rows:', len(main), '| rows with BLANK Loan Type:', blank_ct)
print('\n=== Reconciliation (raw vs report) ===')
print(recon.to_string(index=False))
print('\n=== Charge Offs by Pool x Year ===')
print(co_pool.to_string())
print('\n=== Recoveries by Pool x Year ===')
print(rc_pool.to_string())
print('\n=== By Loan Type ===')
print(by_type.to_string(index=False))
