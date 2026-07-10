"""Consolidate Mountain CU's monthly 'Loan Charge Offs - Recoveries' files into a
single worksheet (one header row) so the CU can validate the charge-off /
recovery data that feeds the report."""
import os, re
import pandas as pd

DD = r'\\EgnyteDrive\tctconsult\Shared\TCT Files\CECL - CM Files\Raw_Uploads\mountain_cu'
OUT = r'\\EgnyteDrive\tctconsult\Shared\TCT Files\CECL - CM Files\Reports\Mountain_CU_ChargeOff_Recovery_Consolidated.xlsx'

MON = {'jan':1,'feb':2,'mar':3,'apr':4,'may':5,'jun':6,'jul':7,'aug':8,
       'sep':9,'oct':10,'nov':11,'dec':12}
KEEP = ['Charge Off Date', 'Loan Suffix', 'Account', 'Account Number',
        'Loan Type', 'Charge Off Amount', 'Recovery Amount']

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
    # drop fully-blank rows (no amounts and no account)
    df = df[~(df.get('Charge Off Amount').isna()
              & df.get('Recovery Amount').isna()
              & df.get('Account Number').isna())]
    df.insert(0, 'Source File', f)
    df.insert(1, 'Period', period_from_name(f))
    frames.append(df)

allrows = pd.concat(frames, ignore_index=True)
allrows = allrows.sort_values(['Period', 'Source File']).reset_index(drop=True)

with pd.ExcelWriter(OUT, engine='openpyxl') as xw:
    allrows.to_excel(xw, sheet_name='CO_Recovery', index=False)

print('WROTE', OUT)
print('files consolidated:', len(files))
print('total rows:', len(allrows))
co = pd.to_numeric(allrows['Charge Off Amount'], errors='coerce').fillna(0).sum()
rc = pd.to_numeric(allrows['Recovery Amount'], errors='coerce').fillna(0).sum()
print('GRAND Charge Off Amount: ${:,.2f}'.format(co))
print('GRAND Recovery Amount:   ${:,.2f}'.format(rc))
print('\nby period:')
g = allrows.copy()
g['CO'] = pd.to_numeric(g['Charge Off Amount'], errors='coerce').fillna(0)
g['RC'] = pd.to_numeric(g['Recovery Amount'], errors='coerce').fillna(0)
for per, sub in g.groupby('Period'):
    print('  {:<8} rows={:>4}  CO ${:>13,.2f}  Recovery ${:>13,.2f}'.format(
        per or '(none)', len(sub), sub['CO'].sum(), sub['RC'].sum()))
