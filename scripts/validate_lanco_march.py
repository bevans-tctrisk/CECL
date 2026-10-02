import os, openpyxl
import cecl_credentials, sqlalchemy as sa, pandas as pd
eng = sa.create_engine(cecl_credentials.get_database_url()); CU = 'Lanco FCU'
df = pd.read_sql(sa.text(
    "SELECT loan_pool, COUNT(*) n, SUM(current_balance) bal FROM monthly_loan_data "
    "WHERE credit_union=:c AND snapshot_date='2026-03-31' GROUP BY loan_pool ORDER BY bal DESC"),
    eng, params={'c': CU})
incl = df[~df['loan_pool'].isin(['Ignore', 'Exclude'])]
print('=== Mar 2026 imported pools ===')
for _, r in df.iterrows():
    print('  {:<22} n={:>4}  ${:,.2f}'.format(r['loan_pool'], int(r['n']), r['bal']))
print('  INCLUDED TOTAL: ${:,.2f}'.format(incl['bal'].sum()))

mb = r'Z:\Shared\TCT Files\CECL - CM Files\Raw_Uploads\lanco_fcu\Monthly Loan Balances by Type.xlsx'
wb = openpyxl.load_workbook(mb, read_only=True, data_only=True); ws = wb['Sheet1']
rows = list(ws.iter_rows(values_only=True)); hdr = list(rows[0]); idx = hdr.index('2026-03-31')
tot = 0
print('=== monthly_balance 2026-03-31 (GL) ===')
for r in rows[1:]:
    if r[0]:
        print('  {:<22} ${:,.2f}'.format(r[0], r[idx] or 0)); tot += r[idx] or 0
print('  GL TOTAL: ${:,.2f}'.format(tot))

co = os.listdir(r'Z:\Shared\TCT Files\CECL - CM Files\Raw_Uploads\lanco_fcu')
cofiles = [f for f in co if 'charge' in f.lower()]
print('=== existing CO files in Raw_Uploads (last 6 of {}) ==='.format(len(cofiles)))
for f in sorted(cofiles)[-6:]:
    print('  ', f)
