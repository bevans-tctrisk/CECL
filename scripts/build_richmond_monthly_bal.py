import os, yaml
os.environ['CECL_WORKSPACE_ROOT'] = r'Z:\Shared\TCT Files\CECL - CM Files'
WS = os.environ['CECL_WORKSPACE_ROOT']
import cecl_credentials, sqlalchemy as sa, pandas as pd
eng = sa.create_engine(cecl_credentials.get_database_url())
cu = 'Credit Union of Richmond'
SHORT = 'credit_union_of_richmond'

df = pd.read_sql(sa.text(
    "SELECT snapshot_date, loan_pool, SUM(current_balance) bal "
    "FROM monthly_loan_data WHERE credit_union=:c AND loan_pool NOT IN ('Ignore','Exclude') "
    "GROUP BY snapshot_date, loan_pool"), eng, params={'c': cu})
df['snapshot_date'] = pd.to_datetime(df['snapshot_date'])
piv = df.pivot(index='loan_pool', columns='snapshot_date', values='bal').fillna(0.0)
piv = piv.sort_index(axis=1)

POOL_ORDER = ['New Auto', 'Used Auto', 'CUDL Auto Loans', 'First Mortgage',
              'Second Mortgages and HELOC', 'Other Secured', 'Other Unsecured', 'Credit Cards']
piv = piv.reindex([p for p in POOL_ORDER if p in piv.index])

# write wide file: row1 header [Pool, date1, date2...]; col A pool, col B.. dates
out_dir = os.path.join(WS, 'Raw_Uploads', SHORT)
os.makedirs(out_dir, exist_ok=True)
out = os.path.join(out_dir, 'Monthly Loan Balances by Type.xlsx')
wide = piv.copy()
wide.columns = [c.strftime('%Y-%m-%d') for c in wide.columns]
wide.insert(0, 'Pool', wide.index)
wide.reset_index(drop=True, inplace=True)
with pd.ExcelWriter(out, engine='openpyxl') as xw:
    wide.to_excel(xw, sheet_name='Sheet1', index=False)
print('WROTE', out)
print('pools:', list(piv.index))
print('months:', len(piv.columns), piv.columns.min().date(), '->', piv.columns.max().date())

# wire into config
cfgp = os.path.join(WS, 'client_configs', SHORT + '.yaml')
cfg = yaml.safe_load(open(cfgp, encoding='utf-8'))
cfg['monthly_balance'] = {
    'source': 'single', 'sheet': 'Sheet1', 'header_row': 1,
    'pool_name_col': 'A', 'first_date_col': 'B',
    'filename': 'Monthly Loan Balances by Type.xlsx', 'saved_path': out,
    'pool_map': {p: p for p in piv.index},
}
with open(cfgp, 'w', encoding='utf-8') as f:
    yaml.safe_dump(cfg, f, sort_keys=False, allow_unicode=True)
print('config monthly_balance wired')
