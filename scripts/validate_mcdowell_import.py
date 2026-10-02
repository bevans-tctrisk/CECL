import os
os.environ['CECL_WORKSPACE_ROOT'] = r'Z:\Shared\TCT Files\CECL - CM Files'
import cecl_credentials, sqlalchemy as sa, pandas as pd
eng = sa.create_engine(cecl_credentials.get_database_url())
CU = 'McDowell Cornerstone CU'

for snap, gl in [('2025-12-31', 22537352.48), ('2026-03-31', 22218927.00)]:
    df = pd.read_sql(sa.text(
        "SELECT loan_pool, COUNT(*) n, SUM(current_balance) bal FROM monthly_loan_data "
        "WHERE credit_union=:c AND snapshot_date=:s GROUP BY loan_pool ORDER BY bal DESC"),
        eng, params={'c': CU, 's': snap})
    incl = df[~df['loan_pool'].isin(['Ignore','Exclude'])]
    tot = incl['bal'].sum()
    print(f"\n=== {snap} ===")
    for _, r in df.iterrows():
        print(f"   {r['loan_pool']:<22} n={int(r['n']):>4}  ${r['bal']:,.2f}")
    print(f"   INCLUDED TOTAL: ${tot:,.2f}  | GL Loans To Members: ${gl:,.2f}  | diff ${tot-gl:,.2f} ({(tot-gl)/gl*100:+.3f}%)")

# snapshot range + fico coverage on Mar 2026
snaps = pd.read_sql(sa.text("SELECT DISTINCT snapshot_date FROM monthly_loan_data WHERE credit_union=:c ORDER BY snapshot_date"), eng, params={'c': CU})
print("\nsnapshots:", list(snaps['snapshot_date'].astype(str)))
fc = pd.read_sql(sa.text("SELECT loan_pool, SUM(CASE WHEN current_fico_score>0 THEN 1 ELSE 0 END) scored, COUNT(*) n FROM monthly_loan_data WHERE credit_union=:c AND snapshot_date='2026-03-31' AND loan_pool NOT IN ('Ignore','Exclude') GROUP BY loan_pool"), eng, params={'c': CU})
print("FICO coverage Mar2026:")
for _, r in fc.iterrows():
    print(f"   {r['loan_pool']:<22} {int(r['scored'])}/{int(r['n'])} = {r['scored']/r['n']*100:.0f}%")
