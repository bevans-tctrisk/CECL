"""Audit WSSC FCU pools: for the 2025-12-31 snapshot, report per-pool how many
loans carry a valid credit score vs none. Pools with ZERO scored balance should
NOT be broken out by grade (risk_rated: false / single line)."""
import os
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
from sqlalchemy import create_engine, text
from cecl_credentials import get_database_url

CU = "WSSC FCU"
SNAP = "2025-12-31"

eng = create_engine(get_database_url())
with eng.connect() as c:
    # confirm available snapshots
    snaps = [r[0] for r in c.execute(text(
        "SELECT DISTINCT snapshot_date FROM monthly_loan_data WHERE credit_union=:cu ORDER BY snapshot_date"), {"cu": CU})]
    print("Snapshots:", snaps)
    rows = c.execute(text("""
        SELECT loan_pool,
               COUNT(*)                                         AS n_loans,
               SUM(CASE WHEN original_fico_score > 0 THEN 1 ELSE 0 END) AS n_scored,
               SUM(current_balance)                             AS bal_total,
               SUM(CASE WHEN original_fico_score > 0 THEN current_balance ELSE 0 END) AS bal_scored
        FROM monthly_loan_data
        WHERE credit_union=:cu AND snapshot_date=:s
        GROUP BY loan_pool
        ORDER BY loan_pool
    """), {"cu": CU, "s": SNAP}).fetchall()

print(f"\nPer-pool credit-score coverage for {CU} @ {SNAP}:")
print(f"{'Pool':<28}{'#loans':>8}{'#scored':>9}{'bal_total':>16}{'bal_scored':>16}  {'grade?'}")
for pool, n, ns, bt, bs in rows:
    ns = ns or 0; bt = float(bt or 0); bs = float(bs or 0)
    grade = "SINGLE-LINE" if ns == 0 else "by-grade"
    print(f"{str(pool):<28}{n:>8}{ns:>9}{bt:>16,.2f}{bs:>16,.2f}  {grade}")
