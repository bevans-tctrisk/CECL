import os
from sqlalchemy import create_engine, text
import cecl_credentials

eng = create_engine(cecl_credentials.get_database_url())
TARGET = {
    "Auto": 244727688.65, "Share Secured": 420414.07, "Unsecured": 17263009.43,
    "Real Estate": 66473681.21, "Credit Card": 17965019.15,
    "Off Books Real Estate": 70245795.42, "Business Credit Card": 246920.88,
    "Negative Share": 25955.05,
}
with eng.connect() as c:
    for snap in ("2025-11-30", "2026-02-28"):
        print(f"\n==== {snap}  (Erie FCU)")
        rows = c.execute(text(
            "SELECT loan_pool, COUNT(*) n, SUM(current_balance) bal "
            "FROM monthly_loan_data WHERE credit_union='Erie FCU' AND snapshot_date=:d "
            "GROUP BY loan_pool ORDER BY SUM(current_balance) DESC"), {"d": snap}).fetchall()
        tot = 0.0
        for pool, n, bal in rows:
            bal = float(bal or 0)
            tot += bal
            t = TARGET.get(pool)
            diff = f"  WARM={t:,.2f} diff={bal-t:+,.2f}" if (t and snap == '2025-11-30') else ""
            print(f"   {pool:26} {n:6d}  ${bal:,.2f}{diff}")
        print(f"   {'TOTAL':26} {'':6}  ${tot:,.2f}")
