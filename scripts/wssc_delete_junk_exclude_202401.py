"""Delete the single bogus 'Exclude' totals row in WSSC 2024-01-31 (member_number
IS NULL, ~$15.6M). The report already drops Exclude rows, so this is cosmetic DB
cleanup. Scoped, verified before/after."""
import os
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
from sqlalchemy import create_engine, text
from cecl_credentials import get_database_url
eng = create_engine(get_database_url())
CU = "WSSC FCU"

with eng.begin() as c:
    before = c.execute(text(
        "SELECT COUNT(*), COALESCE(SUM(current_balance),0) FROM monthly_loan_data "
        "WHERE credit_union=:c AND snapshot_date='2024-01-31' "
        "AND loan_pool IN ('Exclude','Ignore') AND member_number IS NULL"),
        {"c": CU}).fetchone()
    print(f"Matching junk rows: n={before[0]} total=${float(before[1]):,.2f}")
    res = c.execute(text(
        "DELETE FROM monthly_loan_data "
        "WHERE credit_union=:c AND snapshot_date='2024-01-31' "
        "AND loan_pool IN ('Exclude','Ignore') AND member_number IS NULL"),
        {"c": CU})
    print("Rows deleted:", res.rowcount)

with eng.connect() as c:
    tot = c.execute(text("SELECT COUNT(*), SUM(current_balance) FROM monthly_loan_data "
                         "WHERE credit_union=:c AND snapshot_date='2024-01-31'"),
                    {"c": CU}).fetchone()
    print(f"2024-01-31 now: rows={tot[0]} total=${float(tot[1] or 0):,.2f}")
