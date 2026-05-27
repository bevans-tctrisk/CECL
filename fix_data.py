from dotenv import load_dotenv
import os
load_dotenv()
from sqlalchemy import create_engine, text

engine = create_engine(os.getenv('DATABASE_URL'))
with engine.begin() as conn:
    r = conn.execute(
        text("UPDATE monthly_loan_data SET credit_union = :cu, snapshot_date = :sd WHERE credit_union IS NULL AND snapshot_date IS NULL"),
        {"cu": "Ontario Public Employees FCU", "sd": "2025-12-31"}
    )
    print(f"Updated {r.rowcount} rows")

# Verify
with engine.connect() as conn:
    r = conn.execute(text("SELECT DISTINCT credit_union, snapshot_date FROM monthly_loan_data"))
    for row in r:
        print(f"  {row[0]} | {row[1]}")
    r = conn.execute(text("SELECT count(*) FROM vw_cecl_calculations WHERE credit_union IS NOT NULL"))
    print(f"View now returns {r.fetchone()[0]} rows with data")
