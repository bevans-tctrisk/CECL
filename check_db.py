from dotenv import load_dotenv
import os
load_dotenv()
from sqlalchemy import create_engine, text

engine = create_engine(os.getenv('DATABASE_URL'))
with engine.connect() as conn:
    print("=== TABLES ===")
    r = conn.execute(text("SELECT table_name FROM information_schema.tables WHERE table_schema='public'"))
    for row in r:
        print(f"  {row[0]}")

    print("\n=== VIEWS ===")
    r = conn.execute(text("SELECT table_name FROM information_schema.views WHERE table_schema='public'"))
    for row in r:
        print(f"  {row[0]}")

    print("\n=== monthly_loan_data ===")
    r = conn.execute(text("SELECT count(*) FROM monthly_loan_data"))
    print(f"  Row count: {r.fetchone()[0]}")

    r = conn.execute(text("SELECT DISTINCT credit_union, snapshot_date FROM monthly_loan_data"))
    rows = r.fetchall()
    if rows:
        print("  CU / Date combos:")
        for row in rows:
            print(f"    {row[0]} | {row[1]}")
    else:
        print("  No data found!")

    print("\n=== vw_cecl_calculations ===")
    try:
        r = conn.execute(text("SELECT count(*) FROM vw_cecl_calculations"))
        print(f"  Row count: {r.fetchone()[0]}")
    except Exception as e:
        print(f"  Error: {e}")
