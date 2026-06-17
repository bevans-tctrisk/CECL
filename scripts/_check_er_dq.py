"""Quick inspection of ER's DQ history table state."""
import sys
from pathlib import Path
sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from cecl_credentials import get_database_url
from sqlalchemy import create_engine, text

e = create_engine(get_database_url())
with e.connect() as c:
    print("=== Summary by CU (ILIKE '%emergency%') ===")
    r = c.execute(text("""
        SELECT cu, COUNT(*) AS n_rows, MIN(as_of_date) AS earliest, MAX(as_of_date) AS latest,
            SUM(CASE WHEN total_balance IS NOT NULL THEN 1 ELSE 0 END) AS has_bal,
            SUM(CASE WHEN total_balance > 0 THEN 1 ELSE 0 END) AS pos_bal,
            SUM(CASE WHEN dq_amount IS NOT NULL AND dq_amount > 0 THEN 1 ELSE 0 END) AS has_dq
        FROM loan_code_delinquency_history
        WHERE cu ILIKE '%emergency%'
        GROUP BY cu
    """))
    for row in r:
        print(row)

    print("\n=== Sample rows for 2025-12-31 (first 20) ===")
    r2 = c.execute(text("""
        SELECT as_of_date, loan_code, dq_amount, total_balance, dq_pct, source
        FROM loan_code_delinquency_history
        WHERE cu ILIKE '%emergency%' AND as_of_date='2025-12-31'
        ORDER BY loan_code LIMIT 20
    """))
    for row in r2:
        print(row)

    print("\n=== Distinct sources ===")
    r3 = c.execute(text("""
        SELECT source, COUNT(*) FROM loan_code_delinquency_history
        WHERE cu ILIKE '%emergency%' GROUP BY source ORDER BY source
    """))
    for row in r3:
        print(row)

    print("\n=== Distributed (loan_code_history) summary for ER ===")
    r4 = c.execute(text("""
        SELECT source, COUNT(*) AS n, MIN(as_of_date), MAX(as_of_date)
        FROM loan_code_history
        WHERE cu ILIKE '%emergency%' AND source LIKE '5300-distributed%%'
        GROUP BY source ORDER BY source LIMIT 5
    """))
    for row in r4:
        print(row)
