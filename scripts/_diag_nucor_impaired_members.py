"""Diagnose Nucor impaired-loan member key mismatch (Turn 8)."""
import os
os.environ.setdefault('CECL_WORKSPACE_ROOT', r'Z:\Shared\TCT Files\CECL - CM Files')
os.environ.setdefault('PYTHONPATH', r'c:\Dev\CECL')
import sys
sys.path.insert(0, r'c:\Dev\CECL')
from cecl_credentials import get_database_url
os.environ['DATABASE_URL'] = get_database_url()
from sqlalchemy import create_engine, text
eng = create_engine(os.environ['DATABASE_URL'])
with eng.connect() as c:
    rows = c.execute(
        text("SELECT member_number, current_fico_score, loan_pool, current_balance "
             "FROM monthly_loan_data WHERE credit_union=:cu AND snapshot_date=:d "
             "AND member_number LIKE :p"),
        {"cu": "Nucor Emp CU", "d": "2026-06-30", "p": "%10557%"},
    ).fetchall()
    print("10557 rows:")
    for r in rows:
        print(repr(r))
    print("---")
    rows = c.execute(
        text("SELECT DISTINCT member_number FROM monthly_loan_data "
             "WHERE credit_union=:cu AND snapshot_date=:d LIMIT 15"),
        {"cu": "Nucor Emp CU", "d": "2026-06-30"},
    ).fetchall()
    print("Sample member_numbers:")
    for r in rows:
        v = r[0]
        print(f"  {v!r} len={len(v)}")
