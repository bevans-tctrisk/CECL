import sys
sys.path.insert(0, r'c:\Dev\CECL')
from sqlalchemy import create_engine, text
from cecl_credentials import get_database_url
eng = create_engine(get_database_url())
with eng.begin() as cx:
    rs = cx.execute(text("SELECT DISTINCT cu FROM loan_code_delinquency_history ORDER BY cu")).fetchall()
    print("DISTINCT cu in loan_code_delinquency_history:")
    for r in rs:
        print(f"  {r[0]!r}")
    print()
    rs = cx.execute(text("SELECT DISTINCT cu FROM loan_code_history WHERE cu ILIKE '%emergency%' OR cu ILIKE '%responder%' ORDER BY cu")).fetchall()
    print("ER-like cu in loan_code_history:")
    for r in rs:
        print(f"  {r[0]!r}")
    print()
    rs = cx.execute(text("SELECT cu, COUNT(*) FROM loan_code_delinquency_history WHERE cu ILIKE '%emergency%' OR cu ILIKE '%responder%' GROUP BY cu")).fetchall()
    print("ER-like cu in loan_code_delinquency_history with counts:")
    for r in rs:
        print(f"  {r[0]!r}: {r[1]}")
