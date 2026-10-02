import sys
sys.path.insert(0, r'c:\Dev\CECL')
from sqlalchemy import create_engine, text
from cecl_credentials import get_database_url
eng = create_engine(get_database_url())
with eng.begin() as cx:
    # Check whether loan_code_delinquency_history has populated total_balance
    rs = cx.execute(text("""
        SELECT as_of_date, loan_code, dq_amount, total_balance, dq_pct, source
        FROM loan_code_delinquency_history
        WHERE cu='Emergency Responders CU'
          AND loan_code IN ('Junior Liens','First Liens','All Other Secured Non-RE')
          AND COALESCE(dq_amount, 0) > 0
        ORDER BY as_of_date, loan_code
        LIMIT 20
    """)).fetchall()
    print(f"{'Date':<12} {'Code':<28} {'DQ$':>12} {'TotalBal$':>14} {'DQ%':>8}  Source")
    for r in rs:
        d, code, dq, bal, pct, src = r
        bal_s = f"${float(bal):>12,.0f}" if bal else "        NULL  "
        pct_s = f"{float(pct):>7.4f}" if pct else "    NULL"
        print(f"{str(d):<12} {code:<28} ${float(dq or 0):>11,.0f}  {bal_s} {pct_s}  {src}")
    print()
    # Also check distinct sources
    rs = cx.execute(text("""
        SELECT source, COUNT(*) AS n,
               COUNT(*) FILTER (WHERE total_balance IS NOT NULL) AS with_bal
        FROM loan_code_delinquency_history
        WHERE cu='Emergency Responders CU'
        GROUP BY source ORDER BY source
    """)).fetchall()
    print("Sources for ER DQ rows:")
    for r in rs:
        print(f"  {r[0]!r}: {r[1]} rows, {r[2]} with total_balance")
