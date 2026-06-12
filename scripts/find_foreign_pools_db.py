"""Check monthly_loan_data for any CU using the 4 foreign pool names."""
import sys
sys.path.insert(0, r"c:\Dev\CECL")
from cecl_credentials import get_database_url
import psycopg2

url = get_database_url()
conn = psycopg2.connect(url)
cur = conn.cursor()

cur.execute(
    "SELECT credit_union, COUNT(*) FROM monthly_loan_data GROUP BY credit_union ORDER BY credit_union"
)
print("CUs with rows in monthly_loan_data:")
for r in cur.fetchall():
    print(f"  {r[0]}: {r[1]}")

cur.execute(
    "SELECT credit_union, loan_pool, COUNT(*) FROM monthly_loan_data "
    "WHERE loan_pool IN ('Unsecured','Secured','1st Mortgage','2nd Mortgage/HE') "
    "GROUP BY credit_union, loan_pool ORDER BY credit_union, loan_pool"
)
print("\nCUs using foreign pool names:")
for r in cur.fetchall():
    print(f"  {r[0]} | {r[1]} | {r[2]}")

cur.execute(
    "SELECT credit_union, COUNT(DISTINCT loan_pool), array_agg(DISTINCT loan_pool ORDER BY loan_pool) "
    "FROM monthly_loan_data GROUP BY credit_union ORDER BY credit_union"
)
print("\nDistinct loan_pool values per CU:")
for r in cur.fetchall():
    print(f"\n  {r[0]} ({r[1]} pools):")
    for p in r[2] or []:
        print(f"     - {p}")

cur.close()
conn.close()
