import sys
sys.path.insert(0, r'c:\Dev\CECL')
from sqlalchemy import create_engine, text
from cecl_credentials import get_database_url
eng = create_engine(get_database_url())
with eng.begin() as cx:
    rs = cx.execute(text("SELECT table_name, column_name FROM information_schema.columns WHERE table_name IN ('loan_code_delinquency_history','loan_code_history') ORDER BY table_name, ordinal_position")).fetchall()
    for r in rs:
        print(r[0], '|', r[1])
