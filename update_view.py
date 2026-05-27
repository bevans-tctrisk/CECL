"""
Update the vw_cecl_calculations database view to use proper credit grade
configurations instead of hardcoded 1.5% reserve rate.

Run once: python update_view.py
"""
from dotenv import load_dotenv
import os
load_dotenv()
from sqlalchemy import create_engine, text

engine = create_engine(os.getenv('DATABASE_URL'))

new_view_sql = """
CREATE VIEW vw_cecl_calculations AS
SELECT
    m.credit_union,
    m.snapshot_date,
    m.member_number,
    m.loan_pool,
    m.current_balance,
    m.current_fico_score,
    m.original_fico_score,
    CASE WHEN m.current_fico_score > 0 THEN COALESCE(cg.grade_label::text, 'Not Reported')
         ELSE 'Not Reported' END AS current_grade,
    CASE WHEN m.original_fico_score > 0 THEN COALESCE(og.grade_label::text, 'Not Reported')
         ELSE 'Not Reported' END AS original_grade,
    m.current_balance * CASE WHEN m.current_fico_score > 0 THEN COALESCE(cg.reserve_rate, 0.0050)
                             ELSE 0.0050 END AS expected_loss_amount,
    CASE WHEN m.current_fico_score > 0 THEN COALESCE(cg.reserve_rate, 0.0050)
         ELSE 0.0050 END AS reserve_rate,
    CASE
        WHEN m.current_fico_score = 0 OR m.original_fico_score = 0 THEN 'Unchanged'
        WHEN cg.grade_id IS NULL OR og.grade_id IS NULL THEN 'Unchanged'
        WHEN cg.grade_id < og.grade_id THEN 'Improved'
        WHEN cg.grade_id > og.grade_id THEN 'Deteriorated'
        ELSE 'Unchanged'
    END AS migration_status
FROM monthly_loan_data m
LEFT JOIN credit_grade_configs cg
    ON m.current_fico_score BETWEEN cg.min_score AND cg.max_score
    AND cg.is_active = true
    AND m.current_fico_score > 0
LEFT JOIN credit_grade_configs og
    ON m.original_fico_score BETWEEN og.min_score AND og.max_score
    AND og.is_active = true
    AND m.original_fico_score > 0
"""

with engine.begin() as conn:
    conn.execute(text("DROP VIEW IF EXISTS vw_cecl_calculations"))
    print("  Old view dropped.")

with engine.begin() as conn:
    conn.execute(text(new_view_sql))
    print("View vw_cecl_calculations created successfully.")

# Verify in a separate connection so errors don't rollback the create
with engine.connect() as conn:
    r = conn.execute(text("SELECT count(*), count(DISTINCT migration_status) FROM vw_cecl_calculations"))
    row = r.fetchone()
    print(f"  View returns {row[0]} rows with {row[1]} distinct migration statuses")

    r = conn.execute(text("""
        SELECT current_grade, count(*), sum(current_balance), sum(expected_loss_amount)
        FROM vw_cecl_calculations
        GROUP BY current_grade
        ORDER BY current_grade
    """))
    print("\n  Grade Distribution:")
    for row in r:
        print(f"    {str(row[0]):15s}  {row[1]:5d} loans  bal={float(row[2]):,.2f}  reserve={float(row[3]):,.2f}")
