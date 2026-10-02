"""Smoke-test auto_setup against Census FCU's raw folder."""
from cecl_ui.services import auto_setup

CENSUS = r"Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access\Clients\Census FCU 5641\Portfolio Management\CECL Migration IDLR"

r = auto_setup.scan_folder_for_setup(CENSUS, snapshot_yyyymm="2025-12")

print("=" * 70)
print("SCAN RESULT")
print("=" * 70)
print(f"OK: {r['ok']}")
print(f"Errors: {r['errors']}")
print()
print("Messages:")
for m in r["messages"]:
    print(f"  - {m}")
print()

cls = r["classification"]
print("Classification:")
for k in [
    "loan_data_files",
    "annual_balance_files",
    "monthly_detail_balance_files",
    "monthly_files",
    "warm_files",
    "impaired_files",
    "credit_pull_files",
    "co_files",
    "recov_files",
    "five_thirtythousand_files",
    "other_files",
]:
    vals = cls.get(k) or []
    print(f"  {k}: {len(vals)}")

print()
print("Loan extracts found:")
for e in cls.get("loan_data_files") or []:
    print(f"  - {e['name']}  period={e.get('period')}")

print()
print("Annual balance sheets:")
for e in cls.get("annual_balance_files") or []:
    print(f"  - {e['name']}  year={e.get('year')}")

print()
print("Sample analysis:")
s = r.get("sample")
if s:
    print(f"  headers (first 6): {s.get('headers', [])[:6]}")
    print(f"  pool_codes (first 10): {s.get('pool_code_suggestions', [])[:10]}")
    print(f"  column_suggestions: {s.get('column_suggestions')}")
    print(f"  file_pattern: {s.get('file_pattern')}")
    print(f"  date_pattern: {s.get('date_pattern')}")
else:
    print("  (no sample analysis — no loan extracts detected)")

print()
print("Annual balance layout:")
ab = r.get("annual_bal")
if ab:
    print(f"  sheet: {ab['layout'].get('sheet')}")
    print(f"  label_col: {ab['layout'].get('label_col')}")
    print(f"  header_row: {ab['layout'].get('header_row')}")
    print(f"  pool_labels: {ab.get('pool_labels')[:12]}")
else:
    print("  (no annual balance layout)")

print()
print("Companion files:")
for k, label in [("impaired_file", "impaired"), ("credit_pull_file", "credit_pull"), ("co_file", "CO"), ("recov_file", "recov")]:
    e = r.get(k)
    if e:
        print(f"  {label}: {e['name']}")
    else:
        print(f"  {label}: (not found)")

print()
print("=" * 70)
print("STATE POPULATION TEST")
print("=" * 70)

# Synthesize a minimal default state for testing apply_findings_to_state.
state = {
    "credit_union": "Census FCU (Auto)",
    "short_name": "census_fcu_auto",
    "charter_number": "5641",
    "report_period": "2025-12",
    "raw_data_folder": CENSUS,
    "economic_data": {"state": "", "county": ""},
    "column_mappings": {
        "member_number": "MEMBER_ID",
        "loan_suffix": "",
        "current_balance": "BALANCE",
        "original_fico_score": "FICO_SCORE",
        "current_fico_score": "",
        "loan_pool_code": "LOAN_TYPE",
        "days_delinquent": "DQ_DAYS",
        "interest_rate": "INT_RATE",
        "open_date": "OPEN_DATE",
        "original_loan_amount": "ORIG_AMT",
        "total_available_credit": "",
    },
    "pool_map": {},
    "default_pool": "Ignore",
    "pool_code_split": "",
    "monthly_bal": {},
    "has_warm_files": None,
    "credit_pull": {},
    "sample_uploads": {},
    "mgmt_adj": {},
    "reports": {"tct": True, "vizo": False, "vizo_supp": False, "impdet": False},
    "credit_grades": [],
    "file_pattern": "",
    "date_pattern": "",
    "has_header": False,
}

WORKSPACE = r"Z:\Shared\TCT Files\CECL - CM Files"

report = auto_setup.apply_findings_to_state(state, r, WORKSPACE)
print()
print("Per-step auto-fill report:")
for step_key, msgs in report.items():
    print(f"  [{step_key}]")
    for m in msgs:
        print(f"    - {m}")

print()
print("=" * 70)
print("HIL NEEDS")
print("=" * 70)
needs = auto_setup.compute_hil_needs(state)
for n in needs:
    reasons = "; ".join(n["reasons"])
    print(f"  [{n['severity']:11s}] {n['step_key']:14s} -- {reasons}")

print()
print(f"State keys after apply: {sorted(state.keys())}")
print(f"file_pattern: {state.get('file_pattern')}")
print(f"date_pattern: {state.get('date_pattern')}")
print(f"data_directory: {state.get('data_directory')}")
print(f"reports: {state.get('reports')}")
print(f"mgmt_adj: {state.get('mgmt_adj')}")
print(f"monthly_bal source: {state.get('monthly_bal', {}).get('source')}")
print(f"monthly_bal year_files: {len(state.get('monthly_bal', {}).get('year_files') or [])}")
print(f"monthly_bal pool_labels: {len(state.get('monthly_bal', {}).get('parsed_pool_labels') or [])}")
print(f"pool_map size: {len(state.get('pool_map') or {})}")
print(f"has_warm_files: {state.get('has_warm_files')}")
print(f"sample_uploads keys: {sorted((state.get('sample_uploads') or {}).keys())}")
