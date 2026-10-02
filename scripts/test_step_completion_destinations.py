"""Verify compute_step_completion + _wizard_ctx-style status logic
against a synthetic state that matches the Destinations CU situation:
user has verified pools (Step 2) and mapped codes to pools (Step 3)
after running the Step 1 auto-scan.
"""
from cecl_ui.services import auto_setup

# Mimic the state after: auto-scan ran, user verified 7 pools on Step 2,
# user mapped 50 of 75 loan codes to real pools on Step 3.
state = {
    "credit_union": "Destinations Credit Union",
    "report_period": "2025-12",
    "_auto_scan_completed": True,
    "_auto_scan_report": {
        "loan_pools": ["Detected 7 pools from Vizo Balance Sheets workbook."],
        "pools": ["Mirrored from loan_pools."],
        "files": ["Auto-detected file_pattern."],
        "columns": ["Auto-detected member_number column."],
        "economic": ["Auto-detected state/county."],
        "mgmt_adj": ["Set Vizo defaults."],
        "reports": ["Vizo defaults."],
    },
    "pool_settings": [
        {"name": "Unsecured", "risk_rated": False, "acl_months": 36},
        {"name": "Used Vehicle", "risk_rated": False, "acl_months": 36},
        {"name": "New Vehicle", "risk_rated": False, "acl_months": 36},
        {"name": "1st Lien Real Estate", "risk_rated": True, "acl_months": 84},
        {"name": "Junior Lien Real Estate", "risk_rated": True, "acl_months": 84},
        {"name": "HELOC", "risk_rated": True, "acl_months": 84},
        {"name": "Visa", "risk_rated": False, "acl_months": 36},
    ],
    # Simulate user mapping 50/75 codes to real pools, 25 still 'Ignore'.
    "pool_map": {
        **{f"CODE{i:03d}": "Unsecured" for i in range(20)},
        **{f"CODE{i:03d}": "Used Vehicle" for i in range(20, 30)},
        **{f"CODE{i:03d}": "Visa" for i in range(30, 50)},
        **{f"CODE{i:03d}": "Ignore" for i in range(50, 75)},
    },
    "file_pattern": "(?i)^.*aires.*\\.xlsx$",
    "date_format": "MMDDYY",
    "column_mappings": {
        "member_number": "Member Number",
        "current_balance": "Current Balance",
    },
    "economic_data": {"state": "MD", "county": "Baltimore"},
    "mgmt_adj": {"ltv_baseline": 0.9, "probability_factor": 0.35},
    "reports": {"vizo": True, "vizo_supp": True, "tct": False, "impdet": True},
    "credit_grades": [
        {"label": "A", "min_score": 740, "max_score": 850, "reserve_rate": 0.0025},
    ],
    "monthly_bal": {"source": None, "parsed_pool_labels": [], "pool_map": {}},
    "loan_data_extracts": [],
    "loan_data_files": [],
    "warm": {},
    "balance_title_map": {},
    "hist_extracts": {},
    "sample_uploads": {},
    "credit_pull": {},
    "impaired": {},
}

print("=" * 60)
print("BEFORE-FIX behaviour (auto only when in _auto_scan_report):")
print("=" * 60)
needs = auto_setup.compute_hil_needs(state)
need_keys = {n["step_key"] for n in needs}
report = state["_auto_scan_report"]
old_status = {}
for key in ["loan_pools", "pools", "monthly_bal", "grades", "credit_pull",
            "impaired", "files", "columns", "economic", "mgmt_adj"]:
    if key in need_keys:
        old_status[key] = "hil   ⚠"
    elif key in report:
        old_status[key] = "auto  ✓"
    else:
        old_status[key] = "pending"
for k, v in old_status.items():
    print(f"  {k:<14s} -> {v}")

print()
print("=" * 60)
print("AFTER-FIX behaviour (with compute_step_completion override):")
print("=" * 60)
done = auto_setup.compute_step_completion(state)
required_keys = {n["step_key"] for n in needs if str(n.get("severity")).lower() == "required"}
new_status = {}
for key in ["loan_pools", "pools", "monthly_bal", "grades", "credit_pull",
            "impaired", "files", "columns", "economic", "mgmt_adj"]:
    if key in done and key not in required_keys:
        new_status[key] = "auto  ✓ "
    elif key in need_keys:
        new_status[key] = "hil   ⚠ "
    elif key in report:
        new_status[key] = "auto  ✓ (legacy)"
    else:
        new_status[key] = "pending"
for k, v in new_status.items():
    changed = " <-- CHANGED" if old_status[k] != new_status[k] else ""
    print(f"  {k:<14s} -> {v}{changed}")

print()
print(f"Step completion set: {sorted(done)}")

print()
print("EXPECTED: loan_pools=auto, pools=auto (user mapped some codes).")
assert "loan_pools" in done, "FAIL: loan_pools should be complete (7 named pools)"
assert "pools" in done, "FAIL: pools should be complete (50 codes mapped)"
print("PASS: both Step 2 and Step 3 will now show green checkmark.")
