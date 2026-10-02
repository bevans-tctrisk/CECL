"""Build & save the Curis-Palmetto CECL config by driving the wizard backend.

Runs the wizard's auto-scan, applies the human-in-the-loop corrections
(pool mappings, column mappings, file pattern, CO/recovery columns), then
serialises via config_service.build_yaml_from_wizard and saves the YAML +
wizard draft — exactly the artifacts the web wizard produces.
"""
import os
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
WS = os.environ["CECL_WORKSPACE_ROOT"]

import yaml as _yaml
from cecl_ui.routes.setup import _default_state
from cecl_ui.services import auto_setup, config_service, wizard_drafts

FOLDER = (r"Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access\Clients"
          r"\Curis CU-Palmetto Health CU 61260\Portfolio Management\CECL Migration IDRL")
SNAPSHOT = "2026-03"
SHORT = "curis_palmetto_health_cu"

# ---- 1. Identity + auto-scan (same as wizard Step 1 "Scan & auto-fill") ----
state = _default_state()
state["credit_union"] = "Curis CU-Palmetto Health CU"
state["short_name"] = SHORT
state["charter_number"] = "61260"
state["report_period"] = SNAPSHOT
state["raw_data_folder"] = FOLDER
state["model"] = "migration"
state["has_warm_files"] = "no"
state["economic_data"].update({
    "state": "South Carolina",
    "county": "Richland County",
    "unemployment_rate": 0.041,
    "foreclosures": 0,
    "bankruptcies": 0,
    "population": 415759,
})

findings = auto_setup.scan_folder_for_setup(FOLDER, snapshot_yyyymm=SNAPSHOT)
assert findings.get("ok"), findings.get("error")
auto_setup.apply_findings_to_state(state, findings, WS)
state["_auto_scan_completed"] = True

# ---- 2. HIL corrections ----------------------------------------------------
# 2a. File pattern: match BOTH "CECL Loan Upload" (2024-25) and
#     "CECL Loan Update" (2026). date_format MMYYYY for MM-YYYY filenames.
LOAN_PATTERN = (r"(?i)^(?:0[1-9]|1[0-2])[-_]20\d{2}[\s_\-]+CECL[\s_\-]+Loan"
                r"[\s_\-]+(?:Upload|Update)\.(xlsx|xls)$")
state["file_pattern"] = LOAN_PATTERN
state["date_pattern"] = r"(\d{2})-(20\d{2})"
state["date_format"] = "MMYYYY"
state["account_suffix_length"] = 3
state["member_account"] = {"mode": "fixed_suffix", "suffix_length": 3, "delimiter": "-"}
state["has_header"] = True
state["pool_code_split"] = "/"

# 2b. Column mappings (corrected to real headers).
CM = {
    "member_number": "Account + ID",
    "loan_suffix": "",
    "current_balance": "Current Bal",
    "original_fico_score": "Credit Score",
    "current_fico_score": "",          # only one score column (current snapshot)
    "loan_pool_code": "Loan Code",     # FIX: was "Loan" (the 'L' marker col)
    "days_delinquent": "",             # DQ cols are event counts, not current DPD
    "interest_rate": "Int Rate",
    "open_date": "Loan Date",          # FIX: was placeholder
    "original_loan_amount": "Orignial Amt",  # header is misspelled in source
    "total_available_credit": "Credit Limit",
}
state["column_mappings"] = dict(CM)

# 2c. Pool map — alpha codes (loan file) + numeric codes (loan file & CO files,
#     per the CU's Lookups "Assets" table), grouped by collateral type.
POOL_MAP = {
    # alpha (loan file)
    "UNS": "Unsecured", "LOC": "LOC", "UVE": "Used Auto", "NVE": "New Auto",
    "HEQ": "HELOC", "REC": "Recreational", "SHR": "Share Secured",
    "1MT": "Real Estate", "2MT": "Real Estate",
    # numeric (lookup table)
    "0": "Unsecured", "1": "LOC", "2": "Unsecured", "3": "Unsecured",
    "4": "Unsecured", "5": "Unsecured", "6": "Unsecured", "7": "Unsecured",
    "8": "Unsecured", "9": "Unsecured", "10": "LOC", "11": "Unsecured",
    "12": "Unsecured", "14": "Unsecured", "16": "Unsecured",
    "20": "Share Secured", "22": "New Auto", "23": "Used Auto", "24": "New Auto",
    "25": "Used Auto", "27": "Used Auto", "28": "Used Auto", "29": "Used Auto",
    "30": "New Auto", "31": "Used Auto", "32": "Recreational", "33": "Recreational",
    "34": "New Auto", "35": "Used Auto", "36": "Used Auto", "39": "Used Auto",
    "40": "HELOC", "41": "HELOC", "42": "Real Estate", "43": "Real Estate",
    "44": "Real Estate", "45": "Real Estate", "48": "HELOC", "60": "Used Auto",
    "65": "Unsecured", "69": "Unsecured", "71": "LOC", "72": "New Auto",
    "75": "Used Auto", "78": "Real Estate", "80": "Used Auto", "86": "Unsecured",
    "87": "New Auto", "88": "Used Auto", "90": "Unsecured", "91": "Unsecured",
    "93": "Unsecured", "95": "Used Auto", "97": "Unsecured", "99": "Unsecured",
    "101": "Recreational", "999": "Unsecured",
}
state["pool_map"] = POOL_MAP
state["default_pool"] = "Ignore"

# 2d. Pool settings (per-pool WARM/ACL months + risk rating).
def pool(name, acl, risk=True, mgmt=False):
    return {"name": name, "risk_rated": risk, "acl_months": acl,
            "excluded": False, "use_default_mgmt_adj": mgmt, "brr": False}
state["pool_settings"] = [
    pool("Unsecured", 36),
    pool("LOC", 36),
    pool("New Auto", 36),
    pool("Used Auto", 36),
    pool("Recreational", 60),
    pool("Share Secured", 36),
    pool("HELOC", 84, mgmt=True),
    pool("Real Estate", 84, mgmt=True),
    {"name": "Ignore", "risk_rated": False, "acl_months": None,
     "excluded": False, "use_default_mgmt_adj": False, "brr": False},
]

# 2e. Loan-data extracts: keep ONLY the real loan file (drop the misclassified
#     "Loan Rates and FICO Scores" rate sheet), fix its mappings/pattern.
su = state.setdefault("sample_uploads", {})
ldf = su.get("loan_data_files") or []
keep = [e for e in ldf if "CECL Loan" in (e.get("name") or "")]
for e in keep:
    e["file_pattern"] = LOAN_PATTERN
    e["column_mappings"] = dict(CM)
    e["member_account"] = {"mode": "fixed_suffix", "suffix_length": 3, "delimiter": "-"}
    e["has_header"] = True
su["loan_data_files"] = keep

# 2f. Charge-off / Recovery columns (combined CO+Recovery transaction files:
#     Recovery Date | Account & ID | Loan Type Code | Charge Off Date |
#     Charge Off Amount | Recovery Amount ...). Shared account/code/date +
#     distinct amount cols => report's combined-file mode parses both halves.
ma3 = {"mode": "fixed_suffix", "suffix_length": 3, "delimiter": "-"}
state["co_columns"] = {
    "has_header": True, "skip_rows": 0, "account_col": 1, "code_col": 2,
    "amount_col": 4, "date_col": 3, "member_account": dict(ma3),
}
state["recov_columns"] = {
    "has_header": True, "skip_rows": 0, "account_col": 1, "code_col": 2,
    "amount_col": 5, "date_col": 3, "member_account": dict(ma3),
}

# 2g. data_directory -> per-CU Raw_Uploads.
raw_dir = str(config_service.raw_uploads_dir(WS) / SHORT)
state["data_directory"] = raw_dir

# 2h. Impaired-loan provision types (standard defaults).
state["impaired_loans"] = {
    "types": [
        {"name": "Delinquent Loans", "provision_pct": 0.35},
        {"name": "Known Losses", "provision_pct": 1.0},
        {"name": "Repossessions", "provision_pct": 1.0},
        {"name": "Bankruptcy", "provision_pct": 1.0},
    ]
}

# ---- 3. Serialise + save ---------------------------------------------------
cfg = config_service.build_yaml_from_wizard(state)
if not cfg.get("data_directory"):
    cfg["data_directory"] = raw_dir

target = config_service.save_client_config(WS, SHORT, cfg, overwrite=True)
print("Saved config ->", target)

# Save the wizard draft too (so the CU shows up in the wizard UI).
try:
    wizard_drafts.save_draft(WS, state, active_step="review", model="migration")
    print("Saved wizard draft.")
except Exception as e:
    print("draft save skipped:", type(e).__name__, e)

print("\n===== GENERATED YAML =====")
print(_yaml.safe_dump(cfg, sort_keys=False, allow_unicode=True))
