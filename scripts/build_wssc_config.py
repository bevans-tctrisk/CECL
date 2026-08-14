"""Build WSSC FCU config via the wizard serializer (config_service.build_yaml_from_wizard).
Multi-extract CU: Loan ARIES, Credit Card AIRES, ODP Accounts, CUMA mortgages.
Target: December 2025 report (ignore 2026 files)."""
import os
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
WS = os.environ["CECL_WORKSPACE_ROOT"]
import yaml
from cecl_ui.routes.setup import _default_state
from cecl_ui.services import config_service, wizard_drafts

SHORT = "wssc_fcu"

DATE_PATTERN = r"(20\d{2})\.(\d{2})"
LOAN_FP = r"(?i)^20\d{2}\.(?:0[1-9]|1[0-2])\.\d{2}\s+Loan\s+ARIES\.(xlsx|xls)$"
CC_FP   = r"(?i)^20\d{2}\.(?:0[1-9]|1[0-2])(?:\.\d{2})?\s+Credit\s+Card\s+-?\s*A[IR]{2}ES.*\.(xlsx|xls)$"
ODP_FP  = r"(?i)^20\d{2}\.(?:0[1-9]|1[0-2])\.\d{2}\s+ODP\s+Accounts\.(xlsx|xls)$"
CUMA_FP = r"(?i)^20\d{2}\.(?:0[1-9]|1[0-2])\s+CUMA\s+CECL\s+Report\.(xlsx|xls)$"

MA0 = {"mode": "fixed_suffix", "suffix_length": 0, "delimiter": "-"}
MA_DASH = {"mode": "delimiter", "suffix_length": 0, "delimiter": "-"}

loan_aries_cm = {
    "member_number": "Account Number", "loan_suffix": "",
    "current_balance": "Current Balance", "original_fico_score": "Credit Score",
    "current_fico_score": "", "loan_pool_code": "Loan Type Code",
    "days_delinquent": "Days Delinquent", "interest_rate": "Interest Rate",
    "open_date": "Date of Loan", "original_loan_amount": "Original Loan Amount",
    "total_available_credit": "Credit Limit",
}
cc_cm = {
    "member_number": "Card Account Number", "loan_suffix": "",
    "current_balance": "Current Balance Amount", "original_fico_score": "",
    "current_fico_score": "", "loan_pool_code_static": "CC",
    "days_delinquent": "Delinquent Days Count", "interest_rate": "Interest Rate (APR)",
    "open_date": "Card Open Date", "original_loan_amount": "Original Loan Amount",
    "total_available_credit": "Credit Line Amount",
}
odp_cm = {
    "member_number": "Account", "loan_suffix": "",
    "current_balance": "Current Balance", "original_fico_score": "",
    "current_fico_score": "", "loan_pool_code": "Loan Type Code",
    "days_delinquent": "", "interest_rate": "", "open_date": "",
    "original_loan_amount": "", "total_available_credit": "ODP Limit",
}
cuma_cm = {
    "member_number": "Member #", "loan_suffix": "",
    "current_balance": "Prin Bal", "original_fico_score": "",
    "current_fico_score": "", "loan_pool_code": "Loan Type Code",
    "days_delinquent": "", "interest_rate": "Int Rate",
    "open_date": "Date of Note", "original_loan_amount": "Loan Amt",
    "total_available_credit": "",
}

extracts = [
    {"name": "Loan ARIES", "file_pattern": LOAN_FP, "column_mappings": loan_aries_cm,
     "member_account": dict(MA0), "has_header": True},
    {"name": "Credit Card AIRES", "file_pattern": CC_FP, "column_mappings": cc_cm,
     "member_account": dict(MA0), "has_header": True},
    {"name": "ODP Accounts", "file_pattern": ODP_FP, "column_mappings": odp_cm,
     "member_account": dict(MA0), "has_header": True},
    {"name": "CUMA CECL Report", "file_pattern": CUMA_FP, "column_mappings": cuma_cm,
     "member_account": dict(MA_DASH), "has_header": True},
]

RE = "All Other Real Estate"
CO_ = "Consumer - Other"
VEH = "Vehicle Loans"
pool_map = {
    "70198": RE, "70199": RE, "HE": RE, "HE-A": RE, "HEL": RE, "HELA": RE,
    "HELB": RE, "HELC": RE, "HELD": RE, "HELE": RE, "HELF": RE, "HELO": RE,
    "RE": RE, "Mortgages": RE,
    "HEIA": "Insured HELOCS", "HELI": "Insured HELOCS",
    "HIMP": "Insured HIMP Loans",
    "C19": CO_, "CB": CO_, "CD": CO_, "DEBT": CO_, "EPR": CO_, "LOC": CO_,
    "LOCB": CO_, "LOCC": CO_, "LOCE": CO_, "LOCF": CO_, "LOCV": CO_,
    "OD": CO_, "ODF": CO_, "ODV": CO_, "ODVB": CO_, "ODVC": CO_, "ODVD": CO_,
    "ODVE": CO_, "OLOC": CO_, "SAV": CO_, "SS": CO_, "SG": CO_, "SIG": CO_,
    "SL": CO_, "Signature": CO_, "WA": CO_, "RTS": CO_, "RT Merger Loans": CO_,
    "Unsecure": CO_, "Secure": CO_,
    "NA": VEH, "UA": VEH, "RTNA": VEH, "RTUA": VEH, "Used Auto": VEH,
    "CC": "Credit Card", "Visa": "Credit Card", "Credit Cards": "Credit Card",
    "HL": "Holiday 365", "Holiday365": "Holiday 365", "Holiday Loans": "Holiday 365",
    "Neg Shares": "Overdraft Privilege", "ODLN": "Overdraft Privilege", "ODP": "Overdraft Privilege",
    "SGSG": "St. Gabriel's FCU",
}

def pool(name, months, mgmt=False, rr=True):
    return {"name": name, "risk_rated": rr, "acl_months": months,
            "excluded": False, "use_default_mgmt_adj": mgmt, "brr": False}
pool_settings = [
    pool(CO_, 36), pool(VEH, 36), pool(RE, 84, True), pool("Insured HELOCS", 84, True),
    pool("Insured HIMP Loans", 84, True), pool("Credit Card", 36), pool("Holiday 365", 36),
    pool("Overdraft Privilege", 12), pool("St. Gabriel's FCU", 36),
    {"name": "Ignore", "risk_rated": False, "acl_months": None, "excluded": False,
     "use_default_mgmt_adj": False, "brr": False},
]

co_cols = {"has_header": True, "skip_rows": 0, "member_col": None, "account_col": 1,
           "code_col": 4, "amount_col": 2, "date_col": 0, "member_account": dict(MA0)}
recov_cols = dict(co_cols)

# Multi-format CO/Recovery: WSSC ships three different CO/recovery layouts.
co_recov_formats = [
    {   # Separate "Loan Charge-offs" and "Loan Recoveries" files, same layout.
        "name": "Loan CO/Recovery",
        "file_pattern": r"(?i)Loan\s+(?:Charge-?offs?|Recoveries)",
        "co_columns":    {"has_header": True, "skip_rows": 0, "account_col": 1,
                          "code_col": 4, "amount_col": 2, "date_col": 0},
        "recov_columns": {"has_header": True, "skip_rows": 0, "account_col": 1,
                          "code_col": 4, "amount_col": 2, "date_col": 0},
    },
    {   # Combined "Credit Card Charge-offs and Recoveries" file. No loan-code
        # column (every row is a credit card) -> code_static. CO amount col 4,
        # recovery amount col 3, shared account 0 / date 5.
        "name": "Credit Card CO/Recovery",
        "file_pattern": r"(?i)Credit\s+Card\s+Charge-?offs?\s+and\s+Recoveries",
        "co_columns":    {"has_header": True, "skip_rows": 0, "account_col": 0,
                          "code_static": "CC", "amount_col": 4, "date_col": 5},
        "recov_columns": {"has_header": True, "skip_rows": 0, "account_col": 0,
                          "code_static": "CC", "amount_col": 3, "date_col": 5},
    },
    {   # Combined "ODP Charge-offs and Recoveries" file. code col 1 ("ODP"),
        # CO amount col 3, recovery amount col 2, shared account 0 / date 4.
        # Pattern also matches the older "ODP Charge Off and Recovery" naming.
        "name": "ODP CO/Recovery",
        "file_pattern": r"(?i)ODP\s+Charge[\s_-]?offs?\s+and\s+Recover(?:y|ies)",
        "co_columns":    {"has_header": True, "skip_rows": 0, "account_col": 0,
                          "code_col": 1, "amount_col": 3, "date_col": 4},
        "recov_columns": {"has_header": True, "skip_rows": 0, "account_col": 0,
                          "code_col": 1, "amount_col": 2, "date_col": 4},
    },
    {   # Pre-2023 combined loan + credit-card files: "ChargeOff_Recovery File
        # YYYY.MM" and "CrdtCrd ChrgeOff Rec YYYY.MM". 8-col layout:
        # FileName0, Date1, Account2, LoanNo3, Member4, CO Amt5, Recovery
        # Amt6, Loan Type7 (SIG/HL/UA/RTS/CC...). Extends loss history to 2018.
        "name": "Legacy Combined CO/Recovery (pre-2023)",
        "file_pattern": r"(?i)(?:ChargeOff_Recovery\s+File|CrdtCrd\s+ChrgeOff\s+Rec)",
        "co_columns":    {"has_header": True, "skip_rows": 0, "account_col": 2,
                          "code_col": 7, "amount_col": 5, "date_col": 1},
        "recov_columns": {"has_header": True, "skip_rows": 0, "account_col": 2,
                          "code_col": 7, "amount_col": 6, "date_col": 1},
    },
]

state = _default_state()
state.update({
    "credit_union": "WSSC FCU", "short_name": SHORT, "charter_number": "16268",
    "report_period": "2025-12", "model": "migration", "has_warm_files": "no",
    "file_pattern": LOAN_FP, "date_pattern": DATE_PATTERN, "date_format": "YYYY-MM",
    "account_suffix_length": 0, "member_account": dict(MA0), "has_header": True,
    "column_mappings": loan_aries_cm, "pool_map": pool_map, "default_pool": "Ignore",
    "pool_code_split": "", "pool_settings": pool_settings,
    "co_recov_formats": co_recov_formats,
    "data_directory": str(config_service.raw_uploads_dir(WS) / SHORT),
})
state["economic_data"]["state"] = "Maryland"
state["economic_data"]["county"] = "Prince George's County"
state["reports"] = {"tct": False, "vizo": True, "vizo_supp": True, "impdet": True}
state.setdefault("sample_uploads", {})["loan_data_files"] = extracts

cfg = config_service.build_yaml_from_wizard(state)
target = config_service.save_client_config(WS, SHORT, cfg, overwrite=True)
try:
    wizard_drafts.save_draft(WS, state, active_step="review", model="migration")
except Exception as e:
    print("draft save warn:", e)
print("Saved config ->", target)
print("\nhistorical_file_formats:", list((cfg.get('historical_file_formats') or {}).keys()))
print("loan_data_extracts:", [e['label'] for e in cfg.get('loan_data_extracts', [])])
print("pools:", [p['name'] for p in cfg.get('pools', [])])
print("\n--- full yaml ---")
print(yaml.safe_dump(cfg, sort_keys=False, default_flow_style=False)[:1500])
