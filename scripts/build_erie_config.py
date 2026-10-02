"""Build client_configs/erie_fcu.yaml for Erie FCU (TCT/WARM conversion).

Erie is a TCT client converted from the analyst "CECL-Migration-WARM" workbook
to the wizard/config system. The consumer loan portfolio is delivered as a
pipe-delimited ``TCT.EXPORT_YYYYMM.txt`` extract (no header, 11 positional
columns).  Pool balances, tier-grade distribution and credit migration are all
driven by that extract.  Commercial pools (Business, Off-Books Real Estate,
Participations, Courtesy Pay) come from other files and are handled separately.

Run:  set PYTHONPATH=C:\\dev\\CECL ; set CECL_WORKSPACE_ROOT=Z:\\... ;
      python scripts/build_erie_config.py
"""
import os
import openpyxl
import yaml

WS_ROOT = os.environ.get("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
CFG_PATH = os.path.join(WS_ROOT, "client_configs", "erie_fcu.yaml")
POOLMAP = r"Z:\Shared\Clients\Erie FCU\Client Access\Portfolio Management (CM, ID, Warm)\Erie FCU - Pool Mapping.xlsx"

# ---- collateral code (int) -> pool, from the CU's reference workbook ----
wb = openpyxl.load_workbook(POOLMAP, read_only=True, data_only=True)
coll = {}          # int code -> pool
coll_alpha = {}    # alpha code -> pool
for row in wb["Loan Codes"].iter_rows(min_row=2, values_only=True):
    if row[0] is None or str(row[0]).strip() == "Collateral Code":
        continue
    pool = str(row[1]).strip()
    try:
        coll[int(row[0])] = pool
    except (TypeError, ValueError):
        coll_alpha[str(row[0]).strip()] = pool
wb.close()

# The TCT.EXPORT loan_type codes are zero-padded 4-char strings (e.g. "0101"),
# so map_pool_codes (string match) needs string keys. Emit BOTH the raw
# zero-padded form and the stripped form for every numeric collateral code.
pool_map = {}
for code_int, pool in coll.items():
    pool_map[str(code_int)] = pool          # stripped: "101"
    pool_map[f"{code_int:04d}"] = pool       # padded:   "0101"
for code_a, pool in coll_alpha.items():
    pool_map[code_a] = pool

# ---- validated overrides (see erie_codes.py analysis vs WARM ACL Env tab) ----
# 0513 = serviced / off-books first mortgages -> excluded from on-book RE.
pool_map["0513"] = "Off Books Real Estate"
pool_map["513"] = "Off Books Real Estate"
# 0210 sits with the Unsecured personal codes (0200-0209).
pool_map["0210"] = "Unsecured"
pool_map["210"] = "Unsecured"
# Alpha card codes.
pool_map["MC"] = "Credit Card"
pool_map["BMC"] = "Business Credit Card"
# Negative shares.
pool_map["9051"] = "Negative Share"

# ---- tier credit grades (from 'Grade Ranges & Loan Codes') ----
credit_grades = [
    {"label": "Tier 1", "min_score": 730, "max_score": 900, "reserve_rate": 0.0011},
    {"label": "Tier 2", "min_score": 690, "max_score": 729, "reserve_rate": 0.0011},
    {"label": "Tier 3", "min_score": 660, "max_score": 689, "reserve_rate": 0.0011},
    {"label": "Tier 4", "min_score": 620, "max_score": 659, "reserve_rate": 0.0011},
    {"label": "Tier 5", "min_score": 600, "max_score": 619, "reserve_rate": 0.0011},
    {"label": "Tier 6", "min_score": 300, "max_score": 599, "reserve_rate": 0.0011},
]

# ---- pools (consumer = tier/risk_rated; commercial + off-book handled later) ----
consumer_pools = ["Auto", "Share Secured", "Unsecured", "Real Estate", "Credit Card"]
other_pools = [
    "Off Books Real Estate", "Business", "Business Credit Card", "Student Loans",
    "Participation Loans", "Courtesy Pay", "Negative Share",
]
single_line = {"Student Loans", "Participation Loans", "Courtesy Pay", "Negative Share",
               "Business", "Business Credit Card"}

# Commercial pools whose allowance is an analyst-set per-loan specific reserve
# (from the 'Comm Ln Allowance Loan Loss' matrix schedule) that the firm-wide
# FICO/loss-rate model cannot reproduce. The report uses the WARM workbook's
# already-computed allowance for these.
WARM_ALLOWANCE_POOLS = ["Business", "Business Credit Card"]

# Per-pool loss window (months) from WARM 'Display CO-Recov-DQ' "ACL Quarters":
# consumer 12q=36mo, Real Estate/Off-Books RE 28q=84mo, Business 20q=60mo.
ACL_MONTHS = {
    "Real Estate": 84, "Off Books Real Estate": 84, "Business": 60,
}

pools = []
for name in consumer_pools + other_pools:
    pools.append({
        "name": name,
        "risk_rated": name not in single_line,
        "acl_months": ACL_MONTHS.get(name),
        "excluded": False,
        "use_default_mgmt_adj": False,
    })

not_risk_rated = sorted(single_line) + ["Ignore"]

cfg = {
    "credit_union": "Erie FCU",
    "data_directory": "erie_fcu",
    "file_pattern": r"TCT\.EXPORT_\d{6}\.txt$",
    "date_pattern": r"(20\d{2})(\d{2})",
    "date_format": "YYYYMM",
    "has_header": False,
    "account_suffix_length": 0,
    "member_account": {"mode": "fixed_suffix", "suffix_length": 0},
    "column_mappings": {
        "member_number": 0,
        "loan_pool_code": 2,
        "current_balance": 3,
        "original_fico_score": 4,
        "current_fico_score": 5,
        "open_date": 6,
        "days_delinquent": 7,
        "interest_rate": 8,
        "original_loan_amount": 9,
        "total_available_credit": 10,
    },
    "balance_format": {"remove_chars": ["$", ","], "accounting_negatives": True},
    "pool_code_split": None,
    "pool_map": pool_map,
    "default_pool": "Other/Uncategorized",
    # The analyst's CECL-Migration-WARM workbook is the authoritative source
    # for historical balances, charge-off/recovery history and impaired loans.
    # The report loads it from data_dir first, then this folder.
    "credit_pull": {
        "file_pattern": None,
        "fallback_report_folder": r"Z:\Shared\Clients\Erie FCU\Portfolio Management (CM, ID, Warm)",
        "fallback_report_pattern": r"CECL-Migration.*\.xlsx$",
        "fallback_sheet_pattern": "Credit Pull",
        "fallback_member_col": 0,
        "fallback_score_col": 1,
    },
    "credit_grades": credit_grades,
    "no_score_label": "Not Reported",
    # CU-specific distribution factors (percentages, Tier 1-6 + Not Reported)
    # from the WARM 'Grade Ranges & Loan Codes' Distribution Factor column.
    "distribution_factors": [
        10.521341711028578, 22.921341711028578, 45.14134171102858,
        95.83134171102858, 131.59134171102858, 143.45134171102858,
        164.7613417110286,
    ],
    "reports": {"tct": True, "vizo": False, "vizo_supp": False, "impdet": False},
    "economic_data": {
        "state": "Pennsylvania",
        "county": "Erie",
        "unemployment_rate": 0.046,
        "foreclosures": 5809,
        "bankruptcies": 12666,
        "population": 13078751,
    },
    "mgmt_adj": {"ltv_baseline": 0.9, "probability_factor": 0.35},
    "mgmt_adj_by_pool": {},
    "acl_months_by_pool": ACL_MONTHS,
    "warm_allowance_pools": WARM_ALLOWANCE_POOLS,
    # Secured pools with no charge-off history fall back to the grade
    # reserve_rate x distribution factor (WARM "Default Loss Rate").
    "reserve_rate_fallback_pools": ["Real Estate", "Off Books Real Estate"],
    "pools": pools,
    "pool_order": consumer_pools + other_pools,
    "not_risk_rated": not_risk_rated,
}

os.makedirs(os.path.dirname(CFG_PATH), exist_ok=True)
with open(CFG_PATH, "w", encoding="utf-8") as fh:
    yaml.safe_dump(cfg, fh, sort_keys=False, allow_unicode=True)
print("wrote", CFG_PATH)
print("pool_map keys:", len(pool_map), "| pools:", len(pools))
