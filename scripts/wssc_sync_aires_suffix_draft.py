"""Sync the WSSC wizard draft to the (correct) live config so the wizard shows
the AIRES member/account = combined fixed-suffix length 2, and so that ANY
future wizard save regenerates an equivalent config instead of regressing it.

Only the loan-suffix / file-pattern / provenance staleness is corrected:
  * top-level member_account -> fixed_suffix, suffix_length=2  (AIRES layout)
  * account_suffix_length    -> 2
  * top-level file_pattern    -> loosened pattern from config
  * loan_data_files[Loan ARIES].member_account -> fixed_suffix, 2
  * every loan_data_files[*].file_pattern -> loosened pattern from config
  * hist_balance_provenance   -> rich record from config (avoid save-downgrade)
Credit Card / ODP / CUMA member_account are LEFT AS-IS (validated: 16-digit card
PANs, bare 6-digit member numbers, and Member#-Mtg delimiter respectively — none
use a 2-digit loan suffix).
"""
import json, copy, shutil, datetime
import yaml

D = r"Z:\Shared\TCT Files\CECL - CM Files\wizard_drafts\wssc_fcu__migration.json"
CFG = r"Z:\Shared\TCT Files\CECL - CM Files\client_configs\wssc_fcu.yaml"

draft = json.load(open(D, encoding="utf-8"))
cfg = yaml.safe_load(open(CFG, encoding="utf-8"))

stamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
bak = f"{D}.bak.phase_aires_suffix_sync_{stamp}"
shutil.copy2(D, bak)
print("draft backup ->", bak)

AIRES_MA = {"mode": "fixed_suffix", "suffix_length": 2, "delimiter": "-"}

changes = []

# --- top-level ---
if draft.get("member_account") != AIRES_MA:
    draft["member_account"] = copy.deepcopy(AIRES_MA); changes.append("top member_account -> fixed_suffix/2")
if draft.get("account_suffix_length") != 2:
    draft["account_suffix_length"] = 2; changes.append("account_suffix_length -> 2")
if draft.get("file_pattern") != cfg.get("file_pattern"):
    draft["file_pattern"] = cfg.get("file_pattern"); changes.append("top file_pattern -> loosened")

# --- per loan_data_file: file_pattern from config (by label); AIRES gets suffix 2 ---
cfg_ext = {(e.get("label") or e.get("name")): e for e in (cfg.get("loan_data_extracts") or [])}
ldf = (draft.get("sample_uploads") or {}).get("loan_data_files") or []
for lf in ldf:
    label = lf.get("name")
    ext = cfg_ext.get(label)
    if not ext:
        continue
    if ext.get("file_pattern") and lf.get("file_pattern") != ext["file_pattern"]:
        lf["file_pattern"] = ext["file_pattern"]; changes.append(f"[{label}] file_pattern -> loosened")
    if label == "Loan ARIES" and lf.get("member_account") != AIRES_MA:
        lf["member_account"] = copy.deepcopy(AIRES_MA); changes.append(f"[{label}] member_account -> fixed_suffix/2")

# --- hist_balance_provenance: keep the rich config record ---
if cfg.get("hist_balance_provenance") and draft.get("hist_balance_provenance") != cfg["hist_balance_provenance"]:
    draft["hist_balance_provenance"] = copy.deepcopy(cfg["hist_balance_provenance"])
    changes.append("hist_balance_provenance -> rich config record")

json.dump(draft, open(D, "w", encoding="utf-8"), indent=2)
print("changes applied:")
for c in changes:
    print("  -", c)
if not changes:
    print("  (none — draft already in sync)")
