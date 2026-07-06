"""Apply member_account fixed_suffix suffix_length=2 to WSSC's Loan ARIES extract
(and the mirrored top-level), so account = member(base) + 2-digit loan suffix.
Groups a member's loans together -> Option 3 most-recent-loan derivation yields
real credit migration. Updates YAML + draft; backs up both."""
import os, shutil, datetime, json
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
WS = os.environ["CECL_WORKSPACE_ROOT"]
import yaml
STAMP = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

# ---- YAML ----
cpath = os.path.join(WS, "client_configs", "wssc_fcu.yaml")
shutil.copy2(cpath, cpath + f".bak.phase_suffix2_{STAMP}")
cfg = yaml.safe_load(open(cpath, encoding="utf-8"))
cfg.setdefault("member_account", {}).update({"mode": "fixed_suffix", "suffix_length": 2})
cfg["account_suffix_length"] = 2
for e in cfg.get("loan_data_extracts") or []:
    if e.get("label") == "Loan ARIES":
        e.setdefault("member_account", {}).update({"mode": "fixed_suffix", "suffix_length": 2})
with open(cpath, "w", encoding="utf-8") as fh:
    yaml.safe_dump(cfg, fh, sort_keys=False, default_flow_style=False, allow_unicode=True)
print("YAML member_account:", cfg["member_account"])
print("YAML Loan ARIES member_account:",
      next(e["member_account"] for e in cfg["loan_data_extracts"] if e["label"] == "Loan ARIES"))

# ---- draft ----
dpath = os.path.join(WS, "wizard_drafts", "wssc_fcu__migration.json")
shutil.copy2(dpath, dpath + f".bak.phase_suffix2_{STAMP}")
d = json.load(open(dpath, encoding="utf-8"))
d.setdefault("member_account", {}).update({"mode": "fixed_suffix", "suffix_length": 2})
d["account_suffix_length"] = 2
for e in (d.get("sample_uploads") or {}).get("loan_data_files") or []:
    if e.get("name") == "Loan ARIES":
        e.setdefault("member_account", {}).update({"mode": "fixed_suffix", "suffix_length": 2})
json.dump(d, open(dpath, "w", encoding="utf-8"), indent=2)
print("draft updated.")
