"""Restore Central Susquehanna column mappings that reverted via the
wizard's re-derivation, and apply the user's current_fico_score blank.

Fixes (top-level column_mappings AND every loan_data_extracts[*] copy,
in BOTH the report YAML and the wizard draft JSON):

    member_number        : Member Name    -> Account Number
    original_loan_amount : ORIG NBR PYMTS  -> LOAN ORIGINAL AMNT
    current_fico_score   : CREDIT SCORE    -> (blank / None)

Idempotent. Backs up both files with a .bak.phase_mapping_restore suffix.
Read the summary it prints to confirm every occurrence was updated.
"""
from __future__ import annotations

import json
import os
import shutil

import yaml

WS = os.environ.get("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
YAML_PATH = os.path.join(WS, "client_configs", "central_susquehanna_comm_fcu.yaml")
DRAFT_PATH = os.path.join(
    WS, "wizard_drafts", "central_susquehanna_comm_fcu__migration.json"
)
BAK = ".bak.phase_mapping_restore"

# desired: key -> value (None means blank/unmapped)
FIXES = {
    "member_number": "Account Number",
    "original_loan_amount": "LOAN ORIGINAL AMNT",
    "current_fico_score": None,
}


def _apply_to_mapping(cm: dict) -> list[str]:
    changed = []
    for key, val in FIXES.items():
        if cm.get(key) != val:
            changed.append(f"{key}: {cm.get(key)!r} -> {val!r}")
            cm[key] = val
    return changed


def fix_yaml() -> None:
    if not os.path.exists(YAML_PATH):
        print("YAML not found:", YAML_PATH)
        return
    shutil.copy2(YAML_PATH, YAML_PATH + BAK)
    y = yaml.safe_load(open(YAML_PATH, encoding="utf-8"))
    total = []
    if isinstance(y.get("column_mappings"), dict):
        c = _apply_to_mapping(y["column_mappings"])
        total += [f"[top-level] {x}" for x in c]
    for i, e in enumerate(y.get("loan_data_extracts") or []):
        if isinstance(e, dict) and isinstance(e.get("column_mappings"), dict):
            c = _apply_to_mapping(e["column_mappings"])
            total += [f"[loan_data_extracts[{i}]] {x}" for x in c]
    with open(YAML_PATH, "w", encoding="utf-8") as fh:
        yaml.safe_dump(y, fh, sort_keys=False, default_flow_style=False,
                       allow_unicode=True, width=4096)
    print(f"YAML backed up -> {os.path.basename(YAML_PATH)}{BAK}")
    print("YAML changes:", *(["  " + x for x in total] or ["  (none)"]),
          sep="\n")


def fix_draft() -> None:
    if not os.path.exists(DRAFT_PATH):
        print("Draft not found:", DRAFT_PATH)
        return
    shutil.copy2(DRAFT_PATH, DRAFT_PATH + BAK)
    d = json.load(open(DRAFT_PATH, encoding="utf-8"))
    total = []
    if isinstance(d.get("column_mappings"), dict):
        c = _apply_to_mapping(d["column_mappings"])
        total += [f"[top-level] {x}" for x in c]
    for i, e in enumerate(d.get("loan_data_extracts") or []):
        if isinstance(e, dict) and isinstance(e.get("column_mappings"), dict):
            c = _apply_to_mapping(e["column_mappings"])
            total += [f"[loan_data_extracts[{i}]] {x}" for x in c]
    with open(DRAFT_PATH, "w", encoding="utf-8") as fh:
        json.dump(d, fh, indent=1, ensure_ascii=False)
    print(f"Draft backed up -> {os.path.basename(DRAFT_PATH)}{BAK}")
    print("Draft changes:", *(["  " + x for x in total] or ["  (none)"]),
          sep="\n")


if __name__ == "__main__":
    fix_yaml()
    print()
    fix_draft()
