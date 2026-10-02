"""Verify apply_pool_grouping_choice rebuilds pool_settings + pool_map cleanly."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from cecl_ui.services import auto_setup

FOLDER = (
    r"Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access\Clients"
    r"\Destinations CU 66333\Portfolio Management\CECL Migration IDLR"
)

# Seed an empty state from the auto-scan.
findings = auto_setup.scan_folder_for_setup(FOLDER, "202512")
state = {"credit_union": "Destinations CU", "short_name": "destinations_cu"}
auto_setup.apply_findings_to_state(state, findings, r"c:\Dev\CECL")

print("=== Initial pool_settings (after auto-seed) ===")
for p in state.get("pool_settings") or []:
    print(f"  {p['name']!r}  acl_months={p['acl_months']}")
print(f"  -> {len(state.get('pool_settings') or [])} pool(s)")

print("\n=== Test 1: keep all (default) ===")
choices = {p["name"]: "keep" for p in state["_pool_seed"]["pools"]}
result = auto_setup.apply_pool_grouping_choice(state, choices)
print(f"  result: {result}")
for p in state["pool_settings"]:
    print(f"  {p['name']!r}")

print("\n=== Test 2: split 'All Other Unsecured Loans/LOC' ===")
choices = {p["name"]: "keep" for p in state["_pool_seed"]["pools"]}
choices["All Other Unsecured Loans/LOC"] = "split"
result = auto_setup.apply_pool_grouping_choice(state, choices)
print(f"  result: {result}")
for p in state["pool_settings"]:
    print(f"  {p['name']!r}")

print("\n=== monthly_bal.pool_map after split ===")
pm = (state.get("monthly_bal") or {}).get("pool_map") or {}
for label, pool in sorted(pm.items()):
    marker = "  ** " if pool != label and pool == "All Other Unsecured Loans/LOC" else "     "
    marker = "  ** " if pool == label else marker
    print(f"{marker}{label!r} -> {pool!r}")

print("\n=== Test 3: switch back to keep ===")
choices = {p["name"]: "keep" for p in state["_pool_seed"]["pools"]}
result = auto_setup.apply_pool_grouping_choice(state, choices)
print(f"  result: {result}")
print(f"  pool count back to: {len(state['pool_settings'])}")
print("  'All Other Unsecured Loans/LOC' present:",
      any(p["name"] == "All Other Unsecured Loans/LOC" for p in state["pool_settings"]))
pm = state["monthly_bal"]["pool_map"]
print("  'Signature Loans' ->", pm.get("Signature Loans"))
print("  'Courtesy Pay' ->", pm.get("Courtesy Pay"))

print("\n=== Test 4: split EVERY pool with sub-cats ===")
choices = {p["name"]: "split" for p in state["_pool_seed"]["pools"]}
result = auto_setup.apply_pool_grouping_choice(state, choices)
print(f"  result: {result}")
print(f"  pool count: {len(state['pool_settings'])}")
print("  Expected: 24 (all sub-categories become pools)")
pm = state["monthly_bal"]["pool_map"]
print("  Every label maps to itself?",
      all(label == pool for label, pool in pm.items()))

print("\n=== Test 5: preserve user-edited ACL months on rebuild ===")
state["pool_settings"][0]["acl_months"] = 99
print(f"  Set first pool acl_months=99 (was {99})")
first_name = state["pool_settings"][0]["name"]
print(f"  First pool name: {first_name!r}")
choices = {p["name"]: "keep" for p in state["_pool_seed"]["pools"]}
auto_setup.apply_pool_grouping_choice(state, choices)
# Find that pool in the rebuilt list -- it should still have acl=99 if it's a parent name
match = next((p for p in state["pool_settings"] if p["name"] == first_name), None)
if match:
    print(f"  After rebuild: {match['name']!r} acl_months={match['acl_months']}")
else:
    print(f"  {first_name!r} not in rebuilt list (sub-cat became orphan after switching back to keep)")
