"""Smoke test: Destinations CU monthly_balances_by_pool double-count fix.

Runs auto-scan + apply_findings against the live Destinations CU folder,
then calls balance_check.monthly_balances_by_pool and reports both the
pool totals and the rows that were excluded as parent/summary rows.
"""
from __future__ import annotations

from cecl_ui.services import auto_setup, balance_check


def main() -> None:
    src = (
        r"Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access\Clients"
        r"\Destinations CU 66333\Portfolio Management\CECL Migration IDLR"
    )
    findings = auto_setup.scan_folder_for_setup(src, snapshot_yyyymm="2025-12")

    state = {
        "credit_union": "Destinations CU",
        "short_name": "destinations_cu",
        "charter_number": "66333",
        "report_period": "2025-12",
        "has_warm_files": "no",
        "monthly_bal": {},
        "pool_settings": [],
    }
    auto_setup.apply_findings_to_state(state, findings, workspace_root=r"C:\Dev\CECL")

    mb = state.get("monthly_bal") or {}
    # Simulate the worst-case: user (or hist_pool_map propagation) also
    # mapped the parent pool labels to themselves. Without exclude_labels
    # this would double the totals because the parent row holds the
    # workbook's pre-summed parent total.
    pm = mb.setdefault("pool_map", {})
    for parent_name in (mb.get("exclude_labels") or []):
        pm[parent_name] = parent_name

    print("hist_balance_source:", state.get("hist_balance_source"))
    print("monthly_bal source:", mb.get("source"))
    print("saved_path:", bool(mb.get("saved_path")))
    print("pool_map size (incl parent self-maps):", len(pm))
    print("exclude_labels:", mb.get("exclude_labels"))
    print()

    print("=== With exclude_labels (correct) ===")
    result = balance_check.monthly_balances_by_pool(state)
    total_with = 0.0
    for pool, bal in (result.get("by_pool") or {}).items():
        print(f"  {pool}: ${bal:,.2f}")
        total_with += bal
    print(f"  TOTAL: ${total_with:,.2f}")
    print()

    # Now reset exclude_labels to empty and re-run to show what would
    # happen WITHOUT the fix (the bug the user reported).
    mb["exclude_labels"] = []
    print("=== Without exclude_labels (BUG: doubled) ===")
    result_no_excl = balance_check.monthly_balances_by_pool(state)
    total_no = 0.0
    for pool, bal in (result_no_excl.get("by_pool") or {}).items():
        print(f"  {pool}: ${bal:,.2f}")
        total_no += bal
    print(f"  TOTAL: ${total_no:,.2f}")
    print()
    if total_no > total_with * 1.5:
        print(
            f"FIX CONFIRMED: bug-mode total is "
            f"${total_no - total_with:,.2f} ({total_no/total_with:.2f}x) "
            "higher than fixed-mode."
        )


if __name__ == "__main__":
    main()
