"""Regression check: Census FCU (non-hierarchical balance workbook) is unaffected.

Census uses an annual 'YYYY Balance Sheets.xlsx' layout with no parent/sub
hierarchy, so exclude_labels should remain empty and pool totals unchanged.
"""
from __future__ import annotations

from cecl_ui.services import auto_setup, balance_check


def main() -> None:
    src = (
        r"Z:\Shared\Clients\Vizo Financial Corporate CU\Client Access\Clients"
        r"\Census FCU 5641\Portfolio Management\CECL Migration IDLR"
    )
    findings = auto_setup.scan_folder_for_setup(src, snapshot_yyyymm="2025-12")
    state = {
        "credit_union": "Census FCU",
        "short_name": "census_fcu",
        "charter_number": "5641",
        "report_period": "2025-12",
        "has_warm_files": "no",
        "monthly_bal": {},
        "pool_settings": [],
    }
    auto_setup.apply_findings_to_state(state, findings, workspace_root=r"C:\Dev\CECL")
    mb = state.get("monthly_bal") or {}
    print("hist_balance_source:", state.get("hist_balance_source"))
    print("monthly_bal source:", mb.get("source"))
    print("pool_map size:", len(mb.get("pool_map") or {}))
    print("exclude_labels:", mb.get("exclude_labels"))
    print()
    result = balance_check.monthly_balances_by_pool(state)
    print("ok:", result.get("ok"), " period:", result.get("period"))
    print("error:", result.get("error"))
    if result.get("ok"):
        total = 0.0
        for pool, bal in (result.get("by_pool") or {}).items():
            print(f"  {pool}: ${bal:,.2f}")
            total += bal
        print(f"  TOTAL: ${total:,.2f}")


if __name__ == "__main__":
    main()
