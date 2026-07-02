"""One-shot fix for Central Susquehanna Comm FCU: correct
monthly_bal.pool_map for labels that were misassigned by the
keyword-heuristic auto-mapper.

Before Phase (exact-match precedence) landed, the auto-mapper matched
labels 'New Auto' and 'Used Auto' to pools 'Indirect New Auto' and
'Indirect Used Auto' respectively (both classify as new_vehicle /
used_vehicle, and the Indirect variants came first in pool_settings).
That inflated the Indirect pool totals and left the plain New/Used
pools blank on the Balance Adjustment page.

This script:
  1. Backs up the draft to .bak.phase_exact_match.
  2. Runs selfheal_auto_map_balance_labels again with the new
     exact-match precedence to correct any label whose name matches a
     pool name case-insensitively.
  3. Explicitly fixes any residual bad values by rewriting the mapping
     to prefer the exact-match pool name.
"""
from __future__ import annotations

import json
import sys
from pathlib import Path

DRAFT = Path(
    r"Z:\Shared\TCT Files\CECL - CM Files\wizard_drafts"
    r"\central_susquehanna_comm_fcu__migration.json"
)


def main() -> int:
    sys.path.insert(0, r"c:\Dev\CECL")
    from cecl_ui.services.auto_setup import selfheal_auto_map_balance_labels

    data = json.loads(DRAFT.read_text(encoding="utf-8"))
    state = data.get("state", data)
    mb = state.get("monthly_bal") or {}
    pm = mb.get("pool_map") or {}
    pool_settings = state.get("pool_settings") or []
    pool_names = [
        (ps.get("name") or "").strip()
        for ps in pool_settings
        if isinstance(ps, dict) and (ps.get("name") or "").strip()
    ]
    pool_names_ci = {nm.strip().casefold(): nm for nm in pool_names}

    backup = DRAFT.with_suffix(DRAFT.suffix + ".bak.phase_exact_match")
    if not backup.exists():
        backup.write_bytes(DRAFT.read_bytes())
        print(f"Backed up to: {backup}")

    # Fix: for every label in pool_map, if the label matches a pool name
    # exactly (case-insensitive), route it to that pool.
    changed: list[tuple[str, str, str]] = []
    for lab, cur in list(pm.items()):
        exact = pool_names_ci.get((lab or "").strip().casefold())
        if exact and (cur or "").strip() != exact:
            pm[lab] = exact
            changed.append((lab, cur or "(blank)", exact))
    mb["pool_map"] = pm
    state["monthly_bal"] = mb

    # Also run the (now-updated) auto-mapper for any blank labels.
    msgs = selfheal_auto_map_balance_labels(state)

    DRAFT.write_text(json.dumps(data, indent=2), encoding="utf-8")

    print(f"Changed {len(changed)} label mapping(s):")
    for lab, before, after in changed:
        print(f"  {lab!r}: {before!r} -> {after!r}")
    if msgs:
        for m in msgs:
            print(f"  Self-heal: {m}")

    # Show current MB pool_map after fix
    print("\nMonthly Balance pool_map after fix:")
    for k, v in (state.get("monthly_bal") or {}).get("pool_map", {}).items():
        print(f"  {k!r} -> {v!r}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
