"""Phase 9.35d live remediation: seed missing top-level scalar defaults
into Shuford's adopted draft so the user can navigate Step 12 (File format)
without hitting KeyError: 'pool_code_split'.

Idempotent. Backs up the draft to `.bak.phase935d` before writing.
"""
from __future__ import annotations

import json
import os
import shutil
import sys
from pathlib import Path

# Match wizard env conventions
os.environ.setdefault(
    "CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files"
)
sys.path.insert(0, r"c:\Dev\CECL")

from cecl_ui.services.auto_setup import (  # noqa: E402
    selfheal_adopt_yaml_schema_into_state,
)


def main(short: str = "shuford_fcu") -> int:
    ws = Path(os.environ["CECL_WORKSPACE_ROOT"])
    draft = ws / "wizard_drafts" / f"{short}__migration.json"
    if not draft.exists():
        print(f"ERROR: draft not found at {draft}")
        return 1

    backup = draft.with_suffix(draft.suffix + ".bak.phase935d")
    if not backup.exists():
        shutil.copy2(draft, backup)
        print(f"Backed up draft -> {backup.name}")
    else:
        print(f"Backup already exists at {backup.name}; not overwriting.")

    payload = json.loads(draft.read_text())
    # wizard_drafts payloads wrap state under a 'state' key in newer versions,
    # but earlier shapes wrote state at top level. Handle both.
    if "state" in payload and isinstance(payload["state"], dict):
        state = payload["state"]
        wrapper = payload
    else:
        state = payload
        wrapper = None

    before = set(state.keys())
    msgs = selfheal_adopt_yaml_schema_into_state(state)
    after = set(state.keys())
    new_keys = sorted(after - before)

    print(f"Self-heal returned {len(msgs)} message(s):")
    for m in msgs:
        print(f"  - {m}")
    print(f"Seeded {len(new_keys)} new top-level key(s): {new_keys}")

    if wrapper is not None:
        wrapper["state"] = state
        draft.write_text(json.dumps(wrapper, indent=2))
    else:
        draft.write_text(json.dumps(state, indent=2))
    print(f"Wrote updated draft -> {draft.name}")
    return 0


if __name__ == "__main__":
    sys.exit(main(*sys.argv[1:]))
