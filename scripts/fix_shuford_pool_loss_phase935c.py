"""Phase 9.35c — restore wizard-state schema on adopted-from-config drafts.

Targets any wizard migration draft whose ``_draft_meta.adopted_from_config``
is True AND whose ``pool_settings`` field is empty / missing while
``pools`` is non-empty. Calls
``auto_setup.selfheal_adopt_yaml_schema_into_state`` on the loaded JSON
and writes it back through ``wizard_drafts.save_draft`` so the existing
timestamped-backup convention is preserved.

Idempotent: drafts whose ``pool_settings`` is already populated are
skipped.

Usage::

    python scripts\fix_shuford_pool_loss_phase935c.py            # all CUs
    python scripts\fix_shuford_pool_loss_phase935c.py shuford_fcu  # one CU

Environment::

    set CECL_WORKSPACE_ROOT=Z:\Shared\TCT Files\CECL - CM Files
    set PYTHONPATH=c:\Dev\CECL
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

# Ensure imports resolve when run as a standalone script.
_REPO_ROOT = Path(__file__).resolve().parents[1]
if str(_REPO_ROOT) not in sys.path:
    sys.path.insert(0, str(_REPO_ROOT))

os.environ.setdefault(
    "CECL_WORKSPACE_ROOT",
    r"Z:\Shared\TCT Files\CECL - CM Files",
)

from cecl_ui.services import auto_setup, wizard_drafts  # noqa: E402


def _draft_path(workspace: str, short_name: str) -> Path:
    return Path(workspace) / "wizard_drafts" / f"{short_name}__migration.json"


def restore_one(workspace: str, short_name: str) -> dict:
    """Return ``{short, ok, healed, msgs, before, after, error}``."""
    out: dict = {"short": short_name, "ok": False, "healed": False, "msgs": []}
    path = _draft_path(workspace, short_name)
    if not path.is_file():
        out["error"] = f"Draft not found at {path}"
        return out
    try:
        raw = path.read_text(encoding="utf-8")
        state = json.loads(raw)
    except Exception as exc:  # noqa: BLE001
        out["error"] = f"Failed to read draft: {exc}"
        return out

    out["before"] = {
        "pool_settings_n": len(state.get("pool_settings") or []),
        "pools_n": len(state.get("pools") or []),
        "monthly_bal_keys": sorted((state.get("monthly_bal") or {}).keys()),
    }

    msgs = auto_setup.selfheal_adopt_yaml_schema_into_state(state)
    out["msgs"] = list(msgs or [])
    out["healed"] = bool(msgs)

    out["after"] = {
        "pool_settings_n": len(state.get("pool_settings") or []),
        "pools_n": len(state.get("pools") or []),
        "monthly_bal_keys": sorted((state.get("monthly_bal") or {}).keys()),
    }

    if not msgs:
        out["ok"] = True
        return out

    # save_draft preserves the timestamped .bak_<ts> backup convention so we
    # don't need a separate hand-rolled backup file.
    try:
        wizard_drafts.save_draft(
            workspace, state,
            active_step=(state.get("_draft_meta") or {}).get("active_step") or "review",
            model="migration",
        )
        out["ok"] = True
    except Exception as exc:  # noqa: BLE001
        out["error"] = f"Failed to save: {exc}"

    return out


def list_adopted_drafts(workspace: str) -> list[str]:
    """Return short_names whose migration draft has adopted_from_config=True."""
    out: list[str] = []
    folder = Path(workspace) / "wizard_drafts"
    if not folder.is_dir():
        return out
    for p in folder.glob("*__migration.json"):
        try:
            d = json.loads(p.read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001
            continue
        meta = d.get("_draft_meta") or {}
        if meta.get("adopted_from_config"):
            out.append(p.name.split("__", 1)[0])
    return sorted(out)


def main(argv: list[str]) -> int:
    workspace = os.environ.get("CECL_WORKSPACE_ROOT") or ""
    if not workspace:
        print("ERROR: CECL_WORKSPACE_ROOT not set")
        return 2

    targets = argv[1:]
    if not targets:
        targets = list_adopted_drafts(workspace)
        if not targets:
            print("No adopted drafts found.")
            return 0
        print(f"Found {len(targets)} adopted draft(s): {', '.join(targets)}")

    healed = 0
    skipped = 0
    failed = 0
    for short in targets:
        result = restore_one(workspace, short)
        print(f"\n=== {short} ===")
        print(f"  before: {result.get('before')}")
        print(f"  after:  {result.get('after')}")
        if result.get("error"):
            print(f"  ERROR: {result['error']}")
            failed += 1
            continue
        if not result.get("healed"):
            print("  SKIP: already healed or nothing to translate")
            skipped += 1
            continue
        for m in result.get("msgs") or []:
            print(f"  msg: {m}")
        print(f"  OK: saved")
        healed += 1

    print(f"\nDone. healed={healed} skipped={skipped} failed={failed}")
    return 0 if failed == 0 else 1


if __name__ == "__main__":
    sys.exit(main(sys.argv))
