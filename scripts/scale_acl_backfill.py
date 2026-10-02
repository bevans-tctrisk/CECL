r"""Walk every generated SCALE workbook on disk and upsert its
``Scale Calculation!U29..U33`` cached values into the
``scale_acl_history`` PostgreSQL table.

Usage::

    python scripts/scale_acl_backfill.py             # all CUs
    python scripts/scale_acl_backfill.py greensboro_cu
    python scripts/scale_acl_backfill.py --dry-run

The backfill only succeeds for workbooks that have ALREADY been
opened-and-saved by Excel (so cached formula values exist on disk),
or that were generated on a machine where pywin32 was installed and
the runner's post-run recalc populated them. Workbooks whose cached
values are still ``None`` are reported as "skipped (no cached values
-- open in Excel + Save)" so the analyst knows what to do.

Idempotent. Re-running picks up newly-saved cached values without
clobbering rows the runner already wrote.
"""
from __future__ import annotations

import os
import sys
from pathlib import Path

# Ensure standalone-script env vars are set BEFORE importing cecl_ui.
os.environ.setdefault(
    "CECL_WORKSPACE_ROOT",
    r"Z:\Shared\TCT Files\CECL - CM Files",
)
sys.path.insert(0, r"c:\Dev\CECL")

from cecl_ui.services.scale import acl_history_db, runs_service


_PERIOD_DIR_RX = __import__("re").compile(r"^(20\d{2})-(0[1-9]|1[0-2])$")


def _workspace_root() -> Path:
    return Path(os.environ["CECL_WORKSPACE_ROOT"]).resolve()


def _generated_root() -> Path:
    return _workspace_root() / "Generated_Reports"


def _iter_cu_dirs(filter_short: str | None = None):
    root = _generated_root()
    if not root.exists():
        return
    for child in sorted(root.iterdir()):
        if not child.is_dir():
            continue
        if filter_short and child.name != filter_short:
            continue
        yield child


def _iter_period_dirs(cu_dir: Path):
    for child in sorted(cu_dir.iterdir()):
        if not child.is_dir():
            continue
        if not _PERIOD_DIR_RX.match(child.name):
            continue
        yield child


def _pick_workbook(period_dir: Path) -> Path | None:
    """Prefer the Vizo variant; otherwise the newest .xlsx."""
    candidates = [
        p for p in period_dir.glob("*SCALE*.xlsx")
        if "~$" not in p.name  # skip Excel lock files
    ]
    if not candidates:
        return None
    vizo = [p for p in candidates if "_Vizo" in p.name]
    if vizo:
        return max(vizo, key=lambda p: p.stat().st_mtime)
    return max(candidates, key=lambda p: p.stat().st_mtime)


def main(argv: list[str]) -> int:
    args = [a for a in argv[1:] if a]
    dry_run = "--dry-run" in args
    args = [a for a in args if not a.startswith("--")]
    filter_short = args[0] if args else None

    workspace_root = _workspace_root()
    if not _generated_root().exists():
        print(f"Generated_Reports not found at {_generated_root()}.")
        return 1

    print(f"Workspace: {workspace_root}")
    if dry_run:
        print("(dry-run -- will report what would be captured but not write to DB)")
    if filter_short:
        print(f"Filter: {filter_short!r}")

    try:
        acl_history_db.ensure_table()
    except Exception as exc:  # noqa: BLE001
        print(f"Could not ensure scale_acl_history table: {exc}")
        return 2

    total_captured = 0
    total_skipped = 0
    total_failed = 0

    for cu_dir in _iter_cu_dirs(filter_short):
        cu_short = cu_dir.name
        print(f"\n=== {cu_short} ===")
        for period_dir in _iter_period_dirs(cu_dir):
            period = period_dir.name  # YYYY-MM
            wb = _pick_workbook(period_dir)
            if wb is None:
                print(f"  {period}: no SCALE workbook found")
                continue
            # Read cached values via runs_service helper.
            read = runs_service.read_prior_acl_values(wb)
            if not read.get("ok") or not read.get("values"):
                err = (read.get("error") or "no cached values").strip()
                print(f"  {period}: skipped {wb.name} ({err})")
                total_skipped += 1
                continue
            values = read["values"]
            label_bits = []
            for coord, label in (
                ("AY2", "TEL"), ("AY3", "ACL"), ("AY4", "Adj"),
                ("AY5", "Ratio"),
            ):
                v = values.get(coord)
                if v is None:
                    label_bits.append(f"{label}=N/A")
                elif coord == "AY5":
                    label_bits.append(f"{label}={v:.4%}")
                else:
                    label_bits.append(f"{label}=${v:,.2f}")
            preview = ", ".join(label_bits)
            if dry_run:
                print(f"  {period}: would capture from {wb.name} ({preview})")
                total_captured += 1
                continue
            up = acl_history_db.upsert(
                cu_short, period, values,
                source=f"backfill-script:{wb.name}",
            )
            if up.get("ok"):
                total_captured += 1
                print(f"  {period}: captured from {wb.name} ({preview})")
            else:
                total_failed += 1
                print(
                    f"  {period}: FAILED to upsert from {wb.name}: "
                    f"{up.get('error') or ''}"
                )

    print(
        f"\nSummary: captured={total_captured} "
        f"skipped={total_skipped} failed={total_failed}"
    )
    return 0 if total_failed == 0 else 3


if __name__ == "__main__":
    raise SystemExit(main(sys.argv))
