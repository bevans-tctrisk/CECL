"""Phase SCC.credit-join — Central Susquehanna Comm FCU Dec 2025 credit-movement fix.

Root cause: YAML mapped `member_number: Member Name` (a text names column
like `TARIQ UMAR`) instead of `Account Number` (values like `10053L03`).
The credit-pull loader loaded 592 scores keyed by numeric member number,
but `member_numbers` in `import_file` came from coercing names to int
(=0 for every row), so `member_numbers.map(credit_pull_scores)` matched
zero rows and the priority chain fell through to `current_fico = original_fico`
for all 534 loans.

Fix: rewrite top-level + per-extract column_mappings.member_number from
`Member Name` to `Account Number`. The `member_account.mode: delimiter`
(delimiter=L, suffix_length=3) already handles account splitting so no
other config changes are needed.

Idempotent: backs up YAML to `.bak.phase_credit_join`, patches only when
mapping is currently `Member Name`.
"""
from __future__ import annotations

import io
import os
import shutil
import sys
from pathlib import Path

import yaml

os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")
sys.path.insert(0, r"c:\Dev\CECL")

YAML_PATH = Path(r"Z:\Shared\TCT Files\CECL - CM Files\client_configs\central_susquehanna_comm_fcu.yaml")
BACKUP_PATH = YAML_PATH.with_suffix(YAML_PATH.suffix + ".bak.phase_credit_join")

BAD_COL = "Member Name"
GOOD_COL = "Account Number"


def patch_yaml() -> bool:
    if not YAML_PATH.exists():
        print(f"ERROR: YAML not found at {YAML_PATH}")
        return False

    if not BACKUP_PATH.exists():
        shutil.copy2(YAML_PATH, BACKUP_PATH)
        print(f"Backup: {BACKUP_PATH}")
    else:
        print(f"Backup already exists: {BACKUP_PATH}")

    cfg = yaml.safe_load(YAML_PATH.read_text(encoding="utf-8"))
    changes: list[str] = []

    top_cm = cfg.get("column_mappings") or {}
    if top_cm.get("member_number") == BAD_COL:
        top_cm["member_number"] = GOOD_COL
        changes.append(f"top-level column_mappings.member_number: {BAD_COL!r} -> {GOOD_COL!r}")

    for i, ex in enumerate(cfg.get("loan_data_extracts") or []):
        ex_cm = ex.get("column_mappings") or {}
        if ex_cm.get("member_number") == BAD_COL:
            ex_cm["member_number"] = GOOD_COL
            changes.append(
                f"loan_data_extracts[{i}].column_mappings.member_number: {BAD_COL!r} -> {GOOD_COL!r}"
            )

    if not changes:
        print("No changes required — YAML already patched (idempotent).")
        return True

    YAML_PATH.write_text(yaml.safe_dump(cfg, sort_keys=False), encoding="utf-8")
    print("Changes applied:")
    for c in changes:
        print(f"  {c}")
    return True


def run_import_and_reports() -> None:
    import cecl_credentials

    os.environ["DATABASE_URL"] = cecl_credentials.get_database_url()
    from cecl_ui.services import pipeline_service

    print("\n=== Import ===")
    buf = io.StringIO()
    import contextlib

    with contextlib.redirect_stdout(buf):
        n = pipeline_service.run_import("central_susquehanna_comm_fcu")
    output = buf.getvalue()
    print(output)
    print(f"\nImported: {n}")

    # Verify DB movement
    from import_data import engine
    from sqlalchemy import text

    print("\n=== DB verification (2025-12-31) ===")
    with engine.begin() as conn:
        row = conn.execute(
            text(
                """
                SELECT
                    COUNT(*) AS total,
                    SUM(CASE WHEN original_fico_score IS NOT NULL
                          AND current_fico_score IS NOT NULL
                          AND original_fico_score <> current_fico_score
                        THEN 1 ELSE 0 END) AS differ,
                    SUM(CASE WHEN original_fico_score = current_fico_score
                        THEN 1 ELSE 0 END) AS same
                FROM monthly_loan_data
                WHERE credit_union = 'Central Susquehanna Comm FCU'
                  AND snapshot_date = '2025-12-31'
                """
            )
        ).mappings().one()
    print(f"total={row['total']} differ={row['differ']} same={row['same']}")

    if int(row["differ"] or 0) == 0:
        print("\nWARNING: no credit movement yet — investigating further needed.")
        return

    print("\n=== Reports ===")
    reports = pipeline_service.run_reports(
        "central_susquehanna_comm_fcu",
        "2025-12-31",
        ["vizo", "vizo_supp"],
    )
    print(reports)


if __name__ == "__main__":
    if patch_yaml():
        run_import_and_reports()
