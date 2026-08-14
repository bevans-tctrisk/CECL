"""Phase 9.34 Nucor Emp CU — restore Mar 2026 extract from Archive,
wildcard stale per-extract file_patterns, bump report_period, re-import.

Mirror of Phase 9.33 NOVA fix. Idempotent.
"""
from __future__ import annotations

import os
import re
import shutil
from pathlib import Path

os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")

import yaml  # noqa: E402

WS = Path(os.environ["CECL_WORKSPACE_ROOT"])
YAML_PATH = WS / "client_configs" / "nucor_emp_cu.yaml"
RAW = WS / "Raw_Uploads" / "nucor_emp_cu"
ARCH = WS / "Archive" / "nucor_emp_cu"


# Match: Aires Loans <month> <2-or-4-digit-year> V<n>.xlsx (any separator,
# any version, any extension xlsx|xls). Handles space, underscore, dash.
_AIRES_LOANS_RX = (
    r"(?i)^Aires[\s_]+Loans?[\s_]+"
    r"(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
    r"jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|"
    r"nov(?:ember)?|dec(?:ember)?)[\s_\-\.]+\d{2,4}[\s_]+[Vv]\d{1,2}"
)

# Match: any Nucor Visa file — either "Nucor Visa <month> <year> V<n>" or
# "Nucor Emp CU <month> <year> Visa File V<n>" (Vizo IDLR variants).
_VISA_RX = (
    r"(?i)^Nucor[\s_]+(?:Emp[\s_]+CU[\s_]+)?"
    r"(?:Visa[\s_]+)?"
    r"(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|"
    r"jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|"
    r"nov(?:ember)?|dec(?:ember)?)[\s_\-\.]+\d{2,4}[\s_]+"
    r"(?:Visa[\s_]+(?:File[\s_]+)?)?[Vv]\d{1,2}"
)


def main() -> None:
    print("=" * 78)
    print("Phase 9.34 — Nucor Emp CU credit pull / file_pattern fix")
    print("=" * 78)

    # ----- backup YAML -----
    bak = YAML_PATH.with_suffix(YAML_PATH.suffix + ".bak.phase934_nucor")
    if not bak.exists():
        shutil.copy2(YAML_PATH, bak)
        print(f"[backup] YAML backed up to {bak.name}")
    else:
        print(f"[backup] YAML backup already exists: {bak.name}")

    cfg = yaml.safe_load(YAML_PATH.read_text())

    # ----- 1) report_period -----
    old_rp = cfg.get("report_period")
    if old_rp != "2026-03":
        cfg["report_period"] = "2026-03"
        print(f"[yaml ] report_period: {old_rp!r} -> '2026-03'")
    else:
        print("[yaml ] report_period already '2026-03'")

    # ----- 2) per-extract patterns -----
    extracts = cfg.get("loan_data_extracts") or []
    new_aires = _AIRES_LOANS_RX
    new_visa = _VISA_RX
    for i, ex in enumerate(extracts):
        old_pat = ex.get("file_pattern") or ""
        label = ex.get("label", "?")
        if "Aires" in label or "aires" in old_pat.lower():
            if old_pat != new_aires:
                ex["file_pattern"] = new_aires
                print(f"[yaml ] extract[{i}] '{label}' file_pattern wildcarded")
            else:
                print(f"[yaml ] extract[{i}] '{label}' already wildcarded")
        elif "Visa" in label or "visa" in old_pat.lower():
            if old_pat != new_visa:
                ex["file_pattern"] = new_visa
                print(f"[yaml ] extract[{i}] '{label}' file_pattern wildcarded")
            else:
                print(f"[yaml ] extract[{i}] '{label}' already wildcarded")

    YAML_PATH.write_text(yaml.dump(cfg, sort_keys=False, default_flow_style=False))
    print(f"[yaml ] saved {YAML_PATH.name}")

    # ----- 3) restore archived files to Raw_Uploads -----
    RAW.mkdir(parents=True, exist_ok=True)
    restored = 0
    if ARCH.exists():
        for p in sorted(ARCH.iterdir()):
            if not p.is_file():
                continue
            # Only restore loan extracts (Aires Loans + Visa). Skip the
            # December baseline (already imported with credit pull working).
            name_l = p.name.lower()
            if "aires" in name_l and "loans" in name_l:
                # restore only March-or-later (don't disturb Dec 2025)
                if "mar" in name_l and "2026" in name_l:
                    dest = RAW / p.name
                    if not dest.exists():
                        shutil.copy2(p, dest)
                        print(f"[copy ] Archive -> Raw_Uploads: {p.name}")
                        restored += 1
                    else:
                        print(f"[copy ] already in Raw_Uploads: {p.name}")
    print(f"[copy ] restored {restored} file(s) total")

    # ----- 4) validate the new patterns match the restored files -----
    aires_rx = re.compile(new_aires)
    visa_rx = re.compile(new_visa)
    print("\n[validate] match check against Raw_Uploads:")
    for p in sorted(RAW.iterdir()):
        if not p.is_file():
            continue
        ok_aires = "MATCH" if aires_rx.match(p.name) else "no   "
        ok_visa = "MATCH" if visa_rx.match(p.name) else "no   "
        print(f"  aires={ok_aires}  visa={ok_visa}  {p.name}")


if __name__ == "__main__":
    main()
