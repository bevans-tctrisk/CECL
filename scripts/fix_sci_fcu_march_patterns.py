"""Phase SCI FCU (2026-07-01) — Relax file_pattern for both Aires and
Credit Card extracts so subsequent-quarter files match.

Root cause: Vizo IDLR renamed both extracts between Dec 2025 baseline
and Jan/Feb/Mar 2026 refreshes. The wizard-generated patterns had literals
(`PII`, month-name-only date, IDLR job ID `TS00491933`, literal
`12-31-2025`) that only matched the Dec baseline. Same class of bug as
Phase 9.12 / 9.31 / 9.33 / 9.35.

Change:
  * Aires: accept BOTH `Aires Monthly Loans without Credit Cards[_PII]`
    AND `Aires Loans for Vizo`; accept BOTH month-name+year AND
    MM-DD-YYYY / MM_DD_YYYY / MM DD YYYY dates.
  * Credit Card: accept any filename containing `Credit Card` and/or
    `CECL Report` with an MM-DD-YYYY date. IDLR job ID `TS00491933`
    optional and generalized to `TS\\d+`.

Idempotent (backs up to .bak.phase_sci_march_patterns). No-op if the new
patterns are already present.
"""
from __future__ import annotations

import shutil
import sys
import re
from pathlib import Path

import yaml

WS = Path(r"Z:\Shared\TCT Files\CECL - CM Files")
YAML_PATH = WS / "client_configs" / "sci_fcu.yaml"
BACKUP = YAML_PATH.with_suffix(YAML_PATH.suffix + ".bak.phase_sci_march_patterns")

# Aires: two accepted naming families (old + new March "for Vizo" form);
# date is either MM[-_ ]DD[-_ ]YYYY or Month YYYY. Trailing v[.NN?].xlsx.
AIRES_PATTERN = (
    r"(?i)^Aires[\s_\-]+"
    r"(?:Monthly[\s_\-]+Loans[\s_\-]+without[\s_\-]+Credit[\s_\-]+Cards"
    r"|Loans[\s_\-]+for[\s_\-]+Vizo)"
    r".*?"
    r"(?:\d{1,2}[\-_\s]\d{1,2}[\-_\s]20\d{2}"
    r"|(?:jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may"
    r"|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?"
    r"|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?)[\s_\-]+20\d{2})"
    r"v(?:\.\d{1,3})?\.(xlsx|xls)$"
)

# Credit Card: filename must START with `Credit Card` or `CECL Credit
# Card Report` or `CECL Report`, followed by any middle content
# (optional IDLR job-ID `TS\d+`), MM-DD-YYYY date, and trailing
# v[.NN?].xlsx. Anchored to `^` so Aires files (which contain
# "Credit Cards" mid-name) don't get double-matched.
CREDIT_CARD_PATTERN = (
    r"(?i)^(?:Credit[\s_\-]+Card|CECL[\s_\-]+(?:Credit[\s_\-]+Card[\s_\-]+)?Report)"
    r".*?"
    r"\d{1,2}[\-_\s]\d{1,2}[\-_\s]20\d{2}"
    r"v(?:\.\d{1,3})?\.(xlsx|xls)$"
)


def main() -> int:
    if not YAML_PATH.exists():
        print(f"ERROR: {YAML_PATH} not found", file=sys.stderr)
        return 1
    cfg = yaml.safe_load(YAML_PATH.read_text(encoding="utf-8"))
    if not cfg:
        print("ERROR: YAML is empty", file=sys.stderr)
        return 1

    if not BACKUP.exists():
        shutil.copy2(YAML_PATH, BACKUP)
        print(f"Backup written: {BACKUP}")
    else:
        print(f"Backup already exists (skipping): {BACKUP}")

    extracts = cfg.get("loan_data_extracts") or []
    changes: list[str] = []

    for i, ex in enumerate(extracts):
        old = ex.get("file_pattern", "")
        # Aires extract: starts with Aires
        if re.match(r"\(\?i\)\^Aires", old):
            if old != AIRES_PATTERN:
                ex["file_pattern"] = AIRES_PATTERN
                changes.append(f"extract[{i}] Aires: pattern relaxed")
            else:
                changes.append(f"extract[{i}] Aires: already up to date")
        # Credit Card extract: contains Credit Card or CECL
        elif re.search(r"(?i)credit.*card|cecl", old):
            if old != CREDIT_CARD_PATTERN:
                ex["file_pattern"] = CREDIT_CARD_PATTERN
                changes.append(f"extract[{i}] Credit Card: pattern relaxed")
            else:
                changes.append(f"extract[{i}] Credit Card: already up to date")

    # Also patch top-level file_pattern to Aires (that's the "primary" extract)
    if cfg.get("file_pattern") != AIRES_PATTERN:
        cfg["file_pattern"] = AIRES_PATTERN
        changes.append("top-level: pattern relaxed to Aires (permissive)")
    else:
        changes.append("top-level: already up to date")

    # Write it out
    YAML_PATH.write_text(
        yaml.safe_dump(cfg, sort_keys=False, allow_unicode=True, width=1000),
        encoding="utf-8",
    )

    print()
    print("Changes:")
    for c in changes:
        print(f"  - {c}")

    # Validate: audit Raw_Uploads against new patterns
    raw = WS / "Raw_Uploads" / "sci_fcu"
    print()
    print(f"Raw_Uploads audit ({raw}):")
    if not raw.exists():
        print("  (does not exist)")
        return 0
    patterns = [(cfg["file_pattern"], "top-level")]
    for i, ex in enumerate(cfg.get("loan_data_extracts", [])):
        patterns.append((ex["file_pattern"], f"extract[{i}]"))
    for f in sorted(raw.iterdir()):
        if not f.is_file():
            continue
        matched = [name for pat, name in patterns if re.search(pat, f.name)]
        tag = ",".join(matched) if matched else "NO MATCH"
        print(f"  {f.name}  ->  {tag}")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
