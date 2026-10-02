"""Fix the Calc tab fill error in the SCALE template workbooks.

Rows 75-80 of the Calc tab use a running-total pattern
    =<col><row-17> + <nextcol><samerow>
but AC75/AD75/AC76/AD76 shipped with a stray SUM() range that
double-counts the AE column. Normalize those four cells to the
pattern used by every other cell in the block.

Surgical zip-level XML replacement so cell styles, the logo drawing,
and all other content are preserved byte-for-byte.
"""
from __future__ import annotations

import shutil
import zipfile
from pathlib import Path

TEMPLATE_DIR = Path("cecl_ui/data/scale/templates")

# old formula text -> new formula text (as stored inside <f>...</f>)
REPLACEMENTS = {
    "AC58+SUM(AD75:$AE75)": "AC58+AD75",
    "AD58+SUM(AE75:$AE75)": "AD58+AE75",
    "AC59+SUM(AD76:$AE76)": "AC59+AD76",
    "AD59+SUM(AE76:$AE76)": "AD59+AE76",
}


def _target_sheet_part(zf: zipfile.ZipFile) -> str | None:
    for name in zf.namelist():
        if name.startswith("xl/worksheets/sheet") and name.endswith(".xml"):
            data = zf.read(name).decode("utf8")
            if "AC58+SUM(AD75:$AE75)" in data:
                return name
    return None


def fix_template(path: Path) -> bool:
    with zipfile.ZipFile(path, "r") as zf:
        part = _target_sheet_part(zf)
        if part is None:
            print(f"  SKIP {path.name}: buggy formula not found")
            return False
        text = zf.read(part).decode("utf8")

    for old, new in REPLACEMENTS.items():
        count = text.count(old)
        if count != 1:
            raise SystemExit(
                f"ABORT {path.name}: expected exactly 1 occurrence of "
                f"{old!r} in {part}, found {count}"
            )
        text = text.replace(old, new)

    backup = path.with_suffix(path.suffix + ".bak_calcfix")
    if not backup.exists():
        shutil.copy2(path, backup)

    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path, "r") as zin, zipfile.ZipFile(
        tmp, "w", zipfile.ZIP_DEFLATED
    ) as zout:
        for item in zin.infolist():
            if item.filename == part:
                zout.writestr(item, text.encode("utf8"))
            else:
                zout.writestr(item, zin.read(item.filename))
    tmp.replace(path)
    print(f"  FIXED {path.name} (sheet part {part})")
    return True


def main() -> None:
    templates = sorted(TEMPLATE_DIR.glob("*_CECL_SCALE_template.xlsx"))
    if not templates:
        raise SystemExit(f"No templates found under {TEMPLATE_DIR}")
    for tpl in templates:
        print(f"Processing {tpl.name}")
        fix_template(tpl)


if __name__ == "__main__":
    main()
