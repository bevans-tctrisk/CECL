"""One-off: fix the Impaired Loans ASC 310-10 provision-percentage (col O)
formula in the SCALE output templates.

The template formula skipped $A$8 (Foreclosed Real Estate) and instead
referenced the empty $A$10, so Foreclosed Real Estate loans resolved to a
0% provision instead of 100%.

Buggy   : ...IF(A{r}=$A$7,$B$7,IF(A{r}=$A$9,$B$9,IF(A{r}=$A$10,$B$10,0)))...
Correct : ...IF(A{r}=$A$7,$B$7,IF(A{r}=$A$8,$B$8,IF(A{r}=$A$9,$B$9,0)))...

Edits the worksheet XML directly (inside the .xlsx zip) so charts,
images and formatting are preserved. Run with --apply to write; default
is a dry-run scan.
"""
from __future__ import annotations

import glob
import io
import os
import re
import shutil
import sys
import zipfile
from datetime import datetime

TEMPLATE_GLOB = "cecl_ui/data/scale/templates/*_CECL_SCALE_template.xlsx"

# Match a full O-column formula cell's buggy tail for a given row and
# rewrite it. Row number is captured so the A-refs stay consistent.
BUGGY = re.compile(
    r"IF\(A(\d+)=\$A\$9,\$B\$9,IF\(A\1=\$A\$10,\$B\$10,0\)\)"
)


def _fix_row(m: re.Match) -> str:
    r = m.group(1)
    return f"IF(A{r}=$A$8,$B$8,IF(A{r}=$A$9,$B$9,0))"


def _sheet_targets(z: zipfile.ZipFile) -> dict[str, str]:
    wbxml = z.read("xl/workbook.xml").decode("utf8")
    rels = z.read("xl/_rels/workbook.xml.rels").decode("utf8")
    relmap = dict(re.findall(r'Id="([^"]+)"[^>]*Target="([^"]+)"', rels))
    # some writers order Target before Id
    relmap.update(dict(
        (rid, tgt) for tgt, rid in
        re.findall(r'Target="([^"]+)"[^>]*Id="([^"]+)"', rels)
    ))
    out: dict[str, str] = {}
    for name, rid in re.findall(
        r'<sheet[^>]*name="([^"]+)"[^>]*r:id="([^"]+)"', wbxml
    ):
        tgt = relmap.get(rid, "")
        if not tgt.endswith(".xml"):
            continue
        path = tgt.lstrip("/")
        if not path.startswith("xl/"):
            path = "xl/" + path
        out[name] = path
    return out


def process(tpl: str, apply: bool) -> int:
    with zipfile.ZipFile(tpl) as z:
        targets = _sheet_targets(z)
        names = z.namelist()
        edits: dict[str, str] = {}
        total = 0
        for name, path in targets.items():
            if path not in names:
                continue
            data = z.read(path).decode("utf8")
            new, n = BUGGY.subn(_fix_row, data)
            if n:
                total += n
                print(f"   sheet {name!r} ({path}): {n} cells")
                edits[path] = new
        if not apply or not edits:
            return total
        # Rewrite the zip in memory, replacing only edited parts, then
        # overwrite the file in place (os.replace can raise Access Denied
        # on cloud-synced / junctioned dirs even when the file is writable).
        bak = f"{tpl}.bak.provfix_{datetime.now():%Y%m%d%H%M%S}"
        shutil.copy2(tpl, bak)
        print(f"   backup -> {os.path.basename(bak)}")
        buf = io.BytesIO()
        with zipfile.ZipFile(tpl) as zin, \
                zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                raw = zin.read(item.filename)
                if item.filename in edits:
                    raw = edits[item.filename].encode("utf8")
                zout.writestr(item, raw)
        with open(tpl, "wb") as f:
            f.write(buf.getvalue())
    return total


def main() -> None:
    apply = "--apply" in sys.argv
    grand = 0
    for tpl in sorted(glob.glob(TEMPLATE_GLOB)):
        print("===", os.path.basename(tpl))
        grand += process(tpl, apply)
    mode = "APPLIED" if apply else "DRY-RUN (use --apply to write)"
    print(f"\nTotal buggy cells: {grand}  [{mode}]")


if __name__ == "__main__":
    main()
