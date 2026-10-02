# Vendored report assets

Snapshot of the analyst-owned files in
`<CECL_WORKSPACE_ROOT>/Sample Reports/` that `report_vizo.py` needs to build a
correct Vizo Migration report. That folder is gitignored (`.gitignore:45`), so
before this directory existed a clean clone could not produce a complete
report — the narrative tabs, the Vizo colour theme and the info icons were all
silently omitted, with no error.

| File | Used for | Source |
|---|---|---|
| `Vizo Narrative Tabs - Template.xlsx` | the approved `Introduction-Vizo` and `Executive Summary-Vizo` verbiage, copied cell-by-cell | `Sample Reports/` |
| `vizo_theme1.xml` | the Vizo colour theme (accent1 teal etc.) | `xl/theme/theme1.xml` inside `YYYY-MM CECL-Migration-WARM - Template Credit Union with Vizo.xlsx` |
| `assets/info_*.png` | info icons embedded in report tabs | `Sample Reports/assets/` |

Note this is a *different* theme from the repo-root `vizo_theme.xml`, which is
not interchangeable with it.

## Precedence

The workspace copy wins where it exists, so the analyst workspace remains the
edit point for approved copy; these files are the fallback that makes a clean
clone work. **When the approved text or branding changes on the share, refresh
this snapshot** — otherwise a machine without the share silently renders stale
verbiage.

The 7 MB master template and the two sample PDFs are deliberately not vendored:
nothing in the code path reads them beyond the theme extracted above.
