# Scripts

## `generate_scale_mapping.py`

Generates a new SCALE quarterly mapping CSV by advancing the cell column
of an existing one. This is how new quarters get added without hand-editing
each of the 75 rows.

The hand-maintained `cecl_ui/data/scale/maps/YYYY_MM_mapping.csv` files
only differ from one quarter to the next by the cell column each row
writes to (column AY = 12/31/2025, AZ = 3/31/2026, BA = 6/30/2026, ...).
The header row already embeds the full quarter-to-column lookup, so the
new quarter can be derived deterministically from any existing one.

### Add the next quarter

From the repo root:

```powershell
# Default: source = newest existing mapping
.\.venv\Scripts\python.exe scripts\generate_scale_mapping.py 2026-09

# Show what would be written, don't touch disk
.\.venv\Scripts\python.exe scripts\generate_scale_mapping.py 2026-09 --dry-run

# Pick an explicit source
.\.venv\Scripts\python.exe scripts\generate_scale_mapping.py 2026-09 --source 2026-06

# Overwrite an existing target
.\.venv\Scripts\python.exe scripts\generate_scale_mapping.py 2026-09 --force
```

The script validates that:
- `field_code` and `sheet` columns are untouched (rows match the source 1:1)
- Each row's existing cell column matches the source column letter
- The source's `Report Date` / `Column` metadata is internally consistent

If any row is inconsistent, the script aborts loudly rather than producing
a half-rewritten file.

### Per-quarter workflow

1. **Generate the mapping**: run the script above for the new quarter.
2. **Copy the template**: save the prior period's
   `cecl_ui/data/scale/templates/{YYYY_MM}_CECL_SCALE_template.xlsx` as
   the new quarter's template (rename only — the structure is identical
   quarter-to-quarter; the SCALE workbook itself looks up data from the
   Historical Data sheet via column-letter formulas).
3. **Run reports**: in the wizard, generate the SCALE run for the new
   quarter. The runner will pick up the new mapping automatically via
   `cecl_ui/services/scale/template_loader.resolve_map()`.
