"""Generate a SCALE quarterly mapping CSV for a new period.

The hand-maintained `cecl_ui/data/scale/maps/YYYY_MM_mapping.csv`
files only differ from one quarter to the next by the cell column
each row writes to (column AY = 12/31/2025, AZ = 3/31/2026, BA =
6/30/2026, ...). Each mapping file already embeds the full quarter-
to-column lookup in its header row, so we can derive a new quarter
deterministically from any existing one.

Usage (from repo root):

    # Generate the 2026-06 map from the newest existing mapping
    python scripts/generate_scale_mapping.py 2026-06

    # Use a specific source instead of "newest"
    python scripts/generate_scale_mapping.py 2026-06 --source 2026-03

    # Show what would be written without touching disk
    python scripts/generate_scale_mapping.py 2026-06 --dry-run

    # Overwrite an existing target
    python scripts/generate_scale_mapping.py 2026-06 --force

The script never modifies the source file. It only writes
``cecl_ui/data/scale/maps/<year>_<month>_mapping.csv``.
"""
from __future__ import annotations

import argparse
import re
import sys
from datetime import date
from pathlib import Path

MAPS_DIR = Path(__file__).resolve().parents[1] / "cecl_ui" / "data" / "scale" / "maps"

# Map quarter-end month -> last day of that month (no leap-year quarter-ends).
_QUARTER_END_DAY = {3: 31, 6: 30, 9: 30, 12: 31}


def _parse_period(s: str) -> tuple[int, int]:
    """Parse 'YYYY-MM' into (year, month). Month must be 3/6/9/12."""
    m = re.fullmatch(r"(\d{4})-(\d{2})", s.strip())
    if not m:
        raise ValueError(f"period must be YYYY-MM, got {s!r}")
    year, month = int(m.group(1)), int(m.group(2))
    if month not in _QUARTER_END_DAY:
        raise ValueError(f"month must be 03, 06, 09, or 12; got {month:02d}")
    return year, month


def _period_to_header_date(year: int, month: int) -> str:
    """Build the m/d/yyyy string used in the CSV header row."""
    return f"{month}/{_QUARTER_END_DAY[month]}/{year}"


def _map_path(year: int, month: int) -> Path:
    return MAPS_DIR / f"{year}_{month:02d}_mapping.csv"


def _list_existing_periods() -> list[tuple[int, int]]:
    out: list[tuple[int, int]] = []
    for p in MAPS_DIR.glob("*_mapping.csv"):
        m = re.fullmatch(r"(\d{4})_(\d{2})_mapping\.csv", p.name)
        if m:
            out.append((int(m.group(1)), int(m.group(2))))
    out.sort()
    return out


def _newest_existing_period() -> tuple[int, int] | None:
    periods = _list_existing_periods()
    return periods[-1] if periods else None


def _split_csv_lines(text: str) -> list[list[str]]:
    """Split CSV text into rows of cells without depending on csv module.

    The mapping files use plain comma separation with no quoting or
    embedded commas, so a naive split is correct and preserves
    trailing empty cells exactly.
    """
    lines = text.splitlines(keepends=False)
    return [ln.split(",") for ln in lines]


def _join_csv_rows(rows: list[list[str]], line_ending: str) -> str:
    return line_ending.join(",".join(r) for r in rows) + line_ending


def _detect_line_ending(text: str) -> str:
    # Preserve the source file's line endings (\r\n on Windows, \n on Unix).
    if "\r\n" in text:
        return "\r\n"
    return "\n"


def _find_date_column_index(header: list[str], target_date: str) -> int:
    for idx, cell in enumerate(header):
        if cell.strip() == target_date:
            return idx
    raise ValueError(
        f"Header row does not contain date {target_date!r}. "
        f"Found dates: {[c for c in header if '/' in c]}"
    )


def _column_letter_at(row: list[str], idx: int) -> str:
    if idx >= len(row):
        raise ValueError(
            f"Source row only has {len(row)} cells but we need index {idx}"
        )
    val = row[idx].strip()
    if not re.fullmatch(r"[A-Z]+", val):
        raise ValueError(
            f"Expected a column letter at lookup index {idx}, got {val!r}"
        )
    return val


def _swap_cell_column(cell: str, old_col: str, new_col: str) -> str:
    """Replace the column letters in an A1-style coordinate.

    Asserts the cell starts with ``old_col``; raises if not, so a
    botched source file is caught loudly rather than silently
    half-rewritten.
    """
    m = re.fullmatch(r"([A-Z]+)(\d+)", cell.strip())
    if not m:
        raise ValueError(f"Cell {cell!r} is not in A1 format")
    col, row = m.group(1), m.group(2)
    if col != old_col:
        raise ValueError(
            f"Cell {cell!r} uses column {col}, expected {old_col}. "
            "Source mapping may be inconsistent across rows; aborting."
        )
    return f"{new_col}{row}"


def generate(
    target_period: str,
    source_period: str | None = None,
    dry_run: bool = False,
    force: bool = False,
) -> dict:
    """Generate ``target_period`` mapping CSV from ``source_period``.

    Returns a dict describing what happened (or would happen):
    ``{target_path, source_path, source_period, target_period,
    source_column, target_column, rows_rewritten, written}``.
    """
    t_year, t_month = _parse_period(target_period)
    target_date = _period_to_header_date(t_year, t_month)
    target_path = _map_path(t_year, t_month)

    if source_period:
        s_year, s_month = _parse_period(source_period)
    else:
        latest = _newest_existing_period()
        if latest is None:
            raise FileNotFoundError(
                f"No existing mapping files in {MAPS_DIR}; cannot infer source."
            )
        s_year, s_month = latest
    source_date = _period_to_header_date(s_year, s_month)
    source_path = _map_path(s_year, s_month)

    if (s_year, s_month) == (t_year, t_month):
        raise ValueError(
            f"Source and target are the same period ({target_period}); "
            "nothing to do."
        )
    if not source_path.exists():
        raise FileNotFoundError(f"Source mapping not found: {source_path}")
    if target_path.exists() and not force and not dry_run:
        raise FileExistsError(
            f"Target already exists: {target_path}. Use --force to overwrite."
        )

    raw = source_path.read_text(encoding="utf-8-sig")
    line_ending = _detect_line_ending(raw)
    rows = _split_csv_lines(raw)
    if len(rows) < 2:
        raise ValueError(f"Source has no data rows: {source_path}")

    header = rows[0]
    source_idx = _find_date_column_index(header, source_date)
    target_idx = _find_date_column_index(header, target_date)

    # The per-row lookup (cols 19+) gives column position -> column letter.
    # Every data row has the same lookup, so reading row 1 is safe.
    first_data_row = rows[1]
    source_col = _column_letter_at(first_data_row, source_idx)
    target_col = _column_letter_at(first_data_row, target_idx)

    new_rows: list[list[str]] = [header[:]]
    rows_rewritten = 0
    for r_idx, row in enumerate(rows[1:], start=1):
        if not any(c.strip() for c in row):
            new_rows.append(row[:])
            continue
        new_row = row[:]
        # Col 2 = "cell" (A1 coord). Col 5 = "Column" letter. Col 4 =
        # "Report Date" (only populated on the first data row).
        if len(new_row) > 2 and new_row[2].strip():
            new_row[2] = _swap_cell_column(new_row[2], source_col, target_col)
        if len(new_row) > 5 and new_row[5].strip():
            if new_row[5].strip() != source_col:
                raise ValueError(
                    f"Row {r_idx + 1} 'Column' cell is "
                    f"{new_row[5]!r}, expected {source_col!r}; aborting."
                )
            new_row[5] = target_col
        if len(new_row) > 4 and new_row[4].strip():
            if new_row[4].strip() != source_date:
                raise ValueError(
                    f"Row {r_idx + 1} 'Report Date' cell is "
                    f"{new_row[4]!r}, expected {source_date!r}; aborting."
                )
            new_row[4] = target_date
        new_rows.append(new_row)
        rows_rewritten += 1

    output_text = _join_csv_rows(new_rows, line_ending)

    result = {
        "target_path": str(target_path),
        "source_path": str(source_path),
        "source_period": f"{s_year}-{s_month:02d}",
        "target_period": f"{t_year}-{t_month:02d}",
        "source_column": source_col,
        "target_column": target_col,
        "rows_rewritten": rows_rewritten,
        "written": False,
    }

    if dry_run:
        return result

    target_path.parent.mkdir(parents=True, exist_ok=True)
    target_path.write_text(output_text, encoding="utf-8", newline="")
    result["written"] = True
    return result


def _main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        description="Generate a SCALE quarterly mapping CSV from an "
                    "existing one by advancing the cell column."
    )
    parser.add_argument(
        "target_period",
        help="Target quarter as YYYY-MM (e.g. 2026-06).",
    )
    parser.add_argument(
        "--source",
        dest="source_period",
        default=None,
        help="Source quarter as YYYY-MM. Defaults to the newest "
             "existing mapping file in cecl_ui/data/scale/maps/.",
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Show what would be written without modifying disk.",
    )
    parser.add_argument(
        "--force",
        action="store_true",
        help="Overwrite the target file if it already exists.",
    )
    args = parser.parse_args(argv)

    try:
        result = generate(
            target_period=args.target_period,
            source_period=args.source_period,
            dry_run=args.dry_run,
            force=args.force,
        )
    except (ValueError, FileNotFoundError, FileExistsError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 1

    action = "Would write" if args.dry_run else "Wrote"
    print(
        f"{action} {result['target_path']}\n"
        f"  source:  {result['source_path']} ({result['source_period']}, "
        f"col {result['source_column']})\n"
        f"  target:  {result['target_period']} -> col {result['target_column']}\n"
        f"  rows rewritten: {result['rows_rewritten']}"
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(_main())
