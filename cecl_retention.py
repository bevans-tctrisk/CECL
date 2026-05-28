"""
CECL Data Retention Manager — automated cleanup of old data and reports.

Purges database records and archived reports beyond the configured retention
period, reducing PII exposure.

Usage:
    # Preview what would be deleted (dry run, default)
    python cecl_retention.py --dry-run

    # Actually delete old records (requires confirmation)
    python cecl_retention.py --execute

    # Custom retention period (default: 7 years / 84 months)
    python cecl_retention.py --months 60 --dry-run

    # Purge old archived reports
    python cecl_retention.py --execute --include-reports
"""
import os
import sys
import argparse
from datetime import datetime, date
from dateutil.relativedelta import relativedelta

from sqlalchemy import create_engine, text
from cecl_credentials import get_database_url
from cecl_audit_log import get_audit_logger, log_data_retention

# Honour CECL_WORKSPACE_ROOT so retention runs against the shared data root
# rather than against whichever code clone the script happens to live in.
# Falls back to historical layout when the env var is unset.
BASE = os.environ.get('CECL_WORKSPACE_ROOT') or os.path.dirname(os.path.abspath(__file__))
REPORTS_DIR = os.path.join(BASE, 'Reports')
ARCHIVE_DIR = os.path.join(BASE, 'Archive')

DEFAULT_RETENTION_MONTHS = 84  # 7 years


def get_db_snapshot_dates(engine):
    """Return a list of (snapshot_date, credit_union, record_count) tuples from the database."""
    with engine.connect() as conn:
        result = conn.execute(text(
            "SELECT snapshot_date, credit_union, COUNT(*) as cnt "
            "FROM monthly_loan_data "
            "GROUP BY snapshot_date, credit_union "
            "ORDER BY snapshot_date"
        ))
        return [(row[0], row[1], row[2]) for row in result]


def purge_db_records(engine, cutoff_date, dry_run=True):
    """Delete database records older than cutoff_date."""
    audit = get_audit_logger()
    snapshots = get_db_snapshot_dates(engine)
    old_snapshots = [(d, cu, cnt) for d, cu, cnt in snapshots if d < cutoff_date]

    if not old_snapshots:
        print("  No database records older than the retention cutoff.")
        return 0

    total_records = sum(cnt for _, _, cnt in old_snapshots)
    print(f"\n  Database records to purge (older than {cutoff_date}):")
    for snap_date, cu, cnt in old_snapshots:
        print(f"    {snap_date}  {cu:40s}  {cnt:>8,d} records")
    print(f"    {'':40s}  Total: {total_records:>8,d} records")

    if dry_run:
        print("\n  [DRY RUN] No records deleted.")
        return total_records

    with engine.begin() as conn:
        result = conn.execute(
            text("DELETE FROM monthly_loan_data WHERE snapshot_date < :cutoff"),
            {"cutoff": cutoff_date}
        )
        deleted = result.rowcount

    print(f"\n  Deleted {deleted:,d} records from database.")
    log_data_retention("DB_PURGE", f"cutoff={cutoff_date} deleted={deleted}")
    return deleted


def purge_old_reports(cutoff_date, dry_run=True):
    """Delete report files from the Reports directory older than cutoff_date."""
    audit = get_audit_logger()
    if not os.path.isdir(REPORTS_DIR):
        print("  Reports directory not found.")
        return 0

    old_files = []
    for fname in os.listdir(REPORTS_DIR):
        fpath = os.path.join(REPORTS_DIR, fname)
        if not os.path.isfile(fpath):
            continue
        if not fname.endswith(('.xlsx', '.xlsm', '.xls', '.docx')):
            continue
        # Check modification time
        mtime = datetime.fromtimestamp(os.path.getmtime(fpath)).date()
        if mtime < cutoff_date:
            size_kb = os.path.getsize(fpath) / 1024
            old_files.append((fname, mtime, size_kb))

    if not old_files:
        print("  No report files older than the retention cutoff.")
        return 0

    total_size = sum(s for _, _, s in old_files)
    print(f"\n  Report files to purge (modified before {cutoff_date}):")
    for fname, mtime, size_kb in old_files:
        print(f"    {mtime}  {fname:60s}  {size_kb:>8.1f} KB")
    print(f"    Total: {len(old_files)} files, {total_size:,.1f} KB")

    if dry_run:
        print("\n  [DRY RUN] No files deleted.")
        return len(old_files)

    deleted = 0
    for fname, mtime, _ in old_files:
        fpath = os.path.join(REPORTS_DIR, fname)
        os.remove(fpath)
        deleted += 1

    print(f"\n  Deleted {deleted} report files.")
    log_data_retention("REPORT_PURGE", f"cutoff={cutoff_date} deleted={deleted} files")
    return deleted


def purge_old_archives(cutoff_date, dry_run=True):
    """Delete archived source files older than cutoff_date."""
    audit = get_audit_logger()
    if not os.path.isdir(ARCHIVE_DIR):
        print("  Archive directory not found.")
        return 0

    old_files = []
    for root, dirs, files in os.walk(ARCHIVE_DIR):
        for fname in files:
            fpath = os.path.join(root, fname)
            mtime = datetime.fromtimestamp(os.path.getmtime(fpath)).date()
            if mtime < cutoff_date:
                rel = os.path.relpath(fpath, ARCHIVE_DIR)
                size_kb = os.path.getsize(fpath) / 1024
                old_files.append((rel, mtime, size_kb, fpath))

    if not old_files:
        print("  No archived files older than the retention cutoff.")
        return 0

    total_size = sum(s for _, _, s, _ in old_files)
    print(f"\n  Archived files to purge (modified before {cutoff_date}):")
    for rel, mtime, size_kb, _ in old_files:
        print(f"    {mtime}  {rel:60s}  {size_kb:>8.1f} KB")
    print(f"    Total: {len(old_files)} files, {total_size:,.1f} KB")

    if dry_run:
        print("\n  [DRY RUN] No files deleted.")
        return len(old_files)

    deleted = 0
    for _, _, _, fpath in old_files:
        os.remove(fpath)
        deleted += 1

    print(f"\n  Deleted {deleted} archived files.")
    log_data_retention("ARCHIVE_PURGE", f"cutoff={cutoff_date} deleted={deleted} files")
    return deleted


def main():
    parser = argparse.ArgumentParser(
        description="CECL Data Retention Manager — purge old data beyond retention period",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python cecl_retention.py --dry-run                    # Preview (7-year default)
  python cecl_retention.py --dry-run --months 60        # Preview with 5-year retention
  python cecl_retention.py --execute                    # Purge DB records
  python cecl_retention.py --execute --include-reports   # Purge DB + report files
  python cecl_retention.py --execute --include-archives  # Purge DB + archived source files
        """,
    )
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--dry-run', action='store_true', help='Preview what would be deleted (no changes)')
    group.add_argument('--execute', action='store_true', help='Actually delete old data')
    parser.add_argument('--months', type=int, default=DEFAULT_RETENTION_MONTHS,
                        help=f'Retention period in months (default: {DEFAULT_RETENTION_MONTHS})')
    parser.add_argument('--include-reports', action='store_true',
                        help='Also purge old report files from Reports directory')
    parser.add_argument('--include-archives', action='store_true',
                        help='Also purge old archived source files')
    args = parser.parse_args()

    cutoff = date.today() - relativedelta(months=args.months)
    dry_run = args.dry_run

    print(f"{'='*60}")
    print(f"  CECL Data Retention Manager")
    print(f"  Retention period: {args.months} months")
    print(f"  Cutoff date:      {cutoff}")
    print(f"  Mode:             {'DRY RUN (preview only)' if dry_run else 'EXECUTE (deleting data)'}")
    print(f"{'='*60}")

    if args.execute:
        confirm = input("\n  Type 'yes' to confirm deletion: ").strip().lower()
        if confirm != 'yes':
            print("  Aborted.")
            sys.exit(0)

    engine = create_engine(get_database_url())

    purge_db_records(engine, cutoff, dry_run)

    if args.include_reports:
        purge_old_reports(cutoff, dry_run)

    if args.include_archives:
        purge_old_archives(cutoff, dry_run)

    print(f"\n{'='*60}")
    print(f"  Retention cleanup {'preview' if dry_run else 'execution'} complete.")
    print(f"{'='*60}")


if __name__ == '__main__':
    main()
