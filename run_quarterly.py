"""
CECL Quarterly Processing Orchestrator
Runs the full pipeline: import data -> generate reports for all or specific clients.

Usage:
    python run_quarterly.py --client ontario
    python run_quarterly.py --all
    python run_quarterly.py --client ontario --skip-import
    python run_quarterly.py --all --date 2025-12-31
"""
import os
import sys
import argparse
from datetime import datetime

import yaml
from dotenv import load_dotenv

load_dotenv()

# Honour CECL_WORKSPACE_ROOT so the data root can be decoupled from the
# code location; falls back to historical layout when the env var is unset.
BASE_FOLDER = os.environ.get('CECL_WORKSPACE_ROOT') or os.path.dirname(os.path.abspath(__file__))
CONFIG_FOLDER = os.path.join(BASE_FOLDER, 'client_configs')


def load_client_config(client_name):
    config_path = os.path.join(CONFIG_FOLDER, f'{client_name}.yaml')
    with open(config_path, 'r', encoding='utf-8') as f:
        return yaml.safe_load(f)


def list_clients():
    clients = []
    for f in os.listdir(CONFIG_FOLDER):
        if f.endswith('.yaml') and not f.startswith('_'):
            clients.append(os.path.splitext(f)[0])
    return sorted(clients)


def run_pipeline(client_name, snapshot_date=None, skip_import=False):
    """Run the full pipeline for a single client."""
    config = load_client_config(client_name)
    cu_name = config['credit_union']

    print(f"\n{'#'*60}")
    print(f"# CECL Quarterly Pipeline: {cu_name}")
    print(f"# Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'#'*60}")

    # Step 1: Import new data files
    if not skip_import:
        print(f"\n--- Step 1: Import Data ---")
        from import_data import process_client
        process_client(client_name)
    else:
        print(f"\n--- Step 1: Import (SKIPPED) ---")

    # Step 2: Generate report
    print(f"\n--- Step 2: Generate Report ---")
    from generate_report import generate_report
    report_path = generate_report(client_name, snapshot_date)

    if report_path:
        print(f"\n--- Pipeline Complete ---")
        print(f"  Report: {report_path}")
        print(f"  Dashboard: streamlit run dashboard.py")
    else:
        print(f"\n--- Pipeline Complete (no report generated) ---")
        print(f"  No data found. Place files in Raw_Uploads/ and re-run.")

    return report_path


def main():
    parser = argparse.ArgumentParser(
        description="CECL Quarterly Processing Orchestrator",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python run_quarterly.py --client ontario              # Import + Report for Ontario
  python run_quarterly.py --all                         # All clients
  python run_quarterly.py --client ontario --skip-import  # Report only (no import)
  python run_quarterly.py --list                        # Show available clients
  python run_quarterly.py --all --date 2025-12-31       # Force specific date

Adding a new credit union:
  1. Copy client_configs/_template.yaml -> client_configs/<name>.yaml
  2. Edit the YAML for the new CU's file format and pool codes
  3. Create folder: Raw_Uploads/<name>/
  4. Drop their data files in the folder
  5. Run: python run_quarterly.py --client <name>
        """,
    )
    parser.add_argument('--client', help='Client config name (e.g., "ontario")')
    parser.add_argument('--date', help='Force specific snapshot date (YYYY-MM-DD)')
    parser.add_argument('--all', action='store_true', help='Process all configured clients')
    parser.add_argument('--skip-import', action='store_true', help='Skip data import step')
    parser.add_argument('--list', action='store_true', help='List available clients')
    args = parser.parse_args()

    if args.list:
        print("Available Clients:")
        print(f"  {'Config':20s}  {'Credit Union':40s}")
        print(f"  {'-'*20}  {'-'*40}")
        for c in list_clients():
            cfg = load_client_config(c)
            print(f"  {c:20s}  {cfg['credit_union']}")
        print(f"\nTo add a new client, copy client_configs/_template.yaml")
        return

    if args.all:
        clients = list_clients()
        print(f"Processing {len(clients)} client(s): {', '.join(clients)}")
        for client_name in clients:
            run_pipeline(client_name, args.date, args.skip_import)
    elif args.client:
        run_pipeline(args.client, args.date, args.skip_import)
    else:
        parser.print_help()


if __name__ == '__main__':
    main()
