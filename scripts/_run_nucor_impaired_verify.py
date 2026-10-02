"""Rerun Nucor June 2026 Vizo report + verify Spec ID grade alignment."""
import os, sys
os.environ.setdefault('CECL_WORKSPACE_ROOT', r'Z:\Shared\TCT Files\CECL - CM Files')
sys.path.insert(0, r'c:\Dev\CECL')
from cecl_credentials import get_database_url
os.environ['DATABASE_URL'] = get_database_url()

from cecl_ui.services import pipeline_service

print("=== Running Nucor Vizo for 2026-06-30 ===")
res = pipeline_service.run_reports('nucor_emp_cu', '2026-06-30', ['vizo'])
print("Result:", res)
