"""Regenerate Nova CU March 2026 reports after Phase 9.30 DQ balance refill."""
import os
os.environ.setdefault("CECL_WORKSPACE_ROOT", r"Z:\Shared\TCT Files\CECL - CM Files")

from cecl_ui.services import pipeline_service

result = pipeline_service.run_reports(
    "nova_cu",
    "2026-03-31",
    ["vizo", "vizo_supp"],
)
print(result)
