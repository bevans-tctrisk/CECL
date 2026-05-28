"""
CECL Audit Logger — centralized access and activity logging.

Logs who runs which operations, when, and for which client.
Writes to both console and a persistent log file (logs/cecl_audit.log).

Usage:
    from cecl_audit_log import get_audit_logger
    audit = get_audit_logger()
    audit.info("Generated TCT report for Franklin Trust FCU, period 2025-12-31")
"""
import os
import getpass
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime

# Honour CECL_WORKSPACE_ROOT so the audit log lives next to the analyst data
# (centralised across clones), not buried inside whichever code clone wrote it.
# Falls back to historical layout when the env var is unset.
BASE = os.environ.get('CECL_WORKSPACE_ROOT') or os.path.dirname(os.path.abspath(__file__))
LOG_DIR = os.path.join(BASE, 'logs')
LOG_FILE = os.path.join(LOG_DIR, 'cecl_audit.log')

_logger = None


def get_audit_logger():
    """Return the singleton audit logger, creating it on first call."""
    global _logger
    if _logger is not None:
        return _logger

    os.makedirs(LOG_DIR, exist_ok=True)

    _logger = logging.getLogger('cecl_audit')
    _logger.setLevel(logging.INFO)
    _logger.propagate = False

    # Rotating file handler — 5 MB per file, keep 10 backups (50 MB total)
    fh = RotatingFileHandler(
        LOG_FILE, maxBytes=5 * 1024 * 1024, backupCount=10, encoding='utf-8',
    )
    fh.setLevel(logging.INFO)

    # Console handler
    ch = logging.StreamHandler()
    ch.setLevel(logging.WARNING)

    # Format: timestamp | user | level | message
    user = getpass.getuser()
    fmt = logging.Formatter(
        f'%(asctime)s | {user} | %(levelname)s | %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S',
    )
    fh.setFormatter(fmt)
    ch.setFormatter(fmt)

    _logger.addHandler(fh)
    _logger.addHandler(ch)

    return _logger


def log_report_generation(client_name, credit_union, snapshot_date, report_type, output_path, success=True):
    """Log a report generation event."""
    audit = get_audit_logger()
    status = "SUCCESS" if success else "FAILED"
    audit.info(
        "REPORT_GENERATED | client=%s | cu=%s | date=%s | type=%s | file=%s | status=%s",
        client_name, credit_union, snapshot_date, report_type,
        os.path.basename(output_path) if output_path else "N/A", status,
    )


def log_data_import(client_name, credit_union, source_file, record_count, success=True):
    """Log a data import event."""
    audit = get_audit_logger()
    status = "SUCCESS" if success else "FAILED"
    audit.info(
        "DATA_IMPORTED | client=%s | cu=%s | file=%s | records=%s | status=%s",
        client_name, credit_union,
        os.path.basename(source_file) if source_file else "N/A",
        record_count, status,
    )


def log_data_retention(action, details):
    """Log a data retention/cleanup event."""
    audit = get_audit_logger()
    audit.info("DATA_RETENTION | action=%s | %s", action, details)


def log_session_start(script_name, args=None):
    """Log a script execution start."""
    audit = get_audit_logger()
    audit.info("SESSION_START | script=%s | args=%s", script_name, args or "")


def log_session_end(script_name, result="completed"):
    """Log a script execution end."""
    audit = get_audit_logger()
    audit.info("SESSION_END | script=%s | result=%s", script_name, result)
