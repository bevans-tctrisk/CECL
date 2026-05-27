"""Firm-wide / cross-CU default settings (the "Admin" panel).

This is a single-user, on-machine settings store. Values live in
``admin_defaults.yaml`` at the workspace root (gitignored). Anything in
this file is merged on top of the hardcoded system defaults and below
each CU's wizard YAML.

Today only ``default_mgmt_adj`` is exposed. Future fields (default
credit grades, DQ ranges, etc.) should slot in alongside it.
"""
from __future__ import annotations

from pathlib import Path
from typing import Any

import yaml

# Single source of truth for the on-disk filename.
_FILENAME = "admin_defaults.yaml"

# Environmental Factor Ranges seeded from the 2025-06 SCALE template
# (TCT tab "Environmental Factor Ranges"). 17 delinquency rows
# (J9:K25) + 14 economic-stress rows (L9:M22). Each row is
# ``[minimum_decimal, score_decimal]``. Highest range first.
_DEFAULT_DELINQUENCY_RANGES: list[list[float]] = [
    [0.05, 0.20],
    [0.04, 0.17],
    [0.03, 0.12],
    [0.025, 0.08],
    [0.02, 0.04],
    [0.015, 0.025],
    [0.01, 0.015],
    [0.005, 0.0075],
    [-0.0049, 0.0],
    [-0.0099, -0.0075],
    [-0.0149, -0.015],
    [-0.0199, -0.025],
    [-0.0249, -0.04],
    [-0.0299, -0.08],
    [-0.0399, -0.12],
    [-0.0499, -0.17],
    [-0.05, -0.20],
]
_DEFAULT_ECON_STRESS_RANGES: list[list[float]] = [
    [0.25, 0.10],
    [0.24, 0.08],
    [0.22, 0.07],
    [0.20, 0.06],
    [0.18, 0.05],
    [0.16, 0.04],
    [0.14, 0.035],
    [0.12, 0.03],
    [0.10, 0.02],
    [0.08, 0.01],
    [0.06, 0.00],
    [0.04, 0.00],
    [0.02, -0.01],
    [0.00, -0.02],
]

# Hardcoded fallback if the file is missing / unreadable / malformed.
SYSTEM_DEFAULTS: dict[str, Any] = {
    # Industry-wide management adjustment (decimal). 0.0011 = 0.11%.
    # Applied to a pool ONLY when:
    #   - pool.use_default_mgmt_adj is True, AND
    #   - the pool's ACL base loss rate is 0, AND
    #   - the user has NOT entered a manual overlay for the pool.
    "default_mgmt_adj": 0.0011,
    # Environmental Factor Ranges written into every SCALE report at
    # run time. See env_factor_writer.py for the cell layout. Edits
    # take effect on future runs only; existing reports keep the
    # values they were generated with.
    "env_factor_ranges": {
        "delinquency": _DEFAULT_DELINQUENCY_RANGES,
        "econ_stress": _DEFAULT_ECON_STRESS_RANGES,
    },
}


# Expected row counts for the two range tables. Used by the admin UI
# and the writer to keep input arity in sync with the template.
DELINQUENCY_ROW_COUNT = len(_DEFAULT_DELINQUENCY_RANGES)
ECON_STRESS_ROW_COUNT = len(_DEFAULT_ECON_STRESS_RANGES)


def _admin_path() -> Path:
    """Return the path to admin_defaults.yaml at the workspace root."""
    # cecl_ui/services/admin_defaults.py -> workspace root is parents[2].
    return Path(__file__).resolve().parents[2] / _FILENAME


def load() -> dict[str, Any]:
    """Read admin defaults, falling back to SYSTEM_DEFAULTS for missing keys.

    Never raises; bad/missing file -> SYSTEM_DEFAULTS.
    """
    merged: dict[str, Any] = dict(SYSTEM_DEFAULTS)
    p = _admin_path()
    if not p.exists():
        return merged
    try:
        with p.open("r", encoding="utf-8") as fh:
            data = yaml.safe_load(fh) or {}
        if isinstance(data, dict):
            for k, v in data.items():
                merged[k] = v
    except Exception:  # noqa: BLE001 - keep startup robust
        pass
    return merged


def save(values: dict[str, Any]) -> Path:
    """Write the given values to admin_defaults.yaml.

    Only keys present in ``values`` are written (no defensive merging
    against the file's existing contents). Caller is responsible for
    passing the full set it wants persisted.
    """
    p = _admin_path()
    with p.open("w", encoding="utf-8") as fh:
        yaml.safe_dump(values, fh, sort_keys=True)
    return p
