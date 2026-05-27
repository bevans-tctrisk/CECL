"""Q-factor loader for SCALE.

Reads ``cecl_ui/data/scale/qfactors_default.csv`` and merges user
overrides from wizard state.

CSV columns:
    sheet         worksheet name in the SCALE template
    cell          A1-style coordinate
    label         human description for the wizard UI
    default_bps   default value in basis points (100 bps = 1.00%)

Lines whose ``sheet`` cell starts with ``#`` are ignored. Empty rows are
ignored. The CSV may have only the header row, in which case the
wizard step renders a "no Q-factors configured" notice and the runner
no-ops.
"""
from __future__ import annotations

import csv
from pathlib import Path
from typing import Any

_DEFAULTS_PATH = (
    Path(__file__).resolve().parent.parent.parent
    / "data" / "scale" / "qfactors_default.csv"
)


def defaults_path() -> Path:
    return _DEFAULTS_PATH


def _coerce_bps(value: Any) -> float:
    if value is None or value == "":
        return 0.0
    try:
        return float(value)
    except (TypeError, ValueError):
        return 0.0


def load_defaults(path: str | Path | None = None) -> list[dict]:
    """Return list of ``{sheet, cell, label, default_bps, key}``.

    ``key`` is ``"{sheet}|{cell}"`` and is used to look up user overrides
    in wizard state.
    """
    p = Path(path) if path else _DEFAULTS_PATH
    if not p.exists():
        return []
    rows: list[dict] = []
    with p.open("r", encoding="utf-8-sig", newline="") as fh:
        reader = csv.DictReader(fh)
        for raw in reader:
            sheet = (raw.get("sheet") or "").strip()
            if not sheet or sheet.startswith("#"):
                continue
            cell = (raw.get("cell") or "").strip()
            if not cell:
                continue
            rows.append({
                "sheet": sheet,
                "cell": cell,
                "label": (raw.get("label") or "").strip(),
                "default_bps": _coerce_bps(raw.get("default_bps")),
                "key": f"{sheet}|{cell}",
            })
    return rows


def merge_with_overrides(
    defaults: list[dict], overrides: dict[str, Any] | None,
) -> list[dict]:
    """Return a copy of ``defaults`` with each row's ``effective_bps``
    set from ``overrides[key]`` when present, else ``default_bps``."""
    overrides = overrides or {}
    out: list[dict] = []
    for row in defaults:
        eff = row["default_bps"]
        if row["key"] in overrides:
            eff = _coerce_bps(overrides[row["key"]])
        out.append({**row, "effective_bps": eff})
    return out
