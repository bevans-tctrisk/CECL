"""Resolve SCALE template + mapping CSV per period.

Resolution order (per asset):
  1. Per-CU override (``state.scale.template_override_path`` /
     ``map_override_path``) if set and file exists.
  2. App-wide canonical for that exact period
     (``cecl_ui/data/scale/{templates,maps}/<YYYY_MM>_*``).
  3. App-wide canonical for the nearest prior period (templates only —
     maps must match the period exactly because cell layouts change with
     each NCUA form revision).
"""
from __future__ import annotations

from pathlib import Path
from typing import Optional

from flask import current_app


def _data_dir(kind: str) -> Path:
    """Locate ``cecl_ui/data/scale/{templates,maps}``."""
    # This module lives at ``cecl_ui/services/scale/template_loader.py``.
    here = Path(__file__).resolve().parent.parent.parent
    return here / "data" / "scale" / kind


def _period_to_prefix(period: str) -> str:
    """``2025-12`` -> ``2025_12``."""
    return period.replace("-", "_")


# ---------- template ----------

def list_available_template_periods() -> list[str]:
    out: list[str] = []
    d = _data_dir("templates")
    if not d.is_dir():
        return out
    for p in sorted(d.glob("*_CECL_SCALE_template.xlsx")):
        out.append(p.stem.split("_CECL_SCALE_template")[0].replace("_", "-"))
    return out


def resolve_template(period: str, override_path: Optional[str] = None) -> dict:
    """Find a template for ``period``.

    Returns ``{ok, path, source, message}``. ``source`` is one of
    ``override`` / ``canonical`` / ``canonical_fallback``.
    """
    if override_path:
        p = Path(override_path)
        if p.is_file():
            return {"ok": True, "path": str(p), "source": "override",
                    "message": ""}
    d = _data_dir("templates")
    prefix = _period_to_prefix(period)
    exact = d / f"{prefix}_CECL_SCALE_template.xlsx"
    if exact.is_file():
        return {"ok": True, "path": str(exact), "source": "canonical",
                "message": ""}
    # Fallback: newest available template <= requested period.
    available = list_available_template_periods()
    candidates = [a for a in available if a <= period]
    if candidates:
        chosen = max(candidates)
        p = d / f"{_period_to_prefix(chosen)}_CECL_SCALE_template.xlsx"
        return {
            "ok": True,
            "path": str(p),
            "source": "canonical_fallback",
            "message": f"No template for {period}; using {chosen}.",
        }
    return {"ok": False, "path": "", "source": "",
            "message": f"No template available for {period}."}


# ---------- mapping ----------

def list_available_map_periods() -> list[str]:
    out: list[str] = []
    d = _data_dir("maps")
    if not d.is_dir():
        return out
    for p in sorted(d.glob("*_mapping.csv")):
        out.append(p.stem.split("_mapping")[0].replace("_", "-"))
    return out


def resolve_map(period: str, override_path: Optional[str] = None) -> dict:
    if override_path:
        p = Path(override_path)
        if p.is_file():
            return {"ok": True, "path": str(p), "source": "override",
                    "message": ""}
    d = _data_dir("maps")
    prefix = _period_to_prefix(period)
    exact = d / f"{prefix}_mapping.csv"
    if exact.is_file():
        return {"ok": True, "path": str(exact), "source": "canonical",
                "message": ""}
    return {
        "ok": False,
        "path": "",
        "source": "",
        "message": (
            f"No mapping CSV for {period}. NCUA form layouts shift each "
            f"quarter so the period must match exactly."
        ),
    }
