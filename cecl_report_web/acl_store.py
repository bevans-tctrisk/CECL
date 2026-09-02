"""Prior-period ACL snapshot sidecars.

Each report period's ACL shape (pools / order / impaired / totals -- the same
dict ``change_analysis._parse_acl_sheet`` produces) is persisted as a small
JSON file next to the workbook. The Change Analysis and Summary Variance tabs
then diff against the prior quarter by reading this sidecar instead of opening
a prior ``.xlsx``. Periods generated before sidecars existed simply have none,
so the caller falls back to the workbook for those.
"""

from __future__ import annotations

import glob
import json
import os
import re
from pathlib import Path

_SUFFIX = "Vizo_Model.acl.json"
_SHAPE_KEYS = ("pools", "order", "impaired", "totals")


def _safe_cu(cu: str) -> str:
    return cu.replace(" ", "_").replace("/", "-")


def sidecar_path(rpt_dir: str | Path, cu: str, snapshot: str) -> Path:
    return Path(rpt_dir) / f"{snapshot}_CECL_Migration_{_safe_cu(cu)}_{_SUFFIX}"


def write_acl_snapshot(rpt_dir: str | Path, cu: str, snapshot: str,
                       shape: dict) -> Path | None:
    """Persist *shape* for (cu, snapshot). Never raises -- a failed sidecar
    must not break report generation."""
    try:
        os.makedirs(rpt_dir, exist_ok=True)
        path = sidecar_path(rpt_dir, cu, snapshot)
        payload = {"snapshot": snapshot, "credit_union": cu,
                   **{k: shape.get(k) for k in _SHAPE_KEYS}}
        path.write_text(json.dumps(payload, default=float), encoding="utf-8")
        return path
    except Exception as exc:  # noqa: BLE001
        print(f"  ACL sidecar write failed: {exc}")
        return None


def load_prior_snapshot(rpt_dir: str | Path, cu: str, snapshot: str):
    """Return (shape_dict, prior_snap) for the most recent sidecar dated
    strictly before *snapshot*, or (None, None)."""
    safe = _safe_cu(cu)
    pattern = os.path.join(str(rpt_dir), f"*_CECL_Migration_{safe}_{_SUFFIX}")
    rx = re.compile(r"(\d{4}-\d{2}-\d{2})_CECL_Migration_"
                    + re.escape(safe) + r"_" + re.escape(_SUFFIX) + r"$")
    best, best_date = None, None
    for path in glob.glob(pattern):
        m = rx.search(os.path.basename(path))
        if not m:
            continue
        d = m.group(1)
        if d >= snapshot:
            continue
        if best_date is None or d > best_date:
            best, best_date = path, d
    if not best:
        return None, None
    try:
        data = json.loads(Path(best).read_text(encoding="utf-8"))
        shape = {k: data[k] for k in _SHAPE_KEYS if k in data}
        return shape, best_date
    except Exception as exc:  # noqa: BLE001
        print(f"  ACL sidecar read failed ({best}): {exc}")
        return None, None
