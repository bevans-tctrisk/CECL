"""Cross-CU learned column-mapping suggestions.

Persists a small JSON store mapping each system field to the column-headers
that previous credit unions have assigned to it (with counts). When a new
CU's sample headers are inspected we use this store to pre-fill column
mappings the sample parser couldn't auto-detect.

Storage: ``cecl_ui/data/column_header_suggestions.json``. The file is
created on first write. Header keys are stored normalized (stripped,
lowercased, whitespace collapsed); a parallel ``header_display_forms``
dict keeps the first-seen original casing for display purposes.

Schema (v1):
    {
        "version": 1,
        "updated_at": "2026-05-18T...",
        "field_to_header_counts": {
            "<field>": {"<normalized header>": <int count>, ...},
            ...
        },
        "header_display_forms": {
            "<normalized header>": "<original-case header>", ...
        }
    }
"""

from __future__ import annotations

import json
import re
import threading
from datetime import datetime
from pathlib import Path
from typing import Iterable

_STORE_PATH = (
    Path(__file__).resolve().parent.parent / "data"
    / "column_header_suggestions.json"
)

_LOCK = threading.Lock()

# Fields the wizard can map. Keeping this explicit avoids polluting the
# store if callers ever pass stray keys.
KNOWN_FIELDS: tuple[str, ...] = (
    "member_number",
    "loan_suffix",
    "current_balance",
    "original_fico_score",
    "loan_pool_code",
    "days_delinquent",
    "interest_rate",
    "open_date",
    "original_loan_amount",
    "total_available_credit",
)

_WS_RE = re.compile(r"\s+")


def _normalize(header: str) -> str:
    return _WS_RE.sub(" ", (header or "").strip()).lower()


def _empty_store() -> dict:
    return {
        "version": 1,
        "updated_at": None,
        "field_to_header_counts": {},
        "header_display_forms": {},
    }


def load() -> dict:
    """Return the current store (empty skeleton if file is missing/corrupt)."""
    try:
        with _STORE_PATH.open("r", encoding="utf-8") as f:
            data = json.load(f)
    except (FileNotFoundError, json.JSONDecodeError, OSError):
        return _empty_store()
    if not isinstance(data, dict):
        return _empty_store()
    data.setdefault("version", 1)
    data.setdefault("field_to_header_counts", {})
    data.setdefault("header_display_forms", {})
    return data


def _save(store: dict) -> None:
    store["updated_at"] = datetime.utcnow().isoformat(timespec="seconds")
    _STORE_PATH.parent.mkdir(parents=True, exist_ok=True)
    tmp = _STORE_PATH.with_suffix(".json.tmp")
    with tmp.open("w", encoding="utf-8") as f:
        json.dump(store, f, indent=2, sort_keys=True)
    tmp.replace(_STORE_PATH)


def record_mapping(field_to_header: dict[str, str]) -> None:
    """Increment the count for each (field, header) the user just saved.

    Empty / unknown fields are skipped silently. Safe to call from any
    request handler; uses a process-level lock for the read-modify-write.
    """
    if not field_to_header:
        return
    with _LOCK:
        store = load()
        counts = store["field_to_header_counts"]
        displays = store["header_display_forms"]
        changed = False
        for field, header in field_to_header.items():
            if field not in KNOWN_FIELDS:
                continue
            raw = (header or "").strip()
            if not raw:
                continue
            norm = _normalize(raw)
            field_bucket = counts.setdefault(field, {})
            field_bucket[norm] = int(field_bucket.get(norm, 0)) + 1
            displays.setdefault(norm, raw)
            changed = True
        if changed:
            _save(store)


def suggest_for_headers(
    headers: Iterable[str],
    *,
    skip_fields: Iterable[str] = (),
) -> dict[str, str]:
    """Pick the best learned header for each field, restricted to ``headers``.

    Returns ``{field: header_exactly_as_in_headers}``. Fields listed in
    ``skip_fields`` are omitted (use this to avoid overwriting a field the
    sample parser already mapped). Only fields with a learned match are
    included in the result.
    """
    headers_list = [h for h in (headers or []) if h is not None]
    if not headers_list:
        return {}
    # Map normalized -> original casing as it appears in the current sample.
    norm_to_actual: dict[str, str] = {}
    for h in headers_list:
        norm_to_actual.setdefault(_normalize(str(h)), str(h))
    skip = set(skip_fields or ())

    store = load()
    counts = store.get("field_to_header_counts", {})
    out: dict[str, str] = {}
    for field, bucket in counts.items():
        if field in skip or field not in KNOWN_FIELDS:
            continue
        if not isinstance(bucket, dict):
            continue
        # Sort candidates by count desc, take the first one that exists in
        # the current sample headers.
        for norm, _cnt in sorted(
            bucket.items(), key=lambda kv: (-int(kv[1] or 0), kv[0])
        ):
            actual = norm_to_actual.get(norm)
            if actual:
                out[field] = actual
                break
    return out


def top_headers_for_field(field: str, n: int = 5) -> list[tuple[str, int]]:
    """For introspection / debugging: top-N headers ever mapped to ``field``."""
    store = load()
    bucket = store.get("field_to_header_counts", {}).get(field, {})
    displays = store.get("header_display_forms", {})
    if not isinstance(bucket, dict):
        return []
    ranked = sorted(
        ((displays.get(k, k), int(v or 0)) for k, v in bucket.items()),
        key=lambda kv: (-kv[1], kv[0]),
    )
    return ranked[: max(0, int(n))]
