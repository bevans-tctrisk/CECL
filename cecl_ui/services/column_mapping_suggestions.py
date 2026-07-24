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
    "current_fico_score",
    "loan_pool_code",
    "days_delinquent",
    "interest_rate",
    "open_date",
    "original_loan_amount",
    "total_available_credit",
)

# Fields that must never be auto-suggested from learned history, even when a
# matching header exists. ``current_fico_score`` defaults to none because
# most core/AIRES extracts ship only the origination score (the current
# score is supplied by the credit pull). Users can still map it manually.
_NEVER_AUTOSUGGEST: frozenset[str] = frozenset({"current_fico_score"})

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

    ``current_fico_score`` is never auto-suggested: most core/AIRES extracts
    carry only the origination score and the current score comes from the
    credit pull, so it defaults to none (unmapped) unless the user maps it
    manually.
    """
    headers_list = [h for h in (headers or []) if h is not None]
    if not headers_list:
        return {}
    # Map normalized -> original casing as it appears in the current sample.
    norm_to_actual: dict[str, str] = {}
    for h in headers_list:
        norm_to_actual.setdefault(_normalize(str(h)), str(h))
    skip = set(skip_fields or ()) | _NEVER_AUTOSUGGEST

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


# Factory-default placeholder sentinels used by the wizard's blank state.
# These are NOT real headers and must never enter the learned store.
_PLACEHOLDER_HEADERS: frozenset[str] = frozenset({
    "MEMBER_ID", "BALANCE", "FICO_SCORE", "LOAN_TYPE",
    "DQ_DAYS", "INT_RATE", "OPEN_DATE", "ORIG_AMT",
})


def _iter_config_mappings(cfg: dict) -> "list[dict]":
    """Return every column_mappings dict in a client config.

    Includes the top-level ``column_mappings`` plus each
    ``loan_data_extracts[*].column_mappings`` (credit-card / second-extract
    CUs carry their own per-file mappings).
    """
    out: list[dict] = []
    top = cfg.get("column_mappings")
    if isinstance(top, dict):
        out.append(top)
    for ext in (cfg.get("loan_data_extracts") or []):
        if isinstance(ext, dict) and isinstance(ext.get("column_mappings"), dict):
            out.append(ext["column_mappings"])
    return out


def backfill_from_configs(configs: Iterable[tuple[str, dict]]) -> dict:
    """Seed the learned store from already-validated client configs.

    ``configs`` = iterable of ``(config_id, config_dict)``. For each config
    not already processed, record every ``field -> header`` pair where the
    header is a non-empty STRING that is neither a factory placeholder nor an
    integer positional index (headerless configs map fields to int columns,
    which are meaningless as cross-CU header hints). Idempotent: a
    ``config_id`` is recorded at most once, tracked in the store under
    ``backfilled_config_ids``.

    Returns a summary ``{"configs_processed", "configs_skipped",
    "pairs_recorded"}``.
    """
    processed = 0
    skipped = 0
    pairs = 0
    with _LOCK:
        store = load()
        counts = store["field_to_header_counts"]
        displays = store["header_display_forms"]
        done = set(store.get("backfilled_config_ids") or [])
        changed = False
        for config_id, cfg in configs:
            if not config_id or config_id in done:
                skipped += 1
                continue
            if not isinstance(cfg, dict):
                skipped += 1
                continue
            recorded_any = False
            for cmap in _iter_config_mappings(cfg):
                for field, header in cmap.items():
                    if field not in KNOWN_FIELDS:
                        continue
                    # Skip int positional indices (headerless CUs) and
                    # non-string values outright.
                    if not isinstance(header, str):
                        continue
                    raw = header.strip()
                    if not raw or raw in _PLACEHOLDER_HEADERS:
                        continue
                    norm = _normalize(raw)
                    bucket = counts.setdefault(field, {})
                    bucket[norm] = int(bucket.get(norm, 0)) + 1
                    displays.setdefault(norm, raw)
                    pairs += 1
                    recorded_any = True
                    changed = True
            done.add(config_id)
            changed = True
            if recorded_any:
                processed += 1
            else:
                skipped += 1
        if changed:
            store["backfilled_config_ids"] = sorted(done)
            _save(store)
    return {
        "configs_processed": processed,
        "configs_skipped": skipped,
        "pairs_recorded": pairs,
    }


def backfill_from_config_dir(config_dir) -> dict:
    """Convenience wrapper: load every ``*.yaml`` under *config_dir* and
    feed it to :func:`backfill_from_configs`. Configs that fail to parse are
    skipped silently. Returns the same summary dict.
    """
    import yaml  # local import: keep the store module free of a hard dep

    configs: list[tuple[str, dict]] = []
    for path in sorted(Path(config_dir).glob("*.yaml")):
        if path.stem.startswith(("_", "sample", "client", "template")):
            continue  # skip templates / scaffolding configs
        try:
            cfg = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
        except Exception:  # noqa: BLE001
            continue
        configs.append((path.stem, cfg))
    return backfill_from_configs(configs)
