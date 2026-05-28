"""Save / load in-progress wizard state to disk so users can pause work
on one credit union (e.g. while waiting on more data) and resume later
without losing what they've entered.

Each draft is one JSON file under ``<workspace_root>/wizard_drafts/``.
The filename now includes the model so a single CU can have BOTH a
Migration draft and a SCALE draft side-by-side without collision:

    <short_name_slug>__<model>.json

Models: ``migration`` (the original TCT/Vizo CECL wizard) and
``scale`` (the SCALE wizard). Backward compat: legacy files named
just ``<slug>.json`` are still readable and are treated as migration
drafts. The next save migrates them to the new naming.
"""
from __future__ import annotations

import json
import re
from datetime import datetime
from pathlib import Path
from typing import Any


DRAFTS_DIR_NAME = "wizard_drafts"
DRAFT_META_KEY = "_draft_meta"

MODELS = ("migration", "scale")
DEFAULT_MODEL = "migration"


def drafts_dir(workspace_root: str | Path) -> Path:
    d = Path(workspace_root) / DRAFTS_DIR_NAME
    d.mkdir(parents=True, exist_ok=True)
    return d


def _slug(s: str) -> str:
    s = re.sub(r"[^A-Za-z0-9_\-]+", "_", (s or "").strip()).strip("_")
    return s or "_unnamed"


def _normalize_model(model: str | None) -> str:
    m = (model or "").strip().lower()
    if m not in MODELS:
        return DEFAULT_MODEL
    return m


def _file_for(workspace_root: str | Path, slug: str, model: str) -> Path:
    """Return the per-model filename for a draft."""
    return drafts_dir(workspace_root) / f"{slug}__{model}.json"


def _legacy_file(workspace_root: str | Path, slug: str) -> Path:
    """Return the pre-multi-model filename (no model suffix)."""
    return drafts_dir(workspace_root) / f"{slug}.json"


def draft_key_for_state(state: dict[str, Any]) -> str:
    """Pick the on-disk slug for a draft based on state.

    Priority: short_name (slugified) > credit_union slug > "_unnamed".
    """
    return _slug(state.get("short_name") or state.get("credit_union") or "")


def model_for_state(state: dict[str, Any]) -> str:
    """Infer the model from state (``state['model']`` if set)."""
    return _normalize_model(state.get("model"))


def _default(o: Any) -> Any:
    # Fallback for anything json doesn't know how to handle (Path, etc).
    if isinstance(o, Path):
        return str(o)
    return str(o)


def save_draft(
    workspace_root: str | Path,
    state: dict[str, Any],
    *,
    active_step: str = "",
    model: str | None = None,
) -> Path:
    """Write ``state`` to disk and return the saved path.

    The output file is per-CU **and per-model**; a CU can have both a
    Migration draft and a SCALE draft on disk at once. If ``model`` is
    not supplied, it's inferred from ``state['model']`` (defaults to
    ``migration``).

    One-time migration: if a legacy ``<slug>.json`` exists for this CU
    AND we're writing the migration variant, the legacy file is
    removed after the new file is written successfully.
    """
    slug = draft_key_for_state(state)
    m = _normalize_model(model if model is not None else state.get("model"))

    meta = dict(state.get(DRAFT_META_KEY) or {})
    meta["saved_at"] = datetime.now().isoformat(timespec="seconds")
    meta["model"] = m
    if active_step:
        meta["active_step"] = active_step
    elif not meta.get("active_step"):
        meta["active_step"] = "identity"
    state[DRAFT_META_KEY] = meta

    out = _file_for(workspace_root, slug, m)
    out.write_text(
        json.dumps(state, default=_default, indent=2),
        encoding="utf-8",
    )

    # Retire legacy un-suffixed file once we've successfully written
    # the new per-model file (legacy files are implicitly migration).
    if m == "migration":
        legacy = _legacy_file(workspace_root, slug)
        if legacy.exists() and legacy.resolve() != out.resolve():
            try:
                legacy.unlink()
            except OSError:
                pass
    return out


def load_draft(
    workspace_root: str | Path,
    key: str,
    model: str | None = None,
) -> dict[str, Any] | None:
    """Read a draft for the given CU slug and model.

    Falls back to the legacy un-suffixed file when no per-model file
    exists and the requested model is ``migration``.
    """
    slug = _slug(key)
    m = _normalize_model(model)
    p = _file_for(workspace_root, slug, m)
    if not p.exists() and m == DEFAULT_MODEL:
        p = _legacy_file(workspace_root, slug)
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding="utf-8"))
    except Exception:  # noqa: BLE001
        return None


def _parse_draft_filename(stem: str) -> tuple[str, str]:
    """Return ``(slug, model)`` for a draft file's stem.

    Legacy files (no ``__model`` suffix) are treated as migration.
    """
    for m in MODELS:
        suffix = f"__{m}"
        if stem.endswith(suffix):
            return stem[: -len(suffix)], m
    return stem, DEFAULT_MODEL


def list_drafts(workspace_root: str | Path) -> list[dict[str, Any]]:
    """Return a list of draft summaries, newest first.

    One entry per (CU, model) file. ``key`` is the CU slug (no model
    suffix) so resume/delete URLs can pass slug + model separately.
    """
    out: list[dict[str, Any]] = []
    d = drafts_dir(workspace_root)
    for p in d.glob("*.json"):
        try:
            data = json.loads(p.read_text(encoding="utf-8"))
        except Exception:  # noqa: BLE001
            continue
        slug, model = _parse_draft_filename(p.stem)
        meta = data.get(DRAFT_META_KEY) or {}
        out.append({
            "key": slug,
            "model": meta.get("model") or model,
            "credit_union": data.get("credit_union") or slug,
            "short_name": data.get("short_name") or "",
            "saved_at": meta.get("saved_at") or "",
            "active_step": meta.get("active_step") or "identity",
            "has_warm": data.get("has_warm_files") or "",
            "completed_at": meta.get("completed_at") or "",
        })
    out.sort(key=lambda r: r.get("saved_at") or "", reverse=True)
    return out


def mark_completed(
    workspace_root: str | Path,
    key: str,
    model: str | None = None,
) -> bool:
    """Stamp a draft as completed setup. Returns True on success.

    Writes ``_draft_meta.completed_at`` (ISO timestamp) so the home
    dashboard can move the entry from "Resume in-progress setup" to
    "Completed setup". The rest of the state is untouched -- the user
    can still Edit (resume) or Delete it later.
    """
    data = load_draft(workspace_root, key, model=model)
    if not data:
        return False
    meta = dict(data.get(DRAFT_META_KEY) or {})
    meta["completed_at"] = datetime.now().isoformat(timespec="seconds")
    meta["model"] = _normalize_model(model if model is not None else data.get("model"))
    if not meta.get("active_step"):
        meta["active_step"] = data.get("_active_step") or "identity"
    data[DRAFT_META_KEY] = meta
    out = _file_for(workspace_root, _slug(key), meta["model"])
    out.write_text(
        json.dumps(data, default=_default, indent=2),
        encoding="utf-8",
    )
    return True


def delete_draft(
    workspace_root: str | Path,
    key: str,
    model: str | None = None,
) -> bool:
    """Delete a draft. Returns True if at least one file was removed.

    When ``model == 'migration'`` also removes any legacy un-suffixed
    file so resurrected legacy data doesn't reappear after delete.
    """
    slug = _slug(key)
    m = _normalize_model(model)
    removed = False
    p = _file_for(workspace_root, slug, m)
    if p.exists():
        p.unlink()
        removed = True
    if m == DEFAULT_MODEL:
        legacy = _legacy_file(workspace_root, slug)
        if legacy.exists():
            legacy.unlink()
            removed = True
    return removed
