"""Setup pattern registry (Setup Router, Phase 0).

Postgres-backed store of reusable onboarding patterns learned from prior
client setups. Mirrors the report ``schema_registry`` but adds a **trust
tier** the report registry doesn't need, because setup patterns propagate to
*new* clients (larger blast radius) — see docs/setup_router_design.md §4/§10.

Governance (Q4):
- ``source``  : 'mined' (reconstructed from a production config) | 'manual' | 'ai'
- ``trust``   : 'trusted' | 'candidate'
    * mined + manual  -> born 'trusted' (proven in production / human-authored)
    * ai              -> born 'candidate'; promotes to 'trusted' after it has
                         been validated-and-approved on N>=2 distinct clients.
- every pattern carries ``evidence`` (support count + contributing CUs) = audit trail.

Patterns are aggregated: the same (decision_kind, fingerprint, fragment) seen
in K configs is ONE row with ``support=K``, not K rows.
"""

from __future__ import annotations

import hashlib
import json
from typing import Any

from sqlalchemy import text
from sqlalchemy.engine import Engine

_TRUST_RANK = {"candidate": 0, "trusted": 1}
PROMOTE_THRESHOLD = 2  # distinct validated+approved clients to promote an AI pattern


# ---------------------------------------------------------------------------
# helpers
# ---------------------------------------------------------------------------
def canonical_json(obj: Any) -> str:
    return json.dumps(obj, sort_keys=True, separators=(",", ":"), default=str)


def _checksum(decision_kind: str, fingerprint: dict, fragment: dict) -> str:
    payload = canonical_json(
        {"k": decision_kind, "f": fingerprint, "c": fragment}
    )
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


# ---------------------------------------------------------------------------
# schema
# ---------------------------------------------------------------------------
def ensure_tables(engine: Engine) -> None:
    ddl_patterns = """
    CREATE TABLE IF NOT EXISTS setup_patterns (
        id             SERIAL PRIMARY KEY,
        decision_kind  TEXT        NOT NULL,
        fingerprint    JSONB       NOT NULL,
        fragment       JSONB       NOT NULL,
        evidence       JSONB       NOT NULL DEFAULT '{}'::jsonb,
        source         TEXT        NOT NULL DEFAULT 'mined',
        trust          TEXT        NOT NULL DEFAULT 'trusted',
        checksum       TEXT        NOT NULL UNIQUE,
        support        INTEGER     NOT NULL DEFAULT 1,
        approvals      INTEGER     NOT NULL DEFAULT 0,
        version        INTEGER     NOT NULL DEFAULT 1,
        is_active      BOOLEAN     NOT NULL DEFAULT TRUE,
        created_at     TIMESTAMP   NOT NULL DEFAULT now(),
        updated_at     TIMESTAMP   NOT NULL DEFAULT now()
    );
    """
    idx = """
    CREATE INDEX IF NOT EXISTS ix_setup_patterns_kind
        ON setup_patterns (decision_kind, is_active);
    """
    ddl_metrics = """
    CREATE TABLE IF NOT EXISTS setup_metrics (
        id            SERIAL PRIMARY KEY,
        client        TEXT,
        decision_kind TEXT,
        resolved_by   TEXT,        -- 'local' | 'ai' | 'manual'
        used_ai       BOOLEAN     NOT NULL DEFAULT FALSE,
        validated     BOOLEAN     NOT NULL DEFAULT FALSE,
        approved      BOOLEAN     NOT NULL DEFAULT FALSE,
        created_at    TIMESTAMP   NOT NULL DEFAULT now()
    );
    """
    with engine.begin() as c:
        c.execute(text(ddl_patterns))
        c.execute(text(idx))
        c.execute(text(ddl_metrics))


def reset(engine: Engine) -> None:
    """Clear all mined patterns/metrics (regenerable from configs). Used to
    re-seed cleanly after a pattern-shape change. Does NOT drop tables."""
    ensure_tables(engine)
    with engine.begin() as c:
        c.execute(text("DELETE FROM setup_patterns"))
        c.execute(text("DELETE FROM setup_metrics"))


# ---------------------------------------------------------------------------
# write
# ---------------------------------------------------------------------------
def save_pattern(engine: Engine, decision_kind: str, fingerprint: dict,
                 fragment: dict, *, source: str = "mined",
                 trust: str | None = None, support: int = 1,
                 cu: str | None = None,
                 evidence: dict | None = None) -> dict[str, Any]:
    """Idempotent upsert keyed by checksum(decision_kind, fingerprint, fragment).

    On conflict, merges support and the contributing-CU set and keeps the
    higher trust tier. Returns ``{checksum, support, trust, created}``.
    """
    if trust is None:
        trust = "candidate" if source == "ai" else "trusted"
    ck = _checksum(decision_kind, fingerprint, fragment)
    ev = dict(evidence or {})
    cus = set(ev.get("cus", []))
    if cu:
        cus.add(cu)

    with engine.begin() as c:
        row = c.execute(
            text("SELECT id, support, trust, evidence FROM setup_patterns "
                 "WHERE checksum = :ck"),
            {"ck": ck},
        ).mappings().first()

        if row:
            prev_ev = row["evidence"] or {}
            merged_cus = sorted(set(prev_ev.get("cus", [])) | cus)
            # Support = distinct contributing CUs (idempotent on re-seed).
            # Fall back to a running increment only when no CU is tracked
            # (e.g. manual/ai patterns recorded without a source CU).
            if merged_cus:
                new_support = len(merged_cus)
            else:
                new_support = int(row["support"]) + int(support)
            new_trust = row["trust"] if _TRUST_RANK[row["trust"]] >= _TRUST_RANK[trust] else trust
            new_ev = {**prev_ev, **ev, "cus": merged_cus, "support": new_support}
            c.execute(
                text("UPDATE setup_patterns SET support=:s, trust=:t, "
                     "evidence=CAST(:e AS jsonb), updated_at=now() "
                     "WHERE id=:id"),
                {"s": new_support, "t": new_trust,
                 "e": canonical_json(new_ev), "id": row["id"]},
            )
            return {"checksum": ck, "support": new_support,
                    "trust": new_trust, "created": False}

        new_ev = {**ev, "cus": sorted(cus), "support": int(support)}
        eff_support = len(cus) if cus else int(support)
        new_ev["support"] = eff_support
        c.execute(
            text("INSERT INTO setup_patterns "
                 "(decision_kind, fingerprint, fragment, evidence, source, "
                 " trust, checksum, support) VALUES "
                 "(:k, CAST(:f AS jsonb), CAST(:c AS jsonb), CAST(:e AS jsonb), "
                 " :src, :t, :ck, :s)"),
            {"k": decision_kind, "f": canonical_json(fingerprint),
             "c": canonical_json(fragment), "e": canonical_json(new_ev),
             "src": source, "t": trust, "ck": ck, "s": eff_support},
        )
        return {"checksum": ck, "support": eff_support,
                "trust": trust, "created": True}


def record_metric(engine: Engine, *, client: str | None, decision_kind: str,
                  resolved_by: str, used_ai: bool = False,
                  validated: bool = False, approved: bool = False) -> None:
    with engine.begin() as c:
        c.execute(
            text("INSERT INTO setup_metrics "
                 "(client, decision_kind, resolved_by, used_ai, validated, approved) "
                 "VALUES (:cl,:k,:rb,:ua,:v,:a)"),
            {"cl": client, "k": decision_kind, "rb": resolved_by,
             "ua": used_ai, "v": validated, "a": approved},
        )


# ---------------------------------------------------------------------------
# read
# ---------------------------------------------------------------------------
def get_patterns(engine: Engine, decision_kind: str | None = None, *,
                 min_support: int = 1, trusted_only: bool = False,
                 fingerprint_filter: dict | None = None) -> list[dict[str, Any]]:
    """Return patterns (optionally for one decision_kind), best-support first."""
    q = ("SELECT decision_kind, fingerprint, fragment, evidence, source, "
         "trust, support, approvals FROM setup_patterns "
         "WHERE is_active = TRUE AND support >= :ms")
    params: dict[str, Any] = {"ms": min_support}
    if decision_kind:
        q += " AND decision_kind = :k"
        params["k"] = decision_kind
    if trusted_only:
        q += " AND trust = 'trusted'"
    q += " ORDER BY support DESC, updated_at DESC"
    with engine.connect() as c:
        rows = [dict(r) for r in c.execute(text(q), params).mappings().all()]
    if fingerprint_filter:
        rows = [r for r in rows
                if all(r["fingerprint"].get(k) == v
                       for k, v in fingerprint_filter.items())]
    return rows


def summary(engine: Engine) -> dict[str, dict[str, int]]:
    """Counts per decision_kind: patterns, total support, trusted count."""
    with engine.connect() as c:
        rows = c.execute(text(
            "SELECT decision_kind, COUNT(*) n, SUM(support) sup, "
            "SUM(CASE WHEN trust='trusted' THEN 1 ELSE 0 END) trusted "
            "FROM setup_patterns WHERE is_active=TRUE GROUP BY decision_kind "
            "ORDER BY decision_kind"
        )).mappings().all()
    return {r["decision_kind"]: {"patterns": int(r["n"]),
                                 "support": int(r["sup"] or 0),
                                 "trusted": int(r["trusted"] or 0)}
            for r in rows}
