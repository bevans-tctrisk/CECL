"""Self-learning schema storage + cross-CU analytics (PostgreSQL).

Two tables, created idempotently on first use so no separate migration step
is required:

``report_schemas``
    Versioned, validated report schemas per CU + report type. This is the
    "self-learning" store — every schema the router validates (locally or via
    AI) is persisted here so future quarters skip escalation. ``source`` tracks
    the manual -> local -> ai transition; ``is_active`` marks the current
    baseline for a (cu, report_type) pair.

``report_metrics``
    An optional thin fact table for report-derived metrics, keyed by
    ``(cu, snapshot_date, report_type, pool, metric_name)``. Populated on
    demand via :func:`record_metrics`; queried for peer comparisons.

The analytics helpers additionally read the existing ``monthly_loan_data``
table directly so peer comparisons and score-change trends work immediately,
without waiting for ``report_metrics`` to be backfilled.
"""
from __future__ import annotations

from typing import Any

from sqlalchemy import Engine, text

from .schemas import ReportSchema

# ---------------------------------------------------------------------------
# DDL — idempotent. JSONB keeps schemas flexible yet queryable.
# ---------------------------------------------------------------------------
_DDL_SCHEMAS = """
CREATE TABLE IF NOT EXISTS report_schemas (
    id           SERIAL PRIMARY KEY,
    cu           TEXT        NOT NULL,
    report_type  TEXT        NOT NULL,
    version      INTEGER     NOT NULL,
    schema_json  JSONB       NOT NULL,
    checksum     TEXT        NOT NULL,
    source       TEXT        NOT NULL DEFAULT 'local',
    is_active    BOOLEAN     NOT NULL DEFAULT TRUE,
    created_at   TIMESTAMPTZ NOT NULL DEFAULT now()
)
"""

_DDL_SCHEMAS_IDX = """
CREATE INDEX IF NOT EXISTS ix_report_schemas_cu_type
    ON report_schemas (cu, report_type, is_active)
"""

_DDL_METRICS = """
CREATE TABLE IF NOT EXISTS report_metrics (
    id            SERIAL PRIMARY KEY,
    cu            TEXT        NOT NULL,
    snapshot_date DATE        NOT NULL,
    report_type   TEXT        NOT NULL,
    pool          TEXT,
    metric_name   TEXT        NOT NULL,
    metric_value  DOUBLE PRECISION,
    created_at    TIMESTAMPTZ NOT NULL DEFAULT now()
)
"""

_DDL_METRICS_IDX = """
CREATE INDEX IF NOT EXISTS ix_report_metrics_lookup
    ON report_metrics (metric_name, snapshot_date, cu)
"""


def ensure_tables(engine: Engine) -> None:
    """Create the registry + metrics tables and indexes if absent."""
    with engine.begin() as conn:
        conn.execute(text(_DDL_SCHEMAS))
        conn.execute(text(_DDL_SCHEMAS_IDX))
        conn.execute(text(_DDL_METRICS))
        conn.execute(text(_DDL_METRICS_IDX))


# ---------------------------------------------------------------------------
# Schema registry
# ---------------------------------------------------------------------------
def get_active(engine: Engine, cu: str, report_type: str) -> dict[str, Any] | None:
    """Return the active schema row for (cu, report_type), or ``None``."""
    with engine.connect() as conn:
        row = conn.execute(
            text(
                "SELECT id, cu, report_type, version, schema_json, checksum, "
                "source, created_at FROM report_schemas "
                "WHERE cu = :cu AND report_type = :rt AND is_active = TRUE "
                "ORDER BY version DESC LIMIT 1"
            ),
            {"cu": cu, "rt": report_type},
        ).mappings().first()
    return dict(row) if row else None


def list_versions(engine: Engine, cu: str, report_type: str) -> list[dict[str, Any]]:
    """Return the full version history for a (cu, report_type) pair."""
    with engine.connect() as conn:
        rows = conn.execute(
            text(
                "SELECT id, version, checksum, source, is_active, created_at "
                "FROM report_schemas WHERE cu = :cu AND report_type = :rt "
                "ORDER BY version DESC"
            ),
            {"cu": cu, "rt": report_type},
        ).mappings().all()
    return [dict(r) for r in rows]


def save(engine: Engine, schema: ReportSchema, *, source: str | None = None) -> dict[str, Any]:
    """Persist ``schema`` as the new active baseline.

    Idempotent by checksum: if the active schema already matches, no new row
    is written and the existing row is returned. Otherwise a new version is
    inserted and any prior active row for the pair is deactivated.

    Returns ``{version, checksum, source, created: bool}``.
    """
    src = source or schema.source
    checksum = schema.checksum()
    cu = schema.credit_union
    rt = schema.report_type

    with engine.begin() as conn:
        existing = conn.execute(
            text(
                "SELECT version, checksum FROM report_schemas "
                "WHERE cu = :cu AND report_type = :rt AND is_active = TRUE "
                "ORDER BY version DESC LIMIT 1"
            ),
            {"cu": cu, "rt": rt},
        ).mappings().first()

        if existing and existing["checksum"] == checksum:
            return {
                "version": existing["version"],
                "checksum": checksum,
                "source": src,
                "created": False,
            }

        next_version = (existing["version"] + 1) if existing else 1
        # Deactivate the prior baseline so exactly one row stays active.
        conn.execute(
            text(
                "UPDATE report_schemas SET is_active = FALSE "
                "WHERE cu = :cu AND report_type = :rt AND is_active = TRUE"
            ),
            {"cu": cu, "rt": rt},
        )
        conn.execute(
            text(
                "INSERT INTO report_schemas "
                "(cu, report_type, version, schema_json, checksum, source, "
                " is_active) VALUES "
                "(:cu, :rt, :ver, CAST(:js AS JSONB), :ck, :src, TRUE)"
            ),
            {
                "cu": cu,
                "rt": rt,
                "ver": next_version,
                "js": schema.canonical_json(),
                "ck": checksum,
                "src": src,
            },
        )
    return {
        "version": next_version,
        "checksum": checksum,
        "source": src,
        "created": True,
    }


# ---------------------------------------------------------------------------
# Report metrics (optional store, for report-derived numbers)
# ---------------------------------------------------------------------------
def record_metrics(
    engine: Engine,
    cu: str,
    snapshot_date: str,
    report_type: str,
    metrics: dict[str, float],
    *,
    pool: str | None = None,
) -> int:
    """Insert a batch of ``{metric_name: value}`` rows. Returns count."""
    if not metrics:
        return 0
    payload = [
        {
            "cu": cu,
            "snap": snapshot_date,
            "rt": report_type,
            "pool": pool,
            "name": name,
            "value": float(value) if value is not None else None,
        }
        for name, value in metrics.items()
    ]
    with engine.begin() as conn:
        conn.execute(
            text(
                "INSERT INTO report_metrics "
                "(cu, snapshot_date, report_type, pool, metric_name, "
                " metric_value) VALUES "
                "(:cu, :snap, :rt, :pool, :name, :value)"
            ),
            payload,
        )
    return len(payload)


# ---------------------------------------------------------------------------
# Cross-CU analytics — read monthly_loan_data directly so results are
# available without backfilling report_metrics.
# ---------------------------------------------------------------------------
def peer_comparison(
    engine: Engine, peer_group: list[str], snapshot_date: str
) -> list[dict[str, Any]]:
    """Per-CU portfolio summary across a peer group at one snapshot.

    Returns one row per CU: loan count, total/avg balance, avg FICO.
    """
    if not peer_group:
        return []
    with engine.connect() as conn:
        rows = conn.execute(
            text(
                "SELECT credit_union, COUNT(*) AS loan_count, "
                "SUM(current_balance) AS total_balance, "
                "AVG(current_balance) AS avg_balance, "
                "AVG(current_fico_score) AS avg_fico "
                "FROM monthly_loan_data "
                "WHERE credit_union = ANY(:cus) AND snapshot_date = :snap "
                "GROUP BY credit_union ORDER BY credit_union"
            ),
            {"cus": list(peer_group), "snap": snapshot_date},
        ).mappings().all()
    return [dict(r) for r in rows]


def score_change_trend(engine: Engine, cu: str) -> list[dict[str, Any]]:
    """Average current FICO by snapshot for one CU (score-drift over time)."""
    with engine.connect() as conn:
        rows = conn.execute(
            text(
                "SELECT snapshot_date, AVG(current_fico_score) AS avg_fico, "
                "COUNT(*) AS loan_count "
                "FROM monthly_loan_data WHERE credit_union = :cu "
                "GROUP BY snapshot_date ORDER BY snapshot_date"
            ),
            {"cu": cu},
        ).mappings().all()
    return [dict(r) for r in rows]


def pool_distribution(
    engine: Engine, peer_group: list[str], snapshot_date: str
) -> list[dict[str, Any]]:
    """Balance by pool across a peer group at one snapshot (peer mix)."""
    if not peer_group:
        return []
    with engine.connect() as conn:
        rows = conn.execute(
            text(
                "SELECT credit_union, loan_pool, "
                "SUM(current_balance) AS total_balance, COUNT(*) AS loan_count "
                "FROM monthly_loan_data "
                "WHERE credit_union = ANY(:cus) AND snapshot_date = :snap "
                "AND loan_pool IS NOT NULL "
                "GROUP BY credit_union, loan_pool "
                "ORDER BY credit_union, loan_pool"
            ),
            {"cus": list(peer_group), "snap": snapshot_date},
        ).mappings().all()
    return [dict(r) for r in rows]
