"""Local validation — the first stage of the hybrid router.

Deterministic, offline checks of a CU's report schema. Two layers:

1. **Structural** — does the config's report-relevant slice satisfy the typed
   :class:`~cecl_ui.services.hybrid.schemas.ReportSchema` contract?
2. **Rule engine** — business rules Pydantic can't express on its own
   (required column mappings present, at least one non-excluded pool, grades
   cover the FICO range, drift from the registered baseline).

A failed validation — or a CU with no registered baseline yet — is what tells
the router to escalate to the AI stage.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from pydantic import ValidationError

from .schemas import REQUIRED_COLUMN_KEYS, ReportSchema


@dataclass
class ValidationResult:
    """Outcome of local validation for one CU + report type."""

    ok: bool
    schema: ReportSchema | None
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    has_baseline: bool = False
    needs_escalation: bool = False


def _rule_checks(schema: ReportSchema) -> tuple[list[str], list[str]]:
    """Business-rule checks beyond the Pydantic structural contract."""
    errors: list[str] = []
    warnings: list[str] = []

    if not (schema.credit_union or "").strip():
        errors.append("credit_union is blank")

    # Required column mappings.
    missing = [
        k for k in REQUIRED_COLUMN_KEYS
        if not (schema.column_mappings.get(k) or "").strip()
    ]
    if missing:
        errors.append(
            "missing required column mapping(s): " + ", ".join(missing)
        )

    # At least one usable (non-excluded) pool.
    active_pools = [p for p in schema.pools if not p.excluded]
    if not active_pools:
        errors.append("no active (non-excluded) loan pools defined")

    # Risk-rated pools need an ACL horizon.
    for p in active_pools:
        if p.risk_rated and (p.acl_months is None or p.acl_months <= 0):
            warnings.append(
                f"pool '{p.name}' is risk-rated but has no positive acl_months"
            )

    # Credit grades should span a plausible FICO range when any exist.
    if schema.credit_grades:
        lo = min(g.min_score for g in schema.credit_grades)
        hi = max(g.max_score for g in schema.credit_grades)
        if hi < 700:
            warnings.append(
                f"credit grades top out at {hi}; expected coverage up to ~850"
            )
        if lo > 350:
            warnings.append(
                f"credit grades start at {lo}; low scores may be uncovered"
            )
    else:
        warnings.append("no credit grades defined")

    return errors, warnings


def validate(
    cfg: dict[str, Any],
    report_type: str,
    *,
    baseline: ReportSchema | None = None,
) -> ValidationResult:
    """Validate a client config for one report type.

    ``baseline`` is the registered active schema (if any). When absent, the CU
    has no known-good schema yet and — under the proactive policy — escalation
    is requested so the AI can propose an initial baseline.
    """
    has_baseline = baseline is not None

    # Layer 1: structural / typed contract.
    try:
        schema = ReportSchema.from_config(cfg, report_type)
    except ValidationError as exc:
        errors = [
            f"{'.'.join(str(p) for p in e['loc'])}: {e['msg']}"
            for e in exc.errors()
        ]
        return ValidationResult(
            ok=False,
            schema=None,
            errors=errors,
            has_baseline=has_baseline,
            needs_escalation=True,
        )

    # Layer 2: business rules.
    errors, warnings = _rule_checks(schema)

    # Drift detection against the registered baseline (informational).
    if baseline is not None and baseline.checksum() != schema.checksum():
        warnings.append(
            "schema differs from the registered baseline (v-drift); "
            "the new schema will be re-registered on success"
        )

    if errors:
        return ValidationResult(
            ok=False,
            schema=schema,
            errors=errors,
            warnings=warnings,
            has_baseline=has_baseline,
            needs_escalation=True,
        )

    # Valid structurally + by rules. Proactive policy: with no baseline yet we
    # still request escalation so the AI can confirm/enrich the first schema,
    # but the report is never blocked either way.
    return ValidationResult(
        ok=True,
        schema=schema,
        errors=[],
        warnings=warnings,
        has_baseline=has_baseline,
        needs_escalation=not has_baseline,
    )
