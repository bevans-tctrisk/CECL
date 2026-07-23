"""Hybrid router orchestrator.

Ties together local validation, AI escalation, and self-learning storage into
a single :meth:`HybridReportRouter.resolve` call the pipeline invokes *before*
handing off to the existing report engine.

Design guarantees (Option A rollout):
- **Never blocks** report generation. Every failure path degrades to today's
  behavior and is surfaced as an advisory note.
- **Off by default.** The feature flag lives in the client YAML under
  ``hybrid_router.enabled`` and is checked by the caller
  (:func:`pipeline_service.run_reports_hybrid`), not here.
- **Graceful when dependencies are missing** — no Pydantic, no Anthropic SDK,
  no API key, or no network all resolve to a degraded (but successful) decision.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

from cecl_audit_log import get_audit_logger


def is_available() -> bool:
    """True when the router's hard dependency (Pydantic) is importable."""
    try:
        import pydantic  # noqa: F401, PLC0415
        return True
    except ImportError:
        return False


@dataclass
class HybridDecision:
    """The router's per-(cu, report_type) verdict, for logging + UI."""

    cu: str
    report_type: str
    ok: bool
    action: str  # local_valid | ai_escalated | ai_failed | degraded | error
    source: str  # manual | local | ai | none
    schema_version: int | None = None
    used_api: bool = False
    errors: list[str] = field(default_factory=list)
    warnings: list[str] = field(default_factory=list)
    messages: list[str] = field(default_factory=list)

    def summary(self) -> str:
        bits = [f"{self.report_type}: {self.action}"]
        if self.schema_version is not None:
            bits.append(f"schema v{self.schema_version} ({self.source})")
        if self.errors:
            bits.append("errors: " + "; ".join(self.errors))
        return " | ".join(bits)


class HybridReportRouter:
    """Local-first, AI-fallback resolver with a self-learning registry."""

    def __init__(
        self,
        engine: Any,
        *,
        api_key: str | None = None,
        model: str | None = None,
        proactive: bool = True,
    ) -> None:
        self._engine = engine
        self._api_key = api_key
        self._proactive = proactive
        self._audit = get_audit_logger()
        # Imported here so the module loads even when Pydantic is absent;
        # is_available() gates whether a router is ever constructed.
        from .ai_escalation import DEFAULT_MODEL

        self._model = model or DEFAULT_MODEL

    def resolve(
        self, cfg: dict[str, Any], report_type: str
    ) -> HybridDecision:
        """Validate → (escalate) → persist, returning a decision."""
        from . import local_validator, schema_registry
        from .ai_escalation import escalate
        from .schemas import ReportSchema

        cu = cfg.get("credit_union", "")

        # Ensure storage exists (idempotent, cheap).
        try:
            schema_registry.ensure_tables(self._engine)
        except Exception as exc:  # noqa: BLE001 - DB hiccup must not block reports
            self._audit.warning(
                "hybrid: registry unavailable for %s/%s: %s", cu, report_type, exc
            )
            return HybridDecision(
                cu=cu, report_type=report_type, ok=True, action="degraded",
                source="none",
                messages=[f"registry unavailable ({exc}); proceeding local-only"],
            )

        # Load the registered baseline, if any.
        baseline_schema: ReportSchema | None = None
        baseline_row = schema_registry.get_active(self._engine, cu, report_type)
        if baseline_row:
            try:
                baseline_schema = ReportSchema.model_validate(
                    baseline_row["schema_json"]
                )
            except Exception:  # noqa: BLE001 - stale/incompatible baseline
                baseline_schema = None

        # Stage 1: local validation.
        result = local_validator.validate(
            cfg, report_type, baseline=baseline_schema
        )

        if result.ok and not result.needs_escalation:
            return self._register_and_decide(
                schema_registry, result.schema, "local_valid", "local",
                warnings=result.warnings,
            )

        # Stage 2: escalation (either validation failed, or proactive-no-baseline).
        esc = escalate(
            cfg, report_type,
            validation_errors=result.errors,
            api_key=self._api_key,
            model=self._model,
        )
        self._audit.info(
            "hybrid: escalation for %s/%s used_api=%s model=%s",
            cu, report_type, esc.used_api, esc.model,
        )

        if esc.schema is not None:
            decision = self._register_and_decide(
                schema_registry, esc.schema, "ai_escalated", "ai",
                warnings=result.warnings, used_api=True,
            )
            decision.messages.extend(esc.messages)
            return decision

        # Escalation unavailable/failed. Decide based on local result:
        #   - locally valid (proactive-only trigger) -> keep local schema.
        #   - locally invalid -> degrade (proceed as today) with warnings.
        if result.ok and result.schema is not None:
            decision = self._register_and_decide(
                schema_registry, result.schema, "local_valid", "local",
                warnings=result.warnings,
            )
            decision.messages.extend(esc.messages)
            decision.errors.extend(esc.errors)
            return decision

        action = "ai_failed" if esc.used_api else "degraded"
        self._audit.warning(
            "hybrid: %s for %s/%s; proceeding without a validated schema",
            action, cu, report_type,
        )
        return HybridDecision(
            cu=cu, report_type=report_type, ok=True, action=action,
            source="none", used_api=esc.used_api,
            errors=result.errors + esc.errors,
            warnings=result.warnings,
            messages=esc.messages + ["proceeding with existing config (no gate)"],
        )

    def _register_and_decide(
        self,
        schema_registry,
        schema,
        action: str,
        source: str,
        *,
        warnings: list[str] | None = None,
        used_api: bool = False,
    ) -> HybridDecision:
        """Persist a validated schema and build the decision."""
        try:
            saved = schema_registry.save(self._engine, schema, source=source)
            version = saved["version"]
            if saved["created"]:
                self._audit.info(
                    "hybrid: registered %s/%s schema v%s (%s)",
                    schema.credit_union, schema.report_type, version, source,
                )
        except Exception as exc:  # noqa: BLE001 - storage must not block reports
            self._audit.warning(
                "hybrid: could not persist %s/%s schema: %s",
                schema.credit_union, schema.report_type, exc,
            )
            version = None
        return HybridDecision(
            cu=schema.credit_union,
            report_type=schema.report_type,
            ok=True,
            action=action,
            source=source,
            schema_version=version,
            used_api=used_api,
            warnings=list(warnings or []),
        )
