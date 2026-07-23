"""Hybrid AI report router for the CECL wizard.

A three-stage pipeline layered *in front of* the existing report engine:

1. **Local Validation First** (:mod:`local_validator`) — validate a CU's
   report schema against a typed Pydantic contract (:mod:`schemas`) and the
   last known-good schema in the registry. Zero network, deterministic.
2. **API Escalation** (:mod:`ai_escalation`) — when local validation fails, or
   a CU has no registered schema yet, escalate to ``claude-opus-4.8`` via the
   Anthropic SDK to propose a corrected/complete schema.
3. **Self-Learning Storage** (:mod:`schema_registry`) — persist newly
   validated schemas back to PostgreSQL so future quarters skip escalation,
   and expose cross-CU analytics (peer comparisons, score-change trends).

The orchestrator is :class:`router.HybridReportRouter`. Everything is
feature-flagged **off** by default (opt-in per CU via ``hybrid_router.enabled``
in the client YAML) and degrades gracefully to today's behavior whenever the
router, Pydantic, the Anthropic SDK, an API key, or the network is unavailable.
It never blocks report generation.
"""
from __future__ import annotations

__all__ = [
    "HybridReportRouter",
    "HybridDecision",
    "is_available",
]

from .router import HybridDecision, HybridReportRouter, is_available
