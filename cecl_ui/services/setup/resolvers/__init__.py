"""Specialist-resolver architecture for the Setup Router.

Each setup decision is owned by a specialist resolver with a uniform
four-part contract (resolve_local → escalate_ai → validate → Resolution),
sequenced by a small DAG orchestrator. See :mod:`base` for the contract and
:mod:`specialists` for the concrete resolvers.

Design north-star: *router proposes → ground truth validates → human
approves*. A specialist may only be autonomous when it has an oracle.
"""

from __future__ import annotations

from .base import (
    Resolution,
    Resolver,
    ResolverContext,
    overall_tier,
    run_dag,
)
from .specialists import build_default_resolvers

__all__ = [
    "Resolution",
    "Resolver",
    "ResolverContext",
    "run_dag",
    "overall_tier",
    "build_default_resolvers",
]
