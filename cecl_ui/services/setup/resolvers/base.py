"""Specialist-resolver foundation for the Setup Router.

Each setup decision is owned by a **specialist resolver** with a fixed
four-part contract:

1. ``resolve_local(ctx)``  — deterministic heuristics + registry priors + any
   high-trust source already in context (e.g. a parsed WARM). This is the
   default path; the LLM usually never fires.
2. ``escalate_ai(ctx)``    — a *tight, single-decision* AI prompt with only
   this resolver's slice of context + its tools. Fired ONLY on a local miss.
3. ``validate(ctx, res)``  — this decision's own ground-truth oracle
   (tie-to-book, impaired match, pull-join rate, reserve reconciliation).
   Either the resolver self-validates, or validation is deferred to the final
   whole-config probe.
4. output ``Resolution``   — value + tier + source + evidence, written to the
   registry under its ``decision_kind``.

The **orchestrator** (:func:`run_dag`) sequences resolvers in dependency order,
hands each only its slice of context, and accumulates results in
``ctx.resolved`` so downstream specialists can read upstream decisions.

Guardrail: a specialist is only allowed to be *autonomous* when it has an
oracle. A resolver with no self-validation and no place in the final probe
must return ``tier='red'`` / ``source='needs_input'`` and surface to a human.
"""

from __future__ import annotations

import abc
from dataclasses import dataclass, field
from typing import Any, Optional

Tier = str  # "green" | "yellow" | "red"
TIER_RANK = {"green": 2, "yellow": 1, "red": 0}


@dataclass
class Resolution:
    """The output of one specialist for one decision_kind."""
    kind: str
    value: Any = None
    tier: Tier = "red"
    # registry | detector | warm | ai | profile | derived | needs_input
    source: str = "none"
    confidence: float | None = None
    evidence: str = ""
    alternatives: list = field(default_factory=list)
    # True when resolve_local already checked ground truth (e.g. the pool-code
    # detector's tie-to-book). False = validation deferred to the final probe.
    self_validated: bool = False
    ai_used: bool = False

    def is_actionable(self) -> bool:
        return self.value is not None and self.tier in ("green", "yellow")


@dataclass
class ResolverContext:
    """Shared, mutable state threaded through the DAG. Each resolver reads the
    slice it needs and writes its result into ``resolved``."""
    engine: Any = None
    short: str = ""
    delivery_folder: str = ""
    profiles: dict = field(default_factory=dict)      # archetype -> FileProfile
    all_profiles: list = field(default_factory=list)  # every profiled delivery file
    base_cfg: dict = field(default_factory=dict)
    monthly_by_pool: dict | None = None               # the book {pool: balance}
    warm: dict | None = None                          # parsed WARM (if any)
    warm_path: str | None = None
    warm_folder: str | None = None
    snapshot_date: str | None = None
    # derived scratch (filled by the loan-file resolver / orchestrator)
    loan_path: str | None = None
    header_row: int = 1
    balance_col: str | None = None
    resolved: dict = field(default_factory=dict)      # kind -> Resolution
    notes: list = field(default_factory=list)

    def value_of(self, kind: str, default: Any = None) -> Any:
        r = self.resolved.get(kind)
        return r.value if r is not None else default


class Resolver(abc.ABC):
    """Base specialist. Subclasses set ``kind`` + ``depends_on`` and implement
    ``resolve_local``; ``escalate_ai`` and ``validate`` are optional hooks."""

    kind: str = ""
    depends_on: tuple[str, ...] = ()

    def applies(self, ctx: ResolverContext) -> bool:
        """Whether this specialist runs for this context (e.g. participation
        resolver only when a participation file is present)."""
        return True

    @abc.abstractmethod
    def resolve_local(self, ctx: ResolverContext) -> Optional[Resolution]:
        """Deterministic / registry / high-trust-source path. Return None (or a
        red Resolution) to signal a local miss that may trigger AI escalation."""
        raise NotImplementedError

    def escalate_ai(self, ctx: ResolverContext) -> Optional[Resolution]:
        """Specialist AI escalation, fired ONLY on a local miss. Not wired to a
        live model yet — the contract is a tight, single-decision prompt with
        just this resolver's context + tools, returning a candidate Resolution
        (tier='yellow' until validated). Default: no escalation."""
        return None

    def validate(self, ctx: ResolverContext, res: Resolution) -> Resolution:
        """Apply this decision's ground-truth oracle. Default: no-op (deferred
        to the final whole-config probe). Override to self-validate cheaply."""
        return res

    def run(self, ctx: ResolverContext) -> Resolution:
        res = self.resolve_local(ctx)
        if res is None or res.tier == "red":
            ai = self.escalate_ai(ctx)
            if ai is not None:
                ai.ai_used = True
                res = ai
        if res is None:
            res = Resolution(self.kind, tier="red", source="needs_input",
                             evidence="no local result and no AI escalation")
        return self.validate(ctx, res)


# ---------------------------------------------------------------------------
# DAG orchestration
# ---------------------------------------------------------------------------
def _toposort(resolvers: list[Resolver]) -> list[Resolver]:
    by_kind = {r.kind: r for r in resolvers}
    visited: set[str] = set()
    order: list[Resolver] = []

    def visit(r: Resolver, stack: frozenset[str]) -> None:
        if r.kind in visited:
            return
        if r.kind in stack:
            raise ValueError(f"cyclic resolver dependency at {r.kind!r}")
        for dep in r.depends_on:
            d = by_kind.get(dep)
            if d is not None:
                visit(d, stack | {r.kind})
        visited.add(r.kind)
        order.append(r)

    for r in resolvers:
        visit(r, frozenset())
    return order


def run_dag(ctx: ResolverContext,
            resolvers: list[Resolver]) -> dict[str, Resolution]:
    """Run resolvers in dependency order; skip those whose ``applies`` is
    False; accumulate results in ``ctx.resolved``. Returns the result map."""
    for r in _toposort(resolvers):
        if not r.applies(ctx):
            continue
        ctx.resolved[r.kind] = r.run(ctx)
    return ctx.resolved


def overall_tier(resolutions: dict[str, Resolution]) -> Tier:
    """Worst tier across all actionable decisions (a red anywhere = red)."""
    if not resolutions:
        return "red"
    return min((r.tier for r in resolutions.values()),
               key=lambda t: TIER_RANK.get(t, 0))
