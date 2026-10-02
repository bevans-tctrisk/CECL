"""Setup Router — self-learning client-onboarding assistant.

Proposes client configuration from raw delivery files, validates every
proposal against ground truth (balances tie to the monthly book, impaired
loans match), and learns reusable patterns from each successful setup.

See ``docs/setup_router_design.md`` for the full design. Core principle:
**the currency is validation, not model confidence** — a decision is trusted
because the data ties out, never because the model is confident.

This package is metadata-only at the profiling/AI stages: it never emits or
transmits loan-level rows or PII.
"""

from __future__ import annotations

__all__ = [
    "profiler",
    "pool_code_detector",
    "pool_map_bootstrap",
    "warm_parser",
    "config_assembler",
    "setup_registry",
    "seed_from_configs",
    "probe_validator",
    "router",
    "resolvers",
]
