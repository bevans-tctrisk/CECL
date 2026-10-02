# Setup Router — Specialist Resolver Architecture

**Status:** Foundation built + proven on Tongass · **Date:** 2026-07-23
**Package:** `cecl_ui/services/setup/resolvers/`

Decomposes client setup into **specialist resolvers** — one per decision that
needs judgment *and* has its own ground-truth oracle — sequenced by a thin DAG
orchestrator. Replaces the tendency toward one bloated setup agent with focused,
independently-validatable, independently-learning specialists.

---

## The contract (`resolvers/base.py`)

Every specialist owns one `decision_kind` and implements a four-part contract:

| Part | Method | Purpose |
|---|---|---|
| 1. Local | `resolve_local(ctx)` | deterministic heuristics + registry priors + high-trust sources (e.g. a parsed WARM). The default path — the LLM usually never fires. |
| 2. Escalate | `escalate_ai(ctx)` | a **tight, single-decision** AI prompt with only this resolver's slice of context + tools. Fired **only on a local miss**. (Hook defined; live model not yet wired.) |
| 3. Validate | `validate(ctx, res)` | this decision's ground-truth **oracle**. Either self-validate cheaply, or defer to the final whole-config probe. |
| 4. Output | `Resolution` | value + tier (green/yellow/red) + source + evidence + `self_validated` + `ai_used`, written to the registry under its `decision_kind`. |

`ResolverContext` threads shared state (engine, profiles, base_cfg, monthly
book, parsed WARM, loan-file scratch) and accumulates results in `ctx.resolved`
so downstream specialists read upstream decisions.

**Guardrail:** a specialist may only be *autonomous* when it has an oracle. No
oracle ⇒ it must return `red` / `needs_input` and surface to a human.

---

## The decision DAG (dependencies + oracle per node)

```
loan_file
   ├─► pool_code ──► pool_map ──► secondary_extracts
   │                     └──────► reserve_config
member_account
score_source
monthly_balance
```

| Specialist | `kind` | depends_on | Oracle (ground truth) | Status |
|---|---|---|---|---|
| Loan file | `loan_file` | — | balance column resolves | **built** |
| Member/account | `member_account` | — | member/impaired join rate (probe) | **built** |
| Pool-code column | `pool_code` | `loan_file` | **tie-to-book** (self-validated by detector) | **built** |
| Score sources | `score_source` | — | pull-join rate + movement sanity (probe) | **built** |
| Pool map | `pool_map` | `pool_code` | per-pool **tie-to-book** (self-validated by bootstrap) | **built** |
| Monthly balance | `monthly_balance` | — | book totals reconcile | **built** |
| Secondary extracts | `secondary_extracts` | `pool_map` | pool balance gap closes | **built** |
| Reserve config | `reserve_config` | `pool_map` | **reserve reconciles** to WARM | **built** |

The **orchestrator** (`base.run_dag`) topologically sorts by `depends_on`, skips
resolvers whose `applies()` is False, and runs each: `resolve_local` → (on miss)
`escalate_ai` → `validate`. `overall_tier` = worst tier across decisions.

---

## WARM as a high-trust source (not a separate path)

Rather than a monolithic "WARM fast-lane," each specialist checks `ctx.warm`
first and, when present, resolves from it (`source='warm'`, tier green) — but the
value is still exposed for independent validation. This keeps one uniform code
path for WARM and non-WARM CUs; the WARM just short-circuits most `resolve_local`
bodies to a trusted answer.

Proven on Tongass (WARM present): DAG order `loan_file → member_account →
pool_code → score_source`; all four resolve green from the WARM
(`LNDCBAL` / `fixed_suffix len 3` / `LNDALTC` / `LNDCRSC`+credit_pull).

---

## Why this shape (vs. one big agent)

- **Sharper + cheaper AI** — each specialist escalates only its decision with a
  narrow prompt + narrow tools; no mega-prompt.
- **Failure isolation + re-runnability** — a bad `pool_map` doesn't touch
  `member_account`; re-run one specialist without redoing everything.
- **Independent learning** — the registry already partitions on `decision_kind`;
  each specialist promotes on its own evidence.
- **Parallelism** — independent DAG branches can run concurrently.
- **Narrow tool surfaces** — the pool-code specialist gets balance-tie tools; the
  WARM/formula specialist gets formula-tracing tools; nobody carries tools they
  can't misuse.

---

## Migration plan (incremental, low-risk)

1. **[done]** Foundation: `base.py` (contract + DAG runner) + 4 specialists
   (`loan_file`, `member_account`, `pool_code`, `score_source`), wrapping the
   existing detectors/registry — no logic duplication.
2. **[done]** Remaining decisions migrated onto the contract: `pool_map` (wraps
   `pool_map_bootstrap`), `secondary_extracts` (participation), `monthly_balance`,
   `reserve_config` (shares `config_assembler.derive_reserve_config`). All 8
   resolve green from the WARM on Tongass; non-WARM degrades gracefully.
3. **[done]** `setup/router.py` re-pointed: `propose()` builds a
   `ResolverContext` + `run_dag`, converts each `Resolution` to the existing
   `DecisionProposal`, and assembles the draft via `_assemble_draft`. Accepts an
   optional parsed `warm` + `warm_folder`. Also fixed `_profile_delivery` to keep
   the **biggest** file per archetype (so a big Symitar `LND*` extract isn't
   evicted from the `unknown` bucket by a small history file) and added
   `_find_loan_profile` (locate the extract by its known balance column when the
   profiler's keyword test misses it). Proven on Tongass: all 8 green, draft
   fully assembled.
4. Add the final **whole-config probe** (`probe_validator`) as the orchestrator's
   end-to-end gate after assembly. *(wired: `propose(validate=True)` calls
   `probe_validate` when a pool_map + book are present.)*
5. Wire `escalate_ai` to a live specialist prompt per `decision_kind` (opt-in,
   local-first, cached to the registry) once a genuine local miss appears.

Split `specialists.py` into one file per resolver when the set grows.
