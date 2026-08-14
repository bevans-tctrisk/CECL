# Hybrid AI Models — Complete Overview

**Status:** Working reference · **Last updated:** 2026-07-23
**Scope:** Everything built to date around the "hybrid" AI-assisted automation in the CECL
Migration tool — the two routers, the WARM ingestion path, the pattern registries, and the
origination-aware scoring engine.

> **North-star principle (applies to everything below):**
> **Router proposes → ground truth validates → human approves.**
> Nothing is trusted because a model is confident. It is trusted because the data *ties out* —
> pool balances reconcile to the monthly book, impaired loans match the extract, and the
> generated reserve reconciles to the analyst's WARM. Because these configs drive regulatory
> CECL reserves, a validated result plus human sign-off is mandatory before anything goes live.

---

## 0. The big picture

There are **two independent hybrid routers**, operating one stage apart, plus a **WARM
ingestion path** that feeds the setup router for the subset of clients delivered with a
CECL-Migration-WARM workbook.

```
   ┌─────────────────────────────────────────────────────────────────────┐
   │  CLIENT ONBOARDING (once per client)                                │
   │                                                                     │
   │   raw delivery files ──►  SETUP ROUTER  ──►  client config (YAML)   │
   │        + (if present)     cecl_ui/services/setup/                    │
   │        WARM workbook ──►   WARM INGESTION PATH                       │
   │                           (warm_parser + config_assembler)          │
   │                                                                     │
   │   registry: setup_patterns / setup_metrics (Postgres)               │
   └─────────────────────────────────────────────────────────────────────┘
                                   │  config drives every quarter
                                   ▼
   ┌─────────────────────────────────────────────────────────────────────┐
   │  QUARTERLY REPORTING (every period)                                 │
   │                                                                     │
   │   loan extract ──► import_data ──► REPORT ENGINE ──► TCT/Vizo model │
   │                                    generate_report.py               │
   │                                         ▲                           │
   │                       HYBRID REPORT ROUTER validates the schema     │
   │                       cecl_ui/services/hybrid/                      │
   │                                                                     │
   │   registry: report_schemas (Postgres)                               │
   └─────────────────────────────────────────────────────────────────────┘
```

| System | Package | Stage | Registry | Validation currency |
|---|---|---|---|---|
| Hybrid **Report** Router | `cecl_ui/services/hybrid/` | report generation | `report_schemas` | schema matches typed contract + prior good |
| **Setup** Router | `cecl_ui/services/setup/` | client onboarding | `setup_patterns` / `setup_metrics` | balances tie to book; impaired match |
| **WARM ingestion** path | `setup/warm_parser.py` + `setup/config_assembler.py` | onboarding (WARM CUs) | (feeds setup registry) | reserve reconciles to the WARM |

---

## 1. Hybrid Report Router  (`cecl_ui/services/hybrid/`)  — *pre-existing*

A three-stage pipeline layered **in front of** the report engine. Feature-flagged **off** by
default (opt-in per CU via `hybrid_router.enabled` in the client YAML); degrades gracefully to
today's behavior whenever the router, Pydantic, the Anthropic SDK, an API key, or the network is
unavailable. **It never blocks report generation.**

1. **Local validation first** (`local_validator.py`) — validate a CU's report schema against a
   typed Pydantic contract (`schemas.py`) and the last known-good schema in the registry. Zero
   network, deterministic.
2. **API escalation** (`ai_escalation.py`) — only when local validation fails or a CU has no
   registered schema yet, escalate to `claude-opus-4.8` to propose a corrected/complete schema.
3. **Self-learning storage** (`schema_registry.py`) — persist newly validated schemas back to
   Postgres so future quarters skip escalation; expose cross-CU analytics.

Orchestrator: `router.HybridReportRouter`. Entry point wired at
`pipeline_service.run_reports_hybrid`.

**Design lesson that carried into everything else:** local-first + AI-only-on-miss + persist the
result = cheap, deterministic, and self-improving. The setup router is the same shape applied one
stage earlier.

---

## 2. Setup Router  (`cecl_ui/services/setup/`)  — the onboarding automation

Goal: turn the most labor-intensive, judgment-heavy, error-prone stage (client setup) into
**router proposes → ground truth validates → human approves**. Grounded in a full scan of 38
production configs (see `setup_router_design.md` for the evidence base and counts).

### Module map

| Module | Responsibility |
|---|---|
| `profiler.py` | **Metadata-only** structural profiler of delivery files. Detects header row, per-column dtype/range, and classifies each file into an archetype (loan_extract / participation / impaired / credit_pull / monthly_balances / report_other). Emits **no loan-level rows or PII** (only distinct code values when cardinality ≤ 60). |
| `pool_code_detector.py` | Given a loan df + the monthly book, ranks candidate pool-code columns by **tie-to-book** (per-pool sums vs. the monthly balance file) and returns a green/yellow/red decision. This is the "U-vs-V" resolver: pick the code column whose pooled sums match the book. |
| `pool_map_bootstrap.py` | Greenfield path — infer a `code → pool` map for a brand-new client with no prior config, using two-dimensional keyword inference (segment × collateral) over descriptor columns. Surfaces ambiguous codes for human review. |
| `setup_registry.py` | Postgres store of reusable onboarding **patterns** with a **trust tier** (`candidate`/`trusted`) and a `source` (`mined`/`manual`/`ai`/`warm`). Idempotent upsert keyed by checksum; support = distinct contributing-CU count; AI patterns promote to trusted after N≥2 validated+approved CUs. |
| `seed_from_configs.py` | Mines the 38 existing configs into starter patterns (member_account, pool_code_column, pool_map_code priors, score_source, credit_pull_keys, monthly_balance_source, participation_extract). |
| `probe_validator.py` | **The safety net.** Writes a temp config, runs a real `import_data.process_client` under a sentinel CU, and scores the result via `balance_check.compare_run` (pools tie) + `impaired_check_service.verify_for_run` (impaired match). This is how a proposal earns trust. |
| `warm_parser.py` | (see §3) Extracts an entire config from a WARM workbook. |
| `config_assembler.py` | (see §3) Assembles a runnable config + credit-pull file + registry seeds from a parsed WARM. |
| `router.py` | Orchestrator: `SetupRouter(engine).propose(...)` → a `SetupProposal`, combining registry priors + local detectors + probe validation. Proven to autonomously correct a broken Cottonwood config. |

### Governance / trust model (setup_registry)

- `source`: `mined` (reconstructed from a production config) · `manual` · `ai` · `warm`.
- `trust`: `trusted` | `candidate`. `mined`/`manual`/`warm` are born trusted (proven /
  human-authored / analyst-authored). `ai` is born `candidate` and promotes after N≥2 distinct
  validated-and-approved CUs.
- Every pattern carries `evidence` (support count + contributing CUs) as an audit trail.
- Patterns are **aggregated**: the same decision seen in K configs is one row with `support=K`.

---

## 3. WARM ingestion path  — *the major new capability*

**Discovery:** for the subset of clients delivered with a `CECL-Migration-WARM` workbook, the WARM
**is a filled-in configuration spec**. Rather than reverse-engineer config from raw extracts, we
read the authoritative mappings straight out of the WARM's well-known tabs. This is the single
highest-leverage automation built to date.

### 3a. `warm_parser.py` — read the WARM

`parse_warm(path)` returns a dict assembled from these tabs:

| WARM tab | Parser | Yields |
|---|---|---|
| `Grade Ranges & Loan Codes` | `parse_grade_loan_codes` | `pool_map` (code→pool, `**No pool**`→Ignore), `credit_grades` (label/min/max/reserve_rate), default mgmt-adj by grade, BRR legend, LTV baseline / probability factor / not-reported min |
| `BS Data` | `parse_bs_data` | monthly balances by pool (126 months) + ACL/ALLL history |
| `BS CO DQ Data Enter` | `parse_bs_co_dq` | per-pool risk-rated flag, CG/RR basis, ACL months/quarters, pool order, DQ balances **+ CU metadata** (charter #, name, core processor, economic-stress state/county/unemployment, snapshot date) |
| `ACL Env by Pool Mgmt Adj` | `parse_acl_reserve_rates` | per-pool per-grade **allowance factor** (col G = base + mgmt adj), Total blended rate, env factor, allow-before, `balance_only` flag, `max_base` (col E) |
| `Mmm-yy Data` (latest) | `parse_data_tab` | **column_mappings** derived by tracing the tab's *formulas* back to raw extract columns; `member_account`; score sources; `open_date` |
| `Mmm-YY Credit Pull` (all) | `parse_credit_pull` | member→FICO score maps, per pull, with dates |

**The formula-tracer is the clever bit.** The Data tab pastes the raw core extract on the left and
builds the final CECL fields with formulas. `parse_data_tab` reads the header row + first data row
(all refs stay on that row), then recursively follows each final column's formula until it hits a
cell with no formula (a raw pasted value) — that column's header is the source. Example (Tongass /
Symitar):

```
current_balance  ← LNDCBAL      (=IFERROR(R*1,0))
loan_pool_code   ← LNDALTC      (=TRIM(I) → VLOOKUP into Grade Ranges S:T)
original_fico    ← LNDCRSC
member_number    ← LNDACNO      (member = ROUNDDOWN(acct*0.001) → fixed_suffix, len 3)
current_fico     ← credit pull  (XLOOKUP into 'Dec-25 Credit Pull')
```

### 3b. `config_assembler.py` — turn the WARM into a runnable client

- `build_config_from_warm(...)` → a full config dict: credit_union/charter, member_account,
  column_mappings (incl. `open_date`), pool_map, credit_grades, monthly_balance + ACL,
  economic_data, mgmt_adj, pools, `not_risk_rated` (balance-only pools),
  `base_loss_rate_by_pool_grade` (pinned WARM allowance factors), `warm_allowance_pools`
  (analyst-mgmt-adj-driven pools), `credit_pull` (with `origination_aware: true` and
  `fallback_report_folder` → the WARM folder).
- `emit_credit_pull_file(...)` → writes the WARM's embedded pull scores to a standalone
  `Member Number` / `FICO` xlsx the pipeline consumes.
- `seed_from_warm(...)` → seeds high-trust (`source='warm'`) registry patterns (pool_map priors,
  member_account, score sources, etc.).

### 3c. Validation — proven end-to-end on Tongass FCU (June 2026)

1. **Tie-to-book:** the WARM-derived column_mappings + pool_map applied to the *real* AIRES
   quarterly extract reconciled the core loan book to within **~0.08%** ($154.70M vs $154.82M),
   with zero unmapped codes; the remaining gap fully explained (balance-only participations +
   one timing reclass).
2. **Reserve reconciliation:** the *generated TCT report* (built from the assembled config)
   reconciled to the WARM's total allowance within **+0.40%** — from a starting −26.5% — after a
   chain of fixes (see §5).

---

## 4. Origination-aware scoring  (`import_data.py`)

The analyst methodology for original vs. current FICO, now implemented and opt-in via
`credit_pull.origination_aware`:

- **ORIGINAL** = the AIRES score (`LNDCRSC`), *unless* a credit pull dated **before** the loan's
  origination covers the member → use that most-recent pre-origination pull score.
- **CURRENT** = the most-recent pull score, *unless* the member took out a loan **after** the
  most-recent pull → every one of that member's loans takes the **newest post-pull loan's AIRES
  origination score**.

Key functions: `_load_dated_credit_pulls` (reads all dated pull tabs from the WARM, newest-first),
`_origination_aware_scores` (applies the rules), hooked into `import_file` just before the DB
frame is built. Uses the real `LNDOPEN` origination date — which the manual WARM never did (its
Data tab stamped Open Date = snapshot date), so this is *more* accurate than the source.

Configurable levers: `current_chain_older_pulls` (default false = don't resurrect stale pulls);
the pre-origination pull choice; a future "pull within N months counts as current" cutoff.

Tongass June result: movement up 1,640 / down 1,848 / unchanged 1,028; reserve moved to **+1.00%**
of the WARM — the small, expected divergence reflecting the more-accurate current grades.

---

## 5. Report-engine fixes made along the way (generic, benefit all CUs)

The reserve reconciliation surfaced and fixed several real issues in the report engine + config:

- **`not_risk_rated`** must be the **balance-only** pools (no per-grade loan detail — overdraft /
  participations), detected from the ACL tab — *not* the WARM's "Risk Rated Yes/No" flag (which
  wrongly includes graded commercial pools).
- **`base_loss_rate_by_pool_grade`** pins the WARM's **full allowance factor (col G)** per grade —
  needed because loading the WARM zeroes the mgmt-adj carry-forward for graded pools.
- **`warm_allowance_pools`** uses the WARM's `allow_before` verbatim for analyst-mgmt-adj-driven
  commercial pools the firm-wide model can't reproduce; requires the WARM reachable by
  `load_impaired_data` via `credit_pull.fallback_report_folder`.
- **`report_tct.py`** NRR branch now honors a pinned `Total` blended rate for balance-only pools.
- **`pipeline_service.import_warm_as_baseline`** bug fixed (was writing to the module path instead
  of `CECL_WORKSPACE_ROOT`).
- **`import_data`** date parsing needs `date_pattern` present (has a robust month-name fallback);
  the credit pull only loads when `credit_pull.file_pattern` is set.

---

## 6. Data stores (Postgres)

| Table | Owner | Purpose |
|---|---|---|
| `setup_patterns` | setup router | reusable onboarding decisions: `decision_kind`, `fingerprint` (JSONB), `fragment` (JSONB), `evidence`, `source`, `trust`, `checksum` (unique), `support`, `approvals` |
| `setup_metrics` | setup router | per-decision telemetry: client, decision_kind, resolved_by, used_ai, validated, approved |
| `report_schemas` | report router | last known-good report schema per CU (self-learning) |
| `monthly_loan_data` | pipeline | per-loan snapshots (the join target for validation) |

---

## 7. Status — what is proven vs. pending

**Proven / validated:**
- Hybrid report router (pre-existing, feature-flagged).
- Setup router phases 0–2 (registry, seed from 38 configs, pool-code detector, probe validator,
  orchestrator) — corrected a broken Cottonwood config autonomously.
- WARM parser: pool_map, credit grades, monthly balances + ACL, column_mappings (formula-traced),
  member_account, credit pull, pool settings, CU metadata, ACL reserve rates — validated on
  **Tongass** and **Bridgeton** (two very different CUs).
- Config assembler → full runnable config; tie-to-book within 0.08%; generated report reconciles
  to the WARM within **+0.40%** (base) / **+1.00%** (with origination-aware scoring).
- Origination-aware scoring implemented + validated.

**Pending / next:**
- ~~Wire `warm_parser` + `config_assembler` into the UI setup flow~~ **DONE** — WARM upload now
  runs the validated path and auto-derives most of the config (see §9 + `warm_wizard_autoderive.md`).
- ~~Run the remaining WARM CUs~~ **DONE** — all 6 onboarded (§9); Utah at +2.36% (data-refresh
  punch list), the other five tie to ≤ ~1.2%.
- Have onboarding auto-call `import_warm_as_baseline` from the wizard (still manual in scripts).
- Live wizard **save → generated YAML → report-tie** for one CU (in progress).
- Bootstrap/from-scratch path + SCALE path generalization.
- Wire delimiter/split `member_account` detection into the WARM parse (shrinks the wizard's yellow
  member-account cases); surface `origination_aware` as a tie-tested toggle.

---

## 8. Key operating facts (for anyone picking this up)

- Running code = `C:\dev\CECL`. Workspace **data** = `Z:\Shared\TCT Files\CECL - CM Files` via
  env `CECL_WORKSPACE_ROOT`. DB URL from `.env` at the workspace root.
- WARM CUs split their delivery: quarterly loan/impaired files live under
  `<Client>\Client Access\Credit Migration\<YYYY-MM>\`; the WARM workbook lives in the **sibling**
  `<Client>\Credit Migration\` folder.
- The WARM workbook is the source of truth for a WARM CU; the tool's job is to **reproduce** it in
  the standard TCT/Vizo format and **validate** it ties out — not to re-derive its analyst
  judgments from scratch.

---

## 9. Session changelog — 2026-07-23 (WARM onboarding sweep + wizard auto-derive)

### 9a. All 6 WARM CUs onboarded + reconciled to their WARM

| CU | Core | Reserve vs WARM | Notes |
|---|---|---|---|
| Tongass FCU | Symitar | +0.40% (base) / +1.00% (orig-aware) | reference case |
| Honolulu Fire Dept FCU | DEXA/Aires | **−1.83%** | SSN-keyed pull + credit-card 2nd extract |
| Ontario Public Employees FCU | Symitar | **−0.02%** | date-less loan filename → `date_source: path` |
| Maple FCU | descriptive CSV | **+1.23%** | CUDL keep-first pool_map fix |
| Franklin Trust FCU | headerless positional | **−0.01%** | mgmt-adj double-count fix |
| Utah Community FCU | DFCU/BRR ($2.7B) | **+2.36%** | consumer pools tied; remaining = current-quarter impaired refresh + Exclude/LHFS methodology |

### 9b. Engine improvements (generic — surfaced by the onboarding sweep)

- **SSN-keyed credit pull** (`import_data.py`, `warm_parser.py`, resolvers): when the pull is keyed
  by SSN (not member #), loans join it on their SSN column via `credit_pull.loan_join_column`.
  Per-extract opt-out `credit_pull_join_column: ''` disables it for credit-card files (scored from
  a non-delivery source). Honolulu Personal Loans went from −$90k to −$2k.
- **Keep-first `pool_map`** (`warm_parser.parse_grade_loan_codes`): the WARM loan-code table can
  carry a later legacy/CUDL block that re-maps a code; keep the first (reporting) mapping. Fixed
  Maple's Used Vehicle (+$174k → +$4.7k).
- **Pool order = WARM order** (`config_assembler`): emit `pool_order` from the WARM so every report
  tab lists pools as the analyst's WARM does (was falling back to alphabetical on some tabs).
- **`date_source: path` in `reimport_period`** (`pipeline_service.py`): loan files with no date in
  the name (Ontario `AIRESLOANS.xlsx`, Utah `2026Q1_TCT-Loans_v2.csv`) resolve the snapshot from
  the dated folder; staging preserves the folder. Gated on `date_source=='path'` (no regression).
- **`not_risk_rated` filters `HIDE-*`/`Exclude`** placeholder pools (they read as balance-only).
- **`origination_aware` is CU-specific** (tie-decided): Tongass wants it (+0.40→+1.00, date-aware
  pulls); Ontario/Maple do NOT (carry-forward WARMs). Let tie-to-WARM decide per CU.
- **mgmt-adj double-count** lesson: when merging WARM report-fields onto an existing config, drop
  the old `mgmt_adj_by_pool` — it double-counts with the WARM-pinned col-G base loss (Franklin
  +10.7% → −0.01%).

### 9c. Setup-wizard cleanup — the WARM fast lane

Wired the validated resolver/assembler path into the wizard so a WARM upload auto-derives most of
the config, renders the 11 fully-derived steps green (review-only), and produces a **tying** config.
Full detail in **`docs/warm_wizard_autoderive.md`**. New module
`cecl_ui/services/setup/warm_autoderive.py`; changes in `routes/setup.py`,
`services/config_service.py`, `services/auto_setup.py`, `services/setup/config_assembler.py`,
`templates/setup/step2_warm.html`. Validated: all 8 WARM configs reproduce the tying reserve/mapping
fields via the wizard path; the one bug found (HIDE pools in `not_risk_rated`) is fixed.
