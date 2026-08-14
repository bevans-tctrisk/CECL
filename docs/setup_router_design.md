# Setup Router — Corpus-Grounded Design

**Status:** Draft for review · **Date:** 2026-07-22
**Author:** Generated from a full scan of the 38 production client configs + solved-case notes
**Scope:** A self-learning "Setup Router" that accelerates NEW client onboarding by proposing
configuration from raw delivery files, validating every proposal against ground truth, and
learning reusable patterns from each successful setup.

> Companion to the existing **Hybrid Report Router** (`cecl_ui/services/hybrid/`). That router
> validates report *schemas* at generation time. This router operates one stage earlier — at
> **client setup / config construction** — and is a separate concern with its own registry.

---

## 1. Why this, why now

Client setup is the most labor-intensive, judgment-heavy, error-prone stage of the pipeline
(evidenced by the Cottonwood onboarding: pool-code column, participation extract, member/suffix
layout, score sources, ACL row — each a discrete decision requiring investigation). The work is
highly **repetitive across clients** and, critically, **objectively verifiable** (balances must tie
to the monthly book; impaired loans must match the extract). That combination — recurring +
auto-checkable — is what makes setup automation both valuable and *safe*.

### Guiding principle

> **Router proposes → ground truth validates → human approves.**
> No proposal is ever trusted because the model is confident; it is trusted because the imported
> data *ties out*. Because configs drive regulatory CECL reserves, a validated result plus human
> sign-off is mandatory before any config goes live.

---

## 2. Evidence base (full scan of 38 production configs)

Every design claim below is grounded in these counts.

### Member/account layout is a bounded, learnable decision
| `member_account.mode` (top-level) | configs |
|---|---|
| fixed_suffix | 16 |
| delimiter | 10 |
| split | 6 |
| (none) | 6 |

**Perfect signal:** of the **10** configs using `split` anywhere, **10/10** map a `loan_suffix`
column; of the **18** `fixed_suffix`-only configs, **0/18** map a suffix column. A separate suffix
column in the extract ⇔ `split` mode. (This is precisely the Cottonwood miss.)

### Pool-code column has clustered spellings + a long tail
`Loan Type Code` (28), `LOAN TYPE` (12), `Loan Type` (8), `LOAN_TYPE` (4), `col_cd` (4),
`Acct Type` (4), `Loan Code` (3), `CURRENT_APPLICATION_TYPE` (2), `Loan Purpose Code` (2),
`CURRMIACCTTYPCD` (2)… → local fuzzy-cluster match handles the head; AI escalation handles the tail.
**Ground-truth tiebreak:** when multiple candidate code columns exist, pick the one whose per-pool
sums best match the monthly balance file (the U-vs-V resolution).

### Score sources cluster
Original FICO column: `Credit Score` (12), `CREDIT SCORE` (8), `FICO` (8), + tail
(`RISKSC`, `Borr_FICO`, `CB_SCORE`, `Orig Risk Score`…).
Credit-pull `member_column`: `Member Number` (15), `Account Number` (5), + tail;
`score_column`: `FICO` (16), `Credit Score`/`CREDIT SCORE` (9). `use_standalone_file`: 20 true / 5 false / 13 unset.
`credit_pull` present in **36/38** configs — near-universal scaffolding.

### Monthly-balance archetypes are few
`single` (22), (none) (9), `per_month` (4), `per_year` (2), `manual` (1). Four archetypes cover
everything. **ACL/ALLL configured in 27/38.**

### The participation / second-extract archetype recurs
**7/38** configs use a second extract with a static pool and/or row filter:
census, cottonwood, credit_union_of_richmond, honolulu_fire_department_fcu, mcdowell_cornerstone_cu,
shuford_fcu, wssc. Extract-count distribution: 0→12, 1→9, 2→11, 3→2, 4→2, 5→2.

### Date parsing genuinely varies
`YYYY-MM` (14), (none) (12), `MMDDYYYY` (4), `MMDDYY` (3), `YYYYMMDD` (2), `MMYYYY`/`YYYYMM`/`MMYY` (1 ea).
→ format detection from filename samples is a real, per-client task.

Other: BRR in 2, OAC in 6.

---

## 3. Architecture

Mirrors the existing hybrid package so it reuses proven patterns (`schema_registry` → registry,
`ai_escalation` → escalation, `router` → orchestration).

```
cecl_ui/services/setup/
  profiler.py         # read raw delivery files -> FileProfile (sheets, header row,
                      #   columns, dtype/value stats, distinct codes). METADATA ONLY, no rows.
  signals.py          # derive Signals from profiles (has_suffix_col, candidate_code_cols,
                      #   score_cols, participation_sheet, mb_archetype, date_samples...)
  decision_catalog.py # Signal -> candidate config fragments (member_account, mappings,
                      #   extract archetype, mb source, credit_pull, date_format)
  ai_escalation.py    # reuse hybrid pattern: metadata-only prompt, PII-stripped, lazy anthropic,
                      #   graceful degrade; proposes mappings for the long tail / gaps
  validator.py        # THE CLOSED LOOP: build candidate config -> reimport to a scratch snapshot
                      #   -> balance_check.compare_run (pools tie?) + impaired verify (match rate?)
                      #   + participation tie -> ValidationScore
  setup_registry.py   # Postgres: setup_patterns (signal fingerprint -> winning config fragment,
                      #   evidence, source manual|mined|ai, checksum, version), setup_metrics
  router.py           # SetupRouter(engine, api_key).propose(files) -> SetupProposal
                      #   (config_draft, per-decision confidence + evidence, validation report)
```

### Flow for a new client
1. **Profile** the delivered files (loan extract, monthly balances, participation, impaired,
   credit pull) → structural metadata only. No PII leaves the box.
2. **Signals → local match** against `setup_patterns` (fuzzy on column/sheet names + structural
   fingerprints). Produces high-confidence fragments for the common cases.
3. **AI escalation** only for unresolved/ambiguous decisions (metadata-only prompt) → candidate
   fragments with rationale.
4. **Validate against ground truth** — assemble the candidate config, import to a scratch
   snapshot, and run the existing checks:
   - `balance_check.compare_run` → do per-pool sums tie to the monthly book?
   - `impaired_check_service.verify_for_run` → impaired match rate?
   - participation pool ties to its book balance?
   Score candidates; keep the one that ties. (This is what makes it trustworthy — and it's the
   same evidence a human would demand.)
5. **Register** the winning `signal fingerprint → config fragment + validation evidence` so the
   next client with the same signature is resolved locally, no AI needed. Self-improving.
6. **Human review** in the wizard: proposed config + validation report + confidence per decision;
   analyst approves/edits before it is written.

---

## 4. Decision catalog (signal → decision → ground-truth check)

| Decision | Local signal (from corpus) | Ground-truth validation |
|---|---|---|
| `member_account` mode | separate suffix column present ⇒ `split` (10/10); combined member+suffix in one field with a delimiter ⇒ `delimiter` (10); fixed-width suffix ⇒ `fixed_suffix` (16) | account-key uniqueness + impaired match rate |
| pool-code column | fuzzy-match candidate columns to the known cluster; if >1 candidate, defer to check | per-pool sums vs monthly balance (U-vs-V) |
| original vs current FICO | on-file bureau col ⇒ original; credit-pull file ⇒ current | credit movement appears (non-diagonal) |
| credit_pull keys | `Member Number`/`Account Number` + `FICO`/`Credit Score` clusters | pull match count > threshold |
| participation extract | "Remittance"/"Participation" sheet, `Investor UPB`, `Loan Status` | pool ties to participation book; status filter tuned to tie |
| monthly-balance source | one wide file ⇒ `single`; per-month files ⇒ `per_month`; per-year ⇒ `per_year`; hand grid ⇒ `manual` | pools present for snapshot; totals sane |
| ACL/ALLL | "ALLL Balance" row in monthly file | value sane vs prior quarters |
| date_format | infer from filename token samples | snapshot resolves to a real month-end |

**Solved-case "gotchas" to encode** (from setup notes, so the router doesn't relearn them):
leading blank column A shifting header letters; header not on row 1 (`header_row`); status overlays
that reclassify loans out of on-book pools (Cottonwood code 59); charged-off rows carrying residual
UPB (status filter); transposed member-number typos in impaired files (flag, don't fix).

---

## 5. Registry schema (Postgres, mirrors `schema_registry`)

```
setup_patterns(
  id, decision_kind,            -- 'member_account' | 'pool_code' | 'participation' | ...
  signal_fingerprint JSONB,     -- normalized structural signals that triggered it
  config_fragment  JSONB,       -- the YAML fragment that won
  evidence         JSONB,       -- validation result that justified it (ties, match rate)
  source TEXT,                  -- 'mined' | 'manual' | 'ai'
  checksum TEXT, version INT, is_active BOOL, created_at
)
setup_metrics(client, decision_kind, resolved_by, used_ai, validated, ts)
```

Idempotent `save` by checksum; version bump on change; analytics read `monthly_loan_data`
directly (same as the report registry).

---

## 6. Seeding plan (solves cold-start using what we already have)

1. **Mine all 38 configs** → emit one `setup_patterns` row per decision per config, `source='mined'`,
   with the structural fingerprint reconstructed from each config's own mappings. (The scan in §2 is
   the first slice of this.)
2. **Fold in solved-case notes** → encode the gotchas in §4 as high-priority patterns.
3. Result: the registry starts with real coverage for the common head; AI is reserved for the tail.

---

## 7. Safety, privacy, guardrails

- **PII discipline** (reuse hybrid's `_SENSITIVE_KEYS` stripping): only structural metadata —
  column names, sheet names, dtype/value ranges, distinct code lists — ever reaches the AI stage.
  No member rows, names, SSNs.
- **Never auto-commit**: the router writes a *draft* + validation report; a human approves.
- **Validation-gated confidence**: a decision is "high confidence" only when the scratch import
  ties out — model certainty alone never promotes a decision.
- **Human owns judgment**: the router targets the mechanical ~80% (mappings, modes, archetypes);
  provisions allocation, ambiguous reclassifications, and data-quality calls stay with the analyst.
- **Degrade, never block** (as the report router does): missing SDK/key/DB → local-only proposal.

---

## 8. Phased rollout

- **Phase 0 — Corpus mine + registry seed** (foundation; low risk, immediate reuse value).
- **Phase 1 — Pool-code-column detector** (fully auto-validatable; proves propose→validate→register).
- **Phase 2 — Member/suffix + credit-pull mapping** (the 10/10 suffix signal; high hit rate).
- **Phase 3 — Participation + monthly-balance/ACL archetypes** (the 7/38 second-extract pattern).
- **Phase 4 — AI escalation for the long tail** + wizard "review proposed setup" UI.
- **Phase 5 — Self-learning loop live**: each approved setup writes back to `setup_patterns`.

---

## 9. Integration points

- `pipeline_service.py`: add `propose_setup(short_or_files)` → SetupProposal (parallels
  `run_reports_hybrid`); feature-flagged, off by default.
- Wizard (`cecl_ui/routes/`): a "Suggest configuration" step that renders the proposal + evidence
  for analyst approval before the config is saved.
- `cecl_credentials.get_anthropic_api_key()` reused for escalation.

---

## 10. Open questions for review

1. Scope of Phase 1's file inputs — loan extract + monthly balances only, or full delivery set?
2. Should the scratch-validation import run against the live DB (isolated snapshot key) or a
   throwaway schema? (Prefer isolated snapshot key for reuse of existing checks.)
3. Confidence threshold to auto-fill vs. always surface for review.
4. Registry governance — who curates `source='ai'` patterns before they become trusted?
