# 09 — Charge-off by migration status: recovering the scores from snapshot history

> **Revision 2 (2026-09-01).** A second source — the WARM workbooks' `Mmm-YY` tabs —
> was supplied by the product owner after the first pass and is now measured. It is
> **Path A**; what the first pass built and measured is **Path B**. Path A **supersedes
> §3.2 and §5** of this document: the "credit scores don't move" finding is correct about
> the loan extracts and wrong as a general claim. See **§A**, and the inline
> *superseded* markers. Every other section stands as written.

Supersedes the verdict of `07_chargeoff_feasibility.md`. Same scope: `co_by_status` /
`co_by_pool`, the "Charge off by Credit Grade Migration" bar on every `Risk Change` /
`Risk Chg <pool>` tab. The delinquency half shipped separately (`06_dq_derivation.md`).

Written after the product owner's correction to 07's framing:

> "Most credit unions, when we initially set them up, don't have the current credit score
> (also known as credit score at time of charge off), and many don't even have the original
> credit score. So we have to build those graphs over time by pulling the information on
> loans that charge off from previous loan data extracts and credit pulls."

Evidence base: the live CECL Postgres database (34 credit unions, 1.2M snapshot rows), every
`historical_file_formats.chargeoff` block in
`Z:\Shared\TCT Files\CECL - CM Files\client_configs\`, ~400 raw charge-off uploads under
`Raw_Uploads\`, `import_data.py`, `generate_report.py`, `dq_migration_split.py`, and the
82-tab legacy WARM
`Z:\Shared\Clients\Utah Community FCU\Portfolio Management (CM, ID, Warm)\2026-06 CECL-Migration-WARM - Utah Community FCU.xlsx`.

**No delivered report is changed by this pass.** Two new files were added
(`co_migration_split.py`, `scripts/verify_co_migration.py`); nothing is wired into
`generate_report.py`, `report_vizo.py` or `report_tct.py`, and no config on the share was
edited. Every number below came from running the new module read-only against live data.

---

## VERDICT

**The mechanism is real and now proven end to end. The data to feed it exists for every
credit union that has a WARM file, and does not yet exist for any credit union onboarded
through the wizard. The split is not between "enough time has passed" and "not enough" — it
is between two populations with entirely different source material.**

*(Revision 2. The first pass framed this as purely a question of elapsed time. That is true
only for the wizard population. See §A.)*

Three things changed relative to 07:

1. **07's population was wrong, and being wrong made the problem look 4× harder than it is.**
   07 read Utah's $28.95M `co_by_status` as "an ~11-year cumulative figure" going back to
   2015-10. It is not. It is a **trailing 36-month window**. Utah's `CO Data Entry` matrix
   reproduces **to the penny — total and all six live pools —** as "every charge-off on the
   `CO Data` tab dated **2023-07 or later**, excluding the `Exclude` and `HIDE-*` pools",
   against a 2026-06 report date (§1). We need three years of history, not eleven.

2. **Given the two scores, our existing code reproduces the WARM exactly.** Running
   `dq_migration_split.classify_migration` over the WARM's own `Org Grade` / `Curr Grade`
   columns on that population returns
   `1,116,964.01 / 15,151,657.35 / 8,689,857.16 / 3,995,038.50` — the WARM's four numbers,
   to the cent (§1). Banding the raw scores with our own `credit_grades` instead of the
   WARM's ladder is within 1.0pp on every bucket. The rule was never in doubt; this makes it
   measured rather than asserted.

3. **The lookback join works, and it is measurably bounded by how many months of loan
   history each credit union has.** Mountain CU's first snapshot is 2025-12-31 and its
   report date is 2026-06-30 — 7 of the 36 window months are covered, i.e. a ceiling of
   19.4%. Measured recovery: **163 of 840 charged-off loans, 19.4%** (§4). The model
   predicts the measurement to the decimal. Nothing is broken; there is simply not yet
   enough history.

**Where 07 was right:** the charge-off feed carries no scores *for most credit unions*, the
per-loan identity is discarded by `_parse_chargeoff_file`, and bucketing Mountain's `Exclude`
back-book as it stands would produce a ~99%-Unchanged chart that is confidently wrong. All
three hold.

**Where 07 was wrong:**

| 07 said | Actually |
|---|---|
| "$28.95M is an ~11-year cumulative figure" | Trailing **36 months** (2023-07 → 2026-06), exact to the cent (§1) |
| "The CO uploads our CUs send carry account, code, amount and date — and nothing else" | **8 of 28** charge-off feeds carry a credit-score column nothing reads: Utah (`APPLICATION_CREDITSCORE` + `CREDITSCORE_MOSTRECENT`), Curis (`FICO Score` + `FICO Date`), Jackson River / NOVA / WNC (`FICO Score at Loan Orig`), Shuford (`CreditScore`), SCI (`FICO`), United Community (`FICO`), Lanco (`CREDIT SCORE`) (§2.3) |
| "not derivable … not a plumbing gap that a join would close" | The join closes it. Jackson River CU produces a real four-bucket split **today** — I 3% / D 37% / U 39% / NR 21% — off one such column plus the snapshot lookback (§4) |
| "9 of the 10 affected CUs have zero charged-off rows in loan-level data" | True of the *current* snapshot only. Across **all** snapshots, Mountain matches 871 of 923 charged-off loans (94.4%), Curis 119 of 220, WSSC 213 of 417 (§4) |
| Blocker B: "even inside the window, the two scores are the same number" | True **of the loan extracts**, and estate-wide — most CUs apply **one** dated credit pull, so `current_fico_score` is constant across every snapshot (§3.2). But a real dated score time series does exist for every credit union with a WARM, in tabs nothing reads (§A) |
| *(this document, revision 1)* "no time-varying score source exists" | Wrong. The WARM `Mmm-YY Credit Pull` tabs move **93.4–99.5%** by **47–85 points**; the `Mmm-YY Data` tabs move **38.7–92.3%** at loan level. Recovering from them alone reproduces the analyst's own four-bucket split to within a few points on 3 of 4 credit unions tested (§A.3, §A.9) |

## RECOMMENDATION

**Ship the derivation with its refusal guards armed, and ship the empty state.** Today that
populates 1 of the 19 workbooks and correctly declines the other 18. It populates more every
quarter with no further code change, because the history it reads accrues on every import.

Concretely, in priority order:

| # | Action | Effort | Effect |
|---|---|---|---|
| 1 | **Tier-1 empty state** (§6) — replace the blank titled frame with a "not yet available" note. Shared with the DQ fix from `04_blank_charts.md §6`. | 0.5 day | Fixes the *client-facing* defect on all 18 remaining workbooks now |
| 2 | **Wire `co_migration_split.fill_missing_co_migration`** into `generate_report` next to the DQ call (§7). Guarded; refuses rather than guesses. | 0.5 day | Jackson River populates; everyone else keeps the empty state |
| 3 | **Map the score columns already in the charge-off files** — a wizard field plus 8 config edits (§2.3, §7.3) | 1–2 days | Directly recovers one of the two scores for 8 CUs |
| 4 | **Fix the charge-off / loan-extract account-key mismatches** (§4.2). Six CUs join at 0–28% because the two `member_account` blocks disagree. | 1–2 days, analyst | Lifts realised coverage toward the ceiling for Shuford, Richmond, Lanco, McDowell, WNC, Emergency Responders |
| 5 | **Ask for a second dated credit pull** at renewal (§5.3) | client lead time | The only thing that makes the chart *informative* rather than merely populated |

Do **not** persist a new table (§7.2). Do **not** relax the guards to fill charts sooner
(§3.4).

---

## 1. What the chart's population actually is — and the exact-match proof

07 §4 read the `CO Data` tab (16,324 rows, $59.41M, dated from 2015-10) and concluded the
$28.95M matrix was an 11-year cumulative. That inference is wrong: `CO Data Entry`'s
`B7:L17` cells are a **hardcoded paste**, and the analyst pastes a filtered subset.

Searching every month cut for one that reproduces $28,953,517.02:

```
CLOSE 2023-07  28,953,517.02
best: ('2023-07', 28953517.020000003)   target 28953517.02
```

and per pool, against 07 §4's table:

| Pool | this cut | 07's WARM reading | diff |
|---|---:|---:|---:|
| Commercial | 203,002.36 | 203,002.36 | 0.00 |
| Direct Auto | 4,563,949.35 | 4,563,949.35 | 0.00 |
| Indirect Auto | 14,648,072.37 | 14,648,072.37 | 0.00 |
| Mortgage Real Estate | 20,772.35 | 20,772.35 | 0.00 |
| Other Consumer | 13,173.23 | 13,173.23 | 0.00 |
| Unsecured | 9,504,547.36 | 9,504,547.36 | 0.00 |
| **Total** | **28,953,517.02** | **28,953,517.02** | **0.00** |

2023-07 → 2026-06 is **exactly 36 months** ending on the report date. The `Exclude` and
`HIDE-*` pools are dropped. That is the population rule, and it is the one
`co_migration_split.DEFAULT_LOOKBACK_YEARS = 3` implements.

### The classification rule, measured rather than asserted

With the population pinned, run `dq_migration_split.classify_migration` (the same function
the DQ pie uses, itself a transcription of `report_vizo._sheet_risk_change`) over the WARM's
own grade columns:

```
-- our classifier over the WARM's own Org Grade / Curr Grade labels
   Improved         1,116,964.01   3.86%
   Deteriorated    15,151,657.35  52.33%
   Unchanged        8,689,857.16  30.01%
   Not Reported     3,995,038.50  13.80%
   TOTAL           28,953,517.02

-- WARM CO Data Entry (authoritative)
   Improved         1,116,964.01   3.86%
   Deteriorated    15,151,657.35  52.33%
   Unchanged        8,689,857.16  30.01%
   Not Reported     3,995,038.50  13.80%
```

Identical to the cent. And with our own banding of the WARM's raw `Org Score` /
`Curr Scr Hard` instead of its grade ladder:

```
   Improved         1,093,817.48   3.78%   (−0.08pp)
   Deteriorated    14,991,256.43  51.78%   (−0.55pp)
   Unchanged        8,583,756.01  29.65%   (−0.36pp)
   Not Reported     4,284,687.10  14.80%   (+1.00pp)
```

The residual is a grade-boundary difference between Utah's `Grade Ranges & Loan Codes` tab
and `utah_community_fcu.yaml`'s `credit_grades`, not a logic difference.

**Conclusion: the only missing ingredient is the pair of scores.** Everything downstream of
them is already correct and now verified against the one authoritative source we have.

## 2. The raw material

### 2.1 Loan-snapshot history — `monthly_loan_data`

Schema is unchanged from 07 §2: `credit_union, snapshot_date, member_number, loan_pool,
current_balance, current_fico_score, original_fico_score, business_risk_rating`. No primary
key, no index. Every import appends a full month.

Depth per credit union (live, 2026-08):

| Credit union | snapshots | range | rows |
|---|---:|---|---:|
| Test Nova CU | 79 | 2018-01 → 2025-12 | 201,750 |
| WSSC FCU | 29 | 2024-01 → 2026-06 | 51,116 |
| Credit Union of Richmond | 28 | 2024-01 → 2026-06 | 59,768 |
| Curis CU-Palmetto Health CU | 28 | 2024-01 → 2026-04 | 92,318 |
| Lanco FCU | 28 | 2024-01 → 2026-06 | 55,265 |
| McDowell Cornerstone CU | 17 | 2025-01 → 2026-06 | 17,853 |
| Destinations CU / First Area / Mountain / Shuford | 6 | 2025-12 → 2026-06 | — |
| Central Susquehanna / Emergency Responders / NOVA / Nucor / SCI | 5 | 2025-12 → 2026-06 | — |
| Census / Central Keystone / Honolulu / Jackson River / TCP / WNC | 4 | 2025-12/2026-02 → 2026-06 | — |
| Bridgeton Onized / Franklin Trust / Tongass / United Community / Utah | 3 | 2025-12 → 2026-06 | — |
| Erie / Maple / Ontario | 2 | — | — |
| Bridgeton Onized / Census (Scratch) / Sample / Test Nucor | 1 | 2025-12 | — |

Six credit unions carry 17–79 monthly snapshots — the wizard's Historical step back-loads
them. **Every snapshot retains per-loan credit scores**; `current_fico_score` and
`original_fico_score` are populated on exactly the same rows in every case (07 already noted
the cause: `import_data.py:1715` gap-fills `original = current` when original is 0).

Retention is 84 months (`cecl_retention.py`), comfortably beyond the 36-month window.

### 2.2 The loan identifier — `member_number`

`import_data.py:1654` sets `member_number` to `raw_account`, the full account string from
`derive_member_account(df, config, has_header)` — one of three modes (`fixed_suffix`,
`delimiter`, `split`) driven by the config's top-level `member_account` block. It identifies
a **loan**, not a member, and it is stable across snapshots: for most credit unions
`COUNT(*) == COUNT(DISTINCT member_number)` within a snapshot. The exceptions are
double-imported months (Shuford 2025-12: 2,600 rows / 1,372 accounts; Census 2026-03: 1,024
/ 512; McDowell 2026-02…04) — `load_score_history` collapses those by keeping the
highest-scored row per (account, snapshot).

The charge-off files carry the same identifier. `historical_file_formats.chargeoff` has its
**own** `member_account` block plus `account_col` / `member_col`, in 28 of 40 configs. So
the join key exists on both sides by construction — but the two blocks are configured
independently and do not always agree (§4.2). That is the single largest recoverable loss in
this whole assessment.

### 2.3 Score columns already sitting in the charge-off feeds

07 §4 states "The CO uploads our CUs send carry account, code, amount and date — and nothing
else." Scanning the header row of every charge-off file under `Raw_Uploads\` says otherwise:

| Credit union | column | index | what it is |
|---|---|---:|---|
| Utah Community FCU | `Application Creditscore` + `Creditscore Mostrecent` | 2, 3 | **both** scores |
| Curis CU-Palmetto | `FICO Score` (+ `FICO Date`) | 6, 7 | **origination** — `FICO Date` sits 1–15 years before the charge-off date |
| Jackson River Community CU | `FICO Score at Loan Orig` | 10 | origination |
| WNC Community CU | `FICO Score at Loan Orig` | 11, 12 | origination |
| NOVA CU | `FICO Score at Loan Orig` | 7 | origination |
| United Community FCU | `FICO` | 7 | unqualified |
| Shuford FCU | `CreditScore` | 2 | at write-off (sits beside `ChargeOffDateS`) |
| SCI FCU | `FICO` | 4 | unqualified |
| Lanco FCU | `CREDIT SCORE` | 31 | unqualified |

`member_col` is not the only dead config: **the score columns are not even collected.** For
Utah this is decisive — its `TCT-Chargeoff_v2.xlsx` is the machine-readable equivalent of
the WARM's hand-built `CO Data` tab, carrying both scores per loan, and nothing reads
either.

`co_migration_split.detect_score_columns()` implements this detection. It is **advisory by
default** (`chargeoff.auto_detect_score_columns`, off) because guessing wrong replaces a
blank chart with a wrong one — as the scan itself demonstrates: an unqualified `FICO Score`
is *origination* at Curis and *at-write-off* at Shuford, and only the neighbouring date
column tells them apart.

### 2.4 Dated credit pulls (legacy CUs only)

`import_data._load_dated_credit_pulls` reads every `Credit Pull <date>` tab from a CU's WARM:

```
franklin                       3 pulls: 2025-12-31, 2024-12-31, 2022-12-31
honolulu_fire_department_fcu   3 pulls: 2026-03-31, 2025-06-30, 2024-12-31
bridgeton                      4 pulls: 2025-12-31, 2024-03-31, 2023-06-30, 2022-12-31
maple                          6 pulls: 2025-09-30 … 2022-12-31
tongass_fcu                    4 pulls: 2025-12-31, 2024-09-30, 2023-06-30, 2022-12-31
utah_community_fcu             1 pull:  2025-09-30
erie_fcu                       1 pull:  2022-03-31
bridgeton_onized_fcu / ontario / cottonwood_fcu   0
```

These are **member-level**, dated, real bureau scores — the only genuinely time-varying score
source in the estate. Wizard-onboarded credit unions have exactly one
(`credit_pull.uploaded_filename` + `pull_as_of_date`). §3.2 explains why that single pull is
the binding constraint.

## 3. The lookback join — specification

Implemented in `co_migration_split.py`. Three tiers, best source first; the two scores are
sourced **independently**, so a feed that supplies one of them still helps.

### 3.1 The rule

Given a charge-off row `(account A, charge-off date D, amount, loan code → pool)`:

**Population.** Keep the row when `pool` maps to a report pool and
`D ∈ (asof − 3 years, asof]`. Rows with no parseable date fall back to the filename period
(the same fallback `generate_report.load_chargeoff_recovery_history` already uses); rows
still undated are kept, since dropping them would break the dollar tie-out.

**Score at charge-off.** `current_fico_score` from the **latest snapshot with
`snapshot_date <= D`** in which `A` appears with a non-zero score.
*Tie-break:* a snapshot after `D` is **never** used. Cores that leave the written-off loan in
the extract (Mountain's `chargeoff_exclude_column` → `Exclude` pool) still carry it months
later, but its scores are frozen and mutually equal; pairing them manufactures a spurious
`Unchanged`. This is exactly the trap 07 §3 measured, and refusing the post-event snapshot is
what turns Mountain's "84% Unchanged" into an honest "80% Not Reported".

**Original score.** In order:

1. `orig_score_col` on the charge-off file, when configured (§2.3);
2. `original_fico_score` from the recovered row, **when it differs from that row's current
   score** — a difference proves the CU (or its dated pull) really supplied an origination
   score for this loan rather than `import_data`'s gap-fill. Provenance `row_pair`;
3. the loan's **earliest observed** `current_fico_score`, when it differs from the score at
   charge-off. Provenance `lookback_span`. This is Brian's "recover it from the earliest
   snapshot"; §3.2 explains why it almost never fires;
4. otherwise `original = current`. Provenance `gap_fill` — an assumption, counted separately.

**Bucketing.** `cecl_engine.assign_credit_grade` on each score, then
`dq_migration_split.classify_migration(cg, og, grade_index, n_top, no_score)` — the same call
the DQ pie makes, with BRR label sets for `brr: true` pools. Proven exact in §1.

**Join key.** Literal `member_number` first. On a miss, a normalised key
(`re.sub(r'\D','',x).lstrip('0')`) is tried, but **only where that normalised form maps to
exactly one account in the credit union** — `build_alias_index` drops collisions, so the
fallback can never merge two real loans.

### 3.2 Why the "earliest snapshot" idea does not fire — the estate-wide measurement

> **SUPERSEDED IN PART by §A.3.** Everything measured below is correct and reproducible:
> the *loan extract* carries a static score, so `monthly_loan_data` is not a time series.
> The conclusion drawn from it — that no time-varying score source exists — is wrong. One
> does, in the WARM's `Mmm-YY` tabs, where 93–99% of members move by 47–85 points. Read
> this section as "why Path B alone cannot work", not as "why nothing can work".

Brian's model assumes each snapshot is a fresh observation of the loan's score. Measured
across all 34 credit unions — accounts appearing in ≥2 snapshots, and how many ever show a
*different* `current_fico_score`:

| Credit union | accts in ≥2 snaps | ever change | % | max snaps |
|---|---:|---:|---:|---:|
| Erie FCU | 22,065 | 21,275 | **96.4%** | 2 |
| Credit Union of Richmond | 3,108 | 1,253 | **40.3%** | 25 |
| Curis CU-Palmetto | 5,429 | 1,000 | 18.4% | 28 |
| WSSC FCU | 2,494 | 391 | 15.7% | 27 |
| McDowell Cornerstone CU | 759 | 69 | 9.1% | 17 |
| Lanco FCU | 4,921 | 237 | 4.8% | 16 |
| Test Nova CU | 8,128 | 353 | 4.3% | **79** |
| United Community FCU | 2,362 | 90 | 3.8% | 2 |
| Tongass FCU | 4,124 | 123 | 3.0% | 2 |
| Destinations CU | 1,336 | 37 | 2.8% | 6 |
| Franklin Trust FCU | 7,565 | 69 | 0.9% | 3 |
| Cottonwood / Bridgeton Onized | ~4,000 | ~25 | 0.7% | 3 |
| **Utah Community FCU** | **92,014** | **1** | **0.00%** | 3 |
| **Mountain CU** | **14,021** | **0** | **0.0%** | 6 |
| Nucor / SCI / Shuford / Census / Honolulu / Jackson River / NOVA / TCP / Central Keystone / Central Susquehanna / First Area / Emergency Responders / Ontario | all | **0** | **0.0%** | 3–6 |

**Test Nova CU has 79 monthly snapshots spanning eight years and 95.7% of its loans never
move a single point.** That is the whole story in one line. The score in a monthly loan
extract is the score stored on the loan record — set at origination and refreshed only when
the credit union runs a bureau pull. Wizard-onboarded CUs supply **one** pull, applied at
every import, so the per-loan score is a constant. Stacking more snapshots of a constant
yields no more information.

This does **not** invalidate the lookback. It relocates its value:

* the lookback still recovers the pair the loan carried **on its last snapshot before it
  left the portfolio** — which the report can no longer see once the loan is gone. That is
  the `row_pair` provenance, and where the CU's single credit pull covered the member it is
  a genuine origination-vs-current pair;
* it still supplies the *missing half* when the charge-off file carries the other half —
  which is where the two credit unions that work today get their numbers.

### 3.3 Failure modes and what the chart does

| Situation | Provenance | Bucket | Rationale |
|---|---|---|---|
| Loan in no snapshot at all | `no_history` | **Not Reported** | Score genuinely unknown. `Not Reported` is the WARM's own bucket for "no credit score" (`Q9 = SUM(B17:L17, L7:L16)` — the Not-Reported row and column), so this is semantically correct, not a dumping ground |
| Loan appears, but every snapshot post-dates the charge-off | `no_prior` | **Not Reported** | The dominant case for young credit unions. This is the accrual clock |
| Loan appears but is never scored | `unscored` | **Not Reported** | |
| Only one of the two scores recovered | `partial` | **Not Reported** | The WARM folds a missing-either-side loan into Not Reported |
| Both recovered, `original` genuinely supplied | `row_pair`, `co_file*` | classified | |
| Both recovered, movement across snapshots | `lookback_span` | classified | |
| Both recovered, but `original = current` by gap-fill | `gap_fill` | **Unchanged**, counted separately | The documented WARM convention (`import_data.py:1715`), the same one the Risk Change matrix and the DQ pie apply. Defensible per loan; not defensible as the *basis* of a chart — see §3.4 |
| Charge-off date between snapshots | — | uses the last snapshot before it | The loan was open and scored then; that is the intended reading of "score at charge-off" |
| Charge-off date unparseable | — | filename period, else kept undated | Same fallback as the existing by-year charge-off chart |
| Account key collides after normalisation | — | not joined → `no_history` | `build_alias_index` refuses ambiguous keys |
| Loan code maps to `Exclude` / no pool | — | **dropped from the chart entirely** | Matches the WARM: `Exclude` and `HIDE-*` are outside the population (§1) |

The distinction that matters for the brief's question — *`Not Reported` versus excluded
entirely* — is: **a loan is excluded only when it is not part of the reserve population
(unmapped or `Exclude` pool). Everything else that fails stays in the chart as
`Not Reported`, so the four bars always sum to the credit union's actual 36-month charge-off
dollars.** That keeps the chart tied to the charge-off totals printed elsewhere in the
workbook.

### 3.4 Two refusal guards

Replacing a blank chart with a wrong one is worse than the defect. Both guards were driven
by real measurements:

1. `MAX_NOT_REPORTED_SHARE = 0.75` — refuse when Not Reported swallows more than three
   quarters of the dollars. A chart that is 80% "we don't know" is not a chart.
2. `MIN_MEASURED_SHARE = 0.25` — refuse when less than a quarter of the dollars carries a
   **genuinely measured** pair (`co_file`, `row_pair`, `lookback_span`, or a mixed pair).
   `gap_fill` does not count.

Guard 2 is the one that matters and it is the direct answer to 07's Blocker B. Without it,
Curis CU would publish `I 1% / D 0% / U 51% / NR 48%` where the 51% Unchanged is 112 loans
whose "original" is a copy of their "current". Doc 07 called that outcome "confidently,
specifically wrong" and it was right to.

## 4. Coverage, measured

`scripts/verify_co_migration.py` runs the real loader and the real join against live data.
Score columns from §2.3 are wired **in memory only** by its `--wire-scores` flag;
**no config on the share was modified**.
`asof` = each CU's latest snapshot; window = 36 months; rows deduped on
`(account, date, code, amount)`.

### 4.1 The estate

| Credit union | snaps | hist months | CO loans in 36m window | $ | both scores | measured $ | Not Reported $ | strict verdict | permissive I/D/U/NR |
|---|---:|---:|---:|---:|---:|---:|---:|---|---|
| **Jackson River Community CU** | 4 | 7 | 14 | 120,264 | **8** | **79.1%** | 20.9% | **ok** | **3 / 37 / 39 / 21** |
| Curis CU-Palmetto Health CU | 28 | 28 | 166 | 846,322 | 114 | 59.3% | 33.7% | ok | 3 / 0 / 63 / 34 |
| Mountain CU | 6 | 7 | 840 | 3,980,557 | 163 | 1.7% | 80.2% | refused (coverage) | 1 / 0 / 19 / 80 |
| WSSC FCU | 29 | 30 | 282 | 851,002 | 86 | 7.0% | 33.2% | refused (assumed) | 2 / 0 / 64 / 33 |
| Lanco FCU | 28 | 30 | 134 | 880,792 | 37 | 1.5% | 58.5% | refused (assumed) | 1 / 0 / 40 / 59 |
| Emergency Responders CU | 5 | 7 | 36 | 303,764 | 18 | 9.5% | 18.2% | refused (assumed) | 0 / 0 / 81 / 18 |
| Nucor Emp CU | 5 | 7 | 16 | 84,179 | 6 | 0.0% | 75.5% | refused (coverage) | 0 / 0 / 25 / 75 |
| Shuford FCU | 6 | 7 | 39 | 93,239 | 0 | 0.0% | 100.0% | refused (coverage) | 0 / 0 / 0 / 100 |
| Credit Union of Richmond | 28 | 30 | 34 | 79,967 | 0 | 0.0% | 100.0% | refused (coverage) | 0 / 0 / 0 / 100 |
| McDowell Cornerstone CU | 17 | 18 | 20 | 90,460 | 1 | 0.5% | 99.5% | refused (coverage) | 0 / 0 / 1 / 99 |
| Census FCU | 4 | 6 | 9 | 31,732 | 5 | 0.0% | 37.1% | refused (assumed) | 0 / 0 / 63 / 37 |
| SCI FCU | 5 | 7 | 5 | 3,611 | 1 | 0.0% | 18.6% | refused (assumed) | 0 / 0 / 81 / 19 |
| First Area FCU | 6 | 7 | 3 | 2,818 | 0 | 0.0% | 100.0% | refused (coverage) | 0 / 0 / 0 / 100 |
| WNC Community CU | 4 | 7 | 0 | 0 | — | — | — | no rows in window | — |
| Bridgeton Onized FCU | 1 | 1 | 0 | 0 | — | — | — | no rows in window | — |
| Central Keystone / Honolulu / Destinations / Franklin / TCP / United Community / Central Susquehanna / NOVA | 3–5 | — | **0 files found** | — | — | — | — | see §4.3 | — |

*"permissive" = both guards disabled, i.e. what would be published if the `gap_fill` →
`Unchanged` convention were accepted. Shown so the trade-off is visible, not because it is
recommended.*

Reproduced end to end by
`python scripts/verify_co_migration.py --wire-scores jackson_river_community_cu curis_palmetto_health_cu`:

```
### jackson_river_community_cu      --wire-scores {'orig_score_col': 10}
  scores in file : orig 12  curr 0
  RECOVERY       : BOTH 8/15 (53.3%), $95,153.89
  provenance     : {'gap_fill': 7, 'no_history': 7, 'row_pair': 1}
  diag status    : ok    (79.1% scored, 79.1% measured, 3-year window to 2026-06-30)
      Improved             3,088.36    2.57%
      Deteriorated        44,630.30   37.11%
      Unchanged           47,435.23   39.44%
      Not Reported        25,110.56   20.88%
      TOTAL              120,264.45
  co_by_pool: Used Indirect Auto  I 0 / D 44,630 / U 43,103 / NR 23,140
              Used Vehicle        I 3,088 / D 0 / U 0 / NR 606
              Signature           I 0 / D 0 / U 4,332 / NR 738
              VISA                I 0 / D 0 / U 0 / NR 627
```

**Jackson River is what a working chart looks like.** Its charge-off file carries
`FICO Score at Loan Orig`; the snapshot supplies the score the loan carried when it left the
portfolio; the two differ materially, and the result is a genuine 37% Deteriorated. Eight
loans is a small sample and §7.4 proposes a minimum-loan guard, but the mechanism is
demonstrably sound.

**Curis is the cautionary case.** 28 snapshots, an origination FICO on the charge-off file,
100 loans with both scores — and **83 of those 100 have `file_original == snapshot_score`**.
Its extract score *is* the origination score, so the "second" observation is a copy of the
first. Hence 63% Unchanged and `Deteriorated = $972.74`.

### 4.2 Where coverage is lost, in order of size

For Mountain — the largest affected book — of 840 in-window charge-offs:

```
no_prior      606   the charge-off predates every snapshot        ← the accrual clock
gap_fill      145   joined and scored, but original = current     ← the credit-pull problem
no_history     40   never in any snapshot
unscored       31   in a snapshot, no score
row_pair       18   genuine pair recovered
```

Across the estate, the four losses rank:

1. **`no_prior` — not enough elapsed history.** Purely a function of time. Mountain's
   ceiling is `7 months of history / 36 window months = 19.4%`, and the measured recovery is
   163/840 = **19.4%**. The model predicts the measurement exactly.
2. **`gap_fill` — one credit pull.** §3.2. Time does not fix this; a second dated pull does.
3. **`no_history` — the account-key mismatch.** The charge-off `member_account` block and the
   loan-extract `member_account` block are configured independently and disagree:

   | CU | CO key | DB key | literal join | +normalised |
   |---|---|---|---:|---:|
   | Mountain CU | `1054900001` | `909590001` | 871/923 (94%) | +0 |
   | Curis | `1114280102` | `339200004` | 119/220 (54%) | +0 |
   | Credit Union of Richmond | `181915000` | `18310900260` | 76/290 (26%) | +0 |
   | Emergency Responders | `663101` | `00000193102` | 0/36 | **+18 (50%)** |
   | Lanco FCU | `29498` | `1218602` | 32/628 (5%) | +16 |
   | Shuford FCU | `000060196001`, `00016270001`, `101938001` — **three widths in one CU** | `00001525106` | 0/85 | +0 |
   | WNC Community CU | `11278` | `140300020` | 0/9 | +0 |

   Emergency Responders is pure zero-padding and the normalised fallback fixes half of it.
   Shuford's charge-off parse yields three different member widths across its own files, so
   even normalisation misses; Richmond and Lanco are member-only vs member+suffix. **These
   are config defects, not data gaps**, and fixing them is the cheapest coverage available.
4. **`unscored`.** WSSC loses 119 of 282 this way — its extract scores only ~60% of loans.

### 4.3 Eight credit unions where no charge-off file is found at all

Central Keystone, Honolulu, Destinations, Franklin, TCP, United Community, Central
Susquehanna and NOVA return **zero** charge-off rows through this path, because
`load_chargeoff_recovery_history`'s filename gate requires `'charge' and 'off'` (or
`'recov'`) in the name and their files are called `CO & Recs by Loan Type Code 04302026.xlsx`,
`April 2026 Recoveries-Chg-Offs …`, and so on.

Those CUs nevertheless have 29–90 months of charge-off history in
`loan_code_chargeoff_history` (07 §1) because the **wizard** uses a different, wider
discovery path — `cecl_ui/routes/setup._discover_co_recov_files` (recursive glob over
`_MONTHLY_CO_RECOV_GLOBS`, per-layout header signatures) feeding
`monthly_co_recov_aggregator`. The report engine and the wizard disagree about which files
are charge-off files. That is a pre-existing divergence, out of scope here, but it caps this
derivation for 8 credit unions and is worth its own pass — the aggregator's per-file
`account`/`member` column resolution is exactly what a per-loan charge-off reader wants.

## 5. The honest ceiling

> **SUPERSEDED for the 4 credit unions that have a WARM (Franklin, Honolulu, Bridgeton
> Onized, Utah) — see §A.7.** For those the chart does not need to wait for anything; the
> scores are already on the share. The forecast below stands unchanged for the 8 wizard
> credit unions, which have no WARM and therefore no Path A.

### 5.1 The 19 workbooks

`04_blank_charts.md` lists 19 client-facing Vizo Migration workbooks with all-zero CO bars,
across 12 credit unions.

| Today | Count | Which |
|---|---:|---|
| **Populated by this change now** | **1** | Jackson River Community CU — *not in the original 19*; the 19 are the workbooks already delivered, and Jackson River's next one would populate |
| Populated if the guards were relaxed to accept `gap_fill` | 8 | Mountain ×2, Nucor ×2, SCI, Census, Curis ×2 — **not recommended**, all would be ≥60% Unchanged-by-assumption |
| Refused today, will populate as history accrues | 9 | Mountain ×2, Nucor ×2, SCI, Census, Emergency Responders ×2, Shuford |
| Blocked on a config fix first | 5 | Shuford, WNC, Destinations, Central Keystone ×2 |
| Blocked on missing charge-off file discovery (§4.3) | 5 | Honolulu, Bridgeton Onized, Franklin, Destinations, Central Keystone |
| Utah 2025-12 | 1 | Has a WARM; its blank quarter is the `load_impaired_data` early-return bug (07 §5, `04 §5`), not a data gap. **Fixing that early return populates this one workbook immediately.** |

**Stated plainly: of the 19 already-delivered workbooks, this change populates 0 today and 1
via the early-return fix. What it does is make every future quarter better, automatically.**

### 5.2 The accrual clock

The chart's window is 36 months. A charge-off is resolvable only if a snapshot precedes it.
So a credit union's ceiling is `min(36, months of loan history) / 36`, and it reaches 100%
exactly 36 months after its first snapshot.

| First snapshot | Credit unions | ceiling at 2026-06 | 50% at | 100% at |
|---|---|---:|---|---|
| 2024-01 | Curis, WSSC, Richmond, Lanco | 78–83% | passed | 2026-12 (**2 quarters**) |
| 2025-01 | McDowell Cornerstone | 50% | passed | 2027-12 |
| 2025-06 | Honolulu Fire Department | 36% | 2026-11 | 2028-05 |
| 2025-12 | Mountain, Shuford, Destinations, Nucor, SCI, Emergency Responders, NOVA, TCP, WNC, Jackson River, First Area, Central Susquehanna, Franklin, Tongass, United Community, Utah | **19%** | 2027-05 | **2028-11 (10 quarters)** |
| 2026-01 | Census | 17% | 2027-06 | 2028-12 |
| 2026-02 | Central Keystone | 14% | 2027-07 | 2029-01 |

The 2024-01 cohort is two quarters from a full window. The 2025-12 cohort — which is most of
the affected list — is **ten quarters** out, i.e. late 2028, and will be partially useful
from about 2027-05.

Realised coverage sits **below** these ceilings by the §4.2 losses, so the practical read is:
the 2024-01 cohort becomes viable in 2027 once its key mismatches are fixed; the 2025-12
cohort is a 2028 story unless §5.3 changes it.

### 5.3 What would actually move this faster

Accruing snapshots fixes `no_prior`. It does **not** fix `gap_fill`, which is the larger
long-run constraint (§3.2) — and a credit union with a full 36-month window but a single
credit pull will produce a 100%-Unchanged chart, which the guards will correctly refuse.

Two acquisitions change the picture immediately and are worth asking for at renewal:

1. **A second dated credit pull.** One extra pull, a year apart, gives every covered member a
   genuinely different current score. The machinery already exists
   (`_load_dated_credit_pulls`, `_origination_aware_scores`) and five legacy CUs already
   have 3–6 dated pulls.
2. **The origination score on the charge-off file.** Eight CUs already send one (§2.3) and
   nobody reads it. This is a wizard field and eight config edits.

A credit union with (2) plus even six months of history produces a real chart today — that
is exactly what Jackson River demonstrates.

## 6. The empty state

**A newly onboarded credit union legitimately has no charge-off migration chart, and the
report must say so.** Today it renders a titled, legended, empty frame — `patch_dq_pie_zero_labels`
injects `<c:delete val="1"/>` on every zero slice (`04 §Visual symptom`) — which reads as
broken software, not as absent data.

Proposed, in the slot the suppressed bar vacates, at `pcol_start`, `r_co + 6`:

> **Charge-off credit-grade migration is not yet available for this credit union.**
> This chart compares each charged-off loan's credit grade at origination with its grade at
> charge-off. Building it requires a credit score for the loan from before it charged off,
> which accumulates as monthly loan data is collected. Charge-off totals by pool and by year
> are shown on the Charge-off History tab.

and a second line, only when the derivation ran and was refused, so the analyst can see the
clock ticking without the client reading a diagnostic:

> *Coverage to date: 163 of 840 charged-off loans (19%). Expected available 2028-Q4.*

Rules:

* Keep writing the P/Q/R table either way — the zeros are honest and `report_integrity` and
  the TCT model read them. Suppress only the *chart*.
* The same guard, and near-identical wording, is already proposed for the DQ pie in
  `04 §6 Tier 1`. Ship them together; they are the same edit twice.
* When the derivation *succeeds* but a single pool is genuinely zero (as Utah's
  `Construction` is), that is a real zero and the note must **not** appear — key the guard on
  `co_by_status`/`co_by_pool` being absent or refused, not on the four values being zero.

## 7. Proposed implementation

### 7.1 Shape — follows `dq_migration_split.py`, deliberately

| | `dq_migration_split` (shipped) | `co_migration_split` (proposed) |
|---|---|---|
| Schema change | none | **none** |
| Fixes already-delivered quarters | yes | yes, where data allows |
| Reads | loan frame + `days_delinquent` off the extract | charge-off files + `monthly_loan_data` history |
| WARM-supplied data | returned untouched, no file opened | **identical** — verified below |
| Bucketing | `classify_migration` | the **same function**, imported |
| Refusal | alias guard + `MAX_PLAUSIBLE_DQ_SHARE` | `MAX_NOT_REPORTED_SHARE` + `MIN_MEASURED_SHARE` |
| Entry point | `fill_missing_dq_migration(hist, …)` | `fill_missing_co_migration(hist, …)` |

Verified no-op on a WARM-fed credit union:

```
>>> hist = {'impaired': load_impaired_data(cfg, '2026-06-30')}
>>> before = deepcopy(hist['impaired'])
>>> fill_missing_co_migration(hist, cfg, '2026-06-30', cfg['credit_grades'])
False            # returned in 0.000s -- no charge-off file was opened
>>> hist['impaired'] == before
True
>>> 'co_source' in hist['impaired']
False
```

### 7.2 Does the charge-off case need persistence? No — argued, not assumed

The brief flags "accumulating history is a plausible reason the precedent does not hold". It
is plausible, and it turns out not to hold, for three reasons:

1. **The history is already persisted.** `monthly_loan_data` retains every snapshot for 84
   months (`cecl_retention.py`), against a 36-month window. Nothing needs a new table to
   accumulate; the accumulation is the existing import.
2. **The derivation is idempotent and cheap.** A charged-off loan's recovered pair never
   changes once recovered — but re-deriving it costs one indexed-ish read of
   `monthly_loan_data` per credit union (20,538 accounts for Mountain, 0.4s) plus the
   charge-off files the by-year chart already reads. A cache would be an optimisation with a
   staleness failure mode, bought for nothing.
3. **Persistence would freeze today's poor coverage.** This is the decisive one and it is the
   mirror of `06_dq_derivation.md §1`'s argument. If we snapshot the recovered pair at
   charge-off time, a credit union's chart is permanently capped at the coverage it had the
   quarter each loan charged off. Deriving at report time means a **config fix, a key-format
   fix, or a newly mapped score column retroactively improves every quarter's chart** —
   which, given §4.2, is precisely the improvement path we are on.

The one thing that genuinely *is* lost forever is a score for a loan that charged off before
the credit union's first snapshot. No persistence strategy recovers that; only having started
earlier would have. What follows from it is the opposite of a new table: **stop discarding
what we already receive** —

* wire `member_col` through `_parse_chargeoff_file` (07 §5's dead config), so per-loan
  charge-off identity survives;
* collect `orig_score_col` / `curr_score_col` in the wizard's charge-off step (§2.3);
* keep importing every month, since the loan snapshot *is* the persistence layer.

### 7.3 The integration edit (described, not applied)

`generate_report.py`, immediately after the existing DQ block at `:13009-13015` — i.e. after
every `hist['impaired']` mutation and before the report dispatch loop:

```python
    # ── Charge-off by credit-grade migration (fallback) ──
    # Mirror of the DQ block above for ``co_by_status`` / ``co_by_pool``.
    # Recovers each charged-off loan's origination and at-charge-off scores
    # from the charge-off feed and the loan-snapshot history; refuses rather
    # than guesses when coverage is too thin. See
    # docs/pdf_migration/09_chargeoff_score_history.md.
    try:
        from co_migration_split import fill_missing_co_migration
        fill_missing_co_migration(hist, config, snapshot_date, grades,
                                  no_score=no_score)
    except Exception as _coe_exc:  # noqa: BLE001 - never block a report
        print(f"    CO migration split skipped: {_coe_exc}")
```

Eight lines of comment plus seven of code. `grades` and `no_score` are already in scope at
that point (the DQ call uses both). The module builds its own engine from
`cecl_credentials.get_database_url` when none is passed.

Also required, and independent of this module — **move the DQ/CO block above the early
`return result` in `load_impaired_data`** (07 §5, `04 §5`). It currently discards a perfectly
good `CO Data Entry` tab when the WARM lacks `ACL Env by Pool Mgmt Adj`, which is why Utah's
2025-12 workbook is blank.

Config surface added (all optional, all defaulting to today's behaviour):

```yaml
chargeoff:
  derive_migration_split: true        # false disables per CU
  migration_lookback_years: 3         # the validated 36-month window
historical_file_formats:
  chargeoff:
    orig_score_col: 10                # 0-based; the origination score
    curr_score_col: 2                 # 0-based; the score at write-off
    auto_detect_score_columns: false  # advisory detection, opt-in
```

### 7.4 Before shipping

* **`MIN_MEASURED_SHARE` has a known blind spot.** A pair of "orig from the charge-off file,
  current from a gap-filled snapshot" counts as measured, but if the credit union's extract
  score *is* its origination score the two are the same number and the pair is not really
  independent. Curis passes the guard at 59.3% measured and still lands 63% Unchanged with
  `Deteriorated = $972.74`. A stronger guard would require the two scores to differ on some
  minimum share of loans (Curis: 17 of 100; Jackson River: substantially more). Worth adding
  before this fills any client-facing chart.
* **Add a minimum-loan guard.** Jackson River's chart rests on 8 loans; a four-bucket
  percentage split on single digits is noise. Suggest refusing below ~25 classified loans, or
  below ~$100k, whichever the analyst prefers — needs a product call, so it is not in the
  module yet.
* **Deduplicate charge-off rows.** Several credit unions re-ship cumulative charge-off files
  each quarter (Utah: `TCT-Chargeoff.csv` appears in 11 quarter folders). The measurements
  above dedupe on `(account, date, code, amount)`; the module does **not**, because the
  existing by-year charge-off chart does not either and the two must agree. Deciding this is
  a separate call that affects a shipped chart.
* **`report_integrity`**: extend the all-zero check to accept "suppressed with a note" as a
  pass, or it will fail every report that correctly declines.

## A. Path A — the WARM `Mmm-YY` tabs (added after the first pass; supersedes §3.2 and §5)

The first pass measured score movement in `monthly_loan_data` only, found it to be ~0%, and
concluded the derivation was capped by that. **The measurement was right and the conclusion
was over-generalised.** The product owner supplied the missing source:

> "For security purposes the credit pull files are deleted after 6 months. However, credit
> unions that have a WARM file have the credit score pulls, with two columns: the member
> number and FICO/credit score. The tabs in these WARM files have the naming convention
> Mmm-YY. You can use those to determine the credit score at charge off. You should also be
> able to determine the original and current credit scores at charge off by looking at
> previous quarters."

That is Path A. It exists, it is richer than described, and **nothing in the pipeline reads
it in full**. `co_warm_history.py` implements the reader.

### A.1 What is actually in the WARM — two tab families, not one

Confirmed against the real files rather than assumed. A CECL-Migration-WARM workbook carries
**two** distinct `Mmm-YY` tab families:

| Family | Grain | Columns | Read today? |
|---|---|---|---|
| `Mmm-YY Credit Pull` | **member** | `Member Number` + `FICO` (2 cols), or `Member #` + `Current Credit Score` + `Open Date` (3 cols) | Partially — `import_data._load_dated_credit_pulls`, newest WARM only, opt-in path only |
| `Mmm-YY Data` | **loan** | 60–65 cols incl. `Member #-Suffix`, `Original Credit Score`, `Original Credit Grade`, `Current Credit Score`, `Current Credit Grade`, `Loan Pool`, `Current Balance`, `Charge Off Amount`, `Days Delinquent` | **Not at all** |

Brian's two-column description is exactly right for the `Credit Pull` family. The `Mmm-YY
Data` family is the larger find: it is the **full quarterly loan extract archived inside the
workbook, with both credit scores already present and already banded into grades**, going
back years before `monthly_loan_data` begins.

Franklin Trust FCU's 2026-06 WARM, verbatim:

```
Dec-25 Credit Pull:  5,463 rows x 2 cols   ['Member Number', 'FICO']
Dec-22 Credit Pull:  8,629 rows x 3 cols   ['Member #', 'Current Credit Score', 'Open Date']
Dec-22 Data:         8,625 rows x 65 cols  [... 'Member #-Suffix', 'Loan Pool',
                                            'Original Credit Score', 'Original Credit Grade',
                                            'Current Credit Score', 'Current Credit Grade', ...]
```

### A.2 Inventory — every credit union with a WARM

Tab names read straight out of `xl/workbook.xml` in the zip (`co_warm_history.list_tabs`), so
a 90-tab workbook on a network share answers in milliseconds. 98 WARM workbooks across 9
credit-union config folders:

| Credit union | WARMs | `Credit Pull` tabs | pull dates | `Mmm-YY Data` tabs | Data range |
|---|---:|---:|---|---:|---|
| franklin | 14 | 3 | 2022-12, 2024-12, 2025-12 | **15** | 2022-12 → 2026-06 |
| bridgeton | 10 | 4 | 2022-12, 2023-06, 2024-03, 2025-12 | **15** | 2022-12 → 2026-06 |
| tongass_fcu | 10 | 4 | 2022-12, 2023-06, 2024-09, 2025-12 | **15** | 2022-12 → 2026-06 |
| ontario_public_employees_fcu | 14 | 3 | 2022-12, 2023-06, 2023-09 | **15** | 2022-12 → 2026-06 |
| erie_fcu | 14 | 1 | 2022-03 | **15** | 2022-12 → 2026-05 |
| maple | 10 | **6** | 2022-12, 2023-06, 2023-12, 2024-09, 2025-03, 2025-09 | **14** | 2022-12 → 2026-06 |
| utah_community_fcu | 14 | 1 | 2025-09 | **14** | 2022-12 → 2026-03 |
| honolulu_fire_department_fcu | 10 | 3 | 2024-12, 2025-06, 2026-03 | **11** | 2022-12 → 2026-06 |
| united_community_fcu | 1 | 0 | — | **0** | — |

**Convention deviations, as asked:**

* **`united_community_fcu`** has one WARM with **no `Mmm-YY` tabs at all** — a different
  template. It is the one exception to the convention in this population.
* **Column headers are not stable across years.** Franklin's `Dec-22 Credit Pull` is
  `Member #` / `Current Credit Score` / `Open Date`; its `Dec-25 Credit Pull` is
  `Member Number` / `FICO`. The reader matches on a candidate list and falls back to
  "column 0 = member, column 1 = score", which is Brian's stated convention.
* **`Mmm-YY CC Data`** (credit-card data, e.g. `Mar-22 CC Data`) also matches the `Mmm-YY`
  prefix and is **not** a loan-data tab. `co_warm_history.classify_tab` excludes it.
* **Each quarter's WARM carries every earlier quarter's tabs**, so the newest workbook is
  usually sufficient. The loader still unions across all WARMs in the folder, newest-first,
  because a tab occasionally gets dropped on a re-issue — Honolulu carries 11 `Data` tabs
  where its siblings carry 15.
* `import_data._load_dated_credit_pulls` reads only `max(warms, key=os.path.getmtime)` — a
  **single** workbook — and skips any tab without the literal words "credit pull". That is
  the specific reason this history has been invisible.

### A.3 Score movement — the decisive measurement

The same measurement §3.2 ran against `monthly_loan_data`, run against both Path A families,
so the numbers are directly comparable. "Move" = the account/member appears in ≥2 tabs with a
score and the score is not identical in all of them.

| Credit union | source | in ≥2 | ever move | % | mean range |
|---|---|---:|---:|---:|---:|
| **Franklin Trust FCU** | Path A `Mmm-YY Credit Pull` (member) | 6,456 | 6,099 | **94.5%** | **49.5 pts** |
| | Path A `Mmm-YY Data` (loan) | 23,024 | 8,904 | **38.7%** | 19.0 pts |
| | Path B `monthly_loan_data` | 7,565 | 69 | 0.9% | 0.4 pts |
| **Bridgeton Onized FCU** | Path A `Credit Pull` (member) | 3,368 | 3,349 | **99.4%** | **57.6 pts** |
| | Path A `Data` (loan) | 7,525 | 5,871 | **78.0%** | 40.9 pts |
| | Path B `monthly_loan_data` | 3,064 | 25 | 0.8% | 0.3 pts |
| **Honolulu Fire Department FCU** | Path A `Credit Pull` (member) | 2,556 | 2,388 | **93.4%** | **85.4 pts** |
| | Path A `Data` (loan) | 5,632 | 5,196 | **92.3%** | **106.9 pts** |
| | Path B `monthly_loan_data` | 3,081 | **0** | **0.0%** | 0.0 |
| **Tongass FCU** | Path A `Credit Pull` (member) | 9,287 | 8,996 | **96.9%** | **47.1 pts** |
| | Path A `Data` (loan) | 8,384 | 6,939 | **82.8%** | 39.8 pts |
| | Path B `monthly_loan_data` | 4,127 | 123 | 3.0% | 0.9 pts |
| **Maple FCU** | Path A `Credit Pull` (member) | 8,219 | 8,180 | **99.5%** | **56.5 pts** |
| | Path A `Data` (loan) | 6,745 | 5,558 | **82.4%** | 40.6 pts |
| | Path B `monthly_loan_data` | 654 | 7 | 1.1% | 0.8 pts |

**The binding constraint identified in §3.2 dissolves for every credit union that has a
WARM.** Measured on five credit unions, the `Credit Pull` tabs are a genuine dated bureau
time series — **93.4–99.5%** of members move, by **47–85 points** on average. The `Mmm-YY
Data` tabs inherit that movement at loan level (**38.7–92.3%**). `monthly_loan_data` shows
**0.0–3.0%** for the same credit unions over the same period, with a mean range of under one
point. The gap is two orders of magnitude and it is consistent across every credit union
tested.

§3.2's conclusion is therefore restated: **the loan extract carries a static score, and
`monthly_loan_data` is the wrong place to look for movement — but the movement exists, in
the WARM, for the credit unions that have one.**

### A.4 The source hierarchy, revised

| Rank | Source | Grain | Both scores? | Movement | Reaches |
|---|---|---|---|---|---|
| 1 | WARM `CO Data Entry` (existing) | pre-aggregated | — | — | 8 CUs |
| 2 | WARM `CO Data` per-loan rows | loan | **yes**, pre-banded | — | 8 CUs |
| 3 | WARM `Mmm-YY Data` tabs | **loan** | **yes**, pre-banded | 38–92% | 8 CUs |
| 4 | WARM `Mmm-YY Credit Pull` tabs | member | current only | **93–99%** | 8 CUs |
| 5 | Charge-off file score columns (§2.3) | loan | one side each | — | 8 CUs |
| 6 | `monthly_loan_data` lookback (Path B) | loan | original only in practice | 0–4% | all CUs |

Path A occupies ranks 2–4 and is strictly better than Path B wherever it exists.

### A.5 Recency — do monthly WARM tabs beat quarterly extracts?

Brian's example: *"if a loan charges off in February, the loan would have likely been included
in the December data, so the scores could be pulled from that quarter."*

Correct, and worth stating precisely, because the naming misleads: **the `Mmm-YY Data` tabs
are quarterly, not monthly.** Every credit union above carries Mar / Jun / Sep / Dec only.
So the gap between a charge-off and its nearest preceding score observation is:

| Source | cadence | worst-case gap | mean gap |
|---|---|---|---|
| WARM `Mmm-YY Data` | quarterly | 3 months | ~1.5 months |
| WARM `Mmm-YY Credit Pull` | 1–6 pulls over 4 years | up to ~24 months | ~8 months |
| `monthly_loan_data` (Path B) | **monthly** | **1 month** | ~0.5 months |

**So Path B wins on recency and Path A wins on reach and on movement, and the two are
complementary rather than competing.** Where both cover the charge-off, the right choice is:

* **score at charge-off** — take the most recent observation from *any* source, preferring
  a `Credit Pull` or `Mmm-YY Data` value over a `monthly_loan_data` value **of the same or
  older date**, because the Path B value is a static loan-record score that carries no
  information about the loan's condition near write-off, whereas the Path A value is a real
  bureau observation. Recency alone is the wrong tie-break here: a monthly snapshot one month
  before the charge-off is *less* informative than a bureau pull six months before it, because
  the snapshot is not an observation at all.
* **original score** — take the earliest available `Original Credit Score`, which the
  `Mmm-YY Data` tabs carry directly and which Path B mostly gap-fills.

That reverses the naive "nearest date wins" rule, and it is the rule
`co_migration_split.recover_scores` needs to implement when Path A is wired in.

### A.6 Attaching a member-level score to a loan-level charge-off

Ranks 2, 3 and 5 are loan-level (`Member #-Suffix`, e.g. Franklin's `3040-99`) and join
directly — no ambiguity. Rank 4, the `Credit Pull` tabs, is **member-level**, so it has to be
broadcast to the member's loans. Measured on Franklin's `Dec-25 Data` (9,077 loans):

```
distinct members                                    4,507
loans per member                                    mean 2.01, max 7
members with more than one loan                     2,372 (52.6%)
loans belonging to a multi-loan member              6,942 (76.5% of loans)
multi-loan members whose loans carry DIFFERENT
    current  credit scores                            220 of 2,357  ( 9.3%)
    original credit scores                              0 of 2,357  ( 0.0%)
```

**Three quarters of loans belong to a member with more than one loan, so broadcasting is the
normal case, not an edge case.** The ambiguity it introduces is bounded and measurable:

* For the **current / at-charge-off** score, broadcasting is a near-lossless approximation:
  only **9.3%** of multi-loan members carry different current scores across their own loans
  in the WARM's own data, and a bureau pull is a property of the *member*, not the loan, so
  a single value per member is arguably the more correct representation.
* For the **original** score, broadcasting would be wrong in principle — two loans opened
  years apart should have different origination scores — but the WARM's own data shows
  **0.0%** within-member variation in `Original Credit Score`. That is itself a finding: the
  WARM's "original" score appears to be derived from a member-level pull too, not from a
  per-loan origination record. **Do not use a member-level pull for the original side**; take
  it from the loan-level `Mmm-YY Data` tab or the charge-off file, where it is at least
  recorded per loan.
* **Rule for the implementation:** a member-level pull may supply the *score at charge-off*
  only, flagged with its own provenance (`warm_pull`), never the original side; and where a
  loan-level `Mmm-YY Data` value exists for the same or a later date, the loan-level value
  wins.

The residual ambiguity to disclose: for the ~9% of multi-loan members whose loans genuinely
differ, a broadcast pull assigns one member score to loans that the WARM itself scores
differently. On Franklin's Dec-25 tab that is 220 members out of 4,507 — **4.9% of members,
covering at most ~1,000 of 9,077 loans (11%)**. That is the size of the error the member-level
fallback can introduce, and it is the reason it sits at rank 4 rather than rank 2.

### A.7 Reach — which credit unions Path A actually helps

Path A exists only where a CECL-Migration-WARM workbook exists. Tested directly
(`co_warm_history.warm_files` + a tab scan) against every credit union in the affected set,
and against the wider client tree by climbing each config's `credit_pull.source_folder` to
its Portfolio Management root:

| Credit union | WARM files reachable | `Mmm-YY Data` | `Mmm-YY Credit Pull` |
|---|---:|---:|---:|
| franklin (Franklin Trust FCU) | 14 | 15 | 3 |
| honolulu_fire_department_fcu | 10 | 10–11 | 3 |
| bridgeton (BRIDGETON ONIZED FCU) | 10 | 15 | 4 |
| utah_community_fcu | 14–15 | 9–14 | 1 |
| tongass_fcu | 10 | 15 | 4 |
| maple | 10 | 14 | 6 |
| erie_fcu | 14 | 15 | 1 |
| ontario_public_employees_fcu | 14 | 15 | 3 |
| united_community_fcu | 1 | **0** | **0** |
| **bridgeton_onized_fcu** *(wizard config, same CU as `bridgeton`)* | **0** | 0 | 0 |
| mountain_cu | 1 *(impaired-loans template)* | **0** | **0** |
| emergency_responders_cu | 1 *(impaired-loans template)* | **0** | **0** |
| central_susquehanna_comm_fcu | 1 *(impaired-loans template)* | **0** | **0** |
| nucor_emp_cu, sci_fcu, shuford_fcu, census_fcu, central_keystone_fcu, curis_palmetto_health_cu, wnc_community_cu, destinations_cu, jackson_river_community_cu, tcp_cu | **0** | 0 | 0 |

Two things to be careful about here:

* **A file called "WARM" is not necessarily *the* WARM.** Mountain CU, Emergency Responders
  and Central Susquehanna each have a `CECL-WARM with Credit Migration Impaired Loans*.xlsx`
  on the share. It is the impaired-loans template and carries **no `Mmm-YY` tabs at all**.
  Path A does not reach them.
* **Bridgeton Onized FCU has two configs for one credit union**, and the wizard one is blind
  to its own WARM:

  ```
  bridgeton.yaml               cu='BRIDGETON ONIZED FCU'
      fallback_report_folder = 'Z:\Shared\Clients\BRIDGETON ONIZED FCU\Credit Migration'
  bridgeton_onized_fcu.yaml    cu='Bridgeton Onized FCU'
      fallback_report_folder = None
      source_folder          = 'C:\Users\BRIANE~1\AppData\Local\Temp\cecl_ui_samples'
  ```

  `load_impaired_data` under `bridgeton` returns a fully populated `co_by_status`
  (`128,963.93 / 876,752.91 / 763,045.28 / 928,436.38`). Under `bridgeton_onized_fcu` it
  returns `{}` and logs *"no WARM files"*. The two also write to **different
  `credit_union` values**, so they hold separate row sets in `monthly_loan_data`
  (`BRIDGETON ONIZED FCU`: 3 snapshots; `Bridgeton Onized FCU`: 1).

**Collateral finding worth its own ticket:** `credit_pull.source_folder` is
`C:\Users\BRIANE~1\AppData\Local\Temp\cecl_ui_samples` — a **machine-local temp directory** —
in *eleven* wizard configs (bridgeton_onized_fcu, census_fcu, central_keystone_fcu,
curis_palmetto_health_cu, destinations_cu, first_area_fcu, jackson_river_community_cu,
nova_cu, nucor_emp_cu, sci_fcu, shuford_fcu, test_nova_cu). The wizard persists its upload
staging path as the permanent source. That path is unreachable from any other machine and is
subject to temp cleanup, which is a plausible contributing cause of these credit unions
having exactly one static credit pull baked into `monthly_loan_data` and no pull history at
all. This is not a charge-off bug, but it sits directly upstream of one.

### A.8 The revised timeline — the number to act on

`04_blank_charts.md`'s 19 all-zero workbooks span **14 credit unions**. Splitting them by
whether Path A reaches them:

| | Workbooks | Credit unions | Status |
|---|---:|---|---|
| **Path A available today** | **4** | Honolulu 2025-06, Bridgeton Onized 2025-12, Utah Community 2025-12, Franklin Trust 2026-03 | Scores are already on the share. **No waiting, no acquisition.** |
| No Path A — wizard CUs | 15 | Emergency Responders ×2, Nucor ×2, Shuford, Central Keystone ×2, Curis ×2, Census, Mountain ×2, Destinations, SCI, WNC | §5's forecast stands unchanged |

**§5 said "0 of 19 populate today". The correct number is 4 of 19**, and those four need
plumbing rather than data:

* For all four, `load_impaired_data` returns a **populated** `co_by_status` when run **now**
  against the current WARM — Franklin @ 2026-03 gives
  `275,724.64 / 325,167.44 / 1,226,557.97 / 38,133.59`; Honolulu @ 2025-06 gives
  `72,791.84 / 213,728.40 / 494,962.61 / 125,145.02`. So the delivered blanks are a
  **re-run**, not a data-acquisition problem: either the WARM's `CO Data Entry` was
  incomplete when the workbook was generated, or the report ran under a config that could
  not see the WARM (demonstrably the case for Bridgeton Onized).
* All eight WARM credit unions also carry a per-loan `CO Data` tab with both scores already
  present and banded (Franklin 2,410 rows, Bridgeton 2,157, Tongass 513, Honolulu 351,
  Maple 347, Utah 16,324). That is rank 2 in §A.4 and it is a stronger fallback than
  anything the lookback can produce, for the quarter *after* an analyst stops maintaining
  `CO Data Entry`.

**For the other 15 workbooks nothing in Path A changes the forecast.** Those credit unions
have no WARM, so they have no `Mmm-YY` tabs, no pull history, and a static per-loan score.
The §5.2 accrual table stands, and §5.3's conclusion is now much sharper: **a second dated
credit pull is not a nice-to-have that would improve the chart, it is the difference between
a chart and no chart**, and the WARM population proves how much difference it makes
(93–99% movement versus 0–3%).

**Revised recommendation ordering** (replaces the §RECOMMENDATION table's priorities 1–5
where they conflict):

| # | Action | Effect |
|---|---|---|
| 1 | **Re-run the four WARM-CU workbooks** and fix `bridgeton_onized_fcu.yaml` to point at the real WARM folder (or retire the duplicate config) | **4 of 19 workbooks, no new data** |
| 2 | Tier-1 empty state (§6) for the remaining 15 | Fixes the client-facing defect on the rest |
| 3 | **Read `CO Data` per-loan rows** as a rank-2 fallback under `CO Data Entry` (8 CUs, both scores already present) | Insulates WARM CUs against a stale matrix |
| 4 | **Wire `co_warm_history` into the derivation** — `Mmm-YY Data` then `Mmm-YY Credit Pull`, ahead of the `monthly_loan_data` lookback, with the §A.5 tie-break and the §A.6 member rule | Makes Path A the primary source wherever it exists |
| 5 | Fix the wizard's local-temp `credit_pull.source_folder` (§A.7) and ask the 10 wizard CUs for a second dated pull | The only thing that unblocks the remaining 15 |
| 6 | (unchanged) charge-off score columns §2.3, account-key mismatches §4.2, filename gate §4.3 | Incremental coverage |

### A.9 Blind end-to-end test — does Path A actually reproduce the analyst's answer?

Movement (§A.3) proves the source is informative. It does not prove the recovery works. The
test that does:

> Take a WARM credit union's charge-offs from its `CO Data` tab — account, date, amount,
> pool — **discard that tab's own credit-score columns entirely**, recover both scores from
> the `Mmm-YY Data` and `Mmm-YY Credit Pull` tabs alone, bucket them, and compare against
> the WARM's authoritative `CO Data Entry`.

That simulates precisely the credit union we cannot serve today: one that sends charge-offs
with no scores (every wizard CU) but has a pull history (which no wizard CU yet has). The
answer is known, so the recovery can be scored. Population is §1's rule — 36 months to
2026-06-30, `Exclude` and `HIDE-*` dropped. Script: `scripts/verify_co_pathA.py`.

| | Franklin Trust | Bridgeton Onized | Tongass | Maple |
|---|---:|---:|---:|---:|
| charge-offs in window | 1,206 | 883 | 245 | 193 |
| both scores recovered | **98.8%** | **97.7%** | 71.4% | **99.5%** |
| Improved — derived | 11.26% | 5.37% | 6.71% | 3.78% |
| Improved — WARM | 14.74% | 6.83% | 3.60% | 0.09% |
| Deteriorated — derived | 21.11% | 43.38% | 29.59% | **64.13%** |
| Deteriorated — WARM | 17.70% | 37.57% | 26.75% | **63.14%** |
| Unchanged — derived | 66.71% | 48.64% | 20.35% | 18.25% |
| Unchanged — WARM | 65.46% | 25.12% | 13.55% | 21.65% |
| Not Reported — derived | 0.93% | 2.61% | 43.35% | 13.84% |
| Not Reported — WARM | 2.09% | 30.48% | 56.10% | 15.12% |
| total — derived | $1,950,325 | $2,689,017 | $1,278,932 | $1,023,325 |
| total — WARM | $1,880,588 | $1,153,703 | $1,273,345 | $1,023,324 |

**Maple reproduces the analyst almost exactly** — totals agree to $1 (0.0001%), Deteriorated
64.13% vs 63.14%, Unchanged 18.25% vs 21.65%, Not Reported 13.84% vs 15.12%. **Tongass**
agrees on the total to 0.4% and every bucket to within 7pp, and Path A actually *recovers*
scores the analyst left as Not Reported (43.4% vs 56.1%) — i.e. the derivation is here
strictly better than the hand-built matrix. **Franklin** agrees on the total to 3.7% and every
bucket to within 3.4pp.

**Bridgeton is the outlier and it is not a recovery failure:** the derived total is
$2,689,017 against the WARM matrix's $1,153,703. §1 already established that `CO Data Entry`
is a *hardcoded paste*, not a live formula over `CO Data` — Bridgeton's paste is simply stale
and covers a different population. Its per-bucket *shares* still track (Improved 5.4 vs 6.8,
Deteriorated 43.4 vs 37.6).

Provenance, which shows the §A.5 / §A.6 design doing exactly what it was designed to do —
the loan-level `Mmm-YY Data` tab supplying the original score and the member-level
`Mmm-YY Credit Pull` supplying the score at charge-off:

```
franklin  warm_data+warm_pull 1127   warm_data  63   none 14   warm_pull  2
bridgeton warm_data+warm_pull  825   warm_data  37   none 20   warm_pull  1
tongass   warm_data+warm_pull  159   warm_data  12   none 70   warm_pull  4
maple     warm_data+warm_pull  147   warm_pull  45   none  0   warm_data  1
```

**Compare this with Path B on the same problem (§4.1): Mountain recovered 19.4% of loans and
1.7% of dollars as genuinely measured, and was refused. Path A recovers 71–99% and lands
within a few points of the analyst's own answer.** The mechanism is sound; the first pass was
starved of input, not wrong about the arithmetic.

One layout caveat found while running it: **Honolulu's `CO Data` tab has no account column**
that this reader recognises (`Member #-Suffix` / `Member-SFX` / `Loan Id` / `Account Number`
are all absent), so it could not be tested. Its `Mmm-YY` tabs are fine — only the charge-off
tab's layout differs. That needs a look before Honolulu is wired.

## 8. Not done / could not verify

* **Nothing is wired in.** `generate_report.py` is untouched; no report was regenerated.
  Everything in §4 comes from the module driven directly.
* **No config on the share was edited.** The score columns in §2.3 were wired in memory for
  measurement only.
* **The eight credit unions with no discoverable charge-off file (§4.3)** were not chased
  into the wizard's `_discover_co_recov_files` path. Doing so is probably the single largest
  remaining coverage lever and deserves its own pass.
* **Shuford's three-widths-in-one-CU account parse** was identified but not fixed; it needs
  a per-file format entry, which is an analyst config job.
* **The Curis `FICO Score` column was classified as origination** from the neighbouring
  `FICO Date` (years before the charge-off) on three sampled files, not from a client
  statement. Confirm before wiring it.
* **Utah's charge-off files sum to $45–52M over the window against the WARM's $28.95M** even
  after deduplication. Utah's chart comes from its WARM, so this does not affect any
  delivered number, but it means the raw charge-off feed and the analyst's `CO Data` tab do
  not tie, and the *by-year* charge-off chart is built on the untied side. Worth a look.
* **Recoveries** are still out of scope; the WARM has no `Recov Data Entry` tab and the bar
  is gross charge-offs only (07 §1).
* **No regression suite.** Verification was by direct execution against live data
  (34 credit unions at their latest snapshots, plus Utah's WARM as ground truth).

**Added in revision 2 (Path A):**

* **`co_warm_history.py` is not wired into anything.** `co_migration_split` still reads only
  `monthly_loan_data`. The §A.8 item 4 integration — Path A ahead of Path B, with the §A.5
  tie-break and the §A.6 member rule — is designed and measured but **not implemented**.
* **The four WARM-CU workbooks were not re-run.** §A.8 asserts they would populate because
  `load_impaired_data` returns a populated `co_by_status` when called now; no workbook was
  generated to confirm it end to end.
* **Honolulu's `CO Data` tab has an unrecognised account column** (§A.9), so Honolulu is the
  one WARM credit union whose blind test could not run.
* **`utah_community_fcu`, `erie_fcu` and `ontario_public_employees_fcu`** were queued for the
  §A.3 movement measurement but had not finished when this was written. The five that did
  finish agree so closely that the pattern is not in doubt, but their numbers are absent.
* **The `Mmm-YY Data` tabs also carry `Days Delinquent` and `Charge Off Amount`.** Both are
  directly relevant to `06_dq_derivation.md` and to §4.3's missing-charge-off-file problem,
  and neither was pursued here.
* **Nothing was written to the share and no config was edited**, including
  `bridgeton_onized_fcu.yaml` and the eleven configs whose `credit_pull.source_folder` is a
  machine-local temp path (§A.7).

---

### Reproduction

```bash
# Per-CU coverage, the four buckets, and the WARM alongside where one exists
set CECL_WORKSPACE_ROOT=Z:\Shared\TCT Files\CECL - CM Files
python scripts/verify_co_migration.py mountain_cu curis_palmetto_health_cu
python scripts/verify_co_migration.py --all

# The §4 table, i.e. with the unmapped charge-off score columns wired in memory
python scripts/verify_co_migration.py --all --wire-scores
```

```bash
# The 36-month population proof (§1)
python - <<'PY'
import sys, pandas as pd; sys.path.insert(0, r"C:\Dev\CECL")
W = (r"Z:\Shared\Clients\Utah Community FCU\Portfolio Management (CM, ID, Warm)"
     r"\2026-06 CECL-Migration-WARM - Utah Community FCU.xlsx")
d = pd.read_excel(W, sheet_name='CO Data')
d['amt'] = pd.to_numeric(d['CO Amount'], errors='coerce'); d = d[d['amt'].notna()]
dt = pd.to_datetime(d['CO Date for Rpt'], errors='coerce')
keep = d['Loan Pool'].isin({'Mortgage Real Estate','Indirect Auto','Direct Auto',
                            'Unsecured','Other Consumer','Commercial'})
print(d.loc[keep & (dt.dt.to_period('M') >= pd.Period('2023-07')), 'amt'].sum())
# -> 28953517.02   == the WARM's CO Data Entry grand total
PY
```

```bash
# Path A: tab inventory for one credit union, read straight out of the zip (fast)
python -c "import sys,os; sys.path.insert(0,r'C:/Dev/CECL'); "\
  "import co_warm_history as W, generate_report as gr; "\
  "cfg=gr.load_config('franklin'); "\
  "p=sorted(W.warm_files(cfg), key=os.path.getmtime, reverse=True)[0]; "\
  "print(os.path.basename(p)); "\
  "[print(' ', W.classify_tab(t), t, W.sheet_date(t).date()) "\
  "for t in W.list_tabs(p) if W.classify_tab(t)]"
```

```bash
# Path A: the movement measurement (A.3)
python -c "import sys; sys.path.insert(0,r'C:/Dev/CECL'); "\
  "import co_warm_history as W, generate_report as gr; "\
  "loan,member,meta = W.load_warm_tabs(gr.load_config('franklin')); "\
  "print('loan  ', W.score_movement(loan)); "\
  "print('member', W.score_movement(member))"
#   loan   -> (25936, 23024, 8904, 19.0)   38.7% of loans move
#   member -> ( 8543,  6456, 6099, 49.5)   94.5% of members move
```

```bash
# Path A: the blind end-to-end test (A.9) -- CO Data's own scores discarded
python scripts/verify_co_pathA.py franklin 2026-06-30 maple 2026-06-30 \
                                  tongass_fcu 2026-06-30
```

```bash
# The score-immobility measurement (§3.2)
python - <<'PY'
import sys, pandas as pd; sys.path.insert(0, r"C:\Dev\CECL")
from sqlalchemy import create_engine, text
from cecl_credentials import get_database_url
q = text("""
with s as (select credit_union cu, member_number a, snapshot_date d,
                  current_fico_score c
           from monthly_loan_data where current_fico_score > 0),
     agg as (select cu, a, count(distinct d) nsnap, count(distinct c) nscore
             from s group by 1,2)
select cu, sum((nsnap>=2)::int) multi, sum((nsnap>=2 and nscore>=2)::int) moved,
       max(nsnap) maxsnap
from agg group by 1 order by 2 desc""")
print(pd.read_sql(q, create_engine(get_database_url())).to_string(index=False))
# -> Test Nova CU: 8,128 accounts over 79 snapshots, 353 ever move (4.3%)
# -> Mountain CU:  14,021 accounts over  6 snapshots,   0 ever move (0.0%)
PY
```
