# 06 — Deriving the delinquency-by-migration split (Tier 2, DQ half)

Implements the DQ half of Tier 2 in `04_blank_charts.md`: fill
`hist['impaired']['dq_by_status']` and `hist['impaired']['dq_by_pool']` for credit unions
that have no legacy `CECL-Migration-WARM` workbook, so the *Delinquency by Credit Grade
Migration* pie on every `Risk Change` / `Risk Chg <pool>` tab stops plotting literal zeros.

**Scope:** delinquency only. `co_by_status` / `co_by_pool` are untouched — a separate
assessment covers charge-offs.

**Ground truth:** `utah_community_fcu` @ 2026-06-30 has a working WARM-supplied
`DQ Data Entry` tab. The derivation was run against it and compared line by line. Full
comparison in §4.

---

## 1. Design decision — (B) derive at report time from the loan extracts

Both candidates were evaluated. **(B) wins outright**, and (A) would not have satisfied the
brief on its own.

| | (A) persist `days_delinquent` into `monthly_loan_data` at import | (B) read `days_delinquent` off the loan extract at report time |
|---|---|---|
| Schema change | Yes — `ALTER TABLE … ADD COLUMN` | **None** |
| Fixes existing quarters | **No.** Only months re-imported after the change. Mountain CU 2026-06-30 (a delivered report) would stay blank until someone re-imports it. | **Yes, immediately** — verified for Mountain, SCI, Nucor below |
| Data quality | Identical. Both read the *same* `column_mappings['days_delinquent']` off the *same* extract file; (A) just freezes it earlier. A mis-mapped column is mis-mapped in both. | Identical |
| Join risk | None (already on the row) | Join on `member_number`; measured below |
| Proven machinery | New code path | `generate_impdet_report._load_extract_enrichment` already does exactly this join in production |

Decisive point: (A) cannot repair a single already-delivered workbook without a re-import of
every affected snapshot, and the whole reason this defect matters is the 19 workbooks
already in clients' hands. (A) is still worth doing later as a performance/history play, but
it is not the fix.

### The three checks the brief asked for

**1. Is `days_delinquent` reliably mapped for the affected CUs?**

| Config | top-level | per-extract | verdict |
|---|---|---|---|
| `mountain_cu` | `Day Delinquent` | both extracts: `Day Delinquent` | good |
| `sci_fcu` | `Past Due` | `Past Due`; CC extract `Account Delinquent Day Count` | good |
| `nucor_emp_cu` | `Days Deliquent` *(sic — matches the file header)* | `Days Deliquent`; Visa extract `PAST DUE` | good |
| `utah_community_fcu` | `DAYS_PAST_DUE` | `DAYS_PAST_DUE` | good |
| `destinations_cu` | **`Interest Rate`** | `ceclcc1` → `Interest Rate`; `cecloe` → **`ACCTBS`** (the account number!); `ceclce` → `DLQDAY` (correct); `ceclcc1` (Apr-26) → unmapped | **partly bogus** |
| `wnc_community_cu` | unmapped | unmapped | **absent** |

Not universally reliable. Two guards were added (see §2) so a mis-mapped column produces
*nothing* rather than a wrong chart.

**2. Is the extract present at report time?** Yes for every CU checked. `import_data`
archives extracts after import, and `_load_extract_enrichment` already searches
`data_directory`, `archive_directory` (defaulting to `<workspace>/Archive/<client_short>`)
and `loan_file_folder`, then date-pins the match to the report month via
`_resolve_extract_path`. Observed resolutions at 2026-06-30:

```
mountain_cu   Loan File - 6-30-26 .xlsx                    (+ April Loan File - Upload.xlsx)
sci_fcu       Aires Loans for Vizo (Non CC) 6-30-2026 v.2.xlsx
              Credit Card_CECL Report_6-30-2026v.2.xlsx
nucor_emp_cu  Aires Loans Jun 2026 V2.xlsx                 (+ Nucor Emp CU Jan 2026 Visa File V2.xlsx)
utah…_fcu     2026Q2_tct_loan_v2.csv
```

**3. What is the join key?** `monthly_loan_data.member_number`, which `import_data.py:1654`
sets to `raw_account` — the *full account string* returned by
`derive_member_account(df, config, has_header)` (fixed-suffix / delimiter / split modes).
`_load_extract_enrichment` keys its dictionary on the **same** expression from the **same**
config block, so the two sides are identical by construction. This is the join
`generate_impdet_report` already uses to fill the All Loans tab's *Days Delinquent* column.

Measured coverage (loans in the reserve population matched to an extract row):

| CU | matched / total | note |
|---|---|---|
| `mountain_cu` | 13,548 / 13,548 — **100.0%** | |
| `sci_fcu` | 797 / 797 — **100.0%** | |
| `nucor_emp_cu` | 2,154 / 2,154 — **100.0%** | |
| `destinations_cu` | 1,773 / 1,803 — 98.3% | |
| `utah_community_fcu` | 12,765 / 95,513 — 13.4% | see below |

Utah's 13.4% is *not* a join failure: its extract is one row per **member**-level account
(157,253 rows) while `monthly_loan_data` carries 95,513 loan-level rows, and only the
account strings that exist on both sides match. It still recovers **100.00%** of the WARM's
delinquent dollars (§4), i.e. every unmatched row is a non-delinquent one. The module prints
a `*** WARNING` whenever coverage drops below 95% so a genuinely stale extract is loud.

### Bucketing — identical to the matrix, by construction

`classify_migration()` in the new module is a line-for-line transcription of the
per-original-grade loop in `report_vizo._sheet_risk_change` (the loop that produces
`grand_det` / `grand_imp` / `grand_unc`, `report_vizo.py:1752-1772`):

```
either side == no_score_label      -> Not Reported
i > j  and  j < n_top and i-j < 2  -> Unchanged      (top_grades_double_drop exception)
i > j  otherwise                   -> Deteriorated
i < j                              -> Improved
i == j                             -> Unchanged
```

with `i`/`j` indexed into the same label list the sheet uses —
`[g for g in _all_grades(grades, no_score) if not _is_hidden(g)]`, or
`_brr_grade_labels(config, no_score)` for a pool flagged `brr: true` — and the same
`n_top = config.get('top_grades_double_drop', 3)`.

**Verified equivalence, not asserted.** Running both the per-loan classifier and a faithful
replay of the sheet's matrix loop over the whole Utah portfolio:

```
-- TOTAL (FICO labels)          delta imp  -0.00   delta det  0.00   delta (unc+NR)-unc  -0.00
-- POOL Commercial [BRR]        delta imp   0.00   delta det  0.00   delta               0.00
-- POOL Construction            0.00 / 0.00 / 0.00
-- POOL Direct Auto             0.00 / 0.00 / -0.00
-- POOL Indirect Auto           0.00 / -0.00 / 0.00
-- POOL Mortgage Real Estate    -0.00 / -0.00 / -0.00
-- POOL Other Consumer          0.00 / 0.00 / 0.00
-- POOL Unsecured               0.00 / 0.00 / -0.00
```

Improved and Deteriorated reproduce the matrix to floating-point zero on every tab, and the
classifier's `Unchanged + Not Reported` reproduces the matrix's single `Unchanged` (the
matrix folds the unscored cells into Unchanged; the DQ table splits them out into its
fourth row, exactly as the WARM's `DQ Data Entry` does). **The pie cannot disagree with the
table beside it.**

Two deliberate refinements over a naive reading of the sheet:

* **BRR pools use BRR labels.** A loan in a `brr: true` pool is classified against
  `business_risk_ratings`, matching that pool's tab. It is also used that way when rolling
  up the grand total, so `dq_by_status == Σ dq_by_pool`. (The Total tab's *matrix* drops BRR
  loans entirely — its summary absorbs them into the `total - imp - det` residual — but the
  WARM's own grand-total DQ block includes them, and Utah confirms it: the WARM's Commercial
  DQ total is exactly our Commercial DQ total.)
* **Anything still outside its pool's matrix is counted Unchanged**, mirroring
  `unc_bal = total - imp_bal - det_bal` at `report_vizo.py:1932`, so every delinquent dollar
  lands in a slice and the four percentages sum to 100%. The module prints the amount when
  this fires (it was $0.00 on every CU tested).

---

## 2. What changed

### New file — `dq_migration_split.py` (429 lines, CRLF)

Self-contained, no database writes, no schema change.

| Symbol | Purpose |
|---|---|
| `classify_migration(cg, og, grade_index, n_top, no_score)` | the bucketing rule above |
| `_coerce_days(val)` | safe days-past-due parse (see the warning below) |
| `_label_sets(config, grades, no_score)` | FICO / BRR label lists via `report_vizo` helpers |
| `_strip_bogus_dq_mappings(config)` | guard 1 — drops a `days_delinquent` mapping that aliases another mapped column |
| `load_days_delinquent_by_account(...)` | `{full_account: days}` via `generate_impdet_report._load_extract_enrichment` |
| `derive_dq_by_migration(...)` | returns `(dq_by_status, dq_by_pool)` in `_read_migration_blocks` shape |
| `fill_missing_dq_migration(hist, ...)` | the only entry point `generate_report` calls |

`fill_missing_dq_migration` **returns immediately if `dq_by_status` and `dq_by_pool` are
both already populated** — no extract is even opened. WARM-fed credit unions (Utah,
Franklin, Sample, Maple, Ontario, Bridgeton …) are byte-for-byte unaffected; verified by
loading Utah's WARM, calling the function, and asserting the dict is unchanged and no
`dq_source` marker was added.

`pct` is the bucket's share of the **delinquent** total, matching the WARM convention
(confirmed arithmetically against Utah's delivered workbook: 248,344.81 / 10,867,222.67 =
2.29%).

Threshold: `config['delinquency']['dq_threshold']`, default **60 days** — the WARM
convention, and confirmed exactly right by the Utah reconciliation in §4.
`config['delinquency']['derive_migration_split']: false` disables the whole thing per CU.

#### Two guards against a mis-mapped column

Replacing an empty chart with a *wrong* chart would be worse than the defect. Both guards
were driven by `destinations_cu`, whose wizard auto-mapping put `days_delinquent` on
`Interest Rate` and on the account-number column.

1. **Alias guard** — a `days_delinquent` reference that is byte-identical to another mapped
   field (`member_number`, `current_balance`, `interest_rate`, `loan_pool_code`, …) is
   dropped for that extract, with an explanatory log line.
2. **Plausibility guard** — if more than `MAX_PLAUSIBLE_DQ_SHARE` (50%) of matched loans come
   back `>= threshold` days, the whole derivation is refused and the charts stay as they are
   today, with a loud `*** DQ migration split REFUSED` line.

Result for Destinations: all three bogus mappings are named and ignored, the remaining
sources yield no `>= 60` rows, and the function returns `({}, {})` — i.e. today's behaviour,
not a fabricated chart. **`destinations_cu.yaml`'s `days_delinquent` mappings need fixing by
an analyst before Destinations' DQ pies can populate.** Same for `wnc_community_cu.yaml`,
which maps nothing.

### `generate_report.py` — +16 lines, one call site

Inserted immediately after `_compute_balance_adjustments(...)` and before the report
dispatch loop — i.e. after *every* `hist['impaired']` mutation in the function, so nothing
downstream can clobber it:

```python
    try:
        from dq_migration_split import fill_missing_dq_migration
        fill_missing_dq_migration(hist, config, snapshot_date, df, grades,
                                  no_score=no_score, workspace_root=BASE)
    except Exception as _dqe_exc:  # noqa: BLE001 - never block a report
        print(f"    DQ migration split skipped: {_dqe_exc}")
```

Lazy import (`generate_impdet_report` builds a SQLAlchemy engine at import time and needs
`DATABASE_URL`, which the module sets from `cecl_credentials` the way
`_derive_snapshot_dq_from_extracts` already does). Wrapped so it can never fail a report.

Line endings preserved: the file went from 13,138 CRLF / 0 bare-LF to 13,154 CRLF / 0
bare-LF — exactly the 16 lines added. `git diff --numstat` still shows the pre-existing
uncommitted delta and nothing endings-shaped. Both files pass `py_compile`.

Not touched: `report_vizo.py`, `report_tct.py`, `import_data.py`, `cecl_engine.py`,
`report_acl_model.py`, `cecl_report_web/`, `report_integrity.py`. No schema change was made,
so the Tier-2 step-1 `ADD COLUMN` was **not** needed.

---

## 3. Collateral finding — `dq_extract_parser._coerce_days` overstates DQ ~7×

Found while reconciling against Utah, **not fixed** (it feeds a different consumer and needs
its own sign-off). `cecl_ui/services/dq_extract_parser._coerce_days` strips every non-digit
character before parsing:

```python
>>> _coerce_days(60.0)   ->  600
>>> _coerce_days(30.0)   ->  300
>>> _coerce_days(6.0)    ->   60      # counted as 60+ days delinquent
>>> _coerce_days(147.0)  -> 1470
```

Loan extracts come back as float64 whenever a single cell is blank, so on those files **any
loan 6 or more days past due is booked as 60+**. Measured on Utah's
`2026Q2_tct_loan_v2.csv` (`DAYS_PAST_DUE` is float64):

| | loans ≥ 60 days | delinquent balance |
|---|---|---|
| correct numeric parse | 880 | **$12,108,522.68** |
| `_coerce_days` | 3,614 | $86,351,866.11 |
| Utah's WARM `DQ Data Entry` | — | **$12,108,522.68** |

`loan_code_delinquency_history` for Utah @ 2026-06-30 holds the $86,351,866.11 figure,
written by `generate_report._derive_snapshot_dq_from_extracts`, and it flows into
`hist['impaired']['warm_dq_pct']` — the **DQ% environmental factor**. This is a live,
client-facing overstatement on every CU whose extract parses as float.

The new module therefore uses its own `_coerce_days` (numeric first, regex fallback), which
the Utah reconciliation validates to the penny.

Related, same table: Mountain CU's `loan_code_delinquency_history` row for 2026-06-30 is
$15,846,547.03, but **3,995 of those 4,146 loans ($13,886,592.55) carry a non-zero
`C/O Amount`** and are routed to the `Exclude` pool at import — they are already charged off
and are not in the reserve population. The 151 remaining loans / $987,947.81 are exactly what
the new derivation reports. So the `warm_dq_pct` series also double-counts charged-off
balances for any CU with `chargeoff_exclude_column` set. Both issues belong in one follow-up
pass over `_derive_snapshot_dq_from_extracts`.

---

## 4. Ground truth — `utah_community_fcu` @ 2026-06-30

Command:

```bash
CECL_WORKSPACE_ROOT="Z:\Shared\TCT Files\CECL - CM Files" \
  python scratch/verify.py utah_community_fcu 2026-06-30 compare
```

`WARM` = `load_impaired_data(cfg,'2026-06-30')['dq_by_status'|'dq_by_pool']`, read from
`2026-06 CECL-Migration-WARM - Utah Community FCU.xlsx`, `DQ Data Entry`. `DERIVED` = the
new module, WARM data ignored.

### Grand total — the delinquent population matches **to the penny**

```
  status                  DERIVED      pct |             WARM      pct |           diff
  Improved             299,683.42    2.47% |       452,082.16    3.73% |    -152,398.74
  Deteriorated       6,246,532.63   51.59% |     8,919,340.45   73.66% |  -2,672,807.82
  Unchanged          5,364,598.75   44.30% |     1,599,742.18   13.21% |   3,764,856.57
  Not Reported         197,707.88    1.63% |     1,137,357.89    9.39% |    -939,650.01
  TOTAL             12,108,522.68          |    12,108,522.68          |           0.00
```

**$12,108,522.68 vs $12,108,522.68 — diff $0.00.** The extract join, the 60-day threshold and
the balance column are all exactly right; the WARM analyst and this derivation are selecting
the identical set of dollars. That is the strongest available confirmation of the mechanism.

### Per pool — 5 of 7 exact; the 2 misses offset exactly

| Pool | DERIVED total | WARM total | diff |
|---|---|---|---|
| Commercial | 2,062,454.33 | 2,062,454.33 | **0.00** |
| Construction | 0.00 | 0.00 | **0.00** |
| Direct Auto | 1,016,056.03 | 1,152,995.68 | −136,939.65 |
| Indirect Auto | 3,116,045.15 | 2,979,105.50 | +136,939.65 |
| Mortgage Real Estate | 4,465,858.97 | 4,465,858.97 | **0.00** |
| Other Consumer | 0.00 | 0.00 | **0.00** |
| Unsecured | 1,448,108.20 | 1,448,108.20 | **0.00** |

The Direct/Indirect Auto miss is a **single exactly-offsetting ±$136,939.65** — the WARM's
loan-code→pool map and `utah_community_fcu.yaml`'s `pool_map` disagree about one loan code.
That is a pool-mapping difference, not a derivation error, and the report's own `pool_map` is
authoritative for the report's own tabs. (The WARM also carries four pools the report does
not use at all — `Government Guaranteed`, `Student Loans`, `Participation Loans`,
`Deferred Orig Fees**` — all zero.)

### The status split does differ — here is exactly why, in full

The disagreement is **not** in the bucketing rule (proved identical to the matrix in §1) and
**not** in the DQ population (identical to the penny). It is that the WARM's `DQ Data Entry`
tab classifies each loan using the **WARM workbook's own** current/original score columns,
while the report — and therefore this derivation — must classify using the scores in
`monthly_loan_data`. Two concrete, fully-quantified causes:

**(a) Origination scores the extract does not carry — $3,633,073.58 of the $12.1M.**
248 delinquent Utah loans have no `APPLICATION_CREDITSCORE` in `2026Q2_tct_loan_v2.csv`.
`import_data.py:1701-1704` gap-fills `original = current` for those ("WARM convention"),
which makes them **Unchanged** in the report's matrix. The WARM sheet has origination scores
for most of them and calls them Deteriorated or Not Reported. The arithmetic closes exactly:

| Pool | gap-filled DQ balance | equals the pool's discrepancy |
|---|---|---|
| Mortgage Real Estate | 753,252.02 | = WARM `Not Reported` 528,783.78 + WARM `Deteriorated` shortfall 224,468.24 ✔ |
| Commercial | 2,062,454.33 | = the entire Commercial DQ balance ✔ |
| Direct Auto | 342,530.93 | |
| Unsecured | 350,562.58 | |
| Indirect Auto | 124,273.72 | |

Their derived split is `Unchanged 3,435,365.70 / Not Reported 197,707.88` — i.e. essentially
all of the +$3.76M `Unchanged` excess.

**(b) One BRR commercial loan — $1,974,658.61.** Utah's Commercial DQ is four loans. Three
(11,437.08 + 14,770.57 + 61,588.07 = **87,795.72**) are unscored and land in `Not Reported`,
matching the WARM's `Not Reported` **87,795.72 exactly**. The fourth is a $1,974,658.61 loan
rated `Watch` at both 2026-03-31 and 2026-06-30 in `monthly_loan_data`, so the report's own
`Risk Chg Commercial` matrix shows it on the diagonal — Unchanged. The WARM calls it
Deteriorated. Following the WARM here would put the pie in direct contradiction with the
matrix printed beside it.

**Systematic offset, stated plainly:** the derivation's `Unchanged` runs high and
`Deteriorated`/`Not Reported` run low relative to the WARM, by roughly the balance of
delinquent loans whose origination score is missing from the extract (30% of Utah's
delinquent dollars). Neither convention is "wrong" — but only one of them agrees with the
Risk Change matrix on the same page, and that is the derived one.

For completeness, alternative conventions were tested and none reproduces the WARM either
(Σ|diff| across the four buckets):

| variant | Σ abs diff vs WARM |
|---|---|
| as shipped (`n_top=3`, gap-fill honoured) | 7,529,713 |
| `n_top=0` (drop the top-grade exception) | 6,430,312 |
| treat gap-filled loans as `Not Reported` | 5,650,413 |
| both | 4,991,431 |

The residual is irreducible from this side: it is score data that lives only in the WARM.
Chasing it would break agreement with the matrix and still not match.

---

## 5. `mountain_cu` @ 2026-06-30 — where there were nine all-zero pies

Generated end to end (`reports=['vizo']`, `RPT_DIR` redirected to a scratch folder — nothing
in `Z:\…\Reports\` was written), then read back out of the produced workbook, columns M/N/O:

```
    DQ migration split derived from loan extract(s): 151 loan(s) >= 60 days,
      $987,947.81 delinquent (13,548/13,548 loans matched, 100.0% coverage).
    DQ migration split: filled dq_by_status, dq_by_pool (9 pools)
      (no WARM 'DQ Data Entry' source for this snapshot).
```

| Tab | Improved | Deteriorated | Unchanged | Not Reported | Total |
|---|---|---|---|---|---|
| **Risk Change Total** | 48,269.36 (4.9%) | 3,280.18 (0.3%) | 908,297.63 (91.9%) | 28,100.64 (2.8%) | **987,947.81** |
| Risk Chg Consumer Auto Loan-N | 0 | 0 | 0 | 0 | 0 *(genuine)* |
| Risk Chg Consumer Auto Loan-U | 16,035.80 (5.3%) | 0 | 279,723.74 (92.1%) | 8,082.93 (2.7%) | 303,842.47 |
| Risk Chg Consumer Indirect Au | 0 | 0 | 0 | 0 | 0 *(genuine)* |
| Risk Chg Consumer Indirect Au1 | 11,901.57 (3.1%) | 0 | 360,905.26 (94.4%) | 9,329.40 (2.4%) | 382,136.23 |
| Risk Chg Consumer Secured | 0 | 0 | 4,741.49 (100%) | 0 | 4,741.49 |
| Risk Chg Consumer Unsecured | 14,526.55 (9.9%) | 0 | 122,576.73 (83.7%) | 9,342.89 (6.4%) | 146,446.17 |
| Risk Chg Credit Cards | 5,805.44 (11.1%) | 3,280.18 (6.3%) | 41,969.78 (80.1%) | 1,345.42 (2.6%) | 52,400.82 |
| Risk Chg Real Estate | 0 | 0 | 98,380.63 (100%) | 0 | 98,380.63 |

**7 of the 9 previously-blank DQ pies now carry real data.** The two remaining zeros
(`Consumer Auto Loan-New Auto`, `Consumer Indirect Auto-New Auto`) are genuine — those pools
have no loan ≥ 60 days past due in the reserve population. `report_integrity`'s all-zero
warning for this workbook drops from 18 blocks to 11 (9 CO + the 2 genuine DQ zeros).

Independently cross-checked: the June extract has 4,146 rows at ≥ 60 days
($14,874,540.36), of which **3,995 ($13,886,592.55) carry a non-zero `C/O Amount`** and are
routed to the `Exclude` pool by `chargeoff_exclude_column`. The remaining
**151 loans / $987,947.81** is exactly the derivation's output — the DQ pie correctly
reflects the reserve population the Risk Change matrix is built from.

### The other affected CUs

| CU | outcome at 2026-06-30 |
|---|---|
| `sci_fcu` | 6 loans, **$34,332.72** (Improved 51.2% / Unchanged 44.3% / Not Reported 4.5%), 100% join coverage |
| `nucor_emp_cu` | 19 loans, **$226,846.83** (100% Unchanged), 100% join coverage |
| `destinations_cu` | **refused** — all three `days_delinquent` mappings are bogus (§2). Config fix needed. |
| `wnc_community_cu` | skipped — `days_delinquent` not mapped. Config fix needed. |

---

## 6. Not done / could not verify

* **Charge-offs.** `co_by_status` / `co_by_pool` are untouched — out of scope, and 9 CO bars
  per Mountain workbook remain blank.
* **`report_tct.py`** reads the same four keys (`:2387-2398`) from the same `hist`, so TCT
  models pick the derived DQ up for free — but no TCT workbook was regenerated to confirm it,
  because none of the CUs tested has `reports.tct` enabled.
* **`_coerce_days` / charged-off double-count in `_derive_snapshot_dq_from_extracts`** (§3) —
  diagnosed and quantified, deliberately **not** fixed. It changes the DQ% environmental
  factor for many credit unions and needs its own verification pass and analyst sign-off.
* **Utah's residual status-split difference** is explained and quantified but not eliminated,
  and cannot be from this side: it depends on origination scores that exist only inside the
  WARM workbook. See §4(a)/(b).
* **The Direct/Indirect Auto ±$136,939.65 pool swap** was not chased down to the specific
  loan code; it is a `pool_map` vs WARM-loan-code-map difference, not a DQ question.
* **Mountain's second extract resolves to a stale `April Loan File - Upload.xlsx`.** Harmless
  here — `_load_extract_enrichment` prefers the richer record and the June file is processed
  first, and coverage is 100% — but a mixed-month enrichment is a latent hazard worth a
  `file_pattern` cleanup.
* **No regression suite exists** for these paths; verification was by direct execution
  against live data (Utah, Mountain, SCI, Nucor, Destinations, WNC at 2026-06-30) rather
  than by tests.
