# 07 — Charge-off by migration status (`co_by_status` / `co_by_pool`): feasibility

Second half of the `04_blank_charts.md` fix. Scope: **charge-offs only**. The delinquency
half (`dq_by_status` / `dq_by_pool`) is being handled separately.

Evidence base: `generate_report.py`, `import_data.py`, `report_vizo.py`,
`cecl_ui/services/{co_recov_parser,monthly_co_recov_aggregator,chargeoff_hist_processor,
recovery_hist_processor}.py`, the live CECL Postgres database, the client configs under
`Z:\Shared\TCT Files\CECL - CM Files\client_configs\`, Mountain CU's raw CO uploads, and the
full 82-tab legacy WARM
`Z:\Shared\Clients\Utah Community FCU\Portfolio Management (CM, ID, Warm)\2026-06 CECL-Migration-WARM - Utah Community FCU.xlsx`.

No source file was modified. Every number below was read from the live shares/DB.

---

## VERDICT (question 3) — **NOT DERIVABLE from present data**

`co_by_status` cannot be computed from anything the pipeline currently ingests, for any of the
affected credit unions. This is not a plumbing gap that a join would close. Two independent
blockers, either of which alone is fatal:

1. **The charged-off loans are not in the loan-level data.** For **9 of the 10**
   wizard-onboarded CUs affected by this defect, `monthly_loan_data` contains **zero** rows
   for charged-off loans — not in an `Exclude` pool, not anywhere. There is nothing to bucket.

2. **Where the loans *are* present, the grade dimension is degenerate.** Mountain CU is the
   single exception: its 2026-06-30 extract carries the charged-off back-book, and 3,997 rows
   land in `Exclude`. But of those 3,997, only 1,541 carry a credit score at all, and only
   **39** have `original_fico_score != current_fico_score`. Bucketing them with the
   report's own rule would put ~99% of the charge-off dollars in **"Unchanged"** and render a
   single-bar chart that is confidently, specifically wrong. That is a worse outcome than the
   blank chart it replaces.

The migration split is a comparison of *two credit scores for the same loan* — one at
origination, one at/near the charge-off event. The pipeline never captures either one for a
charged-off loan. See §3 for exactly where each is lost.

## RECOMMENDATION (question 5) — **(d), with a narrow (b) substitution**

**Do not attempt to derive `co_by_status`.** Suppress the chart, and put a chart that is
actually true in the slot it vacates:

- **(d)** Suppress the CO migration bar when `co_by_status` is absent — the Tier-1 guard
  already proposed in `04_blank_charts.md §6`, shared verbatim with the DQ specialist's change.
- **(b)** Replace it with **charge-offs by pool and by year**, sourced from
  `loan_code_chargeoff_history`, which **every affected CU already has** with 29–90 months of
  depth (table in §1). This is a real, defensible, already-loaded number. It drops the
  migration dimension — which is precisely the dimension we cannot honestly supply.
- Keep `co_by_status` as an **optional WARM/upload-supplied input**, never a derived one, so
  the legacy CUs that still hand over a full WARM keep their analyst-maintained figures.

Option **(a)** is ruled out by the verdict. Option **(c)** — capture it as wizard input — is
*possible* but is a data-acquisition project, not a software change: the WARM analyst does not
type these numbers in, they paste a grade×grade matrix built inside the workbook from a CO
transaction file that carries origination and current credit scores per loan (§4). To
reproduce that we would have to ask each CU for a charge-off extract they do not currently
send. Recommend deferring (c) until a client actually asks for the chart back.

**Effort estimate**

| Item | Estimate |
|---|---|
| Tier-1 suppression guard, `report_vizo.py` + `report_tct.py` (shared with DQ fix) | 0.5 day |
| Replacement "Charge-offs by pool / by year" chart from `loan_code_chargeoff_history` | 2–3 days |
| Move the DQ/CO block above the early `return` at `generate_report.py:4990` (§5 of doc 04) | 0.5 day |
| *(deferred)* Option (c): wizard CO-with-scores upload + matrix builder + per-CU data ask | 1–2 weeks + client lead time |

Recommended now: the first three, ~3–4 days total.

---

## 1. Where charge-off and recovery amounts live today (question 1)

Five distinct stores. **None carries a credit grade, and only one carries a loan identity.**

| # | Source | Grain | Has member/loan id? | Has credit score/grade? |
|---|---|---|---|---|
| 1 | Raw per-month CO/Recovery upload files | **per loan transaction** | **yes** | **no** |
| 2 | `_parse_chargeoff_file` / `_parse_recovery_file` output | per transaction, `[code, amount, date]` | **no — dropped** | no |
| 3 | `loan_code_chargeoff_history` / `loan_code_recovery_history` (DB) | per (cu, month, **loan code**) | no | no |
| 4 | `hist['chargeoffs']` / `hist['recoveries']` | per (**year**, **pool**) | no | no |
| 5 | `hist['co_monthly']` / `hist['rc_monthly']` | per (**pool**, `YYYY-MM`) | no | no |
| 6 | WARM `CO Data Entry` tab → `co_by_status` / `co_by_pool` | per (pool, **current grade × original grade**) | no (pre-aggregated) | **yes** |

**(1) The raw files.** These are the only place a charge-off has an identity.
`cecl_ui/services/co_recov_parser.py` maps `account_col` / `code_col` / `amount_col` /
`date_col`, and the wizard additionally saves a `member_col`. Mountain CU's
`Charge off and Recoveries Apr2024.xlsx` is representative:

```
Charge Off Date | Loan Suffix | Account | Account Number | Loan Type | Charge Off Amount | Recovery Amount
      (blank)   |    0001     |  51912  |   519120001    |   8001    |      (blank)      |    28734.41
```

Account Number (`519120001`) is member + suffix and **matches
`monthly_loan_data.member_number` byte-for-byte** — the join key exists at the source. There
are no credit-score columns.

**(2) The parser throws the identity away.** `generate_report.py:470` —
`_parse_chargeoff_file` returns a DataFrame with **`['code', 'amount', 'date']` only**. The
account column is read solely to *filter out* total/comment rows
(`cfg_df = cfg_df[acct_numeric]`, `:519`). `member_col` is written into the config by the
wizard (`cecl_ui/routes/setup.py:9450`) and appears in `mountain_cu.yaml`
(`historical_file_formats.chargeoff.member_col: 2`) but is **read by nothing** — a repo-wide
grep finds no consumer in `generate_report.py`, `import_data.py`, or the aggregator. This is
the exact point at which per-loan charge-off identity is lost.

**(3) The database.** `chargeoff_hist_processor.py:17` —
`PRIMARY KEY (cu, as_of_date, loan_code)`, columns `chargeoff_amount, source, updated_at`.
The recovery table is identical in shape. Loan code, not loan. Live depth is good:

| CU | CO months | Range | Total |
|---|---|---|---|
| Mountain CU | 81 | 2018-12 → 2026-04 | $8,025,580.74 |
| Destinations CU | 90 | 2018-11 → 2026-04 | $1,661,880.23 |
| Nucor Emp CU | 52 | 2018-12 → 2026-03 | $623,632.57 |
| SCI FCU | 53 | 2018-12 → 2025-12 | $249,680.26 |
| Emergency Responders CU | 42 | 2018-12 → 2026-06 | $1,046,782.33 |
| Census FCU | 47 | 2018-01 → 2026-03 | $260,508.37 |
| Central Keystone FCU | 40 | 2019-03 → 2026-04 | $191,669.57 |
| WNC Community CU | 30 | 2018-09 → 2026-03 | $114,035.00 |
| Shuford FCU | 29 | 2018-12 → 2025-12 | $194,971.00 |

This is the feed that makes recommendation **(b)** viable: every affected CU already has
multi-year charge-off totals by pool, in the database, today.

**(4)/(5) `hist`.** `generate_report.py:4390` sets `hist['chargeoffs']` /
`hist['recoveries']` as `{year: {pool: amount}}`, and `co_monthly` / `rc_monthly` as
`{pool: {'YYYY-MM': amount}}`. Pool-level roll-ups of (3). No grade dimension.

**(6) The WARM.** The only source that has ever carried the grade split — see §4.

**One clarification on naming:** `co_by_status` is **gross charge-offs only**. The WARM's
`CO Data Entry` tab is titled `Charge off Amount` (cell A3) and the 82-tab workbook has **no
`Recov Data Entry` tab**. Recoveries never entered this chart; the report label
"CO Balance" is accurate.

## 2. Is a charged-off loan in `monthly_loan_data`? (question 2)

**Usually not. The table has no charge-off column at all.** Live schema:

```
credit_union, snapshot_date, member_number, loan_pool,
current_balance, current_fico_score, original_fico_score, business_risk_rating
```

No `days_delinquent`, no charge-off amount — confirming `04_blank_charts.md §3`.

**The `Exclude` routing (`import_data.py:1665-1683`) relabels; it does not drop.**

```python
_co_ref = config.get('chargeoff_exclude_column')
if _co_ref is not None and _co_ref != '':
    _co_series = df[_co_ref] if has_header else df.iloc[:, int(_co_ref)]
    _co_mask = pd.to_numeric(_co_series, errors='coerce').fillna(0) != 0
    if _n_co:
        clean_data.loc[_co_mask, 'loan_pool'] = 'Exclude'
```

The row survives into the DB with its member number and both FICO columns intact. It is
dropped only later, by the universal balance filter
`clean_data = clean_data[clean_data['current_balance'] > 0]` (`:1712`) — so a loan whose
extract balance was zeroed at write-off disappears, while one still showing its pre-write-off
balance persists as `Exclude`.

**But this path is almost never active.** `chargeoff_exclude_column` is set in exactly **two**
of the 39 client configs — `mountain_cu.yaml` and `nova_cu.yaml`. For every other affected CU
it is unset, and charged-off loans are simply absent from the extract the core exports.

Measured across the affected CUs, at each one's latest snapshot:

| CU | rows | `Exclude` rows | `orig != curr` score |
|---|---|---|---|
| **Mountain CU** | 17,545 | **3,997** | 3,785 (21.6%) |
| Nucor Emp CU | 2,154 | **0** | 430 |
| SCI FCU | 797 | **0** | 269 |
| Destinations CU | 1,803 | **0** | 806 |
| WNC Community CU | 1,284 | **0** | 0 |
| Census FCU | 522 | **0** | 46 |
| Central Keystone FCU | 1,352 | **0** | 268 |
| Emergency Responders CU | 1,367 | **0** | 436 |
| Shuford FCU | 1,357 | **0** | 244 |
| Curis CU / Palmetto Health CU | 2,972 | **0** | 560 |

Nine of ten have no charged-off loans in loan-level data whatsoever.

## 3. Why it is not derivable — the two missing scores

The bucketing *rule* is not the problem. `report_vizo.py:1750-1780` already implements it, and
it matches the WARM's formulas exactly (see §4): grade index `i` (current) vs `j` (original),
`n_top = config.get('top_grades_double_drop', 3)` giving the top grades a two-wide "still
unchanged" band, and `Not Reported → Unchanged`. If we had a charge-off matrix we could bucket
it today with existing code.

What is missing is the **(original grade, current grade) pair for each charged-off loan**.

**Blocker A — the loans predate the loan-level data.** Mountain CU's charge-off history runs
2018-12 → 2026-04 (81 months, $8.03M). Its `monthly_loan_data` starts 2025-12-31 (6
snapshots). Only **5 months / $736,852 — 9.2% of the charge-off dollars** — fall inside the
window where loan-level rows exist at all. The other 90.8% happened before the pipeline had
any record of the loans. Every affected CU has this shape: 29–90 months of CO history against
3–6 months of loan snapshots.

**Blocker B — even inside the window, the two scores are the same number.** This is the
decisive one. Empirical test: take the 26 charged-off accounts from Mountain's
`Charge off and Recoveries Apr2024.xlsx` and look them up in `monthly_loan_data`:

```
snapshot     matched  in Exclude  with score
2026-06-30      25        25          22
(no other snapshot matches any of them)
```

They resolve only in the one 2026-06-30 extract that happened to include the back-book, and
every match looks like this:

```
member_number  balance   loan_pool  current_fico  original_fico
221160300      3184.99   Exclude        653           653
410660007      5686.02   Exclude        598           598
505990060      1768.56   Exclude          0             0
651690300      1610.92   Exclude        633           633
805950002      4613.15   Exclude        785           785
908230001     16241.89   Exclude        562           562
```

`original == current` on every row. Across the whole `Exclude` population at 2026-06-30:
**3,997 rows, 1,541 with any score, 39 with `original != current`.**

The cause is structural. Mountain's config maps both scores to the *same* extract column:

```yaml
original_fico_score: Credit Score
current_fico_score: Credit Score
```

The two are separated only afterwards, by `_origination_aware_scores` joining a dated credit
pull (`credit_pull.pull_as_of_date: '2025-12-31'`). That pull covers *currently active
members*; it does not reach back to a loan charged off in April 2024. So for exactly the
population we need to split, the origination-aware step is a no-op and original collapses onto
current. Bucketing that yields ~99% "Unchanged" — a specific, wrong, client-facing claim.

**No third source closes the gap.** The CO files carry no scores. The DB CO tables carry no
member. `monthly_loan_data` carries scores but not the charged-off loans (and not, for these
loans, two distinct scores). There is no fourth store.

## 4. What the healthy path actually contains (question 4)

`load_impaired_data(load_config('utah_community_fcu'), '2026-06-30')` with
`CECL_WORKSPACE_ROOT="Z:\Shared\TCT Files\CECL - CM Files"` loads the WARM and prints
`CO by migration status: 4 categories, Total CO: $28,953,517.02`.

**`co_by_status` (grand total):**

| Status | balance | pct |
|---|---:|---:|
| Improved | 1,116,964.01 | 3.86% |
| Deteriorated | 15,151,657.35 | 52.33% |
| Unchanged | 8,689,857.16 | 30.01% |
| Not Reported | 3,995,038.50 | 13.80% |
| **Total** | **28,953,517.02** | |

**`co_by_pool` (11 pools):**

| pool | Improved | Deteriorated | Unchanged | Not Reported | Total |
|---|---:|---:|---:|---:|---:|
| Mortgage Real Estate | 0.00 | 0.00 | 7,956.17 | 12,816.18 | 20,772.35 |
| Indirect Auto | 608,389.69 | 8,711,525.87 | 4,355,993.88 | 972,162.93 | 14,648,072.37 |
| Direct Auto | 148,625.34 | 2,119,527.66 | 1,523,070.45 | 772,725.90 | 4,563,949.35 |
| Unsecured | 354,140.89 | 4,316,016.31 | 2,789,075.05 | 2,045,315.11 | 9,504,547.36 |
| Other Consumer | 5,808.09 | 4,587.51 | 1,632.83 | 1,144.80 | 13,173.23 |
| Commercial | 0.00 | 0.00 | 12,128.78 | 190,873.58 | 203,002.36 |
| Construction | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 |
| Government Guaranteed | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 |
| Student Loans | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 |
| Participation Loans | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 |
| Deferred Orig Fees** | 0.00 | 0.00 | 0.00 | 0.00 | 0.00 |

Shape: `{pool: {status: {'balance': float, 'pct': float}}}`, `pct` already a fraction of that
block's total. The four `Construction`…`Deferred Orig Fees**` zeros are genuine, matching the
one blank tab noted in `04_blank_charts.md`.

Note the magnitude: **$28.95M** of charge-offs against a $33.3M pooled allowance. This is
**cumulative across the whole migration study period, not one quarter** — see below.

### What the analyst actually enters — `CO Data Entry`

**It is not a human judgement, and it is not four typed numbers.** The tab is a stack of
per-pool **current-grade × original-grade matrices of charge-off dollars**. For each pool:
column A = current grade (rows `A+, A, B, C, D, E, Hide-F…Hide-I, Not Reported`), columns B–L
= original grade, same ladder. Cells `B7:L17` are **hardcoded pasted values**. Mortgage Real
Estate, in full:

```
A5 'Mortgage Real Estate'   B5 'Original Grade'    P5 'Loan Status'  Q5 'total'  R5 '% of Total CO'
D9  = 7956.17     (current B, original C)
L17 = 12816.18    (current Not Reported, original Not Reported)
M18 = 20772.35    total
```

The Improved / Deteriorated / Unchanged / Not Reported cells in P/Q/R — the four values
`_read_migration_blocks` reads — are then **pure formulas over that matrix**:

```
Q6 Improved     =SUM(C7:K7,D8:K8,E9:K9,F10:K10,G11:K11,H12:K12,I13:K13,J14:K14,K15)
Q7 Deteriorated =SUM(B9,B10:C10,B11:E11,B12:F12,B13:G13,B14:H14,B15:I15,B16:J16)
Q8 Unchanged    =SUM(B7:B8,C8:C9,D9:D10,E10,F11,G12,H13,I14,J15,K16)
Q9 Not Reported =SUM(B17:L17,L7:L16)
R6 =IFERROR(Q6/(SUM(Q6:Q9)),0)
```

`Q8`'s two-wide diagonal for the first three grades (`B7:B8`, `C8:C9`, `D9:D10`, then single
`E10`, `F11`, `G12`…) is exactly `top_grades_double_drop = 3` as implemented in
`report_vizo.py:1758-1770`. **The rule already lives in our code; only its input is missing.**

### Where the analyst gets the matrix — `CO Data`

The workbook builds it upstream, on a `CO Data` tab of **16,303 charge-off transactions**:

```
Charge-Off Date | Loan Id | Loan Type | chrg_off_amt | Original Credit Score | Recent Credit Score |
Member | Suffix | Member-SFX | Org Score | Org Score Prev | Curr Scr Hard | Curr Score Prev |
Loan Type | Loan Pool | Pool order | CO Amount | Original Credit Score | Org Grade |
Current Credit Score | Curr Grade | CO Date for Rpt
```

Sample rows (dates are Excel serials; 42278 = 2015-10-31):

```
42278 | 131930  | CARDPLATV | 5806.25 | … | Unsecured |   0 | Not Reported | 648 | B
42278 | 604209  | CARDPLATV | 4853.80 | … | Unsecured | 730 | A+           | 683 | A
42278 | 1196305 | CARDPLATV |  197.60 | … | Unsecured | 699 | A            | 546 | E
```

Every one of the 16,303 rows has **both** `Org Grade` and `Curr Grade`, derived by banding
`Org Score` (score at origination) and `Curr Scr Hard` (a hard credit-pull score) through the
`Grade Ranges & Loan Codes` tab; `Org Score = 0` becomes `Not Reported`. The date range starts
**2015-10** — confirming the $28.95M is an ~11-year cumulative figure, and explaining why the
totals dwarf any single quarter.

**This is the crux.** The WARM's authority for this chart is a per-loan charge-off file
carrying *origination and current credit scores*. That file is what we do not have and do not
ask for. The CO uploads our CUs send carry account, code, amount and date — and nothing else.

## 5. Secondary observations

- **The `member_col` dead end is worth closing either way.** The wizard collects it, the
  config stores it, nothing reads it. Either wire it through `_parse_chargeoff_file` or drop
  it from the wizard — as it stands it implies a per-loan capability the engine does not have.
- **The early-return bug from `04_blank_charts.md §5` still applies to the CO half.**
  `generate_report.py:4985-4990` returns before the CO block at `:5192` when the WARM lacks an
  `ACL Env by Pool Mgmt Adj` tab, silently discarding good `CO Data Entry` data. Fix it in the
  same pass as the DQ change.
- **Asymmetry with the DQ half is real and expected.** DQ *is* derivable — `days_delinquent`
  is mapped in every affected CU's config, and delinquent loans are live, in the extract, with
  both scores. Charge-offs are not, for the reasons above. The two halves of this defect
  should not be expected to land the same way, and the CO half should not be held back waiting
  for a derivation that cannot exist.

---

### Reproduction

```bash
# Utah's real co_by_status / co_by_pool
CECL_WORKSPACE_ROOT="Z:\Shared\TCT Files\CECL - CM Files" python - <<'PY'
import os, sys; sys.path.insert(0, r"C:\Dev\CECL")
import generate_report as gr
r = gr.load_impaired_data(gr.load_config('utah_community_fcu'), '2026-06-30')
print(r['co_by_status'])
PY
```

```bash
# The degenerate grade pair on Mountain's charged-off loans
python - <<'PY'
import os, sys; os.environ['CECL_WORKSPACE_ROOT'] = r"Z:\Shared\TCT Files\CECL - CM Files"
sys.path.insert(0, r"C:\Dev\CECL")
from sqlalchemy import create_engine, text
from cecl_credentials import get_database_url
with create_engine(get_database_url()).begin() as c:
    print(c.execute(text("""
      SELECT COUNT(*), SUM(CASE WHEN COALESCE(current_fico_score,0)>0 THEN 1 ELSE 0 END),
             SUM(CASE WHEN original_fico_score IS DISTINCT FROM current_fico_score THEN 1 ELSE 0 END)
      FROM monthly_loan_data WHERE credit_union='Mountain CU'
        AND snapshot_date='2026-06-30' AND loan_pool='Exclude'""")).fetchone())
    # -> (3997, 1541, 39)
PY
```
