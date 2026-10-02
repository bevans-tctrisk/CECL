# 08 — `_coerce_days` blast radius

Measurement-only follow-up. No source file was modified. Every figure below was read from
the live CECL Postgres database, the client extracts on the Egnyte shares, or the delivered
workbooks in `Z:\Shared\TCT Files\CECL - CM Files\Reports`.

Evidence base: `cecl_ui/services/dq_extract_parser.py`, `generate_report.py`,
`report_vizo.py`, `import_data.py`, `dq_migration_split.py`, all 38 live client configs,
the `loan_code_delinquency_history` table, and the 2026-03-31 / 2026-06-30 Vizo and TCT
model workbooks for the affected credit unions.

---

## BOTTOM LINE

**Two credit unions have corrupted delinquency data from `_coerce_days`. One of them moved
money. The total allowance impact of `_coerce_days` is $4,320.42, and it has been
delivered to a client.**

The headline number in the brief is real but it does **not** move the allowance. Utah
Community FCU's `loan_code_delinquency_history` does hold $86,351,866.11 against a true
$12,108,522.68 — confirmed, reproduced exactly — but Utah's environmental factor is
structurally immune to that corruption (§4). Utah's delivered allowance is unchanged, in
both the Vizo and the TCT model.

Three separate defects were found feeding the same table. They must not be conflated:

| | Defect | CUs hit | Allowance movement | Delivered? |
|---|---|---|---|---|
| **A** | `_coerce_days` strips the decimal point | Utah Community FCU, SCI FCU | **$4,320.42** (SCI only; Utah $0) | **Yes** — SCI Jun-26 |
| **B** | DQ derivation ignores `chargeoff_exclude_column` | Mountain CU | part of $195,500.23 | **Yes** — Mountain Mar-26 + Jun-26 |
| **C** | DQ derivation sums an undated, stale extract into the snapshot | Mountain CU | part of $195,500.23 | **Yes** — Mountain Mar-26 + Jun-26 |

**The money is in B and C, not in A.** Mountain CU is over-reserved by **$195,500.23** on
its 2026-06-30 report (11.1% of its $1,764,129.05 pooled allowance) and by an estimated
**$192,839.67** on 2026-03-31. Mountain's extract parses as `int64` and is completely
untouched by `_coerce_days` — its inflation is the charge-off back-book being counted as
delinquent, plus a stale April upload being added to the June snapshot.

**Dollar range of allowance movement, per CU, all over-reserving:**

| CU | Snapshot | Pooled allowance as delivered | Corrected | Over-reserve | Cause |
|---|---|---:|---:|---:|---|
| Mountain CU | 2026-06-30 | $1,764,129.05 | $1,568,628.81 | **$195,500.23** | B + C |
| Mountain CU | 2026-03-31 | $1,737,994.82 | $1,545,155.15 | **$192,839.67** *(est.)* | B + C |
| SCI FCU | 2026-06-30 | $56,781.65 | $52,461.23 | **$4,320.42** | **A** |
| Utah Community FCU | 2026-06-30 | unchanged | unchanged | **$0.00** | A (inert) |
| Nucor Emp CU | 2026-06-30 | unchanged | unchanged | **$0.00** | none — data correct |
| Lanco FCU | 2026-06-30 | n/a | n/a | **$0.00** | none — data correct |

The Mar-26 Mountain figure is marked *(est.)* because the correction applies the
June-derived corrected DQ to a March report; the exact March figure needs a March extract.
The direction and order of magnitude are certain, the cents are not.

**Has it reached clients? Yes.** Mountain CU's 2026-03-31 and 2026-06-30 Vizo models and
SCI FCU's 2026-06-30 Vizo model all carry inflated `Delinquency Score` values on the
`Env Factor by Pool` tab and inflated allowance dollars on `ACL Env by Pool Mgmt Adj`.
Utah's two workbooks carry a wrong DQ% *table* on `Display CO-Recov-DQ` but a correct
allowance.

**Scope is narrow and bounded.** Only 5 credit unions have any extract-derived DQ rows at
all, at exactly one snapshot date (2026-06-30). Every other row in
`loan_code_delinquency_history` — 17,620 of 17,846 — comes from the NCUA 5300 backfill and
never touches `_coerce_days`.

---

## 1. Which credit unions are affected (question 1)

### Method

The bug bites only when the days-delinquent column parses as `float64` — which pandas does
whenever any cell in the column is blank. So the verdict is per-CU *and* per-extract-file,
and it is empirically testable rather than inferable from config.

For all 38 live configs, each CU's extract(s) for its latest snapshot were resolved with
the pipeline's own `_resolve_extract_path`, read with the same pandas call
`_derive_snapshot_dq_from_extracts` uses, and every row scored twice: once through the
defective `dq_extract_parser._coerce_days`, once through the known-good
`dq_migration_split._coerce_days`. A CU is AFFECTED iff at least one loan changes
delinquency status between the two.

The harness reproduces the stored database totals **exactly** for all 5 CUs that have
extract-derived rows, which is what validates it as a faithful replica of the production
path.

### Verdicts

**AFFECTED — 2 client CUs, 1 test config (3 total)**

| Config | dtype | Stored / as-parsed DQ | True DQ | Divergent loans |
|---|---|---:|---:|---:|
| `utah_community_fcu` | `float64` | $86,351,866.11 | $12,108,522.68 | 2,734 |
| `sci_fcu` | `float64` (CC extract) + `int64` | $100,342.47 | $34,332.72 | 8 |
| `test_nucor_emp_cu` *(test config, not a client)* | `float64` + `int64` | $184,230.85 | $172,470.75 | 11 |

Utah's single extract `2026Q2_tct_loan_v2.csv` has 144,429 blank cells out of 157,253 rows,
which is what makes the column `float64`. Observed coercions from that file:
`6.0 -> 60`, `16.0 -> 160`, `21.0 -> 210`, `36.0 -> 360`.

SCI is affected through one of its two extracts only —
`Credit Card_CECL Report_6-30-2026v.2.xlsx`, `float64` with 2 blanks in 192 rows,
`35.0 -> 350` on 8 credit-card loans. Its main Aires extract is `int64` and clean.

**CLEAN — 21 measured, 0 divergent loans**

`bridgeton` (`object`), `bridgeton_onized_fcu`, `census_fcu`, `census_fcu_scratch`,
`central_keystone_fcu`, `central_susquehanna_comm_fcu`, `cottonwood_fcu`,
`emergency_responders_cu`, `first_area_fcu`, `franklin`, `lanco_fcu`,
`mcdowell_cornerstone_cu`, `mountain_cu`, `nova_cu`, `nucor_emp_cu`, `ontario`,
`shuford_fcu` (`int64`+`object`), `tcp_cu`, `test_nova_cu`, `united_community_fcu`,
`wssc_fcu` — all `int64` (or `object`) on every resolved extract.

**Mountain CU, Lanco FCU and Nucor Emp CU are in this CLEAN list.** All three have
extract-derived rows in the database, and all three parsed identically under both
coercions. Their stored DQ figures are not affected by `_coerce_days`. Mountain is
nonetheless wrong, for the unrelated reasons in §3.

**STRUCTURALLY CLEAN — 6 configs, derivation cannot run**

`curis_palmetto_health_cu`, `maple`, `ontario_public_employees_fcu`, `sample`,
`tongass_fcu`, `wnc_community_cu` — no `days_delinquent` mapping on any extract, so
`_derive_snapshot_dq_from_extracts` returns before reading a file.

**CANNOT DETERMINE — 8 configs**

| Config | Why |
|---|---|
| `credit_union_of_richmond` | 4 extracts configured, none resolved for 2026-06-30 |
| `destinations_cu` | 5 extracts configured, none resolved for 2026-06-30 |
| `erie_fcu` | extract is `TCT.EXPORT_*.txt`; pandas cannot infer an engine |
| `honolulu_fire_department_fcu` | extract is UTF-16; read fails under the pipeline's own call |
| `jackson_river_community_cu` | mapped column `col_X` is absent from the resolved file |
| `client`, `greensboro_cu`, `howard_county_education_fcu` | no `monthly_loan_data` snapshot — never run |

None of these 8 has any extract-derived row in `loan_code_delinquency_history`, so none has
corrupted *stored* data today. The risk is prospective: if one is run and its column
happens to be `float64`, it will corrupt on the spot. The three read failures are worth
noting on their own — those CUs' DQ derivation is silently doing nothing.

### A separate mapping-quality problem, found in passing

Three configs map `days_delinquent` to a column that is not days delinquent:

- `destinations_cu`: `Interest Rate` (and, on another extract, `Delinquent Balance`)
- `shuford_fcu` VISA extract: `Account Number`
- `jackson_river_community_cu`: `col_X`, which does not exist in the file

`06_dq_derivation.md` independently flagged the Destinations mapping. These produce
garbage DQ regardless of the coercion, and no coercion fix will help them.

---

## 2. Which quarters and snapshots are affected (question 2)

**Exactly one: 2026-06-30.** There is no history to unwind.

`loan_code_delinquency_history` holds 17,846 rows across 29 credit unions spanning
2017-12-31 to 2026-06-30. Grouped by `source`:

| Source prefix | Rows | CUs | Date range |
|---|---:|---:|---|
| `5300DQ:*` | 17,620 | 29 | 2017-12-31 to 2026-06-30 |
| `loan_extract_report` | **226** | **5** | **2026-06-30 only** |

Only the 226 `loan_extract_report` rows ever passed through `_coerce_days`. The 17,620
5300 rows come from the NCUA Solr backfill and are untouched by this defect.

| CU | Rows | Stored DQ | `updated_at` |
|---|---:|---:|---|
| Lanco FCU | 37 | $641,407.93 | 2026-07-30 09:12:15 |
| Utah Community FCU | 71 | $86,351,866.11 | 2026-08-04 16:32:20 |
| SCI FCU | 31 | $100,342.47 | 2026-08-14 16:45:56 |
| Nucor Emp CU | 17 | $226,846.83 | 2026-08-18 13:28:31 |
| Mountain CU | 70 | $15,846,547.03 | 2026-08-31 13:19:03 |

**The behaviour change point is the introduction of `_derive_snapshot_dq_from_extracts`**
(`generate_report.py:12079`, called at `:12744`). Before it existed, DQ history came only
from the 5300 backfill and the wizard. The earliest corrupted write is Lanco on
2026-07-30 - and Lanco's data is in fact correct, so the earliest *harmful* write is
Utah on 2026-08-04.

The wizard's own path (`dq_extract_parser.process_files`, reached from
`cecl_ui/routes/setup.py:9838`) writes `source='loan_extract:<filename>'`. **There are zero
such rows in the database** - that path has never successfully written DQ history, so it
contributes nothing to the blast radius despite calling the same defective helper.

Why only one date: the guard at `generate_report.py:12122-12134` skips derivation when the
period already has non-extract rows, so the derivation only ever fills the current
reporting quarter, which the 5300 filings have not yet covered.

---

## 3. How much the allowance actually moves (question 3)

### The linkage, and where it breaks

The brief's chain is `loan_code_delinquency_history` to `warm_dq_pct` to
`_pool_dq_variance` to `dq_score` to `env_factor` to the reserve. That chain is real, but
it has a branch that decides whether a given CU's allowance moves at all:

```python
# report_vizo.py:496-499  (_pool_dq_variance)
dq = hist.get('dq_pct', {}) if hist else {}
if not dq:                                    # <-- only when empty
    _imp = hist.get('impaired', {}) if hist else {}
    dq = _imp.get('warm_dq_pct', {})
```

`_pool_dq_variance` prefers `hist['dq_pct']` and reads `warm_dq_pct` **only when
`hist['dq_pct']` is entirely empty**. The database overlay at
`generate_report.py:12754-12760` writes into `imp['warm_dq_pct']` - and it runs *after*
`_overlay_warm_history_into_hist` (called at `:12444`) has already populated
`hist['dq_pct']` from the WARM. The overlay is never re-merged into `hist['dq_pct']`.

**Consequence - the single most important finding for question 3:**

- A CU **with** a legacy WARM supplying `warm_dq_pct` has a populated `hist['dq_pct']`, so
  `_pool_dq_variance` never looks at the corrupted overlay. Its `Display CO-Recov-DQ` tab
  shows the wrong DQ% (that tab reads `warm_dq_pct` directly), but its **env factor and
  allowance are correct**.
- A CU **without** a WARM has an empty `hist['dq_pct']`, falls through to the corrupted
  `warm_dq_pct`, and its **env factor and allowance are wrong**.

Utah Community FCU is in the first group. Mountain CU, SCI FCU and Nucor Emp CU are in the
second.

### Utah Community FCU - $86.35M of wrong data, $0 of allowance movement

Verified in the delivered workbooks, not inferred. `Env Factor by Pool`, in both the Vizo
model (generated 2026-08-04 16:25:54, *before* the bad rows landed) and the TCT model
(generated 2026-08-04 16:33:21, *after* them):

```
Portfolio Segment      NCC Score   DQ Variance   DQ Score   ES Score   Env Factor
Mortgage Real Estate       0        0.000974        0         -0.01      -0.01
Indirect Auto             0.02     -0.000858        0         -0.01       0.01
Direct Auto                0        0.000454        0         -0.01      -0.01
Unsecured                  0        0.002143        0         -0.01      -0.01
Other Consumer             0       -0.002532        0         -0.01      -0.01
Commercial                 0       -0.000157        0         -0.01      -0.01
Construction               0        0               0         -0.01      -0.01
```

**`DQ Score` is 0 for every pool in both models.** Meanwhile the same workbooks'
`Display CO-Recov-DQ` tab shows the corrupted series - Indirect Auto `YTD 2026` at 5.315%
against a 2025 value of 1.085%, Direct Auto at 4.160%, Commercial at 2.990%.

Had the corruption reached the env factor it would have been severe. Recomputing
`_pool_dq_variance` against the overlay both ways gives:

| Pool | DQ var (corrupt) | DQ var (correct) | Score corrupt | Score correct | delta env |
|---|---:|---:|---:|---:|---:|
| Direct Auto | 3.1203% | 0.3366% | 12 | 0 | -0.12 |
| Indirect Auto | 3.3552% | -0.2493% | 12 | 0 | -0.12 |
| Commercial | 2.4913% | 0.1710% | 4 | 0 | -0.04 |
| Other Consumer | 2.2080% | 0.0000% | 4 | 0 | -0.04 |
| Construction | 1.6789% | 0.0000% | 2.5 | 0 | -0.025 |
| Unsecured | 1.3914% | 0.0169% | 1.5 | 0 | -0.015 |
| **Mortgage Real Estate** | **0.4286%** | **-0.0091%** | **0** | **0** | **0.000** |

Mortgage Real Estate is the banded-threshold case the brief anticipated: a 47x error in the
input produces **zero** movement, because both values land inside the `-0.50% to 0.49%`
band. That is a real result, not a gap.

**Utah's exposure is presentational, not financial.** The DQ% table in two delivered
workbooks is wrong by up to 7x. The allowance is right.

### SCI FCU - the one place `_coerce_days` moved money

SCI has no WARM, so it takes the vulnerable branch, and its corrupted credit-card DQ
(11.373% against a true 0.136%) flows straight through.

Computed DQ variance for Credit Cards is **8.4671%**, matching the delivered workbook's
`0.08467075` to five decimal places - end-to-end confirmation that the replica is faithful.

From `ACL Env by Pool Mgmt Adj` in
`2026-06-30_CECL_Migration_SCI_FCU_Vizo_Model.xlsx`:

| Pool | Allow before | Env factor | Env allow | delta env | Corrected env allow | **delta $** |
|---|---:|---:|---:|---:|---:|---:|
| Unsecured Loans | $10,484.02 | -0.0200 | -$209.68 | 0 | -$209.68 | $0.00 |
| HELOCS | $17.23 | -0.0075 | -$0.13 | 0 | -$0.13 | $0.00 |
| New Vehicles | $10,000.63 | 0.0400 | $400.03 | 0 | $400.03 | $0.00 |
| Used Vehicles | $8,714.88 | 0.0000 | $0.00 | 0 | $0.00 | $0.00 |
| Other Collateral | $164.70 | 0.0100 | $1.65 | 0 | $1.65 | $0.00 |
| **Credit Cards** | **$21,602.11** | **0.2000** | **$4,320.42** | **-0.20** | **$0.00** | **-$4,320.42** |
| Negative Share Draft | $1,285.80 | 0.0000 | $0.00 | 0 | $0.00 | $0.00 |
| **Total** | **$52,269.37** | | **$4,512.28** | | **$191.86** | **-$4,320.42** |

Pooled allowance as delivered **$56,781.65**, corrected **$52,461.23**. SCI is
**over-reserved by $4,320.42, or 7.6% of its pooled allowance**, and 6 of its 7 pools show
zero movement.

### Mountain CU - the largest movement, and not from `_coerce_days`

Mountain's extract is `int64` with zero divergent loans. Its inflation has two other
causes, both in `_derive_snapshot_dq_from_extracts`.

**Defect B - the derivation ignores `chargeoff_exclude_column`.** Mountain's config sets
`chargeoff_exclude_column: C/O Amount` on both extracts, and `import_data.py:1671` honours
it when building `monthly_loan_data`. `_derive_snapshot_dq_from_extracts` never reads the
key. Measured on `Loan File - 6-30-26 .xlsx`:

```
dq, all rows                    $14,874,540.36   (n=4,146)
dq, excluding charged-off rows      $987,947.81   (n=  151)
dq, from charged-off rows only  $13,886,592.55   (n=3,995)
```

**93.4% of Mountain's recorded June delinquency is its charged-off back-book** - already
written off, counted a second time as delinquent.

**Defect C - an undated stale extract is summed into the snapshot.** Mountain resolves two
files for 2026-06-30. `_neg_share_period_from_path` returns `(2026, 6)` for
`Loan File - 6-30-26 .xlsx` but `None` for `April Loan File - Upload.xlsx`, and the period
guard at `generate_report.py:12180-12182` skips only when the period is *known and
different*. April's $972,006.67 is therefore added to June's total.

Stored $15,846,547.03 decomposes exactly as $987,947.81 true June DQ + $13,886,592.55
charge-off back-book + $972,006.67 stale April extract. The brief's "$15.85M recorded vs
$0.99M true" is confirmed to the dollar.

Correcting all three (correct coercion, charge-off exclusion, strict period filter) yields
58 loan codes totalling $987,947.81, and this movement on
`2026-06-30_CECL_Migration_Mountain_CU_Vizo_Model.xlsx`:

| Pool | Allow before | Env factor | Env allow | delta env | **delta $** |
|---|---:|---:|---:|---:|---:|
| Consumer Unsecured | $604,295.47 | 0.1600 | $96,687.27 | -0.2000 | **-$120,859.09** |
| Consumer Indirect Auto-Used Auto | $426,881.36 | 0.1100 | $46,956.95 | -0.1125 | **-$48,024.15** |
| Consumer Auto Loan-Used Auto | $328,025.02 | 0.0300 | $9,840.75 | -0.0400 | **-$13,121.00** |
| Credit Cards | $154,996.76 | 0.0400 | $6,199.87 | -0.0800 | **-$12,399.74** |
| Consumer Indirect Auto-New Auto | $42,213.25 | 0.0050 | $211.07 | -0.0150 | -$633.20 |
| Consumer Secured | $15,814.72 | 0.0150 | $237.22 | -0.0250 | -$395.37 |
| Consumer Auto Loan-New Auto | $1,691.95 | 0.0300 | $50.76 | -0.0400 | -$67.68 |
| Real Estate | $30,329.94 | -0.0100 | -$303.30 | 0.0000 | $0.00 |
| **Total** | **$1,604,248.45** | | **$159,880.59** | | **-$195,500.23** |

Pooled allowance as delivered **$1,764,129.05**, corrected **$1,568,628.81** -
**over-reserved by $195,500.23, or 11.1%**. Consumer Unsecured alone accounts for
$120,859.09, its DQ variance falling from 11.1978% (score 20, the top band) to -0.1479%
(score 0).

Real Estate is a second zero-movement case: 0.3127% to 0.0972%, both inside the same band.

**Mountain's 2026-03-31 report carries the same contamination.** It was regenerated on
2026-08-31 and its `Env Factor by Pool` tab holds identical DQ variances (Consumer
Unsecured 0.111978, Indirect Used 0.034601) - because `_load_dq_history_from_db` buckets
each year to that year's *latest* date, so a March report picks up June-derived rows.
Applying the same deltas gives **$192,839.67** over-reserved against a delivered
$1,737,994.82. This one is an estimate: the correction uses June's corrected DQ, where a
true restatement needs a March extract.

### Nucor Emp CU and Lanco FCU - measured, and correct

Both parse `int64` with zero divergent loans, neither sets `chargeoff_exclude_column`, and
each resolves a single dated extract. Nucor's env-factor contribution is small and
legitimate. No movement, no restatement.

---

## 4. Have affected numbers reached clients (question 4)

**Yes - three workbook pairs carry an inflated allowance, and two more carry an inflated
DQ table with a correct allowance.**

Timing is decisive here, and the database's `updated_at` column dates every corrupted
write precisely. Comparing it against workbook mtimes:

| Workbook | mtime | DQ rows written | Allowance affected? |
|---|---|---|---|
| `2026-06-30_..._Mountain_CU_Vizo_Model.xlsx` | 2026-08-31 13:19 | 2026-08-31 13:19:03 | **Yes - $195,500.23** |
| `2026-03-31_..._Mountain_CU_Vizo_Model.xlsx` | 2026-08-31 (regen) | inherits June rows | **Yes - ~$192,839.67** |
| `2026-06-30_..._SCI_FCU_Vizo_Model.xlsx` | 2026-08-14 16:45 | 2026-08-14 16:45:56 | **Yes - $4,320.42** |
| `2026-06-30_..._Utah_Community_FCU_Vizo_Model.xlsx` | 2026-08-04 16:25 | 2026-08-04 16:32:20 | No - DQ table only |
| `2026-06-30_..._Utah_Community_FCU_TCT_Model.xlsx` | 2026-08-04 16:33 | 2026-08-04 16:32:20 | No - DQ table only |

The Utah pair is the interesting case. The Vizo model was written seven minutes *before*
the corrupted rows existed, the TCT model one minute *after* - yet both show the identical
inflated `Display CO-Recov-DQ` series and the identical clean `Env Factor by Pool`. The
Vizo run derived the same bad numbers in memory and rendered them; the row-level
`updated_at` reflects the later TCT run's upsert. Either way the conclusion holds: Utah's
DQ table is wrong in both delivered models, and Utah's allowance is right in both.

Mountain's `Supplemental` and TCT variants, where present, are generated from the same
`hist` in the same run and carry the same contamination.

**Not affected:** every other CU's reports. Nucor Emp CU and Lanco FCU have extract-derived
rows that are correct. No workbook for any 5300-sourced CU touches this path.

**One caveat I could not close.** I confirmed contamination by reading the `Env Factor by
Pool` and `ACL Env by Pool Mgmt Adj` tabs of the workbooks in
`Z:\Shared\TCT Files\CECL - CM Files\Reports` and `...\Reports\older reports`. I did not
independently verify which of those files were actually *transmitted* to each credit union,
or whether a corrected version was sent afterwards by hand. That distribution record is
outside the repo and the shares I examined. The workbooks exist in the delivery folder in
their contaminated form.

---

## 5. The correct fix and its own blast radius (question 5)

### Is coercing `60.0 -> 60` the whole fix?

For defect A, yes - but it must be done by short-circuiting on numeric types, not by adding
`.` to the character class. `days_delinquent` is a count of days; a fractional value is
never meaningful, and every caller compares it against an integer threshold. There is no
upstream input where the decimal carries information.

**The repo already contains a correct implementation**, written specifically because of
this bug. `dq_migration_split.py:42` carries the reference version and names the defect in
its own docstring:

```
Deliberately *not* ``dq_extract_parser._coerce_days``: that helper strips
every non-digit character, which turns a pandas float cell (``60.0``)
into ``600``.  Loan extracts routinely come back as float64 columns
whenever a single row is blank, so a digits-only strip is unsafe here.
```

The fix is to make the shared helper behave like the reference one. Proposed diff
(**NOT applied**):

```diff
--- a/cecl_ui/services/dq_extract_parser.py
+++ b/cecl_ui/services/dq_extract_parser.py
@@
 def _coerce_days(val: Any) -> int | None:
     """Parse a 'days delinquent' cell to an int. Empty/non-numeric -> None."""
     if val is None:
         return None
+    # Numeric cells short-circuit. pandas yields float64 for the whole
+    # column as soon as one row is blank, so ``60.0`` arrives here as a
+    # float and must never be stringified and digit-stripped -- that
+    # turns 60.0 into 600 and 6.0 into 60.
+    if isinstance(val, bool):
+        return None
+    if isinstance(val, (int, float)):
+        try:
+            f = float(val)
+        except (TypeError, ValueError):
+            return None
+        return None if math.isnan(f) or math.isinf(f) else int(f)
     s = str(val).strip()
-    if not s:
+    if not s or s.lower() in ('nan', 'none', 'null', '-', 'na', 'n/a'):
         return None
-    # Strip non-digit/minus characters (handles "30 days", "30+", etc.)
-    cleaned = re.sub(r"[^0-9\-]", "", s)
-    if not cleaned or cleaned == "-":
-        return None
+    # Plain numeric string, including a float repr such as '60.0'.
     try:
-        return int(cleaned)
+        return int(float(s.replace(',', '').replace('$', '')))
     except ValueError:
-        try:
-            return int(float(s))
-        except ValueError:
-            return None
+        pass
+    # Embedded number: '30 days', '30+', '60 DPD'. Keeps the decimal point.
+    m = re.search(r'-?\d+(?:\.\d+)?', s)
+    if not m:
+        return None
+    try:
+        return int(float(m.group()))
+    except ValueError:
+        return None
```

`import math` must be added at the top of the module.

Truncation vs rounding: `int(float(x))` truncates, so `59.9 -> 59`. That matches the
reference implementation's behaviour once thresholded and is the conservative choice for a
day count. No observed extract contains fractional days.

### Blast radius of the fix itself

`_coerce_days` has exactly **two** callers:

| Caller | Path | Effect of the fix |
|---|---|---|
| `dq_extract_parser.rollup_dataframe:148` | wizard "Historical DQ" upload | None today - zero `loan_extract:*` rows exist |
| `generate_report.py:12227` | `_derive_snapshot_dq_from_extracts` | Corrects Utah and SCI on next run |

The fix is strictly narrowing: it can only ever *reduce* a parsed value or turn it into
`None`. It cannot newly classify a current loan as delinquent. For the 21 CLEAN CUs the
output is bit-identical, because `int64` columns already short-circuit correctly through
the existing `int(cleaned)` path - verified empirically across 29 configs, where 26 showed
zero divergent loans.

A regression test belongs alongside it, asserting the brief's own table:
`'60' -> 60`, `60.0 -> 60`, `6.0 -> 6`, `'90' -> 90`, `90.0 -> 90`, `30.0 -> 30`,
`'30 days' -> 30`, `'' -> None`, `nan -> None`.

### Sweep for the same pattern elsewhere

Every strip-non-digit coercion in the repo was located and classified. **`_coerce_days` is
the only defective one.**

| Site | Function | Verdict |
|---|---|---|
| `cecl_ui/services/dq_extract_parser.py:42` | `_coerce_days` | **DEFECTIVE** - the subject of this document |
| `cecl_ui/services/impaired_parser.py:47` | `_digits` | **SAFE** - explicitly strips a trailing `.0` *before* joining digits, and is used for account identity, not magnitude |
| `cecl_ui/services/auto_setup.py:644` | `_norm_period` | **SAFE** - normalises a `YYYYMM` period string; decimals are meaningless and floats never reach it |
| `cecl_ui/routes/setup.py:1074` | charter number | **SAFE** - identity string |
| `cecl_ui/services/warm_parser.py:870` | charter number | **SAFE** - identity string |
| `generate_report.py:7877`, `report_vizo.py:1633` | pool-name slug | **SAFE** - `[^\w\s-]` on text, not numeric |

`impaired_parser._digits` is worth noting: the codebase already knew about the trailing
`.0` trap and guarded against it in one place while missing it in another.

### The two defects the fix does *not* address

Defect A is the cheapest of the three and the smallest in dollars. **A `_coerce_days` fix
alone leaves $195,500.23 of Mountain over-reserving in place.** Both remaining defects are
in `_derive_snapshot_dq_from_extracts`:

**Defect B** - honour `chargeoff_exclude_column`. The key is per-extract in the config
(`mountain_cu.yaml`, both entries) and is read today only by `import_data.py:1671` and
promoted at `:1964`. The derivation loop at `generate_report.py:12225-12232` should skip
rows whose charge-off column is non-zero, mirroring `import_data.py:1671-1683`. Only
`mountain_cu.yaml` and `nova_cu.yaml` set the key, so the change touches two CUs.

**Defect C** - tighten the period guard. `generate_report.py:12180-12182` currently reads
`if per is not None and per != target: continue`, which admits any file whose period cannot
be parsed from its name. For a snapshot-specific derivation the safe default is the
opposite: skip files whose period is unknown when at least one file with a matching known
period was found. This is what lets `April Loan File - Upload.xlsx` contribute to a June
snapshot.

I would fix all three together. B and C are where the money is, and shipping A alone would
produce a "fixed" report that is still wrong by $195k.

---

## 6. Remediation shape and cost (question 6)

**No re-import is needed. No backfill is needed. Delete 226 rows and regenerate five
reports.**

The corrupted data is confined to `loan_code_delinquency_history`, at one date, for five
CUs, from one source tag. Nothing downstream persists it: `monthly_loan_data` carries no
`days_delinquent` column at all (confirmed in `07_chargeoff_feasibility.md`), and
`warm_dq_pct` is rebuilt from the table on every run. The table is the only store.

`_derive_snapshot_dq_from_extracts` is self-healing once the code is fixed: its guard skips
the period only when non-extract rows exist, so deleting the `loan_extract_report` rows
lets the next report run rewrite them correctly.

### Steps

| # | Step | Cost |
|---|---|---|
| 1 | Fix `_coerce_days` (defect A) + regression test | 0.5 day |
| 2 | Honour `chargeoff_exclude_column` in the derivation (defect B) | 0.5 day |
| 3 | Tighten the unknown-period guard (defect C) | 0.5 day |
| 4 | `DELETE FROM loan_code_delinquency_history WHERE source = 'loan_extract_report'` (226 rows, all 2026-06-30) | minutes |
| 5 | Regenerate Mountain (Mar-26, Jun-26), SCI (Jun-26), Utah (Jun-26), Nucor + Lanco (Jun-26) | 0.5 day |
| 6 | Verify the six regenerated workbooks against the figures in this document | 0.5 day |
| | **Total** | **~2.5 days** |

Step 4 is safe to scope by `source` - it cannot touch a 5300 row or a manually entered one.
Take a table snapshot first regardless.

Nucor and Lanco are included in step 5 only because their rows are deleted in step 4 and
should be rewritten; their numbers will not change.

### What has to be restated to clients

| CU | Report | Action |
|---|---|---|
| Mountain CU | 2026-06-30 | **Restate** - allowance falls $195,500.23 (11.1%) |
| Mountain CU | 2026-03-31 | **Restate** - allowance falls ~$192,839.67 |
| SCI FCU | 2026-06-30 | **Restate** - allowance falls $4,320.42 (7.6%) |
| Utah Community FCU | 2026-06-30 Vizo + TCT | **Reissue, not restate** - DQ% table corrected, allowance unchanged |
| Nucor Emp CU, Lanco FCU | 2026-06-30 | No action |

The Mountain and SCI restatements are downward: both credit unions are holding more reserve
than the model, correctly run, calls for. That is the safer direction to have erred, but it
is still a misstatement of a reported figure.

Utah is a reissue rather than a restatement - no reported financial figure changes, only a
presentational table. Worth confirming with whoever owns the client relationship whether a
DQ table that reads 7x high warrants proactive notice.

### Prospective risk

The 8 CANNOT DETERMINE configs in section 1 are the live exposure. Any of them could parse
as `float64` on its next run and corrupt on the spot. Fixing defect A closes that risk
permanently, which is the argument for shipping it even though its measured dollar impact
is the smallest of the three.

Separately, the three CUs whose extracts fail to read at all
(`erie_fcu`, `honolulu_fire_department_fcu`, `jackson_river_community_cu`) are silently
deriving no DQ. That is a different bug and should be raised on its own.

---

### Reproduction

```bash
# 1. The blast radius is 226 rows at one date for five CUs.
CECL_WORKSPACE_ROOT="Z:\Shared\TCT Files\CECL - CM Files" python - <<'PY'
import sys; sys.path.insert(0, r"C:\Dev\CECL")
from sqlalchemy import create_engine, text
from cecl_credentials import get_database_url
with create_engine(get_database_url()).begin() as c:
    for r in c.execute(text("""
        SELECT cu, as_of_date, count(*), sum(dq_amount), max(updated_at)
        FROM loan_code_delinquency_history
        WHERE source = 'loan_extract_report'
        GROUP BY cu, as_of_date ORDER BY 5""")):
        print(r)
PY
```

```bash
# 2. The defect, against the reference implementation in the same repo.
python - <<'PY'
import sys; sys.path.insert(0, r"C:\Dev\CECL")
from cecl_ui.services.dq_extract_parser import _coerce_days as bad
from dq_migration_split import _coerce_days as good
for v in ('60', 60.0, 6.0, '90', 90.0, 30.0, '30 days'):
    print(f"{v!r:12} bad={bad(v)!r:8} good={good(v)!r}")
PY
# 60.0 -> 600, 6.0 -> 60, 90.0 -> 900, 30.0 -> 300
```

```bash
# 3. Utah's true vs recorded DQ, straight off the extract.
python - <<'PY'
import pandas as pd, sys; sys.path.insert(0, r"C:\Dev\CECL")
from cecl_ui.services.dq_extract_parser import _coerce_days as bad
from dq_migration_split import _coerce_days as good
from cecl_ui.services import extract_hist_processor as ehp
df = pd.read_csv(r"<utah data_directory>\2026Q2_tct_loan_v2.csv")
b = df['BOOK_BALANCE'].map(ehp._clean_balance)
d = df['DAYS_PAST_DUE']
print('dtype', d.dtype, 'blanks', int(d.isna().sum()), 'of', len(df))
print('recorded ', b[d.map(lambda x:(bad(x)  or 0) >= 60)].sum())   # 86,351,866.11
print('true     ', b[d.map(lambda x:(good(x) or 0) >= 60)].sum())   # 12,108,522.68
PY
```

```bash
# 4. Mountain's contamination is charge-offs, not coercion.
#    Both coercions agree; excluding 'C/O Amount' rows is what moves the number.
#    all $14,874,540.36 / excl charge-offs $987,947.81 / from charge-offs $13,886,592.55
```
