# Moving the Migration PDF off Excel

Decision taken 2026-08-31.

**Goal:** the delivered PDF is rendered directly from report data. The .xlsx
keeps being generated, but is demoted off the deliverable path — a workbook
defect becomes a cosmetic issue, not a corrupt client deliverable.

**Explicitly in scope:** CECL Migration (Vizo) reports.
**Explicitly out of scope:** SCALE. It is genuinely template- and
formula-driven (`cecl_ui/services/scale/`, `excel_recalc.py` spawns Excel to
compute values) and works as it should. Leave it alone.

**Fidelity target:** *not* pixel-parity with the Excel output. Same numbers,
same information, better-looking. Excel's constraints are considered part of
the problem being solved.

---

## Why this is worth doing

Not because Excel is unreliable — that was the initial hypothesis and the
evidence contradicts it. Every failure on 2026-08-31 came from our own
Python/OOXML layer (see `../migration_layout_redesign.md`, "Regression found by
the first real pipeline run"). Excel was the only component that noticed the
file was malformed.

The real case is:

- The delivered workbook is **95.9% literal values** — 5,438 literal cells vs
  233 formulas, and all 233 formulas are in tabs added last week. Excel
  computes nothing. It is a rendering surface, not a calculation engine.
- Rendering through openpyxl costs **~966 lines of raw ElementTree surgery**
  (`report_vizo.py:4825-5790`, four `patch_*` entry points and their helpers —
  every `ET.` call site in the repo is in that span) to make the output look
  right. That layer is where the defects live, and it is nearly untestable.
- Layout iteration is slow and indirect: a change has to survive openpyxl, then
  a patcher that may silently override it. Two rounds of chart tuning were lost
  that way.

## Phase A — integrity gate (DONE, 2026-08-31)

Independent of everything below, and worth having regardless.

`report_integrity.py` structurally validates every generated workbook, with no
Excel required: all XML parts well-formed, every relationship target present,
sheet list non-empty, and — the check that catches the 2026-08-31 defect — every
drawing's chart reference is a `c:chart` in the chart namespace whose `r:id`
resolves to a chart part that exists.

Wired into `generate_report.py` immediately after the chart patchers. A failing
report is logged as a failure and named in a loud end-of-run summary. Also
runnable standalone for CI:

    python report_integrity.py <file.xlsx> [--excel]

`--excel` adds the stronger "can Excel actually open it" check where pywin32 is
available; it reports *skipped* rather than failing when it is not, matching
`excel_recalc.py`'s best-effort convention.

Verified against a synthetic reproduction of the real corruption: the gate fails
it with an explanatory message and a non-zero exit.

## Phase D — sequencing

The work does **not** parallelise cleanly from a standing start. Two things
must exist first or parallel work diverges and has to be redone:

1. **The data contract.** The renderer must consume a `ReportData` structure
   rather than a workbook. Until its shape is fixed, every tab renderer is
   guessing at its inputs.
2. **The design system.** Typography, palette, table styles, page furniture.
   Without it, six people build six visually different reports — and "best
   looking report" is the actual goal here, not a nice-to-have.

So:

    Step 1 (serial)    contract + design system          <- blocking
    Step 2 (parallel)  tab renderers, ~6 tracks
    Step 3 (serial)    assembly, page furniture, regression baseline

Running now, ahead of Step 1, is the research that produces the contract —
three parallel read-only investigations:

| # | Question | Output |
|---|---|---|
| 01 | What does `cecl_report_web/` already cover, and how coupled is it to the workbook? | `01_prototype_audit.md` |
| 02 | What data is available at report-build time, and which tabs read *other tabs* rather than data? | `02_data_inventory.md` |
| 03 | What are the 40 charts really — how many distinct archetypes, and which are hard in SVG? | `03_chart_inventory.md` |

Item 02's "tabs that read other tabs" is the one that can wreck the estimate:
anything deriving its values from another sheet needs restructuring, not
porting. Item 03 is the schedule risk; the plan is to spike the 2–3 worst
charts before committing to a date.

## Known constraint

`cecl_report_web/from_workbook.py` reads the **generated .xlsx** and re-renders
it. Keeping that shape would leave openpyxl and the whole patcher layer on the
critical path, and add a workbook-parsing layer on top. Re-pointing it at the data model
is the whole point of the exercise.

---

# Research complete (2026-08-31) — consolidated plan

Three investigations landed: `01_prototype_audit.md`, `02_data_inventory.md`,
`03_chart_inventory.md`. What they change:

## The shape of the problem is a hub, not a list of tabs

13 of 17 visible tabs are already pure functions of `(df, config, grades, hist)`.
The difficulty is concentrated in one place: **`ACL Env by Pool Mgmt Adj` is an
undeclared hub**, and four other tabs recover their data by screen-scraping the
worksheet it just built.

- `_parse_acl_layout` re-reads the freshly built sheet, matching column-A label
  text to recover **row numbers**; pool headers are detected as "a label whose
  Balance column is empty".
- 16 sites emit cross-sheet formulas like `='ACL Env by Pool Mgmt Adj'!K42`.
  On ACL Summary even the pool *name* is a formula.
- Two builders read cell **values** back out and let control flow depend on
  them: `_sheet_mgmt_adj_summary` decides which grade rows exist, and
  `_sheet_impaired_loans` filters its pool list, by reading cells.
- `change_analysis._parse_acl_sheet` applies the same screen-scrape to the
  **prior quarter's .xlsx on disk**. Prior-period numbers exist nowhere else —
  no DB, no serialized store.
- `_sheet_acl_reserve` stashes `_computed_pooled_total_allow` onto
  `hist['impaired']` so `Impr Deter` agrees with it, which forces the
  out-of-order build and the `wb.move_sheet` repair.

Business rules also live in the paint layer: migration direction is decoded
from **Excel fill hex**, and ACL subtotal-vs-detail from **font bold**. Those
rules exist nowhere as data and must be extracted from `_sheet_risk_change`.

## Step 1 is worth doing on its own merits

Extracting `AclEnvironmental` as a pure function, rewriting the four summaries
against it, and serializing it to JSON for prior-period lookup removes, today,
independent of any PDF work:

- the screen-scraping and the row-number archaeology,
- the build-order coupling and the hidden-sheet dance,
- the silent dependency on Excel having recalculated prior workbooks
  (`data_only=True` returns `None` for every formula cell of a prior report
  nobody opened in Excel — the four summary tabs are exactly such formulas).

That last one is a latent failure in today's pipeline, not just an obstacle to
the rewrite.

## Sequencing

    Step 1 (blocking, ~2 wks)   three tracks, run in parallel:
      1a  AclEnvironmental as a pure function + JSON serialization   <- critical path
      1b  design system: type, palette, table styles, page furniture
      1c  chart chassis + signed-baseline spike (A2, then A5)

    Step 2 (parallel, ~3 wks)   six tracks:
      T1  ACL Env + the four summary tabs          (depends on 1a)
      T2  Impr Deter + its four charts             (depends on 1c)
      T3  Risk Change x9 + migration-state rule extraction
      T4  Display HIst Bal, Display CO-Recov-DQ, Env Factor by Pool
      T5  Cover, Report Index, Change Analysis, narrative tabs
      T6  assembly, page furniture, pagination

    Step 3 (serial, ~1 wk)      regression baseline, parallel-run diffing

## Estimate

| Area | Hours |
|---|---|
| Charts (7 archetypes + chassis + integration) | ~100 |
| `AclEnvironmental` extraction + summary rewrite + serialization | 40-60 |
| Remaining tab renderers | 80-120 |
| Assembly, pagination, page furniture, regression re-baseline | ~40 |
| **Total** | **~260-320 h** |

≈7-8 dev-weeks for one person; ≈5-6 calendar weeks with the Step 2 tracks
genuinely parallel after the blocking phase. Assumes "same information, better
looking" — not pixel parity.

The estimate's largest threat is the **design round**, not the code. The Excel
originals are information-poor (see below), so there is no reference to
converge on; budget review cycles with Brian.

## Risks and blockers

1. ~~**`Sample Reports/` is gitignored and absent from the clone.**~~
   RESOLVED 2026-08-31. The folder does exist on the share
   (`<CECL_WORKSPACE_ROOT>/Sample Reports/`) — it was only missing from the
   *clone*, so nothing was lost. The code-required artifacts are now vendored
   into `report_assets/` (24 KB total): the narrative template, the Vizo
   `theme1.xml` extracted from the 7 MB master workbook, and the info icons.
   See `report_assets/README.md`.

   Note the repo-root `vizo_theme.xml` is a **different** theme from the
   template's `theme1.xml` and is not interchangeable with it.

   **Pending code change:** `report_vizo.py`'s three asset paths
   (`:63`, `:120-121`, `:6623-6624`) still resolve only against the workspace,
   and fall through *silently* when it is absent — a clean clone produces a
   report with no narrative tabs, no theme and no icons, and says nothing.
   They need a resolver that prefers the workspace (keeping the analyst
   workspace as the edit point for approved copy) and falls back to
   `report_assets/`. Deferred only to avoid colliding with the in-flight
   `AclEnvironmental` work in the same file.
2. **The diverging stacked bar is only correct by accident.** Excel renders the
   negative "Deteriorated" series as a tornado around zero; the prototype's
   stacked path takes `abs(val)` and accumulates, so the two series would add
   instead of oppose. Same sign bug in Net Change. A naive port ships a wrong
   chart that looks plausible. This is why 1c spikes A2 first.
3. **18 of 40 charts appear to plot all zeros** across three credit unions —
   the DQ pie and charge-off bar source ranges are empty in the delivered June
   reports for Mountain CU, Nucor Emp CU and SCI FCU. Needs its own
   investigation; if it holds, blank framed charts are shipping today.
4. **Every doughnut renders as a near-solid pie.** `report_vizo.py:1980` sets
   `dc.innerRadius = 50`, but openpyxl's attribute is `holeSize`; Python accepts
   the assignment silently and `holeSize` stays `10.0`. Verified. The intended
   design has never shipped.
5. `render_report_pdf` does **24 cold Chromium launches per report**; and
   `regression.py` screenshots HTML, not PDF, so pagination — where all our
   Excel pain has been — is untested.

## Cheap wins available now, independent of the migration

- Fix `innerRadius` -> `holeSize` (one line).
- Delete the dead keys `hist['delinquency']`, `hist['impaired']['items']` and
  the two write-only `_computed_*` stashes.
- `Report Index`'s tab list is static text, not derived from `wb.sheetnames`;
  it already omits the four tabs added this week and will keep drifting.

---

# Status checkpoint — 2026-08-31

## Done

- **Phase A, integrity gate.** `report_integrity.py` + wired into
  `generate_report.py`. Catches the drawing/chart-namespace corruption that
  shipped unopenable workbooks. Verified against a synthetic repro of the real
  defect; live on both Mountain CU runs.
- **Step 1a, `AclEnvironmental`.** `report_acl_model.py` (pure function,
  dataclasses, JSON round-trip) + `scripts/verify_acl_model.py`.
  **Independently re-run: 665 values compared, 0 mismatches**, plus 662 against
  the delivered workbook, with no writes to reports or client YAML. Additive
  only — `_sheet_acl_reserve` is untouched, so no generated report changes.
- **Vendored assets.** `report_assets/` (24 KB): narrative template, Vizo
  `theme1.xml`, info icons. `Sample Reports/` was never lost, only gitignored.

- **Step 1c, chart chassis + spike.** `cecl_report_web/chart_chassis.py`
  (~1,070 lines): scales, tick/axis engine, currency+percent formatting, real
  Calibri text measurement, label placement and collision resolution.
  Signed baseline solved via two independent accumulators plus a symmetric
  scale and a per-series `direction` flag. All 40 charts render, 0 failures.
  Preview: `chart_spike.html` / `.pdf` on the share, each spike shown beside
  its pre-spike output. Write-up in `05_chart_spike.md`.

  **Revised estimate: ~40-46 h remaining** (vs the ~51-58 h the original
  ~100 h implied for the same remainder). A2 was over-estimated and A3 came
  free. Two caveats: a new 6 h radial-chassis line (arcs / explosion / leader
  lines, needed by A4 and A6, untouched by this spike), and the 8 h
  integration line is now the largest unknown because the spike renders into a
  standalone page rather than the real templates.

- **Blank-chart investigation: CONFIRMED.** See `04_blank_charts.md`.
  Independently re-verified: Mountain CU June has all 9 Risk Change tabs with
  both the DQ and CO blocks reading literal `0` (18 of 40 charts), and
  `load_impaired_data` returns an **empty dict** for `mountain_cu` and
  `sci_fcu` while returning all four DQ/CO keys for `utah_community_fcu`.

  Root cause is a missing upstream source, not miswiring. The only producer of
  `hist['impaired']['dq_by_*'|'co_by_*']` (`generate_report.py:5176-5202`)
  requires a **legacy WARM workbook** named
  `<YYYY-MM> CECL-Migration-WARM - <CU>.xlsx` with `DQ Data Entry` / `CO Data
  Entry` tabs. No file -> `{}` -> `.get('balance', 0)` writes zeros silently.
  Wizard-onboarded clients do not produce that file, so **the share of blank
  reports grows with every new client.**

  Blast radius: 19 client-facing Vizo workbooks in the archive are all-zero;
  only 3 are healthy. `report_tct.py:2387-2398` has the identical defect.
  `patch_dq_pie_zero_labels` then deletes all four slice labels, so the output
  reads as *broken* rather than *no data*.

  Not a regression — all-zero since the oldest archived report (Jun-2025).

  **No restatement warranted:** the Risk Change matrices, Net Credit Change and
  the other charts on those tabs are correct and independently computed.

  **One tempting fix is ruled out.** A filesystem sweep found near-miss
  filenames (e.g. `2026Q1_CECL-Migration -WARM Impaired Loans - Utah Community
  FCU.xlsx`, note the space before `-WARM`), but opening one shows the same
  3 tabs as Mountain CU's upload — `Instructions`, ` Impaired Loans`,
  `Management Adjustment` — and no `DQ Data Entry`. Loosening the filename
  regex would match more 3-tab workbooks and change nothing. The determinant is
  whether a **full ~80-tab legacy WARM** exists for that CU at all; those live
  in TCT-internal `Portfolio Management` folders and wizard-onboarded CUs never
  had one. Tier 2 is therefore the only real fix.

  Note the "$15.8M delinquent" log line is a *different* feed
  (`_derive_snapshot_dq_from_extracts`, keyed by loan code) and cannot supply
  this one, which needs DQ split by migration status. Deriving it is blocked:
  `monthly_loan_data` stores no `days_delinquent` (`import_data.py:1651-1661`)
  even though every affected CU's config maps it.

## Tier 2 fix — in flight (2026-08-31)

Brian chose Tier 2 (fix it properly) over Tier 1 (suppress the empty charts).

- **Delinquency — DONE.** See `06_dq_derivation.md`. New `dq_migration_split.py`
  derives the split at report time from the loan extracts (design B), joined on
  `member_number` via the same enrichment the impdet report already ships, and
  bucketed by a line-for-line transcription of the matrix rule. Wired into
  `generate_report.py` (+16 lines, CRLF verified intact) and it returns
  immediately when the WARM already supplied the keys, so healthy CUs are
  untouched.

  Design A (persist `days_delinquent` at import) was rejected on one decisive
  point: it cannot repair an already-delivered workbook without re-importing
  every affected snapshot, and the 19 blank reports already in clients' hands
  are the whole point. **No schema change was made.**

  Utah ground truth: **population matches to the penny** — derived
  $12,108,522.68 vs WARM $12,108,522.68 — validating threshold, join and
  balance column. The *status split* differs by ~30% of delinquent dollars,
  fully explained: 248 delinquent loans have no origination score, so
  `import_data.py:1701-1704` gap-fills `original = current` and our matrix
  calls them Unchanged where the WARM's own score columns call them
  Deteriorated / Not Reported. **Following the WARM would put the pie in
  contradiction with the matrix printed beside it**, so consistency with the
  matrix was chosen. Four alternative conventions were tested; none reproduces
  the WARM.

  Mountain CU 2026-06-30: 100% join coverage, 151 loans / $987,947.81,
  **7 of 9 blank DQ pies now carry data** (the other 2 are genuinely zero).
  Integrity warning drops 18 -> 11 blocks. Also fixed `sci_fcu` and
  `nucor_emp_cu`.

  Two client configs need an analyst fix, and are guarded so they produce
  nothing rather than something wrong: `destinations_cu.yaml` maps
  `days_delinquent` to `Interest Rate` **and** to an account-number column;
  `wnc_community_cu.yaml` maps nothing.
- **Charge-offs — VERDICT: NOT DERIVABLE from present data.**
  See `07_chargeoff_feasibility.md`. Two independent blockers, either fatal:
  1. For **9 of the 10** affected wizard CUs, `monthly_loan_data` holds zero
     rows for charged-off loans — nothing to bucket.
     `chargeoff_exclude_column` is set in only **2 of 39** client configs
     (`mountain_cu`, `nova_cu`) — independently verified.
  2. Where the rows do exist (Mountain), the grade pair is degenerate: of
     ~3,997 `Exclude` rows only 39 have `original_fico_score !=
     current_fico_score`, because Mountain maps *both* scores to the same
     extract column. Bucketing would drop ~99% of charge-off dollars into
     "Unchanged" — **a confidently wrong chart, worse than a blank one.**

  The decisive insight: the WARM's `CO Data Entry` tab is **not** analyst
  judgement. It is per-pool current-grade x original-grade matrices whose four
  status cells are derived by formulas that match `report_vizo.py:1758-1770`
  exactly, `top_grades_double_drop = 3` included. **The bucketing rule already
  lives in our code; only its input is missing** — a per-loan charge-off file
  carrying origination and current scores, which we never receive.

  Utah's real values (verified): Improved $1,116,964.01 (3.86%),
  Deteriorated $15,151,657.35 (52.33%), Unchanged $8,689,857.16 (30.01%),
  Not Reported $3,995,038.50 (13.80%). Gross charge-offs only — there is no
  `Recov Data Entry` tab.

  **Recommendation (needs Brian's approval — changes client output):** suppress
  the CO migration bar and put a chart in the slot that is actually true —
  charge-offs **by pool and by year** from `loan_code_chargeoff_history`, which
  every affected CU already has with 29-90 months of depth. It drops the
  migration dimension, which is precisely the dimension we cannot honestly
  supply. Keep `co_by_status` as an optional WARM-supplied input so legacy CUs
  keep their figures. **~3-4 days.** Wizard capture of a scored charge-off
  extract is a data-acquisition project, not a software change; defer until a
  client asks.

  Also worth fixing while in there: the early `return` at
  `generate_report.py:4990` silently discards good CO data when a WARM lacks an
  `ACL Env by Pool Mgmt Adj` tab.

Scope note: the two halves will not land the same way, and the charge-off half
should not wait on a derivation that cannot exist. Delinquency is derivable
(`days_delinquent` is mapped for every affected CU, and delinquent loans are
live in the extract with both scores); charge-offs are not.

### Recurrence guard (done)

`report_integrity.py` now carries `_check_dq_co_blocks`, which locates the
migration-status blocks **by row label** (column positions differ between credit
unions) and warns when one is entirely zero. Deliberately a *warning*, not an
error: unlike malformed XML this is "almost certainly wrong" rather than
"definitely broken", and should not block delivery.

Calibrated both ways: on Mountain CU June it reports exactly
`18 block(s) across 9 Risk Change tab(s)` and exits 0; on a fixture with the
blocks populated it stays silent. It correctly ignores the two *other*
status-labelled blocks per tab, which are legitimately non-zero.

## Next actions

1. **Asset resolver in `report_vizo.py`** (deferred to avoid collision):
   paths at `:63`, `:120-121`, `:6623-6624` resolve only against the
   workspace and fall through *silently*. Prefer workspace, fall back to
   `report_assets/`. Now unblocked — Step 1a finished.
2. **Step 1b, design system** — needs Brian's taste; the chart spike is now
   the artefact to react to. Two specific decisions are already queued in
   `05_chart_spike.md` and should be settled first:
   - the **Vizo brand palette fails colour-vision validation** (teal `#0D4D5E`
     reads gray, olive vs amber is dE 1.5 under deuteranopia). The spike ships
     the same four hues re-stepped to pass.
   - Excel fills **Improved maroon and Deteriorated teal**; the port uses the
     semantic mapping instead.
3. **Step 2/T1** — rewrite the four summary tabs against `AclEnvironmental`
   instead of screen-scraping the worksheet. This is what retires
   `_parse_acl_layout` and the cross-sheet formulas.
4. **Serialize `AclEnvironmental` to JSON per quarter** — retires reading
   prior-period numbers out of a prior workbook, and with it the latent
   `data_only=True` failure.

## URGENT — collateral finding: delinquency is overstated ~7x

**`cecl_ui/services/dq_extract_parser._coerce_days` strips non-digit characters,
including the decimal point.** Verified directly:

    '60'  -> 60      60.0 -> 600      6.0  -> 60
    '90'  -> 90      90.0 -> 900      30.0 -> 300

Extracts parse as float64 whenever a cell is blank, so **any loan 6+ days past
due is booked as 60+ days delinquent.** Utah's
`loan_code_delinquency_history` holds $86,351,866.11 where the true figure,
confirmed against the WARM, is $12,108,522.68.

This is not cosmetic. The linkage is
`loan_code_delinquency_history` -> `warm_dq_pct` -> `_pool_dq_variance`
(`report_vizo.py:496-505`) -> `dq_score` -> `env_factor = (ncc_score +
dq_score + es_score) / 100` (`:626-628`) -> the allowance. An inflated DQ%
lands in a worse score band, raising the environmental factor and therefore the
reserve. **Direction of error: over-reserving.**

The same table also double-counts charged-off balances for CUs with
`chargeoff_exclude_column` (Mountain: $15.85M recorded vs $0.99M true).

**Deliberately NOT fixed.** It changes the environmental factor, and therefore
delivered allowance numbers, for many credit unions. It needs its own
verification, a blast-radius assessment across every affected CU and quarter,
and Brian's sign-off. `dq_migration_split.py` uses its own safe coercion, which
Utah validates to the penny, so the new charts are unaffected by the bug.

## Caveats carried forward

- Step 1a verified on one CU/quarter only. **No BRR pool and no `Hide-*` grade
  was exercised** — those branches are transcribed but untested. Run a BRR
  client before relying on them.
- `_expand_unfunded_commitment_oac` mutates `config` mid-calculation, so
  `compute_acl_environmental` is not strictly pure; it is idempotent, which is
  the only reason repeated calls agree. Callers wanting purity must pass a copy.
- `balance_label` uses the raw `snap` string on ACL Env but the `mm/dd/yyyy`
  display form on `Impr Deter` — a real inconsistency between the two tabs
  today, reproduced deliberately rather than "fixed".
- `scripts/verify_acl_model.py` does not suppress the DB upsert in
  `_derive_snapshot_dq_from_extracts`; running it writes the same derived DQ
  rows a normal run writes.
