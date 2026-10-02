# 04 — Blank DQ / Charge-off charts on the Risk Change tabs

Follow-up to the open question raised in `03_chart_inventory.md` ("18 of the 40 charts plot
all zeros … this is an upstream data question, and it should be answered before either
archetype is built").

Evidence base: `report_vizo.py`, `generate_report.py`, `import_data.py` (working tree, no
modifications made), every `*CECL_Migration_*_Vizo_Model.xlsx` and `*_TCT_Model.xlsx` in
`Z:\Shared\TCT Files\CECL - CM Files\Reports\` and `…\Reports\older reports\`, plus the
source WARM workbooks on the client shares.

---

## VERDICT — CONFIRMED

The DQ pie (`Delinquency by Credit Grade Migration`) and the charge-off bar
(`Charge off by Credit Grade Migration`) on `Risk Change Total` and on every
`Risk Chg <pool>` tab are drawing from cells that contain literal `0` in the Mountain CU,
Nucor Emp CU and SCI FCU workbooks. Every value in every source range is zero — not
`None`, not a stale cache; an actual written `0`.

It is **not** a chart-wiring bug, **not** a wrong-cell bug, and **not** a pool-name
lookup mismatch. The cells are written correctly; the dictionary they are written from is
empty. The feed (`hist['impaired']['dq_by_status' | 'dq_by_pool' | 'co_by_status' |
'co_by_pool']`) has exactly one producer, and for these credit unions it never runs.

It is also **not new** and **not universal** — some CUs have always had these charts
populated, and the same CU can flip between populated and blank from one quarter to the
next. See *Blast radius* and *History*.

---

## Blast radius

**Per report:** 2 charts per Risk Change tab × (1 total tab + N pool tabs).

| Report (delivered) | Total charts | Blank DQ pies | Blank CO bars | Blank share |
|---|---|---|---|---|
| `2026-06-30_CECL_Migration_Mountain_CU_Vizo_Model.xlsx` | 40 | 9 | 9 | **18 / 40** |
| `2026-03-31_CECL_Migration_Mountain_CU_Vizo_Model.xlsx` | 40 | 9 | 9 | **18 / 40** |
| `2026-06-30_CECL_Migration_Nucor_Emp_CU_Vizo_Model.xlsx` | 32 | 7 | 7 | **14 / 32** |
| `2026-06-30_CECL_Migration_SCI_FCU_Vizo_Model.xlsx` | 32 | 7 | 7 | **14 / 32** |
| `2026-06-30_CECL_Migration_Destinations_CU_Vizo_Model.xlsx` | 28 | 6 | 6 | **12 / 28** |
| `2026-06-30_CECL_Migration_WNC_Community_CU_Vizo_Model.xlsx` | 8 | 1 | 1 | **2 / 8** |
| `2026-06-30_CECL_Migration_Utah_Community_FCU_Vizo_Model.xlsx` | 36 | 0 | 0 | 0 / 36 (**healthy**) |

The suspicion's "roughly 18 of 40" is exact for Mountain CU.

**Across the archive.** Every Vizo *Migration* workbook in `Reports/` and
`Reports/older reports/` was scanned for the `DQ Balance` / `CO Balance` table headers and
the four data rows under each.

All-zero, client-facing:

| Snapshot | Credit union | Risk Chg tabs | Blank charts |
|---|---|---|---|
| 2025-06-30 | Honolulu Fire Department FCU | 7 | 14 |
| 2025-12-31 | Bridgeton Onized FCU | 7 | 14 |
| 2025-12-31 | Emergency Responders CU | 7 | 14 |
| 2025-12-31 | Nucor Emp CU | 7 | 14 |
| 2025-12-31 | Shuford FCU | 12 | 24 |
| 2025-12-31 | Utah Community FCU | 7 | 14 |
| 2026-02-28 | Central Keystone FCU | 2 | 4 |
| 2026-02-28 | Curis CU / Palmetto Health CU | 9 | 18 |
| 2026-03-31 | Census FCU | 7 | 14 |
| 2026-03-31 | Central Keystone FCU | 4 | 8 |
| 2026-03-31 | Curis CU / Palmetto Health CU | 9 | 18 |
| 2026-03-31 | Emergency Responders CU | 7 | 14 |
| 2026-03-31 | Franklin Trust FCU | 11 | 22 |
| 2026-03-31 | Mountain CU | 9 | 18 |
| 2026-06-30 | Destinations CU | 6 | 12 |
| 2026-06-30 | Mountain CU | 9 | 18 |
| 2026-06-30 | Nucor Emp CU | 7 | 14 |
| 2026-06-30 | SCI FCU | 7 | 14 |
| 2026-06-30 | WNC Community CU | 1 | 2 |

Plus internal/scratch runs: `Census_FCU_(Scratch)` Dec-25, `Test_Nova_CU` Dec-25,
`Test_Nucor_Emp_CU` Dec-25.

Healthy: `Sample_Credit_Union` Dec-25 (10/10 tabs populated), `Franklin_Trust_FCU`
Dec-25 (9/12 DQ, 11/12 CO — the empty ones are genuine zeros in the source),
`Utah_Community_FCU` Jun-26 (7/8; `Construction` is a genuine zero).

**The TCT-model family has the identical defect** — `report_tct.py:2387-2398` reads the same
four keys with the same fallback-to-`{}`. All-zero TCT workbooks: Honolulu Jun-25,
Cottonwood Nov-25 / Feb-26 / May-26, Bridgeton Dec-25, Tongass Dec-25, Franklin Mar-26.
So the total client-facing blast radius is larger than the Vizo count alone.

**Visual symptom.** `patch_dq_pie_zero_labels` (`report_vizo.py:5200`) then injects
`<c:delete val="1"/>` on every zero slice. Verified on Mountain CU June —
`xl/charts/chart7.xml` contains 4 `delete val="1"` and no cached values — so Excel renders a
titled frame with a legend and *nothing inside it*. The reader sees an empty chart, not a
chart of zeros, which is why it reads as "broken" rather than "no delinquency".

---

## 1. Exactly which cells the two charts read

`_sheet_risk_change` — `report_vizo.py:1596` (called at `:6577` for the total tab and
`:6610` per pool).

Layout anchors:

- `ncol = 3 + len(gl)` (`:1638`), `pcol_start = ncol + 2` (`:1641`)
- `r_pgt = r_ph2 + 1 + len(gl)` (`:1907`) → `r_nc = r_pgt + 2` (`:1920`)
- **DQ table** header row `r_dq = r_nc + 16` (`:2038`); labels in `pcol_start`,
  balance in `pcol_start+1`, `% of Total` in `pcol_start+2`; four data rows
  `r_dq+1 … r_dq+4` (`:2043-2049`).
- **DQ pie** (`:2051-2056`):
  `cats_dq = Reference(min_col=pcol_start, min_row=r_dq+1, max_row=r_dq+4)`,
  `vals_dq = Reference(min_col=pcol_start+2, min_row=r_dq+1, max_row=r_dq+4)` ← plots the
  **percent** column, not the balance column.
- **CO table** header row `r_co = r_nc + 22` (`:2107`), same three columns, rows
  `r_co+1 … r_co+4` (`:2112-2118`).
- **CO bar** (`:2130-2133`): `cats_co` = `pcol_start`, `vals_co` = `pcol_start+2`, rows
  `r_co+1 … r_co+4`.

Resolved against the delivered files (read out of `xl/charts/chart*.xml`, so this is what
Excel actually binds to, not a re-derivation):

| Workbook | Tab | DQ pie cats / vals | CO bar cats / vals |
|---|---|---|---|
| Mountain CU (Jun-26 & Mar-26) | `Risk Change Total` | `$M$48:$M$51` / `$O$48:$O$51` | `$M$54:$M$57` / `$O$54:$O$57` |
| Mountain CU | `Risk Chg Real Estate` | `$M$48:$M$51` / `$O$48:$O$51` | `$M$54:$M$57` / `$O$54:$O$57` |
| Nucor Emp CU / SCI FCU (Jun-26) | all tabs | `$L$46:$L$49` / `$N$46:$N$49` | `$L$52:$L$55` / `$N$52:$N$55` |
| Franklin Trust FCU (Mar-26) | all tabs | `$K$44:$K$47` / `$M$44:$M$47` | `$K$50:$K$53` / `$M$50:$M$53` |

The column letters move with grade count; the *offsets* are invariant.

## 2. Actual values in those ranges

Read with `openpyxl.load_workbook(..., data_only=True)`.

`2026-06-30_CECL_Migration_Mountain_CU_Vizo_Model.xlsx` — **all 9 tabs identical**
(`Risk Change Total`, `Risk Chg Consumer Auto Loan-N`, `…-U`, `Consumer Indirect Au`,
`Consumer Indirect Au1`, `Consumer Secured`, `Consumer Unsecured`, `Credit Cards`,
`Real Estate`):

```
N47 "DQ Balance"        O47 "% of Total"
M48 Improved       N48 0   O48 0
M49 Deteriorated   N49 0   O49 0
M50 Unchanged      N50 0   O50 0
M51 Not Reported   N51 0   O51 0

N53 "CO Balance"        O53 "% of Total"
M54 Improved       N54 0   O54 0
M55 Deteriorated   N55 0   O55 0
M56 Unchanged      N56 0   O56 0
M57 Not Reported   N57 0   O57 0
```

Nucor Emp CU Jun-26 (7 tabs), SCI FCU Jun-26 (7 tabs), Mountain CU Mar-26 (9 tabs):
same — every balance and every percent is `0`, on every tab, total and per-pool.

The `Net Credit Change` block on the same sheets **is** correct, exactly as reported. It is
computed from `risk_change_matrix(data_df, …)` inside the function
(`report_vizo.py:1740-1782`, `:1920-1950`) and never touches `hist`. That is why one block
is right and the other is empty on the same tab.

Healthy contrast — `2026-06-30_CECL_Migration_Utah_Community_FCU_Vizo_Model.xlsx`,
`Risk Change Total`:

```
L46 Improved       M46    248,344.81   N46  2.29%
L47 Deteriorated   M47  6,309,627.46   N47 58.06%
L48 Unchanged      M48  2,474,020.07   N48 22.77%
L49 Not Reported   M49  1,835,230.33   N49 16.89%

L52 Improved       M52  1,154,254.09   N52  4.22%
L53 Deteriorated   M53 14,490,820.20   N53 53.02%
L54 Unchanged      M54  8,165,050.35   N54 29.88%
L55 Not Reported   M55  3,518,377.46   N55 12.87%
```

So the writer, the references and the chart plumbing all work. Only the input is missing.

## 3. Where the data comes from, and why it is empty

`report_vizo.py:1621-1631`:

```python
_imp = hist.get('impaired', {}) if hist else {}
if pool_name:
    _dq_pool = _imp.get('dq_by_pool', {})
    _co_pool = _imp.get('co_by_pool', {})
    _pool_lc = pool_name.strip().lower()
    _dq_data = next((v for k, v in _dq_pool.items() if k.strip().lower() == _pool_lc), {})
    _co_data = next((v for k, v in _co_pool.items() if k.strip().lower() == _pool_lc), {})
else:
    _dq_data = _imp.get('dq_by_status', {})
    _co_data = _imp.get('co_by_status', {})
```

then `dq_entry.get('balance', 0)` / `dq_entry.get('pct', 0)` (`:2046-2048`, `:2115-2117`).
Every miss degrades silently to `0`.

**These four keys have exactly one producer in the entire codebase:**
`load_impaired_data` → `generate_report.py:5176-5202`.

```
5176  dq_df = pd.read_excel(found, sheet_name='DQ Data Entry', header=None)
5177  dq_grand, dq_by_pool = _read_migration_blocks(dq_df)
5181      result['dq_by_status'] = dq_grand
5185      result['dq_by_pool']   = dq_by_pool
5192  co_df = pd.read_excel(found, sheet_name='CO Data Entry', header=None)
5195      result['co_by_status'] = co_grand
5199      result['co_by_pool']   = co_by_pool
```

`_read_migration_blocks` (`:5140-5174`) walks column P for `Loan Status` headers and reads
the four rows below from columns P/Q/R. Nothing else in `generate_report.py`,
`report_vizo.py`, `report_tct.py`, `cecl_engine.py` or `import_data.py` ever writes these
keys — verified by repo-wide grep. The later `hist['impaired']` mutations
(`generate_report.py:12643` standalone impaired, `:12661` wizard impaired, `:12989` ACL
balance, `_compute_balance_adjustments`) only *overlay* other keys; they neither supply nor
destroy the DQ/CO ones.

So the chain is:

```
<YYYY-MM> CECL-Migration-WARM - <CU>.xlsx        (legacy full WARM workbook)
  └─ tabs 'DQ Data Entry' / 'CO Data Entry'
       └─ load_impaired_data()            generate_report.py:4772
            └─ hist['impaired']['dq_by_status'|'dq_by_pool'|'co_*']   :12436
                 └─ _sheet_risk_change()  report_vizo.py:1621
                      └─ cells P/Q/R  →  pie + bar
```

**Break point.** `load_impaired_data` locates that workbook by name only
(`generate_report.py:4802-4856`): exact `f"{snap[:7]} CECL-Migration-WARM - {cu}.xlsx"`
(or the `_`-separated variant), else regex
`^{YYYY-MM}.*CECL-Migration-WARM.*\.xlsx$` with the CU name in the filename, searched under
`data_directory` and `credit_pull.fallback_report_folder`. No match ⇒
`print("    No CECL-Migration-WARM file found for {snap_prefix}")` and
`return {}` (`:4863-4865`).

Reproduced against the live shares (replicating the search exactly, and by calling
`generate_report.load_impaired_data` directly):

| Client config | Snapshot | WARM found | `dq_by_status` |
|---|---|---|---|
| `mountain_cu` | 2026-06-30 | **none** — `No CECL-Migration-WARM file found for 2026-06` | `{}` |
| `mountain_cu` | 2026-03-31 | **none** | `{}` |
| `nucor_emp_cu` | 2026-06-30 | **none** | `{}` |
| `sci_fcu` | 2026-06-30 | **none** | `{}` |
| `franklin` | 2026-03-31 | `…\Portfolio Management (CM, WARM, ID)\2026-03 CECL-Migration-WARM - Franklin Trust FCU.xlsx` | populated |
| `utah_community_fcu` | 2026-06-30 | `…\Portfolio Management (CM, ID, Warm)\2026-06 CECL-Migration-WARM - Utah Community FCU.xlsx` | populated |

Mountain CU *does* have a WARM-ish upload —
`Raw_Uploads\mountain_cu\CECL-WARM with Credit Migration Impaired Loans - 6-30-26.xlsx` —
but it fails on both counts: the filename does not match the pattern, and the workbook has
only three tabs (`Instructions`, ` Impaired Loans`, `Management Adjustment`). No
`DQ Data Entry`, no `CO Data Entry`, no `ACL Env by Pool Mgmt Adj`. Nucor Emp CU and SCI FCU
have **no** WARM-shaped workbook at all — only standalone
`… CECL Migration Impaired Loans …` files.

Utah Community FCU's WARM, by contrast, is the full legacy ~80-tab workbook including
`DQ Data Entry` and `CO Data Entry`. That is the whole difference.

**Loosening the filename match would NOT fix this.** The client shares are full of
near-miss names — `2026Q1_CECL-Migration -WARM Impaired Loans - Utah Community FCU.xlsx`
(space before `-WARM`), `2025Q4_CECL-Migration -WARM Impaired Loans - …`, and per-quarter
`… CECL-Migration-WARM - <CU>.pdf` — none of which match
`^{YYYY-MM}.*CECL-Migration-WARM.*\.xlsx$`. Checked one: the `2026Q1` Utah file has **3
tabs** (`Instructions`, ` Impaired Loans`, `Management Adjustment`) — the same trimmed
workbook Mountain CU uploads, with no `DQ Data Entry`. That is the file the *credit union*
supplies into `Client Access`; the full ~80-tab WARM that actually carries the DQ/CO tabs
lives in the TCT-internal `Portfolio Management (…)` folder and exists only for CUs TCT
still maintains a legacy WARM for. So the real determinant is not the filename regex — it is
whether a full WARM workbook exists for that CU at all. Relaxing the pattern would only
match more 3-tab files and change nothing.

**The three candidate explanations, resolved:**

- *Wrong cells?* **No.** `xl/charts/chart*.xml` binds to exactly the ranges the code writes,
  and the same code produces correct output for Utah Community FCU.
- *Pool-name key mismatch (whitespace/case)?* **No.** `report_vizo.py:1626-1631` already
  normalises with `.strip().lower()` on both sides. Verified against Utah's WARM: its
  `dq_by_pool` keys (`Mortgage Real Estate`, `Indirect Auto`, `Direct Auto`, `Unsecured`,
  `Other Consumer`, `Commercial`, `Construction`, …) match the report's pool names 1:1, and
  the one zero tab (`Construction`) is a genuine zero in the source. Sheet-name truncation
  (`safe = re.sub(...)[:20]`, `:1632`) does not affect the lookup, which uses the full
  `pool_name`.
- *Upstream source missing?* **Yes.** `hist['impaired']` has no `dq_*`/`co_*` keys at all for
  the affected CUs.

**"But delinquency data demonstrably exists."** It does — and it is a different shape.
`_derive_snapshot_dq_from_extracts` (`generate_report.py:12100-12260`, the source of the
`Snapshot DQ derived from extract(s) … 4326 loan(s) >= 60 days, $15,846,547.03 delinquent
across 70 loan code(s)` line) reads `days_delinquent` straight off the loan extract and
upserts **`loan_code_delinquency_history`, keyed by loan code**. It feeds
`_load_dq_history_from_db` (`:6544`) → `hist['impaired']['warm_dq_pct']` (the env-factor DQ%
series), which is a *different* consumer. Nothing splits delinquency by **migration status**
(Improved / Deteriorated / Unchanged / Not Reported) — that split only ever existed inside
the analyst-maintained `DQ Data Entry` tab.

And it cannot currently be derived at report time: `monthly_loan_data` — built at
`import_data.py:1651-1661` — carries only `credit_union, snapshot_date, member_number,
current_balance, current_fico_score, original_fico_score, loan_pool`
(+ `business_risk_rating`). **No `days_delinquent`, no charge-off amount.** So `data_df`
inside `_sheet_risk_change` has the migration buckets but nothing delinquent to bucket. The
mapping *is* configured (`mountain_cu.yaml:16 days_delinquent: Day Delinquent`,
`nucor_emp_cu.yaml:18`, `sci_fcu.yaml:17`, all alongside `member_number`) — it is read at
import and thrown away.

## 4. History — when did it start

Repo git history is a single squashed commit (`d3e4873`, 2026-05-27), so pre-May code
archaeology is not available. But the behaviour is **not** a code regression; it is per-CU
data availability, and it has been intermittent since the earliest archived report:

- Oldest workbook in the archive (`2025-06-30` Honolulu Fire Department FCU, Vizo **and**
  TCT) is already all-zero.
- Utah Community FCU: **zero** at 2025-12-31 → **populated** at 2026-06-30.
- Franklin Trust FCU: **populated** at 2025-12-31 → **zero** at 2026-03-31 → **populated**
  again at 2026-06-30 (TCT).

The Franklin flip is explained, and it corroborates the root cause rather than complicating
it: `client_configs/franklin.yaml.bak.20260723_145401` shows that before 2026-07-23 the
config read `data_directory: Franklin Trust Files` and
`fallback_report_folder: Franklin Trust Files` — a folder on the CECL share that contains
**no WARM workbook at all**. On 2026-07-23 both were repointed to
`Z:\Shared\Clients\Franklin Trust FCU\…`, where the `2026-03` and `2026-06` WARM files
actually live. The 2026-03-31 report (generated 2026-04-16) therefore could not find a WARM;
re-running it today does — confirmed by calling `load_impaired_data(franklin, '2026-03-31')`,
which now returns DQ `{Improved 0.00, Deteriorated 1,169.72, Unchanged 190,959.90,
Not Reported 1,232.59}` and CO `{275,724.64 / 325,167.44 / 1,226,557.97 / 38,133.59}`.

So: **a report's DQ/CO charts are populated iff a correctly-named legacy WARM workbook
containing `DQ Data Entry` / `CO Data Entry` is reachable from that CU's configured
folders at run time.** Legacy CUs that still hand over the full WARM (Sample, Franklin,
Utah, Maple, Ontario, Bridgeton via the older TCT pipeline) get charts. Every
wizard-onboarded CU — Mountain, Nucor, SCI, Destinations, WNC, Census, Central Keystone,
Curis, Emergency Responders, Shuford — uploads only the trimmed
`CECL-WARM with Credit Migration Impaired Loans*.xlsx` (3 tabs) or a standalone
`Impaired Loans` file, and gets 100% blank DQ/CO charts, permanently. **The share of blank
reports grows with every new client onboarded through the wizard.**

## 5. Latent second failure mode (found while tracing — not the cause here)

`load_impaired_data` returns early at `generate_report.py:4985-4990` when the WARM lacks an
`ACL Env by Pool Mgmt Adj` tab:

```python
try:
    acl_df = pd.read_excel(found, sheet_name='ACL Env by Pool Mgmt Adj', header=None)
except (ValueError, KeyError):
    print(f"    'ACL Env by Pool Mgmt Adj' tab not found")
    return result          # ← DQ/CO parsing at :5176 is never reached
```

Any WARM that has `DQ Data Entry` / `CO Data Entry` but no ACL Env tab silently loses its
DQ/CO data. None of the currently-affected CUs hit this (they have no WARM at all), but it
should be fixed in the same pass — the DQ/CO block belongs above that return, or the return
should become a flag.

---

## 6. Recommended fix (NOT applied)

Three tiers. Tier 1 stops shipping misleading pages this quarter; Tier 2 is the real fix;
Tier 3 is hygiene.

### Tier 1 — never render a chart with no data (`report_vizo.py`, mirror in `report_tct.py`)

Today an absent feed renders an empty titled pie, which a reader interprets as "we measured
this and it is zero". Suppress the chart and say why. In `_sheet_risk_change`, after
`dq_status = _dq_data` (`report_vizo.py:2042`) and `co_status = _co_data` (`:2111`):

```python
    # ─── DQ Pie Chart "Delinquency by Credit Grade Migration" ───
    _dq_has_data = any(
        float((dq_status.get(_l) or {}).get('balance') or 0)
        for _l in ("Improved", "Deteriorated", "Unchanged", "Not Reported")
    )
    if not _dq_has_data:
        _note = ws.cell(
            row=r_dq + 6, column=pcol_start,
            value="Delinquency by migration status unavailable for this "
                  "quarter (no WARM 'DQ Data Entry' source).",
        )
        _note.font = V10
    else:
        dq_pie = PieChart()
        ...                     # existing :2051-2105 body, indented
        ws.add_chart(dq_pie, anc_dq)
```

and the symmetric guard around the `co_bar` block (`:2120-2185`). Keep writing the P/Q/R
tables either way — they are honest zeros and other tooling reads them; it is only the
*chart* that misleads. `patch_dq_pie_zero_labels` needs no change: with the chart absent
there is nothing for it to patch.

Same edit at `report_tct.py:2387-2398` and its chart blocks below.

### Tier 2 — actually produce the numbers for wizard-onboarded CUs

The Improved/Deteriorated/Unchanged/Not-Reported split is computable from data the pipeline
already reads and then discards. Two steps:

1. **Persist per-loan delinquency.** `import_data.py:1651-1661` — add `days_delinquent` to
   the `clean_data` frame when `config['column_mappings']['days_delinquent']` is present (it
   is, for every affected CU), guarded by an idempotent
   `ALTER TABLE monthly_loan_data ADD COLUMN IF NOT EXISTS days_delinquent INTEGER`
   alongside the existing `business_risk_rating` ALTER at `:1727-1730`.

2. **Derive `dq_by_status` / `dq_by_pool` when the WARM tabs are absent.** In
   `generate_report.py`, after the `load_impaired_data` block (`:12435-12445`), add a
   fallback that buckets each loan with the *same* rule `_sheet_risk_change` uses for the
   matrix — grade index `i` (current) vs `j` (original), `n_top =
   config.get('top_grades_double_drop', 3)`, `Not Reported` → Unchanged (mirroring
   `report_vizo.py:1758-1770`) — summing `current_balance` for rows with
   `days_delinquent >= dq_threshold`, both grand-total and per-pool, writing
   `{'balance': …, 'pct': …}` in the same shape `_read_migration_blocks` produces. Set the
   keys only when the WARM did not supply them, so legacy CUs keep the analyst-maintained
   figures.

   `co_by_status` is a bigger job and should be a separate change: charged-off loans are
   routed to the `Exclude` pool at import (`import_data.py:1677-1679`), so they are not in
   `data_df` at all, and the CO extract would have to be joined back on `member_number`.
   Ship the DQ half first — the DQ pie is the more prominent of the two.

### Tier 3 — make the failure loud

- `load_impaired_data`: move the DQ/CO block above the `return result` at `:4990`, and log a
  `WARNING` (not a silent empty dict) when `found` is `None` or when either Data Entry tab
  is missing — the current `No CECL-Migration-WARM file found` line reads like routine
  chatter.
- Add a `report_integrity.py` check: if any `Risk Chg*` tab writes an all-zero DQ or CO
  block, fail the run loudly rather than emitting the workbook.

## 7. Re-issue guidance

The affected pages are not *wrong* — the Risk Change matrices, `Net Credit Change`, the
summary table, the doughnut and the `Risk Change by Grade` bar on those same tabs are all
correct and independently computed. Only the two DQ/CO charts are empty, and they were empty
in the originally delivered files too. Once Tier 1 lands, the affected CUs' reports can be
regenerated to replace 18 (Mountain), 14 (Nucor, SCI), 12 (Destinations), 2 (WNC) empty
frames with an explicit "data unavailable" note — a cosmetic and honesty improvement, not a
restatement. A restatement is only warranted after Tier 2, when the charts carry real
numbers.

---

### Reproduction

```bash
python - <<'PY'
import openpyxl
p = r"Z:\Shared\TCT Files\CECL - CM Files\Reports\2026-06-30_CECL_Migration_Mountain_CU_Vizo_Model.xlsx"
wb = openpyxl.load_workbook(p, data_only=True)
ws = wb["Risk Change Total"]
for r in range(47, 58):
    print(r, [ws.cell(row=r, column=c).value for c in (13, 14, 15)])
PY
```

```bash
python - <<'PY'
import os, sys
os.environ.setdefault('CECL_WORKSPACE_ROOT', r"Z:\Shared\TCT Files\CECL - CM Files")
sys.path.insert(0, r"C:\Dev\CECL")
import generate_report as gr
for cli, snap in [('mountain_cu', '2026-06-30'), ('utah_community_fcu', '2026-06-30')]:
    r = gr.load_impaired_data(gr.load_config(cli), snap)
    print(cli, '->', (r or {}).get('dq_by_status'))
PY
```
