# CECL Migration report — Vizo layout redesign

Applies the layout redesign already shipped for SCALE (see
`cecl_ui/services/scale/vizo_layout.py`) to the **Migration** Vizo workbook
built by `report_vizo.compose_vizo_main`.

Status: **work items 1, 2 and 4 done; item 3 (Report Index) open.** Reconstructed 2026-08-31 from
the approved reference workbook after the original planning conversation was
lost; tab specs confirmed by Brian the same day.

## Reference workbook

`notes/2026-06-30_CECL_Migration_Mountain_CU_Vizo_Model REdone.xlsx` (2026-08-18)
is the hand-edited approved target. It is a **structural skeleton**: it fixes
tab names, order, and visibility, but four tabs are empty placeholders (see
Open questions).

## Target tab order (24 tabs)

```
 0  Vizo Cover                        13  ACL Env by Pool Mgmt Adj
 1  Report Index                      14  Change Analysis            MOVED
 2  Summary Variance          NEW     15  Impaired Loans             NEW
 3  Impr Deter                        16  ACL Summary                NEW
 4  Risk Change Total                 17  Mgmt Adj Summary           NEW
 5  Risk Chg Consumer Auto Loan-N     18  Env Factor by Pool
 6  Risk Chg Consumer Auto Loan-U     19  >Envir Fact Ranges
 7  Risk Chg Consumer Indirect Au     20  Display HIst Bal
 8  Risk Chg Consumer Indirect Au1    21  Display CO-Recov-DQ
 9  Risk Chg Consumer Secured         22  Introduction-Vizo          MOVED
10  Risk Chg Consumer Unsecured       23  Executive Summary-Vizo     MOVED
11  Risk Chg Credit Cards
12  Risk Chg Real Estate
```

Per-pool `Risk Chg *` tabs (5-12) are CU-dependent; the eight above are
Mountain CU's.

## Work items

### 1. Moves (no new content) - DONE

Implemented in `report_vizo.py` as `_VIZO_MAIN_ORDER` + `_reorder_vizo_main`,
called at the end of `compose_vizo_main`. The table is declarative, idempotent,
and tolerant of tabs that do not exist for a given CU, so the four new tabs can
be slotted in simply by building them - no further ordering work.
`"Risk Chg *"` is a wildcard slot holding the CU-dependent per-pool sheets.

This replaced two ad-hoc placements:

- `Change Analysis` was appended **last** ("always last"). It is still *built*
  last, so it can read every other tab, but is now *displayed* at 14.
- `Introduction-Vizo` / `Executive Summary-Vizo` were moved to just after
  `Display CO-Recov-DQ`; that block was removed in favour of the table.

**`report_tct.py:5037` still appends `Change Analysis` last.** No TCT reference
workbook exists, so the TCT tab order was left alone pending confirmation.

### 2. New tabs - DONE (first draft, awaiting review)

Built in `report_vizo.py` as `_sheet_summary_variance`, `_sheet_impaired_loans`,
`_sheet_acl_summary`, and `_sheet_mgmt_adj_summary`, called from
`compose_vizo_main` after the env merge.

**All four are formula views over `ACL Env by Pool Mgmt Adj`, not
recalculations.** A shared `_parse_acl_layout` walks the built ACL sheet and
returns *row numbers* for every block (pool headers, grade rows, pool totals,
pooled totals, impaired categories, OAC rows, and the five closing totals); the
summary tabs then emit `='ACL Env by Pool Mgmt Adj'!K118`-style references. A
summary therefore cannot drift from the tab it summarises, and stays correct if
someone edits the source tab in Excel. This is the same approach SCALE's Vizo
tabs take over their own calc tabs.

The one exception is the Prior block on Summary Variance: those figures live in
a different workbook, so they are written as static values, read through
`change_analysis._find_prior_report` / `_parse_acl_sheet` - the same loader the
Change Analysis tab uses, so the two tabs cannot disagree about which report
"prior" means.

**`Summary Variance` (2)** - build it to look like SCALE's
`Executive Summary-Vizo`. That tab is a 25-row, 4-column (A:D) sheet in three
blocks:

| Block | Rows | Content |
|---|---|---|
| Current ACL | 9-13 | Current ACL, then Total Expected Losses on Loans / Current ACL Balance / Adjustment / ACL/Total Loans |
| Prior ACL | 15-19 | the same four measures, one period back |
| Change | 21-25 | current minus prior, per measure |

In SCALE the prior column is `=OFFSET('Historical Data'!AZn,0,-1)` - a step
left through a history tab. The Migration workbook has no such tab, but
`change_analysis.py` already locates and loads the prior report to build the
Change Analysis tab. **Recommend sourcing "Prior" from that same loader** rather
than adding a history tab, so the two tabs cannot disagree.

**`Impaired Loans` (15)** - build it to work like SCALE's
` Impaired Loans-Vizo`, which is a Vizo presentation layered over the raw
` Impaired Loans ASC 310-10` data:

- Cols A:J mirror the source rows, wrapped in `IF(src="","",src)` so blanks
  stay blank.
- Cols L:S are the calculation block: per-pool `SUMIF` rollups over the loan
  detail (rows 28+), with a `Total` row.
- The Migration source is `generate_report.load_standalone_impaired`. Note the
  provision formula there is **LGD x Provision%**, not Balance x Percentage -
  see `_RECOVERY_STATUS.md`, that bug was deliberately excluded from recovery.

**`ACL Summary` (16)** - one line per pool carrying that pool's Total-row
figures (Balance, Specific Identification, Loan Loss Calc Balance, Allowance
before Env, Env Factor, Env Factor Allowance, Total Allowance), then Pooled
Totals and the closing blocks (impaired categories, Other Allowance
Considerations, Total Allowance Needed, ACL Balance, Adjustment). It is the ACL
tab with the per-grade detail dropped.

**`Mgmt Adj Summary` (17)** - every adjustment applied and what it is worth.
Grade rows appear only where a management adjustment was actually made (column
F non-zero); each pool carries a Total line with its environmental factor and
dollar effect, since that is applied pool-wide.

**Open for review:** on Mountain CU only one pool (Consumer Unsecured) has
grade-level management adjustments; the other seven adjusted pools carry an
environmental factor alone. The tab is consequently mostly header/Total pairs
with three empty columns. Options if that reads badly: drop the per-grade
columns to a note, or split into an environmental-factor table and a much
smaller management-adjustment table.

### 3. Report Index refresh

`Report Index` (tab 1) still describes the old structure and lists none of
the four new tabs. It needs regenerating once the new tabs exist.

## Architecture: how this differs from SCALE

SCALE reports are **seeded from a template workbook** (fresh runs) or **from
the prior quarter's report** (carry-history runs). That second path is why
`vizo_layout.py` exists: a carry-history run inherits the old tab order
forever, so the layout has to be re-applied idempotently after the fact, with
some sheets copied wholesale from the canonical template.

The Migration Vizo workbook has **no such problem**. `compose_vizo_main`
builds it from an empty `Workbook()` on every run, calling `_sheet_*` builders
in sequence. There is no carry-history seeding and no template to drift from.

**Therefore: no `vizo_layout.py` analogue is needed here.** Reordering is a
direct edit to the build/`move_sheet` sequence in `compose_vizo_main`, and the
new tabs are new `_sheet_*` builder functions. Do not port the normalizer.

The one template dependency is narrative text: `Introduction-Vizo` and
`Executive Summary-Vizo` are copied from
`<workspace>/Sample Reports/Vizo Narrative Tabs - Template.xlsx`
(falling back to the master template) at `report_vizo.py:5826`.

## Item 4: merge Env Factor by Pool + >Envir Fact Ranges - DONE

Confirmed: Migration should get the same merge SCALE received on 2026-08-27.
The reference workbook keeps them as two visible tabs only because it predates
that merge (2026-08-18) - staleness, not intent.

**This does not port directly.** In SCALE the merge was done *by hand in the
template workbook*; `vizo_layout.py` just copies the merged sheet wholesale via
`TEMPLATE_SOURCED` and hides the source through `_ALWAYS_HIDDEN_SHEETS` in
`runner.py` (kept, not deleted, because other sheets still reference its cells).

The Migration workbook has no template for this tab - `_sheet_env_factor`
builds it in code, and `_sheet_env_ranges` builds the ranges.

Implemented as `_merge_env_ranges_into_factor` in `report_vizo.py`, called from
`compose_vizo_main` right after `_sheet_env_ranges`. Rather than duplicating the
~200 lines of range/description rendering, it **copies the finished block** from
the built `>Envir Fact Ranges` sheet onto the bottom of `Env Factor by Pool`,
then hides the source. Decisions worth knowing:

- **Destination column widths win.** The ranges sheet uses narrow A/H gutters
  (3.4) that would crush the pool-name column (22.0). Values, fonts, fills,
  borders, alignment, number formats, row heights, and merges are carried over;
  column widths are not.
- **The repeated CU name is dropped.** The ranges sheet titles itself with the
  CU name, which is already at A1 once the block is part of the tab. That row is
  skipped and the gap closed, so the block opens on "Environmental Factor
  Ranges". SCALE's merged tab reads the same way.
- **The source sheet is hidden, not deleted**, so any cell reference into it
  still resolves - matching `_ALWAYS_HIDDEN_SHEETS` in `runner.py`.
- **Idempotent**: re-running detects the block and returns `False`.

`>Envir Fact Ranges` keeps its slot in `_VIZO_MAIN_ORDER`; only its
`sheet_state` changes.

Verified against `notes/2026-06-30_..._REdone.xlsx`: table ends at row 23, block
lands at 26-64, widths intact, second call a no-op.

## Print / page setup

An audit of the exported PDF found several tabs spilling onto extra pages
despite declaring fit-to-one-page. Root cause: **a leftover
`page_setup.scale` (zoom) makes Excel ignore `fitToWidth`/`fitToHeight`.**
`cecl_ui/services/scale/vizo_layout.py` had already learned this for SCALE;
`report_vizo._fit_to_pages` now applies the same incantation (set the fit
counts, clear `scale`, force `pageSetUpPr.fitToPage`) and is used at every
site that wants a fixed page count.

Result on Mountain CU: **33 -> 29 pages.**

| Tab | Before | After | Fix |
|---|---|---|---|
| Report Index | 2 | 1 | stale zoom; 13 rows were spilling sideways |
| Display CO-Recov-DQ | 2 | 1 | stale zoom |
| Display HIst Bal | 4 | 2 | width was unconstrained (`fitToWidth=0`) against an intended 2 pages tall |
| Env Factor by Pool | 2 wide | 1 wide | the ranges merge lengthened the tab but never re-set its print area |

Left multi-page on purpose:

- **ACL Env by Pool Mgmt Adj** (2, was 4 in landscape) - switched to portrait
  on 2026-08-31 with all eleven columns (A:K) on one page wide. 120 rows still
  need two pages tall; the greedy bin-packing deliberately breaks between pool
  blocks so a pool is never split. The row budget moved from 45/40 to
  `ACL_PAGE1_ROWS = 72` / `ACL_OTHER_ROWS = 67`: the horizontal squeeze that
  fits A:K into portrait shrinks the rows too, so a portrait page holds more of
  them. The pagination is now `_paginate_pool_blocks`, shared by the builder
  and callable against an already-built sheet.
- **Env Factor by Pool** (3) - the merged tab is the env table plus the ranges
  table plus the description paragraphs; forcing one page would shrink the body
  text past readability.
- **Display HIst Bal** (2) - that is its declared intent.

Latent, not yet hit: `Change Analysis` carries the same stale zoom and reports
two page-break columns, but printed on one page because Mountain CU has no
prior report and the tab is nearly empty. It sets no page setup of its own in
`change_analysis.py`. Worth a look once a CU with real change-analysis content
is rendered.

## Impr Deter chart tuning

Requested 2026-08-31 off the preview PDF. Three items; one turned out to be a
preview artefact.

1. **Axis percentages above the chart titles - already fixed in production.**
   They appear only because the preview harness skipped
   `patch_impdet_charts`, which sets `valAx tickLblPos=none` via
   `_fix_valax_bar`. Confirmed in the patched chart XML. No source change was
   needed.
2. **Title sitting on the bars** - the title is drawn with `overlay=1`, so the
   plot area has to move rather than the title. `_sheet_impdet`'s bar-chart
   `ManualLayout` now uses `y=0.22, h=0.632` (was `y=0.124, h=0.729`); the plot
   starts lower and still ends where it did.
3. **Not enough white space under the top charts** - `anc0`/`anc3` moved from
   rows 25-39 to 26-40. Same 14-row height, one blank row gained.

### Gotchas worth remembering

- **openpyxl silently drops `y_axis.tickLblPos = 'none'` on a value axis.**
  `_sheet_impdet` sets it and it does not survive; this is exactly why
  `patch_impdet_charts` exists as a post-save XML pass. Do not "fix" the
  builder line - it cannot work.
- **`patch_impdet_charts` assumes openpyxl-authored chart XML.** Run against a
  workbook whose charts Excel wrote (such as the reference file the preview
  harness starts from) it produces a file Excel refuses to open. Production is
  unaffected: it only ever patches what `compose_vizo_main` just built.
- **Excel rewrites chart XML on every save.** Any post-save XML surgery must
  come *after* the recalculate-and-save pass, with the final PDF exported
  read-only, or the save silently reverts it.

## Regression found by the first real pipeline run (2026-08-31)

Running March then June for Mountain CU produced workbooks **Excel refused to
open**. The preview harness never caught it: it starts from a reference
workbook and skips `patch_impdet_charts`. Two distinct bugs, both now fixed.

### 1. `patch_impdet_charts` corrupted the drawing (latent, pre-existing)

`patch_impdet_charts` starts with `ET.register_namespace('', _C_NS)` so chart
XML serialises with the chart namespace as default. The **drawing** was then
serialised under that same mapping, and its default namespace is
spreadsheetDrawing -- so the chart reference came out unprefixed:

| | chart reference in `xl/drawings/drawing2.xml` |
|---|---|
| good | `<c:chart r:id="rId1"/>` |
| broken | `<chart r:id="rId1"/>` |

Unprefixed, it binds to the drawing namespace, Excel cannot resolve the chart,
and the whole workbook is rejected. The fix registers the drawing's own
namespaces (crucially `c`) around that one serialisation and restores after.

This bug was always there. It only fires when `drawing_changed` is true, and
that only happens when the chart anchors differ from what
`_normalize_impdet_anchor` wants -- which nothing had done until item 3 below.

### 2. The patcher silently overrode the builder's chart anchors

`_normalize_impdet_anchor` hard-coded `top_row = 11 if from_row < 25 else 25`,
so the "shift the bottom charts down" change made in `_sheet_impdet` was
normalised straight back out. Anchor rows are now the module constants
`IMPDET_TOP_ROW = 11` / `IMPDET_BOTTOM_ROW = 26`.

**Any future change to the Impr Deter chart grid must be made in both places** --
the builder and `_normalize_impdet_anchor` -- or the patcher wins and (before
fix 1) corrupted the file on the way.

### 3. Summary Variance never found its prior report

`_sheet_summary_variance` resolved the reports folder from
`config['report_dir'] / config['output_dir']`, neither of which production
sets, so it fell back to `'.'` and always reported "no prior report available"
while the Change Analysis tab on the same workbook found March fine. It now
resolves `CECL_WORKSPACE_ROOT` + `Reports` exactly as `change_analysis` does.
This is precisely the coupling flagged in the per-CU publish paths backlog.

### Verified

March and June 2026 regenerated for Mountain CU; both open in Excel (24
sheets); Change Analysis and Summary Variance both populated from the real
prior quarter.

## Remaining open question

`report_tct.py` still appends `Change Analysis` last, and there is no TCT
reference workbook. Should the TCT model get the same reposition (after
`ACL Env by Pool Mgmt Adj`), or is TCT's order staying as-is?

---

## Backlog: per-credit-union publish paths

Requested 2026-08-31, to start **after** the Migration layout work is finished.
Today every generated report lands in one shared folder
(`Z:\Shared\TCT Files\CECL - CM Files\Reports`) and gets moved by hand
afterwards. The goal is a configured destination per CU so nothing needs moving.

Not started. Notes for whoever picks it up:

- **`_find_prior_report` is the coupling to watch.** It globs a single
  directory for `*_CECL_Migration_<safe_cu>_<suffix>.xlsx`
  (`change_analysis.py`). Both the Change Analysis tab and the new Summary
  Variance Prior block depend on it. Once each CU has its own folder, the
  lookup must resolve *that CU's* folder, not a global one - otherwise both
  tabs silently lose their prior comparison and degrade to "no prior report"
  without erroring, which is easy to miss.
- The directory already flows through config: `report_vizo.py` reads
  `config['report_dir'] or config['output_dir']`. A per-CU value would ride
  the same path, so the plumbing largely exists.
- `CECL_WORKSPACE_ROOT` (see `cecl_ui/app.py`) is the existing root-override
  mechanism and is the natural thing to hang per-CU paths off, rather than a
  new independent setting.
- `cecl_retention.py` assumes a single flat `REPORTS_DIR` when ageing files
  out; it would need to walk per-CU folders.
- `pipeline_service.py` writes WARM baselines to `Reports/_warm_baselines`;
  decide whether that stays global (probably yes - it is not a deliverable).
