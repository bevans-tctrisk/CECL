# Chart Inventory & SVG Feasibility — CECL Migration report

Evidence base: `Z:\Shared\TCT Files\CECL - CM Files\Reports\2026-06-30_CECL_Migration_Mountain_CU_Vizo_Model.xlsx`
(unzipped `xl/charts/chart1..40.xml`, cross-referenced through `xl/worksheets/_rels` →
`xl/drawings/*.xml.rels`), plus `report_vizo.py` (chart builders ~L1230–1600, L1960–2185;
patchers `_apply_graduated_transparency` L4926, `patch_dq_pie_zero_labels` L5200,
`patch_impdet_charts` L5339) and the existing prototype `cecl_report_web/charts.py`.
Two other delivered workbooks (Nucor Emp CU, SCI FCU) were sampled for data-shape confirmation.

---

## Bottom line

- **40 charts collapse to 7 archetypes.** 36 of the 40 are four designs repeated across
  nine "Risk Change" tabs; only 4 charts (the `Impr Deter` 2x2 grid) are one-offs. Build
  seven renderers and the whole deliverable is covered.
- **The hard part is three charts, not forty.** One archetype (A2, the Improved/Deteriorated
  stacked bar) is genuinely hard: it is a **stacked bar with a negative series and both axes
  reversed** — i.e. a diverging tornado chart that Excel happens to draw by accident. The
  current prototype renders it **wrong** (it takes `abs()` and stacks in one direction, so
  Improved and Deteriorated add up instead of opposing). A3 has the same sign bug.
- **18 of the 40 charts plot all zeros in every workbook checked.** The DQ pie (A6) and the
  charge-off bar (A7) — two per Risk Chg tab, nine tabs — have `0` in every source cell in
  Mountain CU, Nucor Emp CU and SCI FCU. `patch_dq_pie_zero_labels` then injects
  `<delete val="1"/>` on all four slices, so Excel shows a blank frame with a title.
  **This is an upstream data question (`hist['impaired']['dq_by_pool'/'co_by_pool']` is
  empty), not a rendering question, and it should be answered before either archetype is
  built.** If the feed is dead, in-scope work drops by roughly a quarter.
- **Nothing here needs a charting library.** Every archetype is rectangles, arcs, text and
  a linear scale. No 3-D, no scatter/bubble, no secondary axis, no trendlines, no
  error bars, no theme-colour (`<a:schemeClr>`) references — all fills are literal `srgbClr`
  hex from the Vizo palette. Hand-rolled SVG is the right call and the prototype already
  proves the data-extraction half.
- **Rough effort: ~90–100 hours** (approx. 2.5 dev-weeks) for all seven archetypes plus a
  shared chart chassis, template integration, a Chromium PDF pass and a visual-regression
  harness. ~40 h of that is the shared chassis + integration + design pass, paid once.

---

## 1. Inventory

Sheet mapping is derived from the drawing relationships, so it is authoritative.

| chart*.xml | Worksheet | Anchor | Type / grouping | Plots | Archetype |
|---|---|---|---|---|---|
| 1 | Impr Deter | R12C1 | barChart col / clustered, gap 100 | `Improved %` by credit grade (C72:C77 → F72:F77), top grade excluded | **A1** |
| 2 | Impr Deter | R12C6 | barChart col / clustered, gap 100 | `Deteriorated %` by credit grade (C71:C76 → G71:G76), bottom grade excluded | **A1** |
| 3 | Impr Deter | R27C1 | barChart **bar / stacked**, overlap 100, gap 10 | Improved (+) and Deteriorated (−) share of balance, by loan pool (F46:F56) | **A2** |
| 4 | Impr Deter | R27C6 | barChart bar / clustered, overlap 100, gap 10 | Net Change % by loan pool, mixed sign (I46:I56) | **A3** |
| 5, 9, 13, 17, 21, 25, 29, 33, 37 | Risk Change Total + 8 `Risk Chg <pool>` | R33C3 | doughnutChart, holeSize 10, explosion 16 | Improved / Deteriorated / Unchanged share (B37:B39 → D37:D39) | **A4** |
| 6, 10, 14, 18, 22, 26, 30, 34, 38 | same nine tabs | R30C6 | barChart col / clustered, gap 100, **overlap −10** | Deteriorated vs Improved balance by grade (A7:A13 → M7:M13, N7:N13) | **A5** |
| 7, 11, 15, 19, 23, 27, 31, 35, 39 | same nine tabs | R47C1 | pieChart, explosion 21 | Delinquent balance % by migration status (M48:M51 → O48:O51) | **A6** |
| 8, 12, 16, 20, 24, 28, 32, 36, 40 | same nine tabs | R47C6 | barChart bar / clustered, gap 100 | Charge-off balance % by migration status (M54:M57 → O54:O57) | **A7** |

The nine tabs carrying the repeated set are: `Risk Change Total`, `Risk Chg Consumer Auto
Loan-N`, `... Loan-U`, `... Indirect Au`, `... Indirect Au1`, `... Consumer Secured`,
`... Consumer Unsecured`, `... Credit Cards`, `... Real Estate`. Pool count varies by credit
union (Nucor = 7 tabs, SCI = 7 tabs), so **the chart count is data-dependent**: 40 here,
32 elsewhere. Any estimate keyed to "40 charts" is keyed to the wrong number — key it to
seven archetypes.

**Adjacent, not in this workbook:** the `CECL_Supplemental_*` workbook contains 8
`lineChart` charts (one multi-series grade-balance time series per pool, built at
`report_vizo.py:3943`). `Display HIst Bal` in the Migration workbook has no drawing part
at all in these deliveries. If the Supplemental report joins the same PDF pipeline, budget
archetype **A8** below.

---

## 2. Archetypes, Excel-specific tricks, and SVG difficulty

### A1 — Graduated-transparency column (charts 1, 2 · 2 instances)
**What it is:** single-series column chart, one bar per credit grade, value axis fully
deleted, percentage printed *inside* each bar rotated 90 degrees.

Excel-specific features relied on:
- **Per-point alpha ramp.** `_apply_graduated_transparency()` writes six `<c:dPt>` elements
  with the same `srgbClr` (`0D4D5E` teal / `3D1A1A` maroon) and `<a:alpha>` stepping
  `100000 → 85000 → 70000 → 55000 → 40000 → 25000`. The bars are the same hue at falling
  opacity over a white plot.
- **Value axis deleted** (`<c:delete val="1"/>`) plus `majorGridlines` with `<a:noFill/>` —
  there is no numeric scale; the labels *are* the scale.
- **Rotated data labels** `rot="-5400000"` (−90 degrees), white 9 pt, `dLblPos="inEnd"`.
- **Per-point label overrides.** `_dlbl_point()` + the `T_COL = 0.22` heuristic: any bar
  shorter than 22% of the tallest gets a `<c:dLbl>` override flipping it to `outEnd` in
  **black**, because white text on a short bar would spill onto the white background and
  disappear. chart2 carries 3 such overrides; chart1 carries 0.
- Category-axis `bodyPr rot="-60000000"` — that is −1000 degrees, outside the legal
  ±5400000 range. Excel silently normalises it. **Do not port this.**

**SVG difficulty: Low.** Alpha-over-white is exactly `fill-opacity`, or precompute the
composite hex. The only real work is the label-placement rule, which should be replaced by
a proper *measured* rule (measure the text, compare to bar height) rather than the 22% guess.
The existing prototype approximates the ramp with `_lighten(base, i/(n-1)*0.68)` — close,
but it should read `<a:alpha>` and be exact.

---

### A2 — Diverging stacked bar, reversed axes (chart 3 · 1 instance) — HIGHEST RISK
**What it is:** nominally a horizontal *stacked* bar of Improved and Deteriorated by pool.
In reality the Deteriorated series is **negative** (`-0.0345`, `-0.0412`, ...), so Excel
draws the two series on opposite sides of a zero baseline — a tornado/diverging chart.

Excel-specific features relied on:
- **Stacked grouping with mixed signs.** `overlap=100`, `gapWidth=10`.
- **Both axes reversed** — `catAx` *and* `valAx` carry `<c:orientation val="maxMin"/>`.
  The category reversal is deliberate (so pool #1 appears at the top, matching every other
  tab). The value reversal means **positive values grow leftward**.
- `catAx axPos="r"`, `tickLblPos="high"` → pool names printed down the **right** edge.
- `valAx tickLblPos="none"` + gridlines `noFill` → **no numeric scale whatsoever**.
- **`manualLayout`** on the plot area (`x .0393  y .22  w .9214  h .6324`) purely to push
  the plot down clear of an `overlay="1"` title. This is layout the browser will handle for
  free; it should not be ported.
- **16 per-point data-label overrides** (`_c3_small`, `T_BAR = 0.30`) flipping small
  segments from white/`inBase` to black/`inEnd`. Note the comment in the source: Excel
  *forbids* `outEnd` on stacked bars, so the workaround is `inEnd` — a constraint that
  simply does not exist in SVG.
- Series-name/colour inversion by design: the series named "Improved" is filled MAROON,
  "Deteriorated" is TEAL ("per Brian's edit"). Confirm this is intentional before porting.

**SVG difficulty: High.** This is the one that will eat an estimate. It needs a signed,
centred value scale (both directions), category labels on the far side, sign-aware label
placement, and a design decision about what to do with the missing numeric axis. The
existing prototype's stacked path does `abs(val)` and accumulates, which produces a
*single-direction* bar of |Improved| + |Deteriorated| — **materially the wrong chart.**

---

### A3 — Reversed horizontal bar with negatives (chart 4 · 1 instance)
**What it is:** "Net Change" — one bar per pool, values from −0.021 to +0.084.

Excel-specific features relied on:
- `barDir=bar`, clustered but with `overlap=100`/`gapWidth=10` (single series, so overlap
  is inert — a copy-paste artefact of A2).
- Both axes `maxMin` again; `catAx axPos="r" tickLblPos="high"`; `valAx tickLblPos="none"`.
- Labels `outEnd`, black, `0.0%`. Chart-level `dLbls` are then blanked by
  `patch_impdet_charts()` so only the series labels survive.

**SVG difficulty: Medium.** Shares the whole signed-baseline machinery with A2, so build
A2 first and this is largely free. Standalone it would be Medium; sequenced after A2, Low.
Prototype has the same `abs()` sign bug — the −2.1% pool currently renders as a bar
pointing the same way as the +8.4% pool.

---

### A4 — Outline-only exploded "doughnut" (charts 5, 9, ... 37 · 9 instances)
**What it is:** three arcs — Improved / Deteriorated / Unchanged — drawn as **stroke only**.

Excel-specific features relied on:
- Every `<c:dPt>` is `<a:noFill/>` + `<a:ln w="38100">` (3 pt) in olive `829901`,
  maroon `873A3A`, teal `0D4D5E`. The wedges are hollow outlines, not filled slices.
- `explosion="16"` — slices pulled apart.
- `holeSize="10"` — **a 10% hole, so it renders as a near-solid pie, not a doughnut.**
  The builder sets `dc.innerRadius = 50`, but openpyxl's attribute is `holeSize`;
  `innerRadius` is silently ignored, so the intended 50% ring was never delivered. This is
  a live bug in the Excel output and a clear "the new one should look better" item.
- **No title, no legend, no data labels.** The chart on its own communicates only three
  relative arc lengths; all meaning is carried by the colour-matched table in columns B/D
  next to it. Any redesign must keep that pairing or fold the numbers into the chart.

**SVG difficulty: Low-Medium.** Arc paths with `stroke` / `fill:none`; explosion is a radial
translate of each wedge. The real work is a design decision (keep hollow-and-exploded, or
move to a proper labelled ring), not geometry. Prototype renders it as a *filled* pie with
a hard-coded 0.55 hole and an invented right-side legend — a different chart, though
arguably a better one.

---

### A5 — Outline-only clustered column, no scale (charts 6, 10, ... 38 · 9 instances)
**What it is:** Deteriorated vs Improved **dollar balances** by credit grade, two series.

Excel-specific features relied on:
- Both series `<a:noFill/>` + `<a:ln w="38100">` (3 pt) maroon / olive — hollow columns.
- `overlap="-10"` — a 10% gap between the paired columns.
- `valAx` `<c:delete val="1"/>` and **no data labels at all** → the reader gets bar heights
  and nothing else. These are millions of dollars with no axis, no gridline and no number.
  **This is the biggest information-conveyance gap in the whole deck** and the easiest win
  in the redesign.
- `catAx axPos="l"` on a *column* chart — invalid (should be `b`); Excel tolerates it.
- `legend` at `t` with a `manualLayout` using `wMode`/`hMode` = `factor`, a form that
  `_fix_manual_layout()` strips elsewhere. Layout the browser does for free.

**SVG difficulty: Medium.** Grouped columns are easy; the work is adding the value axis,
gridlines and tick formatting that Excel is missing, and doing it once in the shared
chassis. Prototype draws the grouped hollow bars correctly but has **no value axis or
gridline capability anywhere**, so the output is as uninformative as the original.

---

### A6 — Outline-only exploded pie with deleted labels (charts 7, 11, ... 39 · 9 instances)
**What it is:** delinquent balance % by migration status, four slices.

Excel-specific features relied on:
- Four `<c:dPt>` `noFill` + 3 pt strokes: olive / maroon / teal / amber `FFC000`.
- `explosion="21"`.
- `dLbls` with `dLblPos="outEnd"`, `showLegendKey="1"`, `showLeaderLines="1"`, `0.0%`.
- **`<c:dLbl><c:delete val="1"/></c:dLbl>` per zero slice**, injected by
  `patch_dq_pie_zero_labels()` via raw zip/XML surgery because openpyxl's `DataLabel`
  class has no `delete` support. In this workbook **all four** are deleted on all nine tabs.
- `manualLayout` shrinking the pie to 36% x 48% of the frame to make room for `outEnd`
  labels — a manual fix for a collision problem SVG should solve by measurement.
- Chart area explicitly solid white.

**SVG difficulty: Medium-High**, and almost all of it is uncertainty rather than code.
Leader lines plus outside labels plus collision avoidance on a four-slice pie is real work,
**but there is no real data to lay out**: `O48:O51` is `[0,0,0,0]` on every tab of all three
workbooks sampled. Until the DQ feed is fixed you cannot see, test, or sign off the design.

---

### A7 — Outline-only horizontal bar, per-point colour (charts 8, 12, ... 40 · 9 instances)
**What it is:** charge-off balance % by migration status; four bars, one colour each.

Excel-specific features relied on:
- Four `<c:dPt>` `noFill` + `<a:ln w="50800">` (4 pt) strokes, same four-colour set.
- `valAx delete="1"`; labels `inEnd`, `0.0%`.
- `catAx orientation="minMax"` here — **the opposite of A2/A3**, so the first category sits
  at the bottom. The reversal convention is inconsistent across the deck; pick one.
- Legend at bottom on a single-series chart, so it prints one meaningless entry.
- `manualLayout` inner `x .18  y .20  w .78  h .62`.

**SVG difficulty: Low.** Simplest archetype in the set. Same caveat as A6: all-zero data.

---

### A8 — Multi-series time-series line (Supplemental workbook · 8 instances, out of current scope)
Grade-level balance over time, one chart per pool, built at `report_vizo.py:3943` with
custom `ChartLines` gridlines, 20 pt Calibri titles and up to ~40 date categories. Listed
so it is not a surprise later; not part of the 40.

---

## 3. State of the existing prototype (`cecl_report_web/charts.py`, 456 lines)

Wired in through `assembly.py:61` (`render_charts_for_sheet`), rendered by `render.py`
through headless Chromium with bundled Calibri. The plumbing is real and works.

**What is already good:**
- `read_chart_specs()` walks `ws._charts` and resolves `<c:f>` references to actual values.
  Verified against the live workbook: every series, category, colour and number format on
  both `Impr Deter` and `Risk Change Total` resolves correctly, including negatives and the
  hollow-stroke colours (`filled=False` + stroke hex). **This is the fiddly half and it is done.**
- Correctly distinguishes fill vs stroke (`_fill_hex` / `_stroke_hex`), so the outline-only
  archetypes come through as `fill="none" stroke=...`.
- Number formats flow through `format.excel_format`, including the accounting mask.
- Deterministic, dependency-free output — good for a visual-regression harness.
- Empty-data guard on pies ("No data" ring), which is exactly what A6/A7 need.

**What is missing or wrong:**
1. **Sign handling.** Both the stacked and the horizontal-clustered paths use `abs(val)`.
   A2 renders as a same-direction stack instead of a diverging tornado; A3 loses the sign
   of every negative pool. *Correctness bugs, not polish.*
2. **No value axis, ticks or gridlines — at all.** Nothing in the file draws a scale. A5 in
   particular is unreadable without one.
3. **`scaling.orientation` (`maxMin`) is never read.** A2/A3 currently look right only by
   coincidence (the renderer happens to draw category 0 at the top). A7's opposite
   convention is not honoured either.
4. **Per-point alpha is dropped.** `_fill_hex` returns bare hex; `<a:alpha>` is discarded and
   approximated by `_lighten(..., i/(n-1)*0.68)`. Close, but derive it from the real value.
5. **Per-point `<c:dLbl>` overrides and `<c:delete>` are not read**, so both the
   small-bar-black-label rule (A1, A2) and the zero-slice suppression (A6) are absent.
6. **Pie/doughnut ignores `explosion`, `holeSize`, `firstSliceAng`, and stroke-only styling** —
   it draws filled wedges with a hard-coded 0.55 hole and its own right-hand legend.
7. **Horizontal clustered multi-series is unimplemented** — `s = series[0]`, silently
   dropping any second series. Not hit today; a trap tomorrow.
8. **Fixed 340 x 240 viewBox for everything.** Source charts are 15 x 7.5 cm and 20 x 7.5 cm
   (2:1 and 2.67:1). Aspect ratios will not match the page grid.
9. **Label placement is unmeasured** — fixed offsets and a `[:7]` truncation of long category
   names. Pool names like "Consumer Indirect Auto-Used Auto" will clip.
10. Legend position from `<c:legend legendPos>` is ignored (always bottom-ish or right).

Net: roughly **35–40% of the job is done**, and it is the least risky 35–40%.

---

## 4. Risk ranking and the de-risking spike

Ranked by effort x uncertainty:

| Rank | Archetype | Effort | Uncertainty | Why it is risky |
|---|---|---|---|---|
| 1 | **A2 — diverging stacked bar** | High | High | Negative-series stacking + double axis reversal + 16 per-point label overrides + no value scale. Currently rendered *incorrectly*. Nothing else in the deck exercises signed baselines. |
| 2 | **A6 — exploded pie, deleted labels** | Medium | **Very high** | Outside labels + leader lines + collision avoidance, on data that is **all zeros in every workbook available**. Cannot be validated or signed off until the DQ feed is resolved. |
| 3 | **A5 — hollow clustered column, no scale** | Medium | Medium | Needs the value-axis/tick/gridline engine that does not exist yet, and forces the "how much better than Excel?" design conversation. x9 instances, so any rework multiplies. |
| 4 | A1 — graduated column | Low-Med | Low | Alpha ramp is trivial; rotated in-bar labels and the small-bar rule need measurement. |
| 5 | A4 — exploded hollow doughnut | Low-Med | Medium | Geometry easy; the `holeSize=10` bug means there is no "correct" original to match — pure design decision. |
| 6 | A3 — reversed bar with negatives | Low | Low | Free once A2's signed scale exists. |
| 7 | A7 — hollow horizontal bar | Low | Low | Simplest. Blocked on the same zero-data question as A6. |

**Nominated spike — build these three first, in this order:**

1. **A2 (diverging stacked bar).** It forces the signed-scale, reversed-axis and
   per-point-label machinery into the shared chassis. If A2 lands cleanly, A1, A3 and A5
   are mostly configuration. If it does not, you learn it in week one instead of week three.
2. **A5 (hollow clustered column).** It forces the value-axis / tick / gridline /
   number-format engine — the single largest missing capability — and it is 9 of the 40
   charts. It is also the best "look how much better this is" demo, since the Excel version
   prints millions of dollars with no numbers on it.
3. **A6 (exploded pie).** Not because the code is hard, but because **the spike's real
   deliverable is the answer to "is this data ever non-zero?"** Timebox it: reproduce the
   chart against hand-fabricated non-zero data, then stop and take the question to whoever
   owns `hist['impaired']['dq_by_pool']`. If the answer is "the feed is dead", A6 and A7 —
   18 of 40 charts — become a one-line empty state and the estimate drops by ~12 h.

A1 + A3 + A4 + A7 can then be produced in a single batch against a proven chassis.

---

## 5. Effort estimate

### Assumptions
- One competent Python developer, comfortable with SVG geometry and CSS print layout, not
  necessarily prior to this codebase — allow ~4 h ramp inside the totals.
- **Reuses `read_chart_specs()` as-is** for extraction; the estimate covers fixing its gaps
  (alpha, orientation, per-point labels, `delete`) but not rewriting it.
- Output target is inline SVG rendered by the existing Playwright/Chromium path with
  bundled Calibri — that pipeline already works and is not re-estimated here.
- Success is "conveys the same information, looks better", **not** pixel parity. One design
  review round is budgeted; a second full redesign is not.
- Each archetype ships with a golden-SVG regression test and is eyeballed in the actual PDF.
- Excludes page layout / pagination / template work outside the chart boxes.
- Excludes fixing the DQ/CO data feed.

| Item | Hours | Notes |
|---|---:|---|
| **Shared chart chassis** | **20** | Linear + signed scales, axis/tick/gridline renderer, Excel number-format labels, measured text + label collision avoidance, legend component, palette, sizing/aspect from the source `width`/`height`. Paid once; everything else assumes it. |
| A1 — graduated column | 6 | Exact alpha ramp, rotated in-bar labels, measured small-bar rule. x2 charts. |
| **A2 — diverging stacked bar** | **14** | Signed baseline, reversed axes, right-side category labels, per-point label overrides. Highest variance: +/-5 h. |
| A3 — reversed bar w/ negatives | 6 | Mostly A2 reuse. |
| A4 — exploded hollow doughnut | 5 | Arc paths, explosion offsets, hole-size decision. x9 charts. |
| A5 — hollow clustered column | 8 | Grouped columns + the value scale Excel omits. x9 charts. |
| A6 — exploded pie + labels | 10 | Leader lines, outside-label collision, zero/empty state. +/-4 h; may collapse to 2 h if the feed is dead. |
| A7 — hollow horizontal bar | 4 | Simplest. x9 charts. |
| Template + assembly integration | 8 | Wire seven archetypes into `impr_deter.html` / `risk_change.html`, sizing and page-fit in Chromium. |
| Visual-regression harness | 4 | Golden SVGs per archetype x the three sample CUs; extends `regression.py`. |
| Design pass & review round | 8 | Required, since "better than Excel" is an explicit goal — palette, whether to keep hollow strokes and explosions, what replaces the missing axes. |
| **Total (in scope, 40 charts)** | **93 h** | approx. 2.5 dev-weeks. |
| Contingency on A2 + A6 | +8 h | Recommend quoting **~100 h**. |
| *A8 — Supplemental line charts (optional)* | *10 h* | *8 charts, ~40-point time series, only if the Supplemental report joins the pipeline.* |

### Estimate shape
- **Spike (A2 + A5 + A6 + the chassis they force): ~40 h.** After it, the remaining
  estimate should be accurate to +/-15%.
- **If the DQ/CO feed is confirmed dead:** drop A6/A7 to a shared empty state, −12 h → ~80 h.
- **The biggest schedule threat is not code**, it is the design round on A4/A5/A6. The Excel
  originals are information-poor (no axes, no labels, a doughnut with a 10% hole nobody
  intended), so there is no reference to converge on. Get the target design agreed during
  the spike, not after it.
