# Chart chassis spike — A2 and A5 (Step 1c)

Done 2026-08-31. Follows `03_chart_inventory.md`, which nominated these two as
the spike. Both are built, both render from live cells in a delivered workbook,
and the preview is on the share:

    Z:\Shared\TCT Files\CECL - CM Files\Reports\chart_spike.html
    Z:\Shared\TCT Files\CECL - CM Files\Reports\chart_spike.pdf   (5 pp, landscape)

Regenerate either with:

    python -m cecl_report_web.spike_preview [<report.xlsx>] [<out_dir>]

Code added — all inside `cecl_report_web/`, nothing in the repo root touched:

| File | Lines | What |
|---|---:|---|
| `chart_chassis.py` | ~1070 | The chassis + the A2 and A5 renderers |
| `spike_preview.py` | ~360 | Pulls real data, builds the preview HTML/PDF |
| `charts.py` | +45 | `_to_chassis_spec()` routes A2/A5 off the buggy legacy path |

---

## Bottom line

- **The signed baseline is solved**, and it was not the hard part. The fix is
  `split_stack()` — about twelve lines. Most of A2's estimated cost was Excel
  workarounds that simply do not need porting.
- **A3 fell out for free**, verified, with *zero* new code: `Net Change` renders
  correctly through the same `render_diverging_stacked_bar` call, negatives left,
  positives right. That was predicted; it is now confirmed rather than assumed.
- **The chassis is ~80% of the work and the archetypes are ~20%.** 1,070 lines of
  shared machinery carry two renderers of ~90 and ~80 lines each. This is the
  right shape, and it means the remaining rectilinear archetypes are close to
  configuration.
- **Two cost items the original estimate missed entirely** — no font metrics were
  available for measured label placement (solved: parsed `calibri.ttf` directly),
  and **the Vizo brand palette does not pass a colour-vision validation** (solved,
  but it needs Brian's sign-off).
- All 40 charts in the workbook still render, 0 failures: 10 now go through the
  chassis (1 × A2, 9 × A5), 30 stay on the legacy path untouched.
- **Revised estimate for the remaining five archetypes plus integration:
  ~40–46 h**, against the ~51–58 h the original ~100 h figure implied for the
  same remainder. Detail in §5.

---

## 1. What the chassis provides

`cecl_report_web/chart_chassis.py`. Nothing in it knows about openpyxl or
workbooks — it consumes a plain spec dict, so the same renderers work against a
parsed `.xlsx` today and against `ReportData` after the Step 1a extraction.

| Piece | API | Notes |
|---|---|---|
| **Scales** | `LinearScale`, `BandScale` | Sign-agnostic: the domain may straddle zero and the pixel range may be reversed, so one class serves an upward column axis, a rightward bar axis and a diverging axis. `.zero` is the property every signed chart hangs off. |
| **Axis engine** | `nice_ticks`, `nice_step`, `y_axis`, `x_category_axis` | 1/2/2.5/5/10 steps, optional `symmetric=True` domain, hairline gridlines, tick rules. |
| **Number formatting** | `fmt_pct`, `fmt_currency`, `currency_unit`, `axis_formatter`, `tick_formatter`, `decimals_for_step` | One unit chosen for a **whole** axis, not per tick. Falls through to `format.excel_format` for a raw Excel mask. |
| **Text measurement** | `text_width`, `ellipsize`, `CALIBRI_EM` | **Real** Calibri advance widths (99 glyphs, em-normalised) parsed out of the bundled `static/fonts/calibri.ttf`. |
| **Label placement** | `place_bar_label`, `resolve_1d_collisions` | Inside-vs-outside by measurement; 1-D collision resolution for tick and leader labels. |
| **Signed stacking** | `split_stack`, `stack_extent`, `effective_values` | §2. |
| **Box model** | `Frame` | Margins computed from measured content, not hard-coded. |
| **Chrome** | `header`, `legend`, `legend_width`, `max_tick_width` | Legend always present for ≥2 series. |
| **SVG primitives** | `esc`, `svg_open/close/text/rect/line` | Deterministic, compact output — golden-SVG regression stays stable. |
| **Palette** | `PALETTE`, `SEMANTIC`, `THEME` | §4. |

The spec dict shape is documented in the module docstring and is the same for
every archetype: `kind`, `title`, `subtitle`, `categories`, `series[]`,
`value_format`, `axis_title`, `width`, `height`, `options{}`.

### Two things the chassis fixed that were not on anyone's list

**No font metrics existed.** "Measured label placement" was in the plan, but
`fontTools` is not installed and nothing in the repo could measure a string.
Rather than add a dependency or guess, the spike parses the `hmtx`/`cmap` tables
of the bundled `calibri.ttf` directly and bakes the resulting table into the
module as a constant. Runtime stays dependency-free. Budget ~0 h going forward —
it is done — but it is exactly the kind of item that quietly eats a day.

**`decimals_for_step`, found by accident.** Rendering A3 produced an axis
labelled `2%, 5%, 8%, 10%` for ticks actually spaced 2.5 points apart — Python's
round-half-to-even turns `2.5` into `"2"` and `7.5` into `"8"`, so an evenly
spaced axis reads as an uneven one. Any 2.5-mantissa step on any percent or
currency axis in the deck would have shipped wrong. Tick precision is now derived
from the step; the per-bar labels keep the data's own precision. So A5's axis
reads `$0 / $1M / $2M / $3M / $4M` while its columns are still labelled `$3.9M`.

---

## 2. How the signed baseline works

The workbook stores the "Deteriorated" series **negative** (`H46 = -0.0345`), and
Excel draws a diverging tornado as a side effect of stacking mixed signs. The old
renderer took `abs(val)` and accumulated into one running total:

```python
acc = 0
for v in row:                   # cecl_report_web/charts.py, pre-spike
    w = abs(v) / vmax * pw      # sign discarded here
    draw(acc, acc + w)
    acc += w                    # +6.5% and -3.5% ADD to a 10% bar
```

Consumer Unsecured rendered as a single **22.3%** bar that exists nowhere in the
data (15.2% + |−7.1%|). It looked entirely plausible. The left-hand pane of the
preview shows it.

The fix is **two independent accumulators**:

```python
def split_stack(row):
    pos = neg = 0.0
    out = []
    for v in row:
        if v >= 0:  out.append((pos, pos + v, v)); pos += v
        else:       out.append((neg, neg + v, v)); neg += v
    return out
```

Positives grow rightward from zero, negatives leftward, through accumulators that
can never contaminate each other. It generalises to *n* segments per side, which
is what makes it reusable rather than a two-series special case:

```
>>> split_stack([3.0, 2.0, -1.0, -4.0])
[(0.0, 3.0, 3.0), (3.0, 5.0, 2.0), (0.0, -1.0, -1.0), (-1.0, -5.0, -4.0)]
```

Three further decisions make it a correct chart rather than merely a signed one:

1. **The scale must be symmetric.** `nice_ticks(..., symmetric=True)` forces
   `−M .. +M`, so 7.1% on the left is exactly as long as 7.1% on the right. This
   is not optional — an asymmetric tornado is the classic version of the lie, and
   Excel could not be checked for it because it draws **no value axis at all**.
2. **Axis labels are magnitudes.** `20% 15% 10% 5% 0 5% 10% 15% 20%`, with
   direction carried by side plus legend, as in a population pyramid. A "−7.1%"
   tick would be false: the value is a *share*, and the minus sign is a rendering
   trick, not arithmetic.
3. **The renderer must not trust the feed's sign convention.** Each series may
   declare `direction: +1 / -1`; `effective_values()` then takes the magnitude
   from the data and the *side* from the spec. So a future `ReportData` supplying
   Deteriorated as a positive number renders identically to today's negative
   cells. Without this, the Step 1a data extraction would silently flip this
   chart.

### Excel workarounds that were dropped, not ported

This is where A2's estimate actually went, and most of it evaporated:

| Excel feature | Disposition |
|---|---|
| `valAx orientation="maxMin"` (positives grow **leftward**) | **Dropped.** A reader-hostile artifact. Positives now grow right. |
| `catAx axPos="r" tickLblPos="high"` (pool names down the right edge) | **Dropped.** Names go on the left, in reading order. |
| `manualLayout` on the plot area (`x .0393 y .22 w .9214 h .6324`) | **Dropped.** Pure compensation for an `overlay="1"` title; CSS/SVG layout handles it. |
| 16 per-point `<c:dLbl>` overrides + the `T_BAR = 0.30` ratio | **Dropped.** Replaced by one `place_bar_label()` call. The ratio mis-fires whenever text length varies — it treats `"0.0%"` and `"(1,234,567)"` identically. |
| `inEnd` fallback for small segments | **Dropped.** That existed only because Excel forbids `outEnd` on stacked bars. SVG has no such constraint, so small segments get a real outside label. |
| `catAx bodyPr rot="-60000000"` (−1000°, outside the legal range) | **Dropped**, as `03_chart_inventory.md` instructed. |

Added beyond Excel: a real value axis, gridlines, and an **explicit empty state**
("no migration") for the three all-zero pools that Excel draws as invisible
nothing.

---

## 3. A5 — the value axis Excel deletes

Excel sets `<c:delete val="1"/>` on the value axis, `noFill` on the gridlines and
no data labels, so nine charts plot millions of dollars with no number anywhere
on them. This was correctly called the biggest information gap in the deck.

The chassis adds nice ticks, hairline gridlines, currency labels with one unit per
axis, a left margin measured from the widest tick label, and per-column value
labels drawn **only where they measurably fit** (`text_width(txt) <= bar_w + gap`)
— so the axis carries the rest instead of the chart turning into a wall of
numbers. A zero balance gets a visible baseline tick, because "Platinum improved
= $0" is information, not missing data.

The same call handles every tab: `Risk Change Total` scales to `$0–$4M` and picks
the M unit; `Risk Chg Consumer Unsecured` scales to `$0–$600K` and picks K, with
no configuration change. Both are in the preview.

**Open design question for Brian** (in the preview, side by side): Excel's columns
are hollow 3 pt outlines. The chassis defaults to **filled**, which reads
markedly better at report size; `options={"outline": True}` keeps the hollow look
with the new axis. This needs a decision before the other eight instances are
signed off — it is one of the "no reference to converge on" items the plan warned
about.

---

## 4. The palette does not validate — a real finding

The Vizo brand hexes used by the Excel charts fail an objective colour-vision
check (run with the `dataviz` skill's `validate_palette.js`):

- teal `#0D4D5E` — **below the chroma floor** (0.066; reads gray) and outside the
  lightness band;
- olive `#829901` vs amber `#FFC000` — **ΔE 1.5 under deuteranopia**, and ΔE 9.5
  even with *normal* colour vision. Effectively the same colour to many readers.

The chassis ships the same four hues re-stepped to pass — teal `#0E7E9E`, maroon
`#B4453F`, amber `#E0A400`, olive `#6E8A00`. In that order every check passes
except one: amber sits at 2.16:1 against white, which obliges a visible label or
a table view wherever amber is used. Only A6/A7 reach the amber slot and both
label every slice, so the obligation is met — but re-run the validator if that
changes.

Separately: **the Excel originals fill the series named "Improved" with MAROON
and "Deteriorated" with TEAL** ("per Brian's edit"). That inverts the normal
reading — cool/positive for the good outcome, warm for the bad one. The port uses
the semantic mapping. **This is a deliberate visible change and needs
confirming.**

Both items are palette *decisions*, not code, and they are the sort of thing that
turns into a rework cycle if discovered late. They are now concrete proposals
Brian can react to rather than blank-page research.

---

## 5. Revised estimate

Everything below is *remaining* work, from where the repo now stands.

| Item | Orig. | Revised | Why it moved |
|---|---:|---:|---|
| Shared chassis — rectilinear half | 20 | **done** | Scales, ticks, formatting, metrics, labels, palette, frame, chrome. |
| Chassis — **radial half** (new line) | — | **6** | Arc paths, explosion offsets, leader lines, on-circle label collision. Not built; A4/A6 need it and nothing in the spike exercised arcs. |
| A2 — diverging stacked bar | 14 | **done** | Real cost was ~6 h once the chassis existed; the rest was Excel workarounds that were dropped, not ported. |
| A5 — hollow clustered column | 8 | **done** | ×9 instances, one call. |
| A3 — reversed bar w/ negatives | 6 | **0.5** | **Verified free.** Same renderer, one mixed-sign series, zero new code. Budget is polish only. |
| A1 — graduated column | 6 | **3** | Alpha ramp is `fill-opacity`. Rotated in-bar label + small-bar rule are `place_bar_label` + a rotate, both in the chassis. |
| A4 — exploded hollow doughnut | 5 | **4** | Geometry is cheap once the radial chassis exists; cost is the `holeSize=10` design decision. |
| A6 — exploded pie + labels | 10 | **7** | Leader lines and collision are the work. **Still blocked** on the all-zero DQ feed — the spike did not resolve that, and shouldn't; it is an upstream data question. |
| A7 — hollow horizontal bar | 4 | **2** | Simplest; A2's machinery covers it. Same data blocker. |
| Template + assembly integration | 8 | **8** | Unchanged and **untested**. See §6 — this is now the largest single unknown. |
| Visual-regression harness | 4 | **3** | Output confirmed deterministic; golden SVGs are cheap. |
| Design pass & review round | 8 | **6** | Palette and filled-vs-hollow now have concrete proposals attached, not open questions. |
| **Remaining total** | | **~40 h** | |
| Contingency (A6 + integration) | | **+6** | |
| **Quote for the remainder** | | **~40–46 h** | |

For comparison, the original ~100 h budgeted ~42 h for the chassis + A2 + A5 and
~51–58 h for everything else. The revised remainder is **~40–46 h**, i.e. the
back half came in a little cheaper than feared, mostly because A2 was
over-estimated and A3 is genuinely free.

**If the DQ/CO feed is confirmed dead**, A6 and A7 collapse to a shared empty
state: −6 h → **~34–40 h**.

Confidence: the archetype lines are now good to about ±15%, as the plan hoped.
The integration line is not — see below.

---

## 6. What is still unknown

1. **Integration is untested.** The spike renders into a standalone preview page,
   not into `impr_deter.html` / `risk_change.html`. Chart sizing inside the real
   page grid, and page-fit in Chromium, are unproven. This is now the largest
   single unknown in the chart estimate, and it is the one line that did not move.
2. **The radial archetypes are untouched.** No arc code exists. The 6 h is an
   estimate against an unbuilt component, not a measurement.
3. **The all-zero DQ/CO question is unresolved** and still gates A6/A7 — 18 of the
   40 charts. `03_chart_inventory.md` recommended timeboxing it and taking the
   question to whoever owns `hist['impaired']['dq_by_pool']`. **That should happen
   now**; it is worth ~6 h of scope and it does not depend on any chart work.
4. **Two design decisions need Brian**: filled vs hollow columns (§3), and the
   maroon/teal series inversion (§4).
5. **The legacy path still has the sign bug** for anything the chassis does not
   yet claim. `_to_chassis_spec()` routes only `bar`+`stacked` and multi-series
   `col`+`clustered`. A3 (`bar`+`clustered`, single series) still renders through
   the old `abs()` path in the live prototype even though the chassis renders it
   correctly — wiring that up is part of the 0.5 h A3 line.

## 7. Verification performed

- All **40 charts** across all 24 sheets of the Mountain CU workbook render
  through `render_charts_for_sheet()` with **0 failures**; 10 via the chassis
  (1 × A2, 9 × A5), 30 unchanged on the legacy path.
- Doctests in `chart_chassis` pass (`split_stack`, `decimals_for_step`).
- Tick-engine edge cases exercised: empty, zero-span, all-negative, symmetric,
  1e-4 and 1e9 domains.
- Sign regression asserted directly: `split_stack` puts the two series on
  opposite sides, and `stack_extent` returns ±the individual magnitudes rather
  than their sum.
- `direction: -1` on positive input asserted equal to negative input.
- HTML and PDF both rendered and **visually inspected** (page 1 of the PDF was
  initially near-blank from `break-inside: avoid` on a too-tall section; the print
  CSS was fixed and re-checked).
- `report_vizo.py` and `generate_report.py` **not touched** (confirmed by mtime —
  both predate the first write of this session).
- All three files valid Python and pure CRLF, matching repo convention.
