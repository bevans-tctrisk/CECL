# Prototype Audit — `cecl_report_web` (browser-rendered CECL PDF)

Audit date: 2026-08-31. Read-only review of `cecl_report_web/` (2,324 lines) against
the target architecture: **render the PDF from report DATA, never from the generated
`.xlsx`**.

---

## Bottom line

1. **Coverage is better than expected — and the headline number is misleading.** The
   prototype renders *every* tab of both the Vizo Model (24 tabs) and Vizo Supplemental
   (7 tabs) workbooks end-to-end. But it does so with **four bespoke page renderers plus
   one generic "faithful grid" transliterator** that handles 13 of the 24 main tabs. The
   generic grid is what makes the coverage number look complete, and it is exactly the
   part that **cannot survive** the move to a data model — it has no semantics, it just
   copies cells, fills, fonts and merges out of the sheet. Real *modelled* coverage is
   4 archetypes.

2. **`model.py` is a good contract for 4 pages and a placeholder for the rest.**
   `CoverPage`, `ImprDeterPage`, `RiskChangePage`, `AclEnvPage` are clean, render-agnostic
   dataclasses. Then `GridPage`/`GridCell` (`model.py:173-201`) smuggle presentation
   (hex fills, font sizes, colspans, pre-formatted text) into the "contract", and
   `ReportModel.pages` (`model.py:216-220`) is an untyped `dict[str, Any]` bag. About
   **150 of 220 lines are fit for purpose; ~70 lines are the escape hatch that must be
   replaced with real page nodes.**

3. **`from_workbook.py` (550 lines) is ~100% discardable** — there is no "pure data"
   loader in it. Every function is anchored to sheet names, cell coordinates, or *reads
   business meaning out of cell styling*. The worst instance: whether a migration cell is
   "improved" or "deteriorated" is decided by its **Excel fill hex**
   (`from_workbook.py:222-239`), and whether an ACL line is a subtotal is decided by
   **font bold** (`from_workbook.py:356-361`). Those two facts must become real fields on
   the data model.

4. **Roughly 850 of the 2,324 lines get discarded (~37%); ~1,470 are keepable (~63%)** —
   see the table in §3. The keepable half is the valuable half: Chromium plumbing, print
   CSS, the SVG chart engine's drawing half, the regression harness, and the four good
   templates. The coupling is **localised in two files**, not spread through the renderer.

5. **Real gaps for a production PDF deliverable:** no page numbering in the PDF path
   (the Excel path stamps `Page &P of &N`, `report_vizo.py:6407`); no
   `break-inside: avoid` control (the Excel path breaks between pool blocks,
   `report_vizo.py:6364`); hidden sheets are rendered into the PDF (`assembly.py:96`);
   one Chromium process is launched **per tab** (24+ cold launches per report,
   `assembly.py:130` → `render.py:115-116`); `playwright` and `pypdf` are **not in
   `requirements.txt`**.

6. **The prototype has been dormant since 2026-07-10** — 18 commits over two days, then
   nothing in ~7 weeks. It stopped at "Phase 5a", and the phase it never reached is
   precisely this migration: `from_workbook.py:4-7` says *"Later phases can add a
   compute-time populator with the same model as output — templates don't change."*
   One piece of work was deliberately deleted rather than left undone (§7).

---

## 1. What `cecl_report_web` covers today

### 1.1 Dispatch

Everything routes through `assembly._render_sheet` (`assembly.py:53-87`), which switches
on **lower-cased sheet-name substrings**:

| Match (`assembly.py`) | Renderer | Template | Orientation |
|---|---|---|---|
| `"cover" in low` (:66) | `fw.load_cover` | `cover.html` | portrait (forced, :69) |
| `"impr deter"` / `"improved"`+`"deteriorated"` (:70) | `fw.load_impr_deter` | `impr_deter.html` | portrait |
| `"risk change"` / `"risk chg"` (:74) | `fw.load_risk_change` | `risk_change.html` | landscape (forced) |
| `"acl env"` (:78) | `fw.load_acl_env` | `acl_env.html` | landscape (forced) |
| *anything else* (:80) | `fw.load_grid` | `grid.html` | landscape if the name matches `_LANDSCAPE_HINTS` (:26) **or** the sheet has >12 columns (:27, :82) |

Charts on any sheet are re-rendered as inline SVG and injected (`assembly.py:59-62`,
`charts.render_charts_for_sheet`).

### 1.2 Concrete tab coverage — Vizo Model

Verified against `Z:\...\Reports\2026-06-30_CECL_Migration_Mountain_CU_Vizo_Model.xlsx`
(24 tabs; tab set and order defined by `report_vizo._VIZO_MAIN_ORDER:6518-6536`).

**Bespoke ("modelled") renderers — 11 of 24 tabs, 4 archetypes:**

- `Vizo Cover` → `cover.html`
- `Impr Deter` → `impr_deter.html`
- `Risk Change Total` + 8 × `Risk Chg <pool>` → `risk_change.html`
  (per-pool routing added in `8c5b33c`; `from_workbook.py:295-303` accepts a `sheet` arg)
- `ACL Env by Pool Mgmt Adj` → `acl_env.html`

**Generic grid transliteration — 13 of 24 tabs, no model, no semantics:**

`Report Index`, `Summary Variance`, `Change Analysis`, `Impaired Loans`, `ACL Summary`,
`Mgmt Adj Summary`, `Env Factor by Pool`, `>Envir Fact Ranges` *(hidden!)*,
`Display HIst Bal`, `Display CO-Recov-DQ`, `Introduction-Vizo`, `Executive Summary-Vizo`.

### 1.3 Vizo Supplemental (7 tabs) — 1 bespoke, 6 generic

`Vizo Cover (2)` → `cover.html`; everything else goes through `load_grid`:
`Report Index (2)`, `> Historical Trends Balance`, `> Detail_HIst Balances`,
`>Detail_Charge off Hist`, `Pool_Balance Adjust`, `Appendix_Supplemental` *(hidden!)*.
Commit `8426726` declares "Supplemental renders end-to-end" — that end-to-end is 6/7 tabs
of pure cell transliteration.

### 1.4 What it **cannot** do

| Gap | Evidence |
|---|---|
| **No page archetype for any of the 13 grid tabs.** A data-driven renderer must invent ~9 new page models (index, variance, change analysis, impaired detail, ACL summary, mgmt-adj summary, env-factor table, historical/CO-recov trend tables, narrative). | `from_workbook.py:472-547` is the only "renderer" they have |
| **Impr Deter's migration tables are gone.** `by_pool` / `by_grade` are loaded into the model (`from_workbook.py:163-197`) but the template renders only the CECL box + charts (`impr_deter.html:13-29`). The tables were **deleted** in `698f1c8`; their CSS is still orphaned at `report.css:109-121` (`.mig-table`, `.by-pool`). | dead model fields + dead CSS |
| **Risk-change side panel (Det/Imp/Unch) dropped.** `MatrixRow.side` and `RiskChangeMatrix.side_headers` are always `[]` (`from_workbook.py:286-292`), justified as *"part of the workbook but NOT the printed report"* (`from_workbook.py:245-247`). | dead model fields |
| **Hidden sheets are rendered into the PDF.** `build_pages` iterates `wb.sheetnames` with no `sheet_state` filter (`assembly.py:96`), so `>Envir Fact Ranges` and `Appendix_Supplemental` are included. Excel's print job excludes them. | `assembly.py:94-104` |
| **No page numbers, no TOC/bookmarks, no running header/footer.** `render_pdf` never passes `display_header_footer`/`header_template` (`render.py:127-142`); the `pypdf` merge adds no outline (`assembly.py:128-135`). | vs. `report_vizo.py:6407 _add_page_numbers` |
| **No pagination control inside a tab.** `.page { page-break-after: always }` (`report.css:30-37`) is the only rule; no `break-inside: avoid`, no widow/orphan handling. | vs. `report_vizo.py:6364 _paginate_pool_blocks` |
| **No repeated table headers on the 13 grid tabs.** `acl_env.html:14-18` and `risk_change.html:15-26` use `<thead>` (Chromium *does* repeat those across print pages); `grid.html:6-17` emits bare `<tr><td>` with no `<thead>`. | vs. `report_vizo.py` `print_title_rows` at :2773, :3310, :4061, :4723 |
| **Per-tab failures are swallowed.** Any loader exception becomes a red one-line "Could not render tab" page (`assembly.py:84-87`) — the PDF still "succeeds" with a missing page. | silent-degradation risk for a client deliverable |
| **No xlsx *writer*.** `__init__.py:15-18` and `model.py:5-6` both promise "model → xlsx extract page"; **no such code exists**. If `.xlsx` stays a secondary output it must be written from scratch, or `report_vizo.py` keeps producing it independently. | absent |

---

## 2. `model.py` — shape and fitness as the contract

220 lines, 13 dataclasses, no imports beyond `dataclasses`/`typing`. The docstring states
the intent exactly right (`model.py:8-10`): *"Keep these dataclasses render-agnostic: no
HTML, no openpyxl, no formatting decisions."*

| Dataclass | Lines | Holds | Verdict |
|---|---|---|---|
| `CoverPage` | 19-33 | CU, period, title/subtitle, firm, disclaimer paragraph, footer, **two logos as `data:` URIs** | Good, except `top_logo`/`bottom_logo` are pre-encoded base64 strings — an asset-resolution decision leaking into the model. Should be a logo *identifier*. |
| `KeyValueRow` | 36-41 | label + `float \| None` | Clean, reusable |
| `PoolMigrationRow` | 44-51 | pool, improved, deteriorated, net_change (fractions) | Clean. **Currently unrendered.** |
| `GradeMigrationRow` | 54-61 | grade, balance, improved, deteriorated | Clean. **Currently unrendered.** |
| `ImprDeterPage` | 64-78 | CU, period, `heading_lines`, CECL box, by_pool, by_grade | Good shape. `heading_lines: list[str]` is a pre-composed display artifact, not data. **No chart node** — the page is chart-centric but its charts come from the workbook, not the model. |
| `MatrixCell` | 81-94 | value, **`state`**, `is_pct`, `bold` | `state` is documented as *"read straight from the workbook fill"* (`model.py:85-88`). The **semantics are right** (improved / deteriorated / plain / header) but the **provenance is wrong**. `is_pct` and `bold` are formatting, not data. |
| `MatrixRow` | 97-103 | label, cells, total, side, range_label | Good; `side` unused |
| `RiskChangeMatrix` | 106-114 | corner string, col_headers, rows, side_headers, is_pct | `corner` is a literal display string (`"$ Current Grade"`); should be a `unit` enum |
| `RiskChangePage` | 117-124 | CU, heading_lines, matrices, summary | Good |
| `AclPoolRow` | 127-148 | pool, `kind` (header/grade/total), 10 named numeric columns, `is_total` | **The strongest node in the file** — 10 explicitly-named allowance columns, exactly the compute engine's vocabulary. `kind` and `is_total` overlap (redundant). |
| `AdjustmentRow` | 151-157 | label, value, **`bold`** | `bold` is styling standing in for "this is a subtotal" |
| `AclEnvPage` | 160-170 | CU, heading, col_headers, pool rows, pooled totals, impaired rows, adjustment rows | Good. `col_headers: list[str]` is presentation. |
| `GridCell` | 173-191 | **`text` (already formatted), align, bold, italic, fill hex, color hex, size, colspan, rowspan, wrap** | **Not a data contract — a cell-styling struct.** The docstring admits it: *"styling is copied straight from the workbook"* (`model.py:177-179`). |
| `GridPage` | 194-201 | CU, sheet_name, rows of GridCell, landscape | `sheet_name` and `landscape` are workbook/print concerns |
| `ReportModel` | 204-220 | CU, period, report_type, cover, **`pages: dict[str, Any]`** | The top-level node is a stub: only `cover` is typed; every other page lives in an untyped bag (`model.py:216-220`). **Nothing in the codebase ever instantiates `ReportModel`** — `assembly.py` builds page objects directly. |

### Fitness verdict

**Keep and extend (~150 lines):** `CoverPage`, `KeyValueRow`, `PoolMigrationRow`,
`GradeMigrationRow`, `ImprDeterPage`, `MatrixCell` / `MatrixRow` / `RiskChangeMatrix` /
`RiskChangePage`, `AclPoolRow`, `AdjustmentRow`, `AclEnvPage`. These name real business
quantities and would be populated identically from `df`/`config`/`hist` as from a workbook.

**Replace (~70 lines):** `GridCell`, `GridPage`, `ReportModel.pages`.

**Specific fixes needed before it can serve as the contract:**

- **Purge presentation fields:** `MatrixCell.is_pct` / `.bold`, `AdjustmentRow.bold`,
  `AclPoolRow.is_total`, all of `GridCell`. Where they encode *meaning* (subtotal, unit),
  promote them to explicit semantic fields (`role: "subtotal"`, `unit: "pct"`).
- **`MatrixCell.state` must be computed** from (current_grade, original_grade) ordering,
  not read from a fill. It is the clearest example of business logic living in the
  spreadsheet's paint layer.
- **`heading_lines` / `col_headers` are pre-rendered strings** on four page types. They
  should be structured (CU, snapshot date, report title) with the template composing the
  display lines.
- **`ReportModel` needs typed slots** per page archetype (or a `list[Page]` tagged union)
  plus **report-level ordering** — tab order currently comes from `wb.sheetnames`
  (`assembly.py:96`) and must come from the model instead (`_VIZO_MAIN_ORDER:6518-6536`
  is the authority to port).
- **No node exists for charts.** A `ChartSpec` node (type, series, categories, colors,
  labels) is a mandatory addition. The `charts.py` SVG renderer already consumes a plain
  dict of exactly that shape (`charts.py:189-195`), so the node can be lifted straight
  out of `read_chart_specs`.

---

## 3. `from_workbook.py` coupling — function by function

Legend: **(a)** pure data that could come from anywhere · **(b)** presentation/styling
read out of cells · **(c)** workbook structure (sheet names, cell coordinates).

| Function | Lines (count) | (a) | (b) | (c) | Notes |
|---|---|---|---|---|---|
| `load_cover` | 39-84 (46) | ~5 | ~13 | ~28 | Hard-coded coordinates `A14/A16/A17/B21/B44` (:55-64); sheet found by `"cover" in name` (:46-48); **logos extracted from `ws._images` via `im._data()` and anchor-row ordering** (:66-78) — a private-API dependency on openpyxl. |
| `_find_tab` | 96-101 (6) | — | — | 6 | Substring match over `wb.sheetnames`. |
| `load_impr_deter` | 117-206 (90) | ~12 | 0 | ~78 | Tab by hint (:120); heading taken from rows 1-5 of column A (:132-142); CECL box found by **full-sheet O(rows×cols) label scan then "first numeric cell to the right"** (:146-161); by-pool/by-grade anchored on the literal strings `"Loan Type"` / `"Grade"` with fixed offsets `hc+1 … hc+4` (:163-197). The by-grade offsets skip `hc+2` — an undocumented layout assumption. |
| `_find_label` | 209-216 (8) | — | — | 8 | Full-sheet linear scan; called ~5× per report tab. |
| `_FILL_STATE` + `_fill_state` | 222-239 (18) | 0 | **18** | 0 | **The single worst coupling in the prototype.** Migration direction (improved / deteriorated / header) is decoded from the last 6 hex of `fill.fgColor.rgb` against a literal map `{"0D4D5E": header, "829901": improved, "873A3A": deteriorated}`. A theme change silently turns every cell "plain". |
| `_read_matrix` | 242-292 (51) | ~10 | ~4 | ~37 | Corner at `(hr, 1)`, score range at column 2, grade buckets from column 3; `"Grand Total"` column detected by scanning rows `hr` and `hr-1` (merged-header workaround, :251-260) with fallback `gt_col = 12` (:259-260). Row loop terminates on the literal strings `"Balance Adjustment"` / `"Total in Portfolio"` (:270). |
| `load_risk_change` | 295-342 (48) | ~12 | 0 | ~36 | **Loads the workbook twice** — `data_only=True` for values, `data_only=False` for styles (:301-302). Matrices located by scanning column A for the literals `"$ Current Grade"` / `"% Current Grade"` (:319-323). Summary values taken as *"the rightmost numeric cell on the row"* (:330-336). |
| `_bold` | 356-361 (6) | 0 | **6** | 0 | Font bold used as a data signal. |
| `load_acl_env` | 363-439 (77) | ~20 | ~6 | ~51 | Dual workbook load again (:365-366). **Fixed column map: balance=2, specific_id=3, llc=4, base_loss=5, mgmt_adj=6, factor=7, before_env=8, env_factor=9, env_allow=10, total=11** (:386-393). Section state machine driven by literal labels `"Pooled Totals"` / `"Impaired Loans"` / `"Total"` (:402-425); **row type inferred from "is column 2 empty?"** (:412, :419-422); impaired value pinned to *"column K"* (:427); adjustment vs. impaired split by label prefix + bold (:428-433). |
| `_hex6` | 444-457 (14) | 0 | **14** | 0 | openpyxl color → hex; **drops theme/indexed colors** it cannot resolve (a known fidelity hole). |
| `_align` | 460-469 (10) | 0 | **10** | 0 | Reproduces Excel's "General" alignment rule (numbers right, text left). |
| `load_grid` | 472-547 (76) | 0 | ~45 | ~31 | Pure transliteration: trims the extent by scanning every cell (:487-495); builds a merged-range → colspan/rowspan map from `wsf.merged_cells.ranges` (:497-506); per cell copies **number-formatted text, alignment, bold, italic, fill hex, font color, font size, wrap** (:508-543). Contains **zero** business semantics. |

**Classification summary for `from_workbook.py`: there is no function in category (a).**
Every loader is (c)-dominant with (b) contamination. Even the numeric extraction is
reached only through coordinate or label anchoring.

### Quantification: discarded vs. kept across the 2,324 lines

| Component | Lines | Fate in the target architecture |
|---|---|---|
| `from_workbook.py` | **550** | **DISCARD 100%.** No loader's logic survives; all are coordinate-, label-, or style-anchored. |
| `charts.py:1-196` — `_fill_hex`, `_stroke_hex`, `_split_ref`, `_resolve`, `_ref_numfmt`, `_chart_title`, `read_chart_specs` | **~196** | **DISCARD.** Mines `ws._charts` and resolves `'Sheet'!$A$1:$B$2` refs. Replaced by a `ChartSpec` node emitted by compute. |
| `charts.py:199-455` — `_wrap`, `_legend`, `_svg_bar`, `_svg_pie`, `render_chart_svg` | **~257** | **KEEP.** Already consumes a plain dict spec (`charts.py:189-195`); only the producer changes. |
| `assembly.py:26-42, 53-83, 90-106` — `find_report`, name-substring dispatch, `build_pages` over `wb.sheetnames` | **~60** | **DISCARD.** Dispatch moves to page-node type; ordering moves to the model. |
| `assembly.py` remainder — `_fragment`, `render_report_html`, `render_report_pdf` | **~76** | **KEEP** (needs the multi-launch fix, §4). |
| `format.excel_format:89-137` | **~48** | **DISCARD.** Exists only to interpret Excel `number_format` strings for the grid renderer. |
| `format.py` remainder — `acct0/acct2/pct0/1/2/4/mcell` | **~90** | **KEEP** — encodes the *report's* number conventions, not Excel's. |
| `model.py` | 220 | **KEEP ~150 / REPLACE ~70** (§2). |
| `render.py` | 151 | **KEEP 100%** — Chromium/Playwright plumbing, font embedding, `file://` origin trick. |
| `templates/`: `base.html`, `cover.html`, `impr_deter.html`, `risk_change.html`, `acl_env.html` | ~220 | **KEEP.** |
| `templates/grid.html` | 25 | **DISCARD** (or retain as a debug view). |
| `static/report.css` | 260 | **KEEP ~230** — grid rules (:211-234) and the orphaned `.mig-table` rules (:109-121) go. |
| `regression.py` | 147 | **KEEP** — only the fixture entry point changes (§5). |
| `__init__.py` | 21 | KEEP |

**Totals: ~851 lines discarded (≈37%), ~1,473 kept (≈63%).** The discarded third is
concentrated in **two files** (`from_workbook.py` and the front half of `charts.py`),
which is the good news: the coupling is localised, not spread through the renderer.

---

## 4. `render.py` (Playwright) and `assembly.py` (PDF merge)

### How Playwright is driven (`render.py`)

- `render_pdf` (`render.py:97-145`): `sync_playwright()` → `chromium.launch()` →
  `new_page()` → `page.pdf(...)` → `browser.close()`.
- **The `file://` trick** (`render.py:118-124`): the HTML is written to a temp file and
  loaded via `page.goto(fp.as_uri())` so the page origin is `file://` and the bundled
  `@font-face` TTFs (also `file://`) load same-origin. The comment at `render.py:53-61`
  records that multi-MB base64 data-URI fonts were *silently dropped* by Chromium's
  stylesheet parser — a hard-won detail worth preserving verbatim.
- Determinism steps: `page.evaluate("document.fonts.ready")` (:125) and
  `page.emulate_media(media="print")` (:126).
- **Fonts are bundled Calibri TTFs** (`render.py:26-32`; `static/fonts/*.ttf`, ~5.6 MB,
  gitignored). `render.py:22-25` flags the licensing issue and names **Carlito** as the
  metric-compatible server substitute.
- CSS is inlined into `<style>` (`render.py:79-88`, `base.html:9`), so the HTML is
  entirely self-contained — no static server, no base URL.

### Page-setup capabilities

| Capability | Status |
|---|---|
| **Per-page portrait/landscape** | **Yes** — but via `page.pdf(landscape=…)`, i.e. **one PDF render per orientation**, not per-page CSS. `render.py:131-135` documents *why*: setting `prefer_css_page_size` makes Chromium honour the CSS `@page size` (portrait) and silently ignore the `landscape` flag. `report.css:10-15` therefore declares only `@page { margin: 0 }`. |
| Page size / margins | `format="Letter"`, uniform `0.75in` margin (`render.py:100-103, 136-141`). Not per-page. |
| `print_background` | Yes (`render.py:103`), plus `print-color-adjust: exact` (`report.css:24-25`) — required for the teal/olive/maroon fills. |
| Explicit page breaks | Only `.page { page-break-after: always }` / `break-after: page` (`report.css:28-37`). |
| **`break-inside: avoid`** | **Absent.** Long tables (ACL Env pool blocks, Display HIst Bal) will split arbitrarily mid-block. |
| **Repeated table headers** | **Partial.** `<thead>` in `acl_env.html:14-18` and `risk_change.html:15-26` repeats natively in Chromium print. `grid.html:6-17` has no `<thead>`, so 13 tabs lose their header on page 2+. |
| **Page numbering** | **None.** No `display_header_footer` / `header_template` / `footer_template` anywhere. |
| Bookmarks / outline / page labels | None. |

### PDF assembly (`assembly.py`)

`render_report_pdf` (`assembly.py:122-136`): `build_pages` → for each page,
`R.render_pdf(p["full"], landscape=p["landscape"])` → `pypdf.PdfReader` → append every
page to a single `PdfWriter` → return bytes. That is the whole merge — no outline, no
page labels, no re-numbering after merge.

`render_report_html` (`assembly.py:109-119`) builds the Flask preview by string-splitting
each rendered page on `<body>` / `</body>` (`assembly.py:45-50`) and concatenating the
fragments under one inlined stylesheet.

**Performance flag.** `render_report_pdf` calls `render.render_pdf` once per tab, and each
call does its own `sync_playwright()` + `chromium.launch()` + `browser.close()`
(`render.py:115-144`). For the 24-tab Mountain report that is **24 cold Chromium
launches**. Add the workbook loads — `load_risk_change` (`:301-302`), `load_acl_env`
(`:365-366`) and `load_grid` (`:481-482`) each open the file **twice**, and
`read_chart_specs` (`charts.py:139`) once more — and a single report performs roughly
**50 full `load_workbook` calls**. Both are trivially fixable (one browser + one context
per report), and the second disappears entirely in the data-driven design.

### Wiring

- Flask blueprint `report_web_bp` registered at `/report-web`
  (`cecl_ui/app.py:45, 91`); routes `GET /<short>/<snap>` (HTML preview) and `.../pdf`
  (`cecl_ui/routes/report_web.py:35-60`).
- Dashboard entry point survives on HEAD:
  `cecl_ui/templates/run/client_dashboard.html:201-226` ("Preview report in browser (beta)").
- `assembly.find_report` (`assembly.py:30-42`) locates the generated workbook by filename
  convention — a Phase-0 artefact that **vanishes** once the renderer takes a model.

---

## 5. `regression.py` — what it compares, and re-baselining

### What it actually does

- `render_pages_png` (`regression.py:43-69`): calls `assembly.build_pages`, then for each
  page opens a Chromium page at a **fixed viewport width — 1100 px portrait / 1500 px
  landscape** (`:27-29, 55-57`), `device_scale_factor=1`, waits `document.fonts.ready`,
  `emulate_media("print")`, and takes a **full-page screenshot**.
- **It screenshots the HTML, not the PDF.** Orientation only selects a viewport width;
  pagination, page breaks and the merged-PDF output are **never exercised**.
- `_diff` (`regression.py:72-94`): PIL → numpy RGB arrays; per-pixel max-channel absolute
  delta; a pixel counts as changed above `_PIXEL_TOL = 20` (`:31-32`); a page fails above
  `_FAIL_FRACTION = 0.00005` (`:33-36`) — roughly **5 changed pixels in a 1100×1200
  image**. A dimension mismatch short-circuits to `1.0` (full change) with no diff image
  (`:82-83`).
- Failures write `_current_<sheet>.png` and a red-highlighted `_diff_<sheet>.png`
  (`:136-143`); the runner exits non-zero (`scripts/report_visual_regression.py:88`).
- Fixtures (`scripts/report_visual_regression.py:34-56`) are **3 generated workbooks on
  the `Z:` share**, chosen for shape: WNC (single-line pools), Jackson River (grade-rated
  with real migration), Tongass (many pools). References live in `Reports/_web_refs` and
  are **not committed** (`:5-6`). Commit `9dc273e` records "49/49 pages pass".

### What re-baselining against a data-driven renderer takes

1. **Fixtures change identity.** Today a fixture is `{label, cu, snap, report_path}`
   pointing at an `.xlsx` (`report_visual_regression.py:38-55`). It becomes a *frozen
   `ReportModel`* — ideally serialized JSON checked into the repo — so the harness has
   **no workbook dependency and no `Z:`-drive dependency**. That is a strict improvement:
   the current references are undiffable binaries derived from binaries on a network share.
2. **`build_pages` must be re-pointed.** `regression.py:48` is the *only* coupling to
   `assembly`; swapping it for a `build_pages_from_model(model)` is a one-line change.
   The diff engine (`:72-94`) and the runner (`:97-147`) need **no edits at all**.
3. **Keying changes.** Results are keyed by *sheet name* (`:65, 123`) and files named via
   `_safe(sheet)` (`:39-40, 117`). Keys become page ids from the model.
4. **A full re-baseline is unavoidable and every page will differ.** Pixel-exactness
   cannot survive a renderer whose input changed — column widths, header text and any
   number-format inference will shift. Plan for `--update` on day one, then hold the line.
5. **Recommended additions during the port:**
   (a) a **numeric** regression that diffs `model → JSON` against a baseline model — far
   more meaningful than pixels for a financial report, and it catches problems the pixel
   diff catches late or not at all;
   (b) rasterise the **merged PDF** (`scripts/supp_web_render.py` already demonstrates the
   pattern with PyMuPDF/`fitz`) so pagination and page breaks are actually covered;
   (c) an explicit page-count assertion per fixture.

---

## 6. Concrete blockers and risks

### Blockers — must be solved before the renderer can take a model

1. **Improved/deteriorated state has no data source.** It is decoded from Excel fill hex
   (`from_workbook.py:222-239`). `_sheet_risk_change` (`report_vizo.py:1596-2222`) *knows*
   this — it paints the cells — but never exports the knowledge. The grade-ordering rule
   must be extracted from that ~600-line function into a computed model field.
2. **Impr Deter renders as charts only, and the charts come from the workbook.** The
   template shows the CECL box + `charts` (`impr_deter.html:13-29`), and `charts` is
   `read_chart_specs(report_path, sheet)` (`assembly.py:59-62`). Until a `ChartSpec` node
   exists (§2), **the flagship executive page cannot be rendered from data at all.** The
   already-loaded `by_pool` / `by_grade` tables are the fallback, and their template code
   is recoverable from commit `698f1c8`.
3. **13 of 24 main tabs have no model.** Each needs a page archetype designed from the
   corresponding builder in `report_vizo.py`: `_sheet_report_index:816`,
   `_sheet_summary_variance:6177`, `_sheet_impaired_loans:6097`, `_sheet_acl_summary:5893`,
   `_sheet_mgmt_adj_summary:5998`, `_sheet_env_factor:2777`, `_sheet_loss_factor:2954`,
   `_sheet_co_recov_dq:3313`, `_sheet_env_ranges:3619`, plus the supplemental set
   (`_sheet_hist_trends_bal:3864`, `_sheet_detail_hist_bal:4074`,
   `_sheet_detail_co_hist:4382`, `_sheet_bal_adjust:4601`). This is the bulk of the work
   and it is **not** what the prototype's "covers every tab" status suggests.
4. **Two tabs are not generated code at all.** `Introduction-Vizo` and
   `Executive Summary-Vizo` are *cell-copied from an external template workbook*
   (`report_vizo.py:6636-6690`; `Vizo Narrative Tabs - Template.xlsx` on the `Z:` share).
   The narrative text has no representation in the data layer — a decision is needed:
   port the verbiage into the repo/DB, or keep reading that template.
5. **Formula-driven tabs.** `ACL Summary`, `Mgmt Adj Summary`, `Impaired Loans` and
   `Summary Variance` are *formula views over `ACL Env by Pool Mgmt Adj`*, not
   recalculations (`report_vizo.py:6710-6714` comment; see `_parse_acl_layout:5791`,
   `_summary_title:5864`, `_summary_header:5873`). The prototype only ever sees their
   **cached values** via `data_only=True`. A data-driven renderer must **compute** these
   numbers in Python — logic that exists nowhere today.
6. **`Change Analysis` lives in a separate module** (`change_analysis.append_change_analysis`,
   called at `report_vizo.py:6704`) and writes directly into the workbook. It needs its
   own model node and a value-returning API.

### Risks

7. **7 weeks dormant, single author.** Last prototype commit `8426726` (2026-07-10); HEAD
   is 2026-08-31. Nothing has kept it working against `report_vizo.py` changes since — and
   `report_vizo.py` was modified as recently as 2026-08-31 (`Change Analysis` at `:6704`).
   The audit found no evidence the prototype has ever been run against a workbook
   containing that tab.
8. **Missing dependencies.** `requirements.txt` lists `Jinja2`, `numpy`, `pillow`,
   `openpyxl` — but **not `playwright` and not `pypdf`**. The prototype cannot be
   installed from the manifest today.
9. **Calibri licensing.** `render.py:22-25` is explicit: the bundled TTFs are fine on a
   Windows workstation, not on a redistributable multi-user server. Switching to Carlito
   is claimed to keep pagination identical — **that claim is untested and will invalidate
   every pixel reference.** Do the font swap and the re-baseline in the same change.
10. **Number-format fidelity.** `format.excel_format` (`format.py:95-136`) is explicitly
    *"not a full Excel format engine"* (`:101`). In the target architecture that is
    actually a benefit — formats become explicit per-field template filters — but the port
    must audit every currently-inferred format, because the inference silently absorbs
    cases nobody has enumerated.
11. **Theme/indexed colors are silently dropped.** `_hex6` (`from_workbook.py:444-457`)
    returns `None` for colors it cannot resolve to RGB, so theme-colored cells render
    unstyled today. Side-by-side comparisons against the current PDF will surface these;
    they are prototype bugs, not target-architecture requirements.
12. **Two page-fidelity behaviours have no HTML equivalent yet:** page numbering running
    continuously across the whole report (`report_vizo.py:6407` — Excel treats a whole-
    workbook PDF export as one print job) and pool-block pagination
    (`report_vizo.py:6364-6381`). Both are solvable (Chromium footer templates;
    `break-inside: avoid`) but neither is prototyped.
13. **Silent per-tab failure** (`assembly.py:84-87`) is acceptable for a preview and
    unacceptable for a client deliverable. A strict mode is required before go-live.
14. **Chromium inside the Flask request.** The preview is wired into the app
    (`cecl_ui/app.py:91`); `render_report_pdf` forks 24 Chromium processes inside a single
    request. Needs a worker/queue, or at minimum single-browser reuse, before it is more
    than a beta button.
15. **Tab ordering authority moves.** Order currently comes from `wb.sheetnames`
    (`assembly.py:96`), which `report_vizo._reorder_vizo_main:6539-6557` establishes from
    `_VIZO_MAIN_ORDER:6518-6536` — including the `"Risk Chg *"` wildcard slot for per-pool
    tabs. That ordering logic must be ported into the model, not re-derived.

---

## 7. Git history — intended phasing, and what was abandoned

18 commits, all by `bevans@tctrisk.com`, across **two days**: 2026-07-09 and 2026-07-10.

| Phase | Commit | Delivered |
|---|---|---|
| 1 (POC) | `c1aad07` | Package skeleton, `ReportModel`, Jinja2 + print CSS, Cover → PDF. Solved: bundled fonts over `file://`; Jinja autoescape mangling inline CSS. Explicitly scoped out: *"Impaired + Mgmt Adj stay on the Excel path."* |
| 2 | `75f3604` | Impr Deter (CECL box + by-pool/by-grade tables) |
| 3 | `36df612` | Risk Change migration matrices |
| 3b | `53e71b3` | ACL Env by Pool (allowance calculation) |
| 3c | `7d91bda` | **Generic faithful-grid renderer** ("3 more tabs") |
| 3d | `2a2bb11` | Narrative/index tabs + wrap-text |
| 4 | `99e0407` | Embedded charts → inline SVG |
| **5a** | `1834b50` | Flask preview + merged-PDF assembly; fixed the `prefer_css_page_size` / landscape bug |
| — | `6c71a90` | Dashboard Preview/PDF buttons |
| — | `9dc273e` | Visual-regression harness (49/49 pages pass) |
| polish (07-10) | `6cbcbfc`, `698f1c8`, `242a4c6`, `9d34eb9`, `7fd6f7f`, `8c5b33c`, `8426726` | Cover fidelity, chart-centric Impr Deter, chart colors, matrix robustness, ACL borderless, per-pool Risk Chg routing, Supplemental end-to-end |

**Phasing read.** Phases 1→4 were a *fidelity* ladder (make each archetype look right);
5a was *assembly*. The **"5a" label implies a 5b that never landed** — on the evidence of
what is missing, 5b was production hardening (page numbers, hidden-sheet filter, strict
errors, performance) and/or the switch to a compute-time populator that
`from_workbook.py:3-7` explicitly anticipates:

> *"Phase-0 data source. Populating the model from the already-generated xlsx guarantees
> the PDF numbers match the Excel report exactly… Later phases can add a compute-time
> populator with the same model as output — templates don't change."*

**The target architecture in this migration is the phase the prototype was designed for
and never reached.** The workbook adapter was always labelled Phase 0, not the design.

### Abandoned / reversed work — three items

1. **Impr Deter migration tables were deleted** in `698f1c8` (48 lines removed from
   `impr_deter.html`) when the page went chart-centric. The model fields
   (`ImprDeterPage.by_pool`, `.by_grade`) and the CSS (`report.css:109-121`) survive as
   dead code. **Recoverable from that commit** — and likely *needed*, since the charts
   that replaced them are workbook-sourced.
2. **The Risk Change side panel (Deteriorated/Improved/Unchanged) was dropped** in
   `9d34eb9` — "drop side panel (robust to any grade count)". Justified as not part of the
   printed report (`from_workbook.py:245-247`); worth confirming with the report owner
   rather than silently inheriting the assumption.
3. **No xlsx writer was ever written**, despite being promised in `__init__.py:15-18` and
   `model.py:5-6` as the thing that makes "PDF-only, but leave the Excel door open" cheap.

**Not abandoned:** the Flask blueprint (`cecl_ui/app.py:45, 91`) and the dashboard buttons
(`client_dashboard.html:201-226`) are both still wired in on HEAD.

---

## Appendix — the Excel path being replaced

- `report_vizo.py`: **6,743 lines**. 24 `_sheet_*` builders plus 2 `compose_*` entry
  points (`compose_vizo_main:6560`, `compose_vizo_supp:6720`).
- **Raw ElementTree XML surgery: `report_vizo.py:4825-5790` ≈ 966 lines** — 4 `patch_*`
  entry points (`patch_dq_pie_zero_labels:5200`, `patch_impdet_charts:5339`,
  `patch_drawing_onecell_to_twocell:5577`,
  `patch_remove_chart_borders_and_axis_lines:5657`) over ~20 `_fix_*` / `_ensure_*`
  helpers, all unzipping the `.xlsx` and rewriting `xl/charts/*.xml` and
  `xl/drawings/*.xml` (`:5233, :5318, :5355, :5551, :5603`). Called from
  `generate_report.py:13042-13045`.

  *Note on the count:* the working brief cites ~1,891 lines for this region; the measured
  span at current HEAD is ~966. All 97 `ET.` call sites in the repo are inside it — there
  is **no** XML surgery in `generate_report.py` or `report_tct.py`. The larger figure most
  likely includes the openpyxl chart-construction code inside `_sheet_impdet:1035-1595`,
  which is what the patchers exist to fix up.

- **All of the above is bypassed entirely by the browser renderer** — the single biggest
  structural win of the migration, and the reason the chart work in `charts.py:199-455`
  matters: it replaces both the openpyxl chart builders *and* the XML patchers that clean
  up after them.
