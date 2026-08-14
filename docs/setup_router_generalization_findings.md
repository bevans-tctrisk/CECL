# Setup Router — Generalization Findings (2026-07-22)

**Test:** router run in re-onboard mode (base_cfg = the CU's real config = ground truth) against 4 CUs chosen for maximum diversity vs Cottonwood (the only CU it was built/tuned on).

## ⚠ Correction (2026-07-23) — the WARM/non-WARM labeling was WRONG
The original "WARM / non-WARM" split came from a naive filename scan that flagged the tool's **own generated output** reports (`..._CECL_Migration_..._Vizo_Model.xlsx`) as "WARM." Per the analyst: **Credit Union of Richmond was built from scratch — it has no WARM.** richmond, WNC, and McDowell were built **from scratch** (Vizo AIRES / raw extracts).

**Bridgeton, however, IS a WARM‑built CU** (analyst confirmed). Its WARM migration models — a full quarterly history **through 2026‑06** — live in a **separate `...\Credit Migration\` folder**, NOT the `...\Portfolio Management\` folder that its config's `loan_source_folder` points at and that this test scanned. So the test looked in the wrong directory and never saw the WARM. Cottonwood is the other confirmed WARM‑built CU.

**New finding — multi‑folder delivery:** WARM CUs commonly split delivery across folders (WARM model in `Credit Migration\`, loan/monthly/impaired files elsewhere). Both the router and this test assumed a single delivery folder — that assumption is wrong for WARM CUs and must be handled.

Key lesson: **a WARM‑named file in a folder does NOT mean the config was built from a WARM** (richmond has one analyst file but was built from scratch), and **absence of a WARM in the scanned folder does NOT mean there is none** (bridgeton's WARM was one folder over). Build method + WARM location are the analyst's process knowledge, not reliably auto‑detectable.

**Impact:** this run effectively tested **three from‑scratch/AIRES CUs plus one WARM CU pointed at the wrong folder** — not the intended 2 WARM + 2 non‑WARM. The technical gaps below are all real, but the **WARM side is still effectively uncovered** (Cottonwood aside). A proper WARM re‑test = bridgeton with the router pointed at BOTH its `Credit Migration\` (WARM) and delivery folders.

## Results
| CU | build method | member_account (prop / truth) | pool_code | overall |
|---|---|---|---|---|
| credit_union_of_richmond | from scratch (AIRES) | fixed_suffix / ['delimiter', 'fixed_suffix'] OK | None MISS/RED | red |
| bridgeton | **WARM** (test scanned wrong folder; config incomplete) | None / [] n/a | None MISS/RED | red |
| wnc_community_cu | from scratch (AIRES) | fixed_suffix / ['delimiter'] MISS | None MISS/RED | red |
| mcdowell_cornerstone_cu | from scratch (AIRES) | fixed_suffix / ['fixed_suffix'] OK | None MISS/RED | red |

## Key findings

1. **member_account: delimiter is undetectable (CONFIRMED GAP).** WNC uses `delimiter` member_account; the router proposed `fixed_suffix` because the profiler only detects a *separate suffix column* (=> split) vs none (=> top-support fixed_suffix). It has no signal to tell `delimiter` from `fixed_suffix`. Needs a profiler signal: sample the member column for a consistent delimiter char (e.g. 'L','-').
2. **pool_code path assumes the profiled loan file's columns match base_cfg (CONFIRMED GAP).** Richmond & McDowell (AIRES) crashed on a missing balance column — the profiler picked an AIRES loan file and read it with its OWN detected header row, which disagrees with the extract's real header_row/columns. AIRES puts position numbers on row 1 and headers on row 2, and ships v1/v2 variants. Router must use the matched extract's header_row + verify the mapped columns exist (now degrades gracefully instead of crashing).
3. **pool_code detector needs pool names to align (CONFIRMED GAP).** Even with the universal `load_monthly_balances` fix (now loads single/per_month/per_year), WNC's detector returned needs_input — the monthly pool names / mapped pools didn't align enough to tie. Vizo monthly 'by type' rows map to pools via a balance_title_map the router doesn't consult yet.
4. **Sparse configs (bridgeton) have no ground truth AND the greenfield bootstrap needs descriptors.** Bridgeton's config is nearly empty; AIRES/Full-Loan-File extracts have no PURPOSE/COLLATERAL description columns, so the description-based bootstrap has no signal. Greenfield on these formats needs a different signal (code reference sheet, or tie-guided assignment).

## Verdict
The router generalizes on the pieces that are format-independent (member_account for split/fixed cases, file archetype classification) but is still **coupled to the Cottonwood delivery shape** (rich descriptor extract + manual monthly). It is NOT yet ready for an analyst-facing new-CU run on Vizo/AIRES clients.

## Recommended fixes before a live new-CU test (priority order)
1. Profiler: detect `delimiter` member_account by sampling the member column for a repeated non-numeric separator; add `delimiter` to the registry lookup dimension.
2. Router: in re-onboard mode use the matched extract's `header_row` and validate mapped columns exist; improve AIRES header detection (row of real names, skip position-number rows); handle v1/v2 dupes.
3. pool_code: consult `balance_title_map` / pool aliasing so Vizo 'by type' monthly rows align to pools; then the detector can tie.
4. Greenfield: add a code-reference/description-sheet ingest and/or tie-guided assignment for extracts without descriptor columns.

## Per-CU detail

### credit_union_of_richmond (WARM)
- Vizo AIRES loan file; delimiter top-level / fixed_suffix extract; 4 extracts incl participation
- TRUTH: member_account=['delimiter', 'fixed_suffix'], pool_code=['Loan Purpose Code', 'Loan Type Code']
- FILES: loan_extract=Loan AIRES v2  03-2026.xlsx, report_other=2026-03 Improved Deteriorated Loans - Credit_Union_of_Richmond.xlsx, unknown=2026-03 Management Adjustment Worksheet- Credit Union of Richmond.xlsx
  - [green] member_account -> {'mode': 'fixed_suffix'}  (src=registry) — loan_suffix column absent -> fixed_suffix (registry support 22 CUs)
  - [red] pool_code -> None  (src=profile) — balance column 'Account Balance Total Amount' not found in profiled loan file (header/format mismatch); cols: ['Record Code', 'Account Number', 'Addre

### bridgeton (WARM)
- Non-Vizo "Full Loan File"; config nearly EMPTY (no member_account, 0 extracts) = near-greenfield
- TRUTH: member_account=[], pool_code=[3]
- FILES: report_other=2026-06 Improved-Deteriorated Loans Report.xlsx, unknown=Chg Off Q2 2026 - 2.xlsx

### wnc_community_cu (non-WARM)
- Vizo AIRES; DELIMITER member_account; 1 extract
- TRUTH: member_account=['delimiter'], pool_code=['LOAN TYPE']
- FILES: impaired=20260630 CECL-Credit Migration Impaired Loans.xlsx, loan_extract=20260630 Aries Loans.xlsx, report_other=2026 05 May Chargeoffs and Recoveries.xlsx, unknown=2026-06 Management Adjustment Worksheet- WNC Community CU.xlsx
  - [green] member_account -> {'mode': 'fixed_suffix'}  (src=registry) — loan_suffix column absent -> fixed_suffix (registry support 22 CUs)
  - [red] pool_code -> None  (src=needs_input) — no pool_map yet - needs code->pool bootstrap (use code descriptions). Known code-column names: Loan Type Code, LOAN TYPE, Loan Type, LOAN_TYPE, Acct T

### mcdowell_cornerstone_cu (non-WARM)
- Vizo AIRES; fixed_suffix; 3 extracts incl participation
- TRUTH: member_account=['fixed_suffix'], pool_code=['Loan Type Code']
- FILES: loan_extract=June 2026 Aires Loans.xlsx, report_other=2026-06 Improved Deteriorated Loans - McDowell_Cornerstone_CU.xlsx, unknown=2026-06 Management Adjustment Worksheet- McDowell Cornerstone CU.xlsx
  - [green] member_account -> {'mode': 'fixed_suffix'}  (src=registry) — loan_suffix column absent -> fixed_suffix (registry support 22 CUs)
  - [red] pool_code -> None  (src=profile) — balance column 'Principal Owned at Period End' not found in profiled loan file (header/format mismatch); cols: ['Record Code', 'Account', 'Member Name
