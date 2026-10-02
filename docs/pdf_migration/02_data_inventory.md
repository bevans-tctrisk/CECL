# 02 — Data Inventory for the PDF Migration

Scope: what data is actually available at report-build time, per tab, for
`report_vizo.compose_vizo_main(client, snap, df, config, grades, hist=None)`
(`report_vizo.py:6560`).

Citations use `report_vizo.py` = **rv**, `generate_report.py` = **gr**,
`change_analysis.py` = **ca**.

## Bottom line

**13 of the 17 visible tabs are already pure functions of `(df, config, grades,
hist)`** and could be rendered to PDF from data today. The blockers are a small,
well-defined set:

1. **Four tabs are formula views over a fifth tab.** `ACL Summary`,
   `Mgmt Adj Summary`, `Impaired Loans` and `Summary Variance` contain almost no
   data of their own — they are `='ACL Env by Pool Mgmt Adj'!K42`-style
   cross-sheet references, with row numbers recovered by `_parse_acl_layout()`
   (rv:5791) *re-reading the worksheet that was just built*. Two of them also
   read cell **values** back out and branch on them. See §4.
2. **`Change Analysis` and `Summary Variance` each parse the prior quarter's
   saved `.xlsx` off disk** (ca:280, rv:6212) with `data_only=True` — i.e. they
   depend on Excel having cached formula results in that file. There is no
   serialized store of prior-period results anywhere.
3. **`Introduction-Vizo` and `Executive Summary-Vizo` are not built at all** —
   they are cell-by-cell copies out of `Sample Reports/Vizo Narrative Tabs -
   Template.xlsx`, which is **gitignored and absent from the clone**
   (rv:6621-6685, `.gitignore:45`).
4. **`Impr Deter` is coupled to `ACL Env` through a side-effect stash** on
   `hist['impaired']['_computed_pooled_total_allow']` (rv:2680), which is why
   `compose_vizo_main` builds the two tabs out of display order and then repairs
   the order with `wb.move_sheet` (rv:6570-6581).
5. **Charts are `Reference(ws, ...)` cell ranges**, and four `patch_*` functions
   re-open the **saved file** to rewrite chart XML (gr:13043-13046). None of
   that survives a data-driven renderer; charts must be re-authored — which also
   makes ~700 lines of XML-patching helpers (rv:4834-5790) and
   `report_integrity.check_and_report` obsolete.

Everything else — `Vizo Cover`, `Report Index`, `Impr Deter`, `Risk Change
Total`, per-pool `Risk Chg *`, `ACL Env by Pool Mgmt Adj`, `Env Factor by Pool`,
`>Envir Fact Ranges`, `Display HIst Bal`, `Display CO-Recov-DQ` — is computed
inline in Python from the four inputs plus two PNG logos and `admin_defaults.yaml`.

---

## 1. The four inputs

### 1.1 `df` — pandas DataFrame of loans

Provenance: `df = load_loans(cu, snapshot_date, config)` (gr:12323) →
`df = calculate_cecl(df, grades, no_score, brr_rules=..., ...)` (gr:12364) →
passed to `compose_vizo_main_new` at gr:13016.

`load_loans` (gr:222-233) is a raw `SELECT * FROM monthly_loan_data WHERE
credit_union=:c AND snapshot_date=:s`, so the authoritative schema is the
writer, `import_data.py:1650-1698` (+ `to_sql` at `import_data.py:1782`). There
is no `CREATE TABLE`; the schema was created implicitly by the first `to_sql`,
plus one `ALTER TABLE ... ADD COLUMN IF NOT EXISTS business_risk_rating TEXT`
(`import_data.py:1726-1729`). `_apply_excluded_pools` (gr:236-259) then drops
rows whose `loan_pool` is in `config['excluded_pools'] ∪ {'Exclude','Ignore'}`.

Full schema (DB columns 1-8, derived 9-13):

| # | Column | dtype | Source | Description |
|---|---|---|---|---|
| 1 | `credit_union` | str | `import_data.py:1651` | CU display name; SQL filter key. Not read in rv. |
| 2 | `snapshot_date` | date | `import_data.py:1652` | Report period (month-end); SQL filter key. Not read in rv. |
| 3 | `member_number` | str | `import_data.py:1653` | Member + loan-suffix account key. Not read in rv. |
| 4 | `current_balance` | float | `import_data.py:1654` | Outstanding principal. **The money column.** |
| 5 | `current_fico_score` | int | `import_data.py:1657` | Current score; feeds `current_grade`. Not read directly in rv. |
| 6 | `original_fico_score` | int | `import_data.py:1658` | Origination/prior score; feeds `original_grade`. Not read directly in rv. |
| 7 | `loan_pool` | str | `import_data.py:1659` | Canonical pool name (post `pool_map`). **The universal slicing key.** |
| 8 | `business_risk_rating` | str/None | `import_data.py:1686`, backfilled gr:231 | Analyst rating for commercial pools; consumed by `calculate_cecl`. Not read directly in rv. |
| 9 | `current_grade` | str | `cecl_engine.py:202-204`, BRR override `:224` | Grade band label as of snapshot. |
| 10 | `original_grade` | str | `cecl_engine.py:205-207`, `:246` | Grade at origination — the migration "from" axis. |
| 11 | `migration_status` | str | `cecl_engine.py:252-256` | Improved / Deteriorated / Unchanged. **Not read in rv** — rv re-derives it from the matrix. |
| 12 | `reserve_rate` | float | `cecl_engine.py:257-263` | Per-loan base reserve rate. Not read in rv. |
| 13 | `expected_loss_amount` | float | `cecl_engine.py:265` | `balance × rate`. Not read in rv. |

**Only three columns are read directly in `report_vizo.py`**, plus one more via
`risk_change_matrix`:

| Column | Used for | Citations |
|---|---|---|
| `loan_pool` | pool enumeration and the `pdf = df[df['loan_pool'] == pool]` slice behind every per-pool tab | rv:224, 619, 1147, 2363, 2393, 2891, 3118, 6597, 6609 |
| `current_balance` | every dollar total not sourced from WARM | rv:446, 633, 1617, 2394, 2498, 2671, 3119, 3169 |
| `current_grade` | per-grade balance breakout; the risk-rated test `sub['current_grade'].nunique() > 1` | rv:633, 2497, 3168, 4148, 6600 |
| `original_grade` | migration-matrix columns — read inside `risk_change_matrix` | `cecl_engine.py:298`, called rv:451, 1177, 1611, 1615 |

There is no `df.groupby`, `df.columns`, `iterrows` or `.loc[]` on `df` anywhere
in rv (the single `.loc[]` at rv:558 is on the migration matrix). Note that
`_ordered_pools` (rv:210-254) deliberately **unions** `df`'s pools with pools
appearing only in `hist` — so `df` is *not* the authoritative pool list.

`risk_change_matrix(df, grades, no_score_label, labels=None)`
(`cecl_engine.py:288-303`) is the only non-trivial `df` transform: a square
DataFrame indexed by grade label, values = summed `current_balance`.

### 1.2 `config` — the client YAML

`client_configs/` is gitignored (`.gitignore:47`) and lives under
`$CECL_WORKSPACE_ROOT` (gr:43-45), so no real client YAML is in the clone. The
most complete in-repo generator is `create_sample_cu.py:248-313`; the newer
wizard schema is normalized by `cecl_ui/services/config_service.py:138-190`.

R = affects **rendering** (text/labels/thresholds shown); C = affects computation.

| Key (path) | Default | Controls | R/C | Citations |
|---|---|---|---|---|
| `credit_union` | required | CU name on every sheet header; output filename | R | rv:6561, 6722 |
| `no_score_label` | `'Not Reported'` | label of the no-FICO bucket; also a grade key in mgmt-adj math | R+C | rv:449, 607, 1047, 1598, 2232, 2787, 2957 |
| `top_grades_double_drop` | `3` | WARM rule: top N original grades need a 2+ grade drop to count Deteriorated | C | rv:452, 1178, 1711 |
| `business_risk_ratings[].label` | `[]` | ordered BRR band labels replacing FICO grades on flagged pools | R | rv:270-273 |
| `pools[].name` | — | key for the two per-pool maps | C | rv:294, 352-355 |
| `pools[].brr` | `False` | renders that pool with BRR labels | R+C | rv:293-295 |
| `pools[].use_default_mgmt_adj` | `False` | apply `admin_defaults.yaml → default_mgmt_adj` when base rate is 0 | C | rv:349-356, used 375-377, 393-396 |
| `mgmt_adj_by_pool` | `{}` | `{pool: rate}` manual overlay; highest-precedence mgmt adj | C | rv:2233, used 372, 389 |
| `not_risk_rated` | `[]` | pools rendered as one total line, no grade breakout | R+C | rv:2378 |
| `other_allowance_considerations[]` | `[]` | `{title, balance, percentage, amount}` rows added to Total Allowance Needed | R+C | rv:401-426, used 650-653 |
| `economic_data.{unemployment_rate, population, bankruptcies, foreclosures}` | 0/1 | Economic Stress Index inputs; **overridden by `hist['impaired']['economic_data']`** | C | rv:468-474, 2780-2784 |
| `economic_data._sources` | — | "Data Sources:" footnote block on Env Factor by Pool | R | rv:2926-2943 |
| `acl_balance` | `0` | fallback ACL balance; drives the CECL Adjustment | C | rv:655, 2728 |
| `warm_months` | `{}` (per-pool 36) | `{pool: months}` WARM lookback; lower precedence than `hist['impaired']['acl_months']` | R+C | rv:3150, 3217, 3470, 3512, 3560 |
| `pool_map` | `{}` | read only for its *values*, as the pool-list fallback when WARM gives no `pool_order` | C | rv:4404 |
| `report_dir` / `output_dir` | `None` | test-only override for locating the prior-quarter workbook | C | rv:6206 |
| `warm_allowance_pools` | — | consumed downstream in Change Analysis | C | ca:399, via rv:6706 |

Two config-adjacent inputs are **not** in `config`:
`admin_defaults.yaml → default_mgmt_adj` (currently `0.0011`), read directly by
`_load_admin_default_mgmt_adj()` (rv:334-346); and the `CECL_WORKSPACE_ROOT`
env var, which resolves `_WORKSPACE_BASE` (rv:59).

### 1.3 `grades` — list of grade-band dicts

Passed separately (YAML `credit_grades`), not read off `config`. Each item is
`{'label', 'min_score', 'max_score', 'reserve_rate'}` (rv:452, 637, 639).
`_all_grades()` (rv:256) appends `no_score_label`; `_is_hidden()` (rv:304) drops
any label starting `Hide` (`HIDDEN_GRADES = ['Hide-F','Hide-G','Hide-H','Hide-I']`,
rv:123); `_grade_ranges()` (rv:538-552) turns min/max into the "Score Range"
display string ("700+", "550 or less", "600-649").

### 1.4 `hist` — the historical / WARM dict

Constructed once in `load_historical_data` (gr:4389-4399, def gr:4114) then
mutated throughout `main()` (gr:12369-12995).

**Top level:**

| Key | Shape / contents | Built | Read in rv |
|---|---|---|---|
| `years` | `sorted list[int]`, capped at snapshot year | gr:4392, 4409-4412, trimmed 12424 | rv:482, 2264, 2971, 3321 |
| `chargeoffs` | `{year: {pool: $}}` annual charge-offs | gr:4390, overlaid 6011, 12002, 12596 | rv:479, 2261, 2968, 3318 |
| `recoveries` | `{year: {pool: $}}` annual recoveries | gr:4391, 6022, 12007, 12599 | rv:480, 2262, 2969, 3319 |
| `avg_balances` | `{year: {pool: avg balance}}` — loss-rate denominator | gr:4396, 6081 | rv:481, 2263, 2970 |
| `dq_pct` | `{year: {pool: pct 0-1}}` delinquency rate | gr:4398, 6033 | rv:502, 3320, 3571 |
| `co_monthly` | `{(year, month): {pool: $}}` monthly charge-offs | gr:4393, 6044, 12602 | rv:3356, 3460, 3550 |
| `rc_monthly` | `{(year, month): {pool: $}}` monthly recoveries | gr:4394, 6060, 12606 | rv:3358, 3503, 3552 |
| `impaired` | the big nested sub-dict — see below | gr:12438 and many | rv:201, 219, 504, 574, 1144, 1620, 2247, 2782, … |
| `monthly_balances` | `DataFrame(pool, date, balance)` | gr:4395 | **never read in rv** (gr-only) |
| `alll_by_date` | `{Timestamp: ACL balance}` | gr:4399, 12965 | **never read in rv** — feeds `impaired['acl_balance']` at gr:12991-12993 |
| `delinquency` | `{quarter: {pool: dq balance}}` | gr:4397 | **dead key — never read anywhere** |

**`hist['impaired']`** (aliased `_imp` / `imp` in rv):

| Key | Shape / contents | Built | Read in rv |
|---|---|---|---|
| `pool_order` | `list[str]` canonical WARM pool order (NRR last) | gr:3006, 3487, 5131 | rv:220, 2362, 4103, 6586 |
| `risk_rated` | `{pool: bool}` — decides which pools get a `Risk Chg *` tab and whether NCC is forced to 0 | gr:3007, 3491, 5251, 12885 | rv:221, 617, 1144, 2361, 2888, 3109, 6587 |
| `hist_bal_data` | `{pool: {'dates': [datetime], 'grades': {grade: [float]}, 'total': [float]}}` | gr:3005, 3486, 5486; extended 12479 | rv:222, 2269, 2374, 2989, 4102 |
| `pool_bal_detail` | `{pool: {grade\|'Total': {loan_report_bal, bal_adj, balance_sheet_total[, specific_id]}}}` | gr:5398, 2745 | rv:223, 2358, 4636 |
| `acl_pools` | `{pool: {'grades': {grade: {balance, spec_id, calc_bal, base_rate, mgmt_adj, factor, allow_before}}, 'total': {… , env_factor, env_allow, total_allow}}}` | gr:5130 (built 5013-5080) | rv:2253 |
| `acl_impaired` | `{impairment category: allowance $}` | gr:5132, 5667, 12648 | rv:585, 2254, 2710, 2716 |
| `acl_summary` | `{pooled_balance, pooled_spec_id, pooled_allow_before, pooled_env_allow, pooled_total_allow, total_spec_allow, total_allow_needed, acl_balance, adjustment}` | gr:5133 (built 5086-5125); `acl_balance` overwritten 12991 | rv:584, 2255, 2671, 2685, 2728 |
| `acl_balance` | float, current ACL/ALLL balance | gr:5007, 12993 | rv:655 |
| `acl_months` | `{pool: int}` life-of-loan months (WARM Months) | gr:3489, 5276, 12856 | rv:510, 2265, 2975, 3327, 4104 |
| `spec_id_by_pool` | `{pool: {grade: balance removed}}` | gr:4927, 5726, 12649 | rv:2258 |
| `total_spec_id` | float, last-resort specific-ID total | gr:4896, 12650 | rv:591 |
| `pooled_total_allowance` | float, prior WARM-parsed pooled allowance (fallback) | gr:5001 | rv:603-604 |
| `exec_summary_3` | `{'improved': {grade: $}, 'deteriorated': {grade: $}}` | gr:4980 | rv:1166 (Impr Deter charts) |
| `dq_by_status` / `co_by_status` | `{loan_status: {'balance', 'pct'}}` grand-total blocks | gr:5181, 5195 | rv:1629, 1630 |
| `dq_by_pool` / `co_by_pool` | `{pool: {loan_status: {'balance','pct'}}}` | gr:5185, 5199 | rv:1623, 1624 |
| `economic_data` | `{state, county, unemployment_rate, foreclosures, bankruptcies, population}` | gr:5228-5235 | rv:609, 2248, 2783 |
| `env_ranges` | `{'ncc': [(lo,hi,score)], 'dq': [...], 'es': [...], 'ncc_labels': [...], 'dq_labels': [...], 'es_labels': [...]}` | gr:5319-5326 | rv:202 (used 611, 2249, 2786), rv:3685 |
| `balance_adjustments` | `{pool: adjustment $}` (loan file vs balance sheet) | gr:5397, 2701 | rv:1808 |
| `total_balance_adjustment` | float | gr:5399, 2702 | rv:1812 |
| `total_in_portfolio` | float (loan total + adjustments) | gr:5400, 2703 | rv:1823 |
| `warm_co` / `warm_rc` / `warm_net` | `{year: {pool: $}}` from the prior WARM/TCT workbook | gr:3035-3037, 4707-4709 | rv:3345, 3346, 3347 |
| `warm_co_totals` / `warm_rc_totals` | `{pool: $}` life-of-window totals | gr:3038-3039 | rv:3348, 3349 |
| `warm_net_co` | `{pool: net CO $}` — drives the life-of-loan loss rate | gr:3040, 3608, 4766 | rv:2268, 3016 |
| `warm_co_monthly` / `warm_rc_monthly` | `{(year, month): {pool: $}}` | gr:3060, 3063 | rv:3353, 3354 |
| `warm_dq_pct` | `{year: {pool: pct}}` DQ% from the WARM sheet | gr:3042, 12617, 12760 | rv:505, 3570 |
| `prior_mgmt_adj` | `{pool: {grade: mgmt adj}}` (non-zero only) | gr:3306, 3758 | rv:2256 |
| `prior_env_factor` | `{pool: env factor}` | gr:3307, 3759 | rv:2257 |
| `items` | `{impairment category: provision}` | gr:4896 | **dead — superseded by `acl_impaired`** |
| `warm_snapshot_balances` | `{pool: $}` at the month ≤ snapshot | gr:12477 | **never read in rv** (gr:2580 only) |

**Written, not read — computed at render time and stashed back into the input
dict** (rv:2678-2682):

| Key | Written | Read |
|---|---|---|
| `_computed_pooled_total_allow` | rv:2680 | rv:599-600 (`_compute_acl_totals`, highest precedence) → **Impr Deter** |
| `_computed_grand_allow_before` | rv:2681 | **nowhere — write-only** |
| `_computed_grand_env_allow` | rv:2682 | **nowhere — write-only** |

`report_acl_funding.py:63` passes the same `hist` straight into
`report_vizo._compute_acl_totals` (`report_acl_funding.py:72`), so any change to
the fallback chain affects that report too.

---

## 2. Tab → data dependency map

Sheet names are authoritative from the `create_sheet` calls and
`_VIZO_MAIN_ORDER` (rv:6518-6536), which is also the PDF page order.

| Tab | Builder (line) | `df` | `config` | `grades` | `hist` | Other |
|---|---|---|---|---|---|---|
| Vizo Cover | `_sheet_cover` (rv:685) | — | `credit_union` | — | — | 2 PNG logos; ~13 lines of literal disclaimer (rv:781-793) |
| Report Index | `_sheet_report_index` (rv:816) | — | `credit_union` | — | — | ~60 lines of literal prose incl. a **static tab index** (rv:891-898) |
| Summary Variance | `_sheet_summary_variance` (rv:6177) | — | `report_dir`/`output_dir` | — | — | **formula view over ACL Env + prior-quarter `.xlsx`** (§4) |
| Impr Deter | `_sheet_impdet` (rv:1035) | `loan_pool` | `no_score_label`, `top_grades_double_drop` | ✔ | `impaired` (via `_compute_acl_totals`, `_ordered_pools`, `exec_summary_3`) | 4 charts; **reads back its own cells** (rv:1213-1223, 1383, 1429, 1502-1503); depends on the ACL Env stash |
| Risk Change Total | `_sheet_risk_change(pool_name=None)` (rv:1596) | `current_balance`, `current_grade`, `original_grade` | `no_score_label`, `top_grades_double_drop`, `business_risk_ratings`, `pools` | ✔ | `dq_by_status`, `co_by_status`, `balance_adjustments`, `total_balance_adjustment`, `total_in_portfolio` | 4 charts; 2 info-icon PNGs (rv:2197-2213) |
| Risk Chg * (per pool) | same builder, `pool_name=<pool>` | filtered by `loan_pool` | same + BRR flags | ✔ | `dq_by_pool`, `co_by_pool`, `risk_rated`, `pool_order` | one tab per risk-rated pool (rv:6586-6608) |
| ACL Env by Pool Mgmt Adj | `_sheet_acl_reserve` (rv:2223) | `loan_pool`, `current_grade`, `current_balance` | `no_score_label`, `mgmt_adj_by_pool`, `not_risk_rated`, `pools`, `business_risk_ratings`, `economic_data`, `acl_balance`, `other_allowance_considerations` | ✔ | `years`, `avg_balances`, `chargeoffs`, `recoveries` + ~15 `impaired.*` keys | `admin_defaults.yaml`; **hub tab — five other tabs reference it** |
| Change Analysis | `ca.append_change_analysis` (ca:244) | — | `warm_allowance_pools` | — | — | **parses current ACL Env sheet + prior quarter's `.xlsx`** (§4) |
| Impaired Loans | `_sheet_impaired_loans` (rv:6097) | — | — | — | — | **formula view over ACL Env; also reads its cell values** (§4) |
| ACL Summary | `_sheet_acl_summary` (rv:5893) | — | — | — | — | **pure formula view over ACL Env** (§4) |
| Mgmt Adj Summary | `_sheet_mgmt_adj_summary` (rv:5998) | — | — | — | — | **formula view; reads cell values to filter rows** (§4) |
| Env Factor by Pool | `_sheet_env_factor` (rv:2777) | `loan_pool` | `no_score_label`, `economic_data`, `economic_data._sources` | ✔ | `impaired.economic_data`, `risk_rated`, `env_ranges` | receives the ranges block appended by `_merge_env_ranges_into_factor` (rv:6429) |
| >Envir Fact Ranges | `_sheet_env_ranges` (rv:3619) | — | — | — | `impaired.env_ranges` (else module defaults) | **the largest prose block shipped by Python** (rv:3752-3846); hidden after its content is copied into Env Factor by Pool |
| Display HIst Bal | `_sheet_loss_factor` (rv:2954) | `loan_pool`, `current_grade`, `current_balance` | `no_score_label`, `warm_months` | ✔ | `years`, `avg_balances`, `chargeoffs`, `recoveries`, `impaired.acl_months`, `warm_net_co` | — |
| Display CO-Recov-DQ | `_sheet_co_recov_dq` (rv:3313) | — | `warm_months` | — | `years`, `chargeoffs`, `recoveries`, `co_monthly`, `rc_monthly`, `dq_pct` + `impaired.warm_co*`/`warm_rc*`/`warm_net`/`warm_dq_pct`/`acl_months` | DQ pie patched post-save |
| Introduction-Vizo | *no builder* (rv:6621-6685) | — | — | — | — | **cell-copied from an external template workbook** |
| Executive Summary-Vizo | *no builder* (rv:6621-6685) | — | — | — | — | **cell-copied from an external template workbook** |

`_sheet_introduction` (rv:971) and `_sheet_exec_summary` (rv:1011) exist but are
**never called** — dead code superseded by the template copy. They do, however,
contain the Python-native version of that methodology text (rv:981-988,
996-1002, 1023-1028), which is useful raw material for step 4 of the migration.

---

## 3. Tabs NOT derivable from `(df, config, grades, hist)`

### 3.1 External workbook: the narrative tabs — the biggest content gap

`compose_vizo_main` (rv:6622-6626) opens

- `$CECL_WORKSPACE_ROOT/Sample Reports/Vizo Narrative Tabs - Template.xlsx`, or
- falling back to `.../YYYY-MM CECL-Migration-WARM - Template Credit Union with Vizo.xlsx`

and copies the `Introduction-Vizo` and `Executive Summary-Vizo` sheets
cell-by-cell — values, fonts, borders, fills, number formats, protection,
alignment, merges, column widths, row heights, page margins, orientation
(rv:6628-6685) — then forces each to one page, column A only (rv:6675-6685).
Loaded through `_load_workbook_resilient` (rv:67-98), which retries 3× on
`OSError` and then copies to a local temp file — i.e. it is expected to live on
a flaky SMB share.

`Sample Reports/` is listed in `.gitignore:45` and **does not exist in the
clone**. The approved appendix verbiage is not in version control at all.
Extracting it into a versioned source is a prerequisite, and needs access to the
analyst workspace.

### 3.2 External workbook: theme and icons

- `_apply_vizo_theme` (rv:100-116) unzips `xl/theme/theme1.xml` out of the same
  missing master template and assigns it to `wb.loaded_theme`. Every
  `Color(theme=N)` in the report resolves through it (e.g. rv:6229, 6231,
  3645-3647). A PDF renderer needs the Vizo palette as literal hex.
- `ICON_INFO_DARKRED` / `ICON_INFO_DARKGREEN` (rv:120-121) point into
  `Sample Reports/assets/` — also absent. Both insertion sites (rv:2197-2213)
  are `os.path.isfile`-guarded, so they silently vanish today.

### 3.3 Images that *are* in the repo

`LOGO_VIZO` and `LOGO_TCT` (rv:118-119) resolve to `logos/vizo_financial.png`
and `logos/tct_risk_solutions.png`, both present. `_sheet_cover` PIL-trims the
Vizo logo's drop shadow (2% left/top, 5% right, 8% bottom) into a `BytesIO` PNG
before insert (rv:709-722). Easy to port.

### 3.4 Hard-coded narrative prose

Not external, but not in the data either:

| Builder | Lines | Content |
|---|---|---|
| `_sheet_env_ranges` | rv:3752-3846 | **Largest Python-shipped prose block.** Env Factor overview, a GAAP quotation, the Q&E methodology paragraph (Comptroller's Handbook / regression-ANOVA), plus definitions of Net Credit Change, Delinquency, and Economic Stress Score with named sources (BLS, Realty Trac, federal court records). Row heights hand-tuned to the wrap. This block is then *copied into Env Factor by Pool*. |
| `_sheet_cover` | rv:781-793 | Vizo legal disclaimer (~9 wrapped lines, merged B21:E31); title (rv:751); `© {datetime.now().year} TCT Risk Solutions` (rv:794 — uses *today's* year, not the snap year) |
| `_sheet_report_index` | rv:858-912 (main), 930-965 (supp) | Report-overview paragraph and a **static tab index list** in one multi-line cell (rv:891-898) — not derived from `wb.sheetnames`, so it will drift from the real tab set |
| `_sheet_risk_change` | rv:2183-2192 | The two improved/deteriorated footnote definitions, paired with the info icons |
| `_compute_acl_totals` | rv:667-679 | Label strings returned alongside the numbers, incl. the conditional `'Adjustment (Underfunded)'` / `'(Overfunded)'` |
| `_sheet_summary_variance` | rv:6350-6353 | "No prior report available" note |
| `_sheet_mgmt_adj_summary` | rv:6069-6071 | "No management or environmental adjustments were applied this period." |
| `change_analysis` | ca:171-231 | **Generated prose** — `_explain_pool` / `_explain_impaired` compose English sentences by f-string. Portable as-is (they take dicts, not cells), but row heights are estimated by `14 * ceil(len(n)/92) + 4` (ca:407), a manual wrap calc a PDF renderer replaces |
| `_sheet_appendix_supp` | rv:4770-4806 | Two paragraphs plus three `">insert report details"` placeholders; the tab is built then hidden (rv:6729) because it is unfinished |

### 3.5 Substantial inline computation

These builders hold real model logic, not just layout:

| Function | Lines | Computes |
|---|---|---|
| `_compute_acl_totals` | rv:562-682 | The 4 CECL Adjustment figures, each with a 3-deep fallback chain; includes a **full inline recomputation of pooled allowance** (rv:606-648) when WARM data is absent |
| `_sheet_acl_reserve` life-loss block | rv:2288-2315 | Recomputes per-pool life loss *inline* rather than calling `_pool_life_loss`: annualises `hist_bal_data` into per-year per-grade averages, applies the pool's `acl_months` window, prefers `warm_net_co` matched case-insensitively |
| `_sheet_acl_reserve` allowance loop | rv:2423-2426, 2451-2517, 2548-2568, 2600-2620 | `env_factor = (ncc_score + dq_score + es_score)/100`; per-grade `base_rate = max(0, pool_ll × dist)`, `factor = base_rate + mgmt_adj`; pool and grand totals; a separate NRR branch |
| `_ncc` | rv:443-465 | Net credit change from the matrix, applying the top-N-double-drop WARM rule; returns `(imp_pct, det_pct, net_pct)` |
| `_eco_stress` | rv:468-474 | `unemployment×100 + bankruptcies/pop×100 + foreclosures/pop×100` |
| `_pool_life_loss` | rv:477-493 | Per-year `(|CO| − |RC|)/avg_bal`, then the **simple mean** of yearly rates (not balance-weighted) |
| `_pool_dq_variance` | rv:496-526 | `last_rate − mean(rates)` over the pool's `acl_months` window; needs ≥2 observations |
| `_score` / `_env_ranges` | rv:322, 199-208 | Maps NCC/DQ/ES onto scoring bands; falls back to module constants `NCC_RANGES`/`DQ_RANGES`/`ES_RANGES` |
| `_resolve_mgmt_adj_grade` / `_total` | rv:359-398 | Precedence engine: manual overlay × dist > admin default × dist (only when `use_default` and base rate 0) > **prior-report carry-forward** > 0 |
| `_dist_factor` | rv:329-330 | Grade index → factor from a hard-coded `DIST_FACTORS` list, clamped |
| `_other_allowance_considerations` | rv:401-426 | **Recomputes `amount` from `balance × pct/100` when missing or stale** — the config value is not authoritative |
| `_ordered_pools` | rv:210-254 | Canonical pool ordering; unions `df`, `hist_bal_data`, `pool_bal_detail`; NRR pools forced last |
| `_windowed_year_val` | rv:3371-3410 | Prorates a partial year's CO/recoveries when the pool's ACL window covers only part of it; normalizes sign |
| `_sheet_impdet` | rv:1213-1223 | Improved/deteriorated percentages computed **from cells already written** |

### 3.6 Charts and post-save XML patching

Every chart is an openpyxl `BarChart`/`LineChart`/`PieChart`/`DoughnutChart`
bound to `Reference(ws, ...)` cell ranges — the chart's data source *is* the
worksheet. Several builders write hidden data tables solely to feed charts
(e.g. Impr Deter's red-font table at rows 45+, outside the print area,
rv:1145-1220).

| Tab | Charts |
|---|---|
| Impr Deter | 4 `BarChart` — Improved by grade (rv:1353-1395), Deteriorated by grade (rv:1398-1437), stacked Improved/Deteriorated by pool (rv:1441-1524), Net Change by pool (rv:1528-1582). Per-point label placement is value-dependent (`_small_indices` rv:1331, `_c3_small` rv:1506) |
| Risk Change Total / Risk Chg * | `DoughnutChart` (rv:1963-1990), `BarChart` by grade (rv:1993-2035, explicitly drops the "Not Reported" row at rv:2007), DQ `PieChart` (rv:2051-2104), charge-off `BarChart` (rv:2120-2182). **Instantiated once per risk-rated pool**, so chart count scales with pool count |
| > Historical Trends Balance (supplemental) | one `LineChart` per pool, series per grade, all `Reference`s into a *different* sheet (rv:3943-4053) |

After `wb.save()` (gr:13035), gr:13043-13046 re-opens the saved `.xlsx` and
rewrites chart XML in place — each uses `zipfile` + `ElementTree`, writes to a
`mkstemp` file, and moves it back over the original:

1. `patch_dq_pie_zero_labels` (rv:5200-5338) — reopens the file *twice*: once
   with `load_workbook(data_only=True)` to map `(sheet, col, row) → value` for
   every "Risk Chg"/"Risk Change" sheet (rv:5210-5223), then over
   `xl/charts/chart*.xml` to inject `<c:delete val="1"/>` on labels whose
   backing cell is zero (openpyxl's `DataLabel` has no `delete` support).
2. `patch_impdet_charts` (rv:5339-5576) — the heaviest: resolves the Impr Deter
   drawing through workbook rels, **classifies charts by anchor quadrant**, then
   applies number formats, title/series/axis normalization, manual layout, and
   graduated bar transparency (rv:4926, 5505-5512). Carries a documented
   namespace hazard (rv:5344-5345, 5418-5424).
3. `patch_drawing_onecell_to_twocell` (rv:5577-5656) — rewrites every
   `oneCellAnchor` to `twoCellAnchor` and strips a stale default `xmlns=` left
   by patch 2.
4. `patch_remove_chart_borders_and_axis_lines` (rv:5657-5790) — `noFill` on
   chart-area and axis lines across all charts; normalizes titles.

`report_integrity.check_and_report` exists only to catch bugs these patches
introduce. Rendering charts directly makes all four patches, the ~360-line XML
helper layer (rv:4834-5198), and `report_integrity` obsolete.

### 3.7 Excel print-layout logic

`_paginate_pool_blocks` (rv:6364), `_fit_to_pages` (rv:6383), and
`_add_page_numbers` (rv:6407-6423) encode page-break and fit-to-page behaviour.
`_add_page_numbers` stamps `"Page &P of &N"` and its docstring explicitly notes
it relies on Excel treating the whole-workbook PDF export as one continuous
print job — a direct renderer owns pagination itself.

---

## 4. Tabs that derive from ANOTHER TAB (the hard cases)

Five distinct mechanisms.

### 4.1 `_parse_acl_layout(ws)` — re-reading a built sheet to recover row numbers

rv:5791-5863. Iterates the **already-built** `ACL Env by Pool Mgmt Adj`
worksheet from row 6, matching column-A label text (`"pooled totals"`,
`"impaired loans"`, `"total specifically identified"`, `"other allowance
considerations"`, `"total allowance needed"`, `"allowance for credit loss
balance"`, `"adjustment"`) to build:

```
{"pools": [{"name", "header_row", "grades": [(label, row)], "total_row"}],
 "impaired": [(label, row)], "oac": [(label, row)],
 "totals": {"pooled","spec","oac","needed","balance","adjustment","adjustment_label"}}
```

Pool headers are detected **structurally** — "a label whose Balance column is
empty" (rv:5855-5861). Called by four builders, each of which bails out entirely
if the ACL sheet is absent:

| Caller | `_parse_acl_layout` at | Guard at |
|---|---|---|
| `_sheet_acl_summary` | rv:5901 | rv:5899 |
| `_sheet_mgmt_adj_summary` | rv:6010 | rv:6007 |
| `_sheet_impaired_loans` | rv:6104 | rv:6101 |
| `_sheet_summary_variance` | rv:6193 | rv:6191 |

### 4.2 Cross-sheet Excel formulas written into the summary tabs

`_ACL_REF = "'ACL Env by Pool Mgmt Adj'!"` (rv:5789). Sixteen formula-emitting
sites:

- **ACL Summary** (rv:5920, 5925, 5942, 5957) — even the pool *name* is a
  formula `='ACL Env...'!A{header_row}`; the seven numeric columns map ACL
  columns `B,C,D,H,I,J,K` of each pool's Total row onto destination columns 2-8.
- **Mgmt Adj Summary** (rv:6038, 6047, 6052, 6063, 6084) — grade rows pull
  `B,E,F,G,H`; total rows pull `B,H,I,J`.
- **Impaired Loans** (rv:6117, 6129, 6151, 6155, 6166) — column `K` for
  impairment categories, column `C` for per-pool specific ID.
- **Summary Variance** (rv:6294-6302) — Current block references `K{needed}`,
  `K{balance}`, `K{adjustment}`, and computes ACL/Total-Loans as
  `=IFERROR('ACL Env...'!K{needed}/'ACL Env...'!B{pooled},"")`. The Change block
  is `=IF(C{prior}="","",C{cur}-C{prior})` (rv:6330) — an intra-sheet formula
  over its own Prior block. Grid rows are hard-coded: current 10-13, prior
  16-19, change 22-25 (rv:6275).

The module comment states the intent plainly (rv:5779-5785): "written as Excel
formulas pointing at that sheet rather than as recomputed numbers, so a summary
can never disagree with the tab it summarises."

**Consequence for the PDF work:** these four tabs have no independent data
source, and a `data_only=True` read of the delivered workbook returns `None` for
every one of these cells unless Excel has opened and recalculated the file.

### 4.3 Builders reading cell **values** out of a sheet they did not build

Worse than formulas, because control flow depends on the read.

| Site | Code | Effect |
|---|---|---|
| rv:6025-6027 | `_val(row, col)` → `src_ws[col + str(row)].value` | **Mgmt Adj Summary** decides which grade rows exist at all: `adj_grades = [(g,gr) for g,gr in p["grades"] if _val(gr,"F")]` (rv:6032), and skips whole pools when `not adj_grades and not env_factor` (rv:6034-6035) |
| rv:6137-6139 | `_spec(p)` → `src_ws["C"+str(p["total_row"])].value` | **Impaired Loans** filters to `pools_with_spec = [p for p in lay["pools"] if _spec(p)]` (rv:6141) |
| rv:3884-3928, 4030 | `detail_ws = wb["> Detail_HIst Balances"]`, then `.cell(...).value`, `.font.sz`, `.font.bold` | **`> Historical Trends Balance`** (supplemental) discovers pool blocks by scanning the **formatting** of another sheet — bold + `sz >= 9` means "pool header" (rv:3898) — counts `datetime` cells to find date columns, walks to the `'Total'` row, and pulls series names from it (rv:4030) |
| rv:5210-5223 | `load_workbook(saved_path, data_only=True)`, then `cell_vals[(ws_name, col, row)] = cell.value` across Risk Chg sheets | `patch_dq_pie_zero_labels` reads the **saved file** to locate zero pie slices |
| rv:6448-6494 | `dst = wb["Env Factor by Pool"]`; `src = wb[">Envir Fact Ranges"]`; copies values + styles + merges + row heights into `dst`, reading `dst["A1"].value` (rv:6470) to decide which source rows to drop | `_merge_env_ranges_into_factor` (called rv:6689) — one tab's content is physically relocated into another and the source sheet hidden (not deleted, "so cell references into it keep resolving") |

### 4.4 `_parse_acl_sheet` — the same trick, applied to a *prior quarter's file*

ca:68-124. Structurally the twin of `_parse_acl_layout` but returns **values**
rather than rows, scraping fixed columns B/C/J/K:

```
{"pools": {name: {balance, spec_id, total_allow}}, "order": [...],
 "impaired": {category: allowance},
 "totals": {pooled_balance, pooled_total_allow, total_spec_allow,
            total_allow_needed, acl_balance, adjustment}}
```

Two consumers, both reading a prior quarter's saved `.xlsx` off disk:

- `ca.append_change_analysis` (ca:244) — parses the current workbook's ACL Env
  sheet (ca:246) *and* `load_workbook(prior_path, read_only=True,
  data_only=True)` (ca:280), then diffs per pool (ca:296-310).
- `_sheet_summary_variance` (rv:6200-6218) — imports `_find_prior_report` and
  `_parse_acl_sheet` from `change_analysis` and repeats the same read,
  deliberately, "so this tab and the Change Analysis tab can never disagree
  about which workbook is 'prior'" (rv:6205-6207). **The prior workbook is
  therefore opened twice per run.**

`_find_prior_report` (ca:128-142) globs
`{rpt_dir}/*_CECL_Migration_{safe_cu}_Vizo_Model.xlsx`, regex-parses the
`YYYY-MM-DD` filename prefix, and picks the newest strictly before `snap`.
`rpt_dir` is re-derived from `CECL_WORKSPACE_ROOT` independently in three places
(rv:59, ca:262-263, rv:6205-6210).

**The prior quarter's numbers exist only inside a prior `.xlsx`.** There is no
database or serialized store. And because the read is `data_only=True`, it
returns Excel's *cached* values — a prior report generated but never opened in
Excel yields `None` for every formula cell, including all four summary tabs.

### 4.5 Side-effect coupling through the input dict

`_sheet_acl_reserve` writes `_computed_pooled_total_allow` (rv:2680) back onto
`hist['impaired']`; `_compute_acl_totals` reads it first in its fallback chain
(rv:599-600), which is how `Impr Deter` gets the same "Total Allowance Needed"
as `ACL Env`. Not a worksheet read, but the same class of problem: a tab's
output is another tab's input, expressed as mutation plus build ordering.
`compose_vizo_main:6570-6581` documents the resulting build-order/display-order
divergence. The two sibling stashes (rv:2681-2682) are never read at all.

### 4.6 The dependency graph

```
df, config, grades, hist
        │
        ├──────────────► Vizo Cover, Report Index            (+ logos, literal prose)
        ├──────────────► Risk Change Total, Risk Chg *       (+ info icons)
        ├──────────────► Env Factor by Pool ◄── >Envir Fact Ranges   (cell copy, §4.3)
        ├──────────────► Display HIst Bal, Display CO-Recov-DQ
        │
        └──► ACL Env by Pool Mgmt Adj ── stash on hist['impaired'] ──► Impr Deter
                     │
                     ├── _parse_acl_layout (rows) ──► ACL Summary
                     │                              ├► Mgmt Adj Summary  (+ value reads)
                     │                              ├► Impaired Loans    (+ value reads)
                     │                              └► Summary Variance
                     │
                     └── _parse_acl_sheet (values) ──► Change Analysis
                                                       Summary Variance (Prior block)
                                                            ▲
                                       prior quarter's .xlsx on disk ──┘
                                       (data_only=True → needs Excel's cache)

external template .xlsx ──► Introduction-Vizo, Executive Summary-Vizo,
                            theme1.xml palette, info icons     [NOT IN GIT]
```

---

## 5. A proposed `ReportData` contract

Design goals, in priority order:

1. **`ACL Env by Pool Mgmt Adj` stops being a data source.** Its computation
   produces a structured `AclEnvironmental`; the four summary tabs and
   `Impr Deter` read that object. This alone removes `_parse_acl_layout`
   (rv:5791), all sixteen cross-sheet formula sites, both `_val`/`_spec` cell
   reads, and the `_computed_*` stash.
2. **Prior-period figures become a first-class input.** Emit
   `{snap}_{cu}_acl.json` alongside every report; `_find_prior_report` globs
   that. `_parse_acl_sheet` (ca:68) survives only as a legacy importer for
   periods generated before the change — and stops depending on Excel's cache.
3. **Narrative prose moves out of both the builders and the missing template**
   into versioned templates keyed by tab.

```python
# ── leaves ────────────────────────────────────────────────────────
@dataclass(frozen=True)
class GradeBand:
    label: str
    min_score: int | None
    max_score: int | None
    reserve_rate: float
    hidden: bool                 # label.startswith('Hide')          rv:304
    range_text: str              # "700+" / "600-649"                rv:538

@dataclass(frozen=True)
class AclGradeRow:
    """One grade line inside a pool block on ACL Env (cols A..H)."""
    grade: str                   # A
    balance: float               # B
    specific_id: float           # C
    calc_balance: float          # D
    base_loss_rate: float        # E
    mgmt_adj: float              # F
    allowance_factor: float      # G
    allowance_before_env: float  # H

@dataclass(frozen=True)
class AclPoolBlock:
    name: str
    risk_rated: bool             # <- hist['impaired']['risk_rated']
    brr: bool                    # renders BRR labels, not FICO       rv:264
    grades: list[AclGradeRow]
    total: AclGradeRow           # the pool "Total" row
    env_factor: float            # I
    env_allowance: float         # J
    total_allowance: float       # K
    # env-factor provenance -> also renders Env Factor by Pool
    ncc_pct: float;   ncc_score: float
    dq_variance: float; dq_score: float
    econ_stress: float; es_score: float
    life_loss_rate: float
    acl_months: int
    prior_mgmt_adj: dict[str, float]   # <- impaired['prior_mgmt_adj']
    prior_env_factor: float | None

@dataclass(frozen=True)
class OacRow:
    title: str; balance: float; percentage: float; amount: float   # rv:401

@dataclass(frozen=True)
class AclEnvironmental:
    """Everything ACL Env by Pool Mgmt Adj renders. THE hub object."""
    pools: list[AclPoolBlock]           # already in _ordered_pools order
    impaired: list[tuple[str, float]]   # (category, allowance) <- acl_impaired
    other_considerations: list[OacRow]
    pooled_balance: float
    pooled_specific_id: float
    pooled_calc_balance: float
    pooled_allow_before_env: float
    pooled_env_allowance: float
    pooled_total_allowance: float
    total_specific_id: float
    total_oac: float
    total_allowance_needed: float
    acl_balance: float
    adjustment: float
    adjustment_label: str               # "Adjustment (Underfunded|Overfunded)"
    balance_label: str                  # "...as of 12/31/2025"      rv:667-679
    acl_over_total_loans: float | None  # needed / pooled_balance

@dataclass(frozen=True)
class RiskChangeTab:
    """Risk Change Total, and one per risk-rated pool."""
    pool: str | None                 # None => Risk Change Total
    labels: list[str]                # FICO grades or BRR bands
    score_ranges: dict[str, str]
    matrix: dict[tuple[str, str], float]   # (current, original) -> balance
    improved: float; deteriorated: float; unchanged: float
    dq: dict; charge_offs: dict            # dq_by_pool/status, co_by_pool/status
    balance_adjustment: float | None       # per-pool row
    total_balance_adjustment: float | None # total tab only
    total_in_portfolio: float | None

@dataclass(frozen=True)
class EnvRangeTable:
    kind: Literal['ncc', 'dq', 'es']
    bands: list[tuple[float, float, float]]   # lo, hi, score
    labels: list[str]                         # <- env_ranges['*_labels']

@dataclass(frozen=True)
class HistorySeries:
    """Display HIst Bal + Display CO-Recov-DQ."""
    years: list[int]
    avg_balances: dict[int, dict[str, float]]
    charge_offs: dict[int, dict[str, float]]
    recoveries: dict[int, dict[str, float]]
    co_monthly: dict[tuple[int, int], dict[str, float]]
    rc_monthly: dict[tuple[int, int], dict[str, float]]
    dq_pct: dict[int, dict[str, float]]
    hist_bal: dict[str, dict]        # <- impaired['hist_bal_data']
    warm_window_months: dict[str, int]   # <- impaired['acl_months']
    warm_net_co: dict[str, float]

@dataclass(frozen=True)
class ImprDeter:
    by_grade: list[tuple[str, float, float]]   # grade, improved $, deteriorated $
    by_pool:  list[tuple[str, float, float]]
    # the improved%/deteriorated% now read back out of cells (rv:1213-1223)
    # become derived properties of this object.
    adjustment_box: list[tuple[str, float]]    # 4 CECL Adjustment lines

@dataclass(frozen=True)
class PriorPeriod:
    """Replaces _parse_acl_sheet + _find_prior_report (ca:68, ca:128)."""
    snap: str
    pools: dict[str, dict]        # name -> {balance, spec_id, total_allow}
    order: list[str]
    impaired: dict[str, float]
    totals: dict[str, float]      # pooled_balance, pooled_total_allow,
                                  # total_spec_allow, total_allow_needed,
                                  # acl_balance, adjustment

@dataclass(frozen=True)
class Branding:
    cu_name: str
    vizo_logo: Path; tct_logo: Path
    info_icon_red: Path | None; info_icon_green: Path | None
    theme_colors: dict[str, str]  # literal hex; replaces the theme1.xml unzip, rv:100

@dataclass(frozen=True)
class Narrative:
    """Everything currently hard-coded or copied from the missing template."""
    cover_disclaimer: str                     # rv:781-793
    report_index_blocks: list[tuple[str, str]]# rv:858-912
    env_ranges_prose: list[tuple[str, str]]   # rv:3752-3846 (the big one)
    risk_change_footnotes: tuple[str, str]    # rv:2183-2192
    introduction: list[str]                   # <- Introduction-Vizo template tab
    executive_summary: list[str]              # <- Executive Summary-Vizo tab
    no_prior_report_note: str                 # rv:6350
    no_mgmt_adj_note: str                     # rv:6069

# ── the contract ──────────────────────────────────────────────────
@dataclass(frozen=True)
class ReportData:
    credit_union: str
    snap: date
    branding: Branding
    narrative: Narrative

    grades: list[GradeBand]
    pool_order: list[str]                 # <- _ordered_pools(df, hist)
    risk_rated: dict[str, bool]
    economic_data: dict                   # incl. `_sources` footnote list

    acl: AclEnvironmental                 # hub; every summary derives from this
    impr_deter: ImprDeter
    risk_change_total: RiskChangeTab
    risk_change_by_pool: list[RiskChangeTab]
    env_ranges: list[EnvRangeTable]
    history: HistorySeries

    prior: PriorPeriod | None             # None => "earliest report on file"

    tab_order: tuple[str, ...] = _VIZO_MAIN_ORDER   # rv:6518 — the page order
```

### How each problem tab renders under the contract

| Tab | Source |
|---|---|
| ACL Summary | `acl.pools` one line each + `acl.pooled_*` + `acl.impaired` + `acl.other_considerations` + the four total lines. No formulas. |
| Mgmt Adj Summary | `acl.pools`, filtering on `row.mgmt_adj != 0` and `pool.env_factor != 0` — a field test replacing the `_val(...)` cell reads at rv:6025-6035. |
| Impaired Loans | `acl.impaired` + `[p for p in acl.pools if p.total.specific_id]` — replaces `_spec(...)` at rv:6137. |
| Summary Variance | Current = `acl.{total_allowance_needed, acl_balance, adjustment, acl_over_total_loans}`; Prior = `prior.totals`; Change = subtraction in Python, not `=IF(C16="",…)`. |
| Change Analysis | `acl` vs `prior`, diffed in Python. `_explain_pool` / `_explain_impaired` (ca:171, 218) port unchanged — they already take dicts. |
| Impr Deter | `impr_deter` + `acl.adjustment_box`; the stash at rv:2680 disappears because `acl` is computed once, before any rendering. |
| Env Factor by Pool | `acl.pools` (the `ncc_*`/`dq_*`/`es_*`/`env_factor` provenance) then `env_ranges` + `narrative.env_ranges_prose` — the merge at rv:6429 becomes two sections of one page, not a cell copy, and `>Envir Fact Ranges` stops existing as a tab. |
| Introduction / Executive Summary | `narrative.introduction` / `narrative.executive_summary`. |
| > Historical Trends Balance (supp) | `history.hist_bal` directly — no font-scraping of `> Detail_HIst Balances` (rv:3888-3928). |

### Migration order this implies

1. **Extract `AclEnvironmental`** out of `_sheet_acl_reserve` (rv:2223-2776) as
   a pure `compute_acl(df, config, grades, hist, snap) -> AclEnvironmental`.
   Have the existing builder render *from* it. Delete the `_computed_*` stash
   (rv:2678-2682) and the ordering hack at rv:6570-6581 in the same change.
   `report_acl_funding.py:72` needs updating alongside.
2. **Rewrite the four summary builders** to consume `AclEnvironmental` instead
   of `_parse_acl_layout`. Output should be byte-identical apart from formulas
   becoming literals — a small, verifiable diff given the workbook is already
   95.9% literal.
3. **Serialize `AclEnvironmental` to JSON** next to each saved report; point
   `_find_prior_report` at it, keeping `_parse_acl_sheet` as a fallback
   importer. This removes the `data_only=True` dependency on Excel's cache.
4. **Lift the narrative text** out of `Sample Reports/Vizo Narrative Tabs -
   Template.xlsx` and `theme1.xml` into the repo. **Requires the analyst
   workspace — neither file is in version control** (`.gitignore:45`).
   `_sheet_introduction` (rv:971) and `_sheet_exec_summary` (rv:1011), though
   dead, hold a Python-native version of that text worth diffing against.
5. **Re-author the charts** against `ReportData` series. This retires
   `patch_dq_pie_zero_labels`, `patch_impdet_charts`,
   `patch_drawing_onecell_to_twocell`,
   `patch_remove_chart_borders_and_axis_lines`, the XML helper layer
   (rv:4834-5198), and `report_integrity`.
6. **Build the PDF renderer** against `ReportData`, owning its own pagination
   (replacing `_paginate_pool_blocks`, `_fit_to_pages`, `_add_page_numbers`).

Steps 1-3 are worth doing regardless of the PDF work: they remove the
build-order coupling, the hidden-sheet dance, and the silent dependency on Excel
having recalculated prior workbooks.
