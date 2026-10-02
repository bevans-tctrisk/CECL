# WARM Wizard Auto-Derive & Cleanup

**Status:** Implemented + validated · **Last updated:** 2026-07-23
**Scope:** The setup-wizard streamlining that wires the validated resolver/assembler path
(`parse_warm` → `config_assembler` → resolvers) into the UI so a WARM upload auto-derives most
of the config, marks the derived steps complete, and produces a config that *ties to the WARM*.

> North-star (same as everything else): **derive → tier → review → tie-to-book → save.**
> The auto-derive fills the config; the human confirms yellow/red and the balance-adjustment
> step must tie before save.

---

## 1. Why

Audit finding (`docs/` companion + memory `cecl-wizard-cleanup-audit.md`): the wizard ran on a
**second, older WARM parser** (`cecl_ui/services/warm_parser.py :: analyse_warm_file`, 50 KB)
and a **20-step manual flow**, while the validated path (`cecl_ui/services/setup/warm_parser.py
:: parse_warm` + `config_assembler` + `resolvers/`) derived more, more accurately, and was
tie-validated across 6 WARM CUs. The old parser lacked: column-mapping derivation (Data-tab
formula tracing), reserve col-G pinning, SSN credit-pull detection, the keep-first pool_map fix,
and WARM pool order — so users hand-entered things the resolvers now derive and validate.

Context: ~100 more CUs are coming, **most with WARM files**. The WARM flow must be a fast lane.

---

## 2. The bridge — `cecl_ui/services/setup/warm_autoderive.py`

`derive(warm_path, *, short_name, snapshot_date, reports) -> {ok, warm, config, decisions,
overall_tier, error}`

- Runs `parse_warm` + `config_assembler.build_config_from_warm` (delivery-specific fields left
  blank for the loan-extract step).
- Builds a **fallback-only `credit_pull` block** (the assembler only builds credit_pull when
  given a standalone filename, unknown at WARM-upload time): `fallback_report_folder = <warm
  folder>`, `fallback_report_pattern = ^<YYYY-MM> CECL-Migration.*\.xlsx$` (from the WARM
  filename), member/score cols, `pull_as_of_date = snapshot`, plus the **SSN join**
  (`loan_join_column`, `member_column='SSN'`) when the pull is SSN-keyed. This is what lets the
  report load the WARM's scores/ACL/CO/DQ at report time.
- Returns a **tiered decision list** (`kind, label, value, tier, source, evidence`) for the
  review surface: identity, pools, loan_code_map, credit_grades, column_mappings,
  member_account, monthly_balance, reserve_config, credit_pull, economic, pool_order.
  `overall_tier` = worst of the decisions.

### Tier logic
Green when the WARM confidently resolves the field; **yellow** when it's ambiguous. Notably
`member_account` is **yellow** when the Data-tab trace lands on a computed "Member" helper column
(`member_source ∈ {member, member #, member-sfx, …}`) instead of the raw account column — the
delimiter/split-CU gap (Honolulu `L`, Maple `LOAN SUFFIX`). Evidence tells the user to confirm
the real account column + delimiter on the loan-extract step.

---

## 3. Wiring into the wizard  (`cecl_ui/routes/setup.py`)

- **WARM upload handler** (`step2_warm`): after `_apply_warm_to_state`, calls
  `warm_autoderive.derive(...)`, stores `state["autoderive"] = {ok, overall_tier, decisions}`,
  and calls `_merge_autoderive_into_state`. Then sets
  `state["_warm_autoderived_steps"]` = the 11 fully-covered steps.
- **`_merge_autoderive_into_state(state, cfg)`**: overlays the derivations onto wizard state,
  writing a field only when it's still at its factory default (user edits always win):
  `column_mappings` (Data-tab trace), `member_account`, `pool_map` (keep-first), `credit_grades`,
  and the full `credit_pull` fallback + SSN join. Stashes `state["warm_reserve"] =
  {base_loss_rate_by_pool_grade, warm_allowance_pools, not_risk_rated, monthly_balance}` for the
  YAML emit.
- **`config_service.build_yaml_from_wizard`**: previously did **not** emit the reserve pinning,
  so wizard-built WARM configs could not tie. Now overlays `state["warm_reserve"]` keys before
  `return cfg` (no-op for non-WARM CUs). column_mappings / member_account / pool_map / credit_pull
  already flowed through the existing emit; pool_order / not_risk_rated also via `set_pools` from
  the pool-settings order.

---

## 4. Step collapse (no handler/redirect surgery)

The stepper and its status/HIL banner already support "auto-fill → mark complete → guide to
input" (built for the Step-1 folder scan). We reuse it instead of removing handlers (whose
"Next" redirects are hardcoded):

- `auto_setup.compute_step_completion`: `done |= set(state.get("_warm_autoderived_steps") or [])`
  → the 11 derived steps render **green ✓ (review-only)**.
- `routes/setup.py :: _wizard_ctx` gate widened from `if _auto_scan_completed:` to
  `if _auto_scan_completed or autoderive.ok:` → WARM CUs get the badges + HIL banner without a
  Step-1 folder scan. Recommended-HIL on green steps is auto-filtered; the banner steers the user
  to the interactive steps.

**Auto (green, review-only):** pools, balances, grades, columns, economic, mgmt_adj, files,
dq_hist, co_recov, orig_score, credit_pull.
**Interactive (kept):** identity (confirm), warm (upload+review), sample (loan extract — needs the
delivery file + headerless/positional fallback), monthly_bal (ACL), baseline (WARM→baseline
import action), balance_check (tie-to-book gate), impaired (current quarter), reports, review.

### Review panel
`templates/setup/step2_warm.html` renders an "Auto-derived from WARM" card with a green/yellow/red
overall badge and one row per decision (colored dot + evidence), inline-styled (no CSS dependency).

---

## 5. Bug fixed during validation

`config_assembler.build_config_from_warm` emitted the **`HIDE-Loan Type*` placeholder pools into
`not_risk_rated`** (Honolulu 14, Franklin 9) because they read as "balance-only" (no grade
balances) in the WARM ACL tab. Fixed with a `_real_pool()` filter (drops names starting `hide` /
== `exclude`). Verified: Honolulu 2026-03 → `not_risk_rated = ['Loan Participation', 'Coll in
Process of Liquidation']`, matching the hand-built config.

---

## 6. Validation — all system CUs through the wizard path

39 client configs; **8 have a reachable WARM** (bridgeton, erie_fcu, franklin, honolulu, maple,
ontario, tongass, utah); the other 31 are SCALE/from-scratch (separate flow). Harness: for each
WARM CU, `derive` → fresh `_default_state` → `_merge_autoderive_into_state` →
`build_yaml_from_wizard`, then diff the WARM-responsible fields vs the existing validated config.

Result (5 completed; Bridgeton/Erie/Utah have slow-parsing WARMs): the **core tying fields
reproduce** — `base_loss_rate_by_pool_grade` 21/21 pools on all five, plus credit_grades,
warm_allowance_pools, pool_order (clean CUs), credit_pull SSN join, and member_account/pool_map for
clean CUs. Remaining diffs are (a) the known member-account gaps (delimiter/split/headerless —
yellow-flagged, corrected at the loan-extract step), (b) non-core column mappings filled at the
loan-extract step, and (c) benign period/pool_map differences (harness used the newest WARM vs the
older existing config). Full Flask render of `GET /setup/step/warm` → 200 with the panel + badges.

---

## 7. Files touched

| File | Change |
|---|---|
| `cecl_ui/services/setup/warm_autoderive.py` | **new** — bridge (derive + tiered decisions + credit_pull fallback + tier logic) |
| `cecl_ui/routes/setup.py` | WARM-upload wiring, `_merge_autoderive_into_state`, `_warm_autoderived_steps`, `_wizard_ctx` gate |
| `cecl_ui/services/config_service.py` | `build_yaml_from_wizard` emits `warm_reserve` (base-loss pins / warm-allowance / not_risk_rated / monthly_balance) |
| `cecl_ui/services/auto_setup.py` | `compute_step_completion` honors `_warm_autoderived_steps` |
| `cecl_ui/services/setup/config_assembler.py` | `not_risk_rated` filters `HIDE-*`/`Exclude` (`_real_pool`) |
| `cecl_ui/templates/setup/step2_warm.html` | "Auto-derived from WARM" tiered review panel |

---

## 8. Roadmap

1. **Live wizard save → YAML → report tie** for one CU (next).
2. Wire delimiter/split `member_account` detection into the WARM parse to shrink the yellow cases.
3. Optionally consolidate the 11 green steps into a single "Auto-derived (review)" group to
   shorten the visible stepper further.
4. Surface `origination_aware` as a tie-tested toggle on the credit-pull step (it's CU-specific:
   Tongass wants it, Ontario/Maple don't — decided by which ties).
5. Retire `cecl_ui/services/warm_parser.analyse_warm_file` once the wizard reads everything from
   `parse_warm`.
6. Apply the same treatment to the SCALE path where NCUA-5300/Solr templates allow.
