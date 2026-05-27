# CECL

Tools for setting up and running **CECL (Current Expected Credit Loss)** reports for credit unions, built around a guided web wizard plus a set of report-generation engines.

This repository contains the in-house TCT toolset used to:

- Configure a credit union end-to-end through a Flask wizard (loan pools, balance titles, credit grades, historical data, sample uploads, column mappings, charge-offs/recoveries, impaired loans, economic factors, management adjustments, etc.).
- Backfill historical data from NCUA 5300 Call Report data via Solr.
- Import per-month loan-level data into PostgreSQL.
- Generate the three report flavors: **TCT**, **Vizo**, and **Impairment Detail**.

---

## Repository layout

```
cecl_ui/                  Flask wizard (routes, services, templates, static)
  routes/                 Blueprints (setup wizard, run pipeline, admin, scale)
  services/               Parsers, processors, Solr backfills, etc.
  templates/setup/        Wizard step templates
  data/                   Reference CSVs (NCUA 5300 codes, SCALE maps, etc.)

cecl_engine.py            Core CECL calculation engine
generate_report.py        Main report orchestrator (TCT)
report_tct.py             TCT workbook writer
report_vizo.py            Vizo workbook writer
generate_impdet_report.py Impairment Detail report
import_data.py            Loan-extract -> PostgreSQL importer
run_quarterly.py          Quarterly batch runner
dashboard.py              Dashboard helpers

run_ui.py                 Entry point for the Flask wizard
start_cecl_wizard.ps1     PowerShell launcher (uses .venv)
requirements.txt          Python dependencies
vizo_theme.xml            Excel theme used for Vizo workbooks
```

Per-credit-union artifacts (excluded from git):

```
Raw_Uploads/<cu>/         Source extracts uploaded for each CU
Generated_Reports/<cu>/   Output workbooks (one folder per snapshot date)
client_configs/           Per-CU YAML config files
wizard_drafts/            In-progress wizard state (one file per CU/model)
```

---

## Requirements

- Windows 10/11 (development environment)
- Python 3.11 (in a local `.venv`)
- PostgreSQL instance reachable via `.env` (`cecl_migration_db` is the default DB used in development)
- Git (PortableGit works fine — no admin needed)

Python dependencies are pinned in [requirements.txt](requirements.txt).

---

## First-time setup

```powershell
# 1. Create / activate venv
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 2. Install dependencies
pip install -r requirements.txt

# 3. Create .env with DB credentials and any secrets (NOT committed)
#    Required keys: DATABASE_URL, FLASK_SECRET_KEY, ...

# 4. Launch the wizard
.\start_cecl_wizard.ps1
# or:
python run_ui.py
```

The wizard is served at <http://localhost:5000>.

> Note: the Flask app runs with `use_reloader=False`, so changes to `.py`
> files require a manual restart.

---

## Wizard flow (high level)

The wizard has two paths depending on whether the credit union provides a WARM workbook:

**With WARM upload (20 steps):** identity → warm → balances → baseline → dq_hist → monthly_bal → grades → credit_pull → orig_score → sample → columns → pools → balance_check → co_recov → impaired → files → economic → mgmt_adj → reports → review

**Without WARM (18 steps):** identity → historical → dq_hist → monthly_bal → grades → credit_pull → orig_score → sample → columns → pools → balance_check → co_recov → impaired → files → economic → mgmt_adj → reports → review

Saved progress lives in [wizard_drafts/](wizard_drafts/) as `<cu-slug>__<model>.json` so the Migration and SCALE wizards can be filled in side-by-side for the same CU.

---

## Generating a report

Once a CU is configured (YAML written to `client_configs/<cu>.yaml`) and its extracts are imported, generate a report via the wizard's **Run** page or directly:

```powershell
python run_quarterly.py --cu "Credit Union Name" --snapshot 2026-03-31
```

Output workbooks are written under `Generated_Reports/<cu>/<snapshot>/`.

---

## Working with the repo

```powershell
git add -A
git commit -m "describe the change"
git push
```

VS Code's Source Control panel (Ctrl+Shift+G) works for stage/commit/push as well.

---

## License

Proprietary — internal TCT use. Not for redistribution.
