# Project Rules for Claude

- **Code style:** Standard PEP 8 Python formatting.
- **Tests:** Run tests with `pytest` before declaring a feature complete.
- **Fallback model:** Use `claude-opus-4.8` via the `anthropic` SDK for API escalations.

## Hybrid report router

- The hybrid schema router lives in `cecl_ui/services/hybrid/` and is **opt-in per CU**.
  Enable it in `client_configs/<cu>.yaml`:

  ```yaml
  hybrid_router:
    enabled: true          # default: false (off) — see run_reports_hybrid
    proactive: true        # escalate to the API when a CU has no schema yet
    model: claude-opus-4.8
  ```

- Local validation (Pydantic + rules) runs first; the API is only called when
  validation fails or a CU has no registered schema.
- The router must **never block** report generation — always degrade to the
  existing pipeline on any failure.
- Store secrets via `cecl_credentials` (Windows Credential Manager), never in code.
  DB URL: `get_database_url()`. Anthropic key: `get_anthropic_api_key()`.
- Never send `monthly_loan_data` rows or member PII to the API — schema/column
  metadata only.
