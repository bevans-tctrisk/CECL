"""API escalation — the fallback stage of the hybrid router.

When local validation fails, or a CU has no registered schema yet, this module
asks ``claude-opus-4.8`` (via the Anthropic Python SDK) to propose a
corrected/complete report schema. The model returns **data only** (JSON); that
JSON is re-validated through the Pydantic contract before the router is allowed
to trust or persist it. AI output is never executed.

Every dependency here is optional and lazily imported. When the SDK is not
installed, no API key is configured, or the network call fails, the functions
return a result with ``schema=None`` and an explanatory message so the router
can degrade gracefully to today's behavior.

Privacy: only schema/column *metadata* is sent — never rows from
``monthly_loan_data`` or member PII.
"""
from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from typing import Any

from pydantic import ValidationError

from .schemas import KNOWN_REPORT_TYPES, ReportSchema

DEFAULT_MODEL = "claude-opus-4.8"
_MAX_TOKENS = 4096

_SYSTEM_PROMPT = (
    "You are a CECL reporting schema assistant for a credit-union analytics "
    "tool. Given a raw client configuration and a list of validation errors, "
    "return a corrected report schema as STRICT JSON matching this shape:\n"
    "{\n"
    '  "credit_union": str,\n'
    '  "report_type": str,\n'
    '  "pools": [{"name": str, "risk_rated": bool, "acl_months": int|null, '
    '"excluded": bool, "use_default_mgmt_adj": bool}],\n'
    '  "credit_grades": [{"label": str, "min_score": int, "max_score": int, '
    '"reserve_rate": float}],\n'
    '  "column_mappings": {str: str},\n'
    '  "reports": {str: bool},\n'
    '  "default_pool": str|null\n'
    "}\n"
    "Rules: reserve_rate is a decimal in [0,1]; credit-grade score bands must "
    "not overlap; include at least one non-excluded pool; column_mappings MUST "
    "include member_number and current_balance. Respond with ONLY the JSON "
    "object, no prose, no markdown fences."
)

# Config keys that may contain local paths / uploaded filenames. Stripped
# before the payload leaves the machine.
_SENSITIVE_KEYS = {
    "source_folder",
    "saved_path",
    "fallback_report_folder",
    "loan_source_folder",
    "loan_file_folder",
    "uploaded_filename",
}


@dataclass
class EscalationResult:
    """Outcome of an AI escalation attempt."""

    schema: ReportSchema | None
    used_api: bool
    model: str
    messages: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)


def _sanitize_config(cfg: dict[str, Any]) -> dict[str, Any]:
    """Recursively drop path/PII-ish keys before sending to the API."""
    def _clean(obj: Any) -> Any:
        if isinstance(obj, dict):
            return {
                k: _clean(v)
                for k, v in obj.items()
                if k not in _SENSITIVE_KEYS
            }
        if isinstance(obj, list):
            return [_clean(v) for v in obj]
        return obj

    return _clean(cfg)


def _get_client(api_key: str):
    """Lazily construct an Anthropic client, or return ``None`` if unavailable."""
    try:
        import anthropic  # noqa: PLC0415 - optional dependency
    except ImportError:
        return None
    try:
        return anthropic.Anthropic(api_key=api_key)
    except Exception:  # noqa: BLE001 - never surface SDK init errors upward
        return None


def _extract_json(text: str) -> dict[str, Any] | None:
    """Parse the first JSON object out of the model's response."""
    text = (text or "").strip()
    if not text:
        return None
    # Strip accidental markdown fences.
    fenced = re.search(r"```(?:json)?\s*(\{.*?\})\s*```", text, re.DOTALL)
    candidate = fenced.group(1) if fenced else text
    try:
        return json.loads(candidate)
    except json.JSONDecodeError:
        # Fall back to the outermost braces.
        start, end = candidate.find("{"), candidate.rfind("}")
        if 0 <= start < end:
            try:
                return json.loads(candidate[start : end + 1])
            except json.JSONDecodeError:
                return None
    return None


def escalate(
    cfg: dict[str, Any],
    report_type: str,
    validation_errors: list[str],
    *,
    api_key: str | None,
    model: str = DEFAULT_MODEL,
) -> EscalationResult:
    """Ask Claude to propose a valid schema. Degrades gracefully.

    Returns an :class:`EscalationResult`; ``schema`` is ``None`` (with an
    explanatory message) whenever the SDK, key, network, or model output is
    unusable — the router treats that as "proceed as today".
    """
    if report_type not in KNOWN_REPORT_TYPES:
        return EscalationResult(
            schema=None, used_api=False, model=model,
            errors=[f"unknown report_type '{report_type}'"],
        )

    if not api_key:
        return EscalationResult(
            schema=None, used_api=False, model=model,
            messages=["AI escalation skipped: no Anthropic API key configured "
                      "(set it via cecl_credentials). Proceeding local-only."],
        )

    client = _get_client(api_key)
    if client is None:
        return EscalationResult(
            schema=None, used_api=False, model=model,
            messages=["AI escalation skipped: anthropic SDK unavailable. "
                      "Run 'pip install anthropic'. Proceeding local-only."],
        )

    user_payload = json.dumps(
        {
            "report_type": report_type,
            "validation_errors": validation_errors,
            "config": _sanitize_config(cfg),
        },
        default=str,
    )

    try:
        resp = client.messages.create(
            model=model,
            max_tokens=_MAX_TOKENS,
            system=_SYSTEM_PROMPT,
            messages=[{"role": "user", "content": user_payload}],
        )
    except Exception as exc:  # noqa: BLE001 - network/auth/rate-limit, etc.
        return EscalationResult(
            schema=None, used_api=True, model=model,
            errors=[f"Anthropic API call failed: {exc}"],
            messages=["AI escalation failed; proceeding local-only."],
        )

    text = "".join(
        getattr(block, "text", "") for block in getattr(resp, "content", [])
    )
    data = _extract_json(text)
    if data is None:
        return EscalationResult(
            schema=None, used_api=True, model=model,
            errors=["AI response was not valid JSON."],
        )

    # The model must not override the report type we asked about.
    data["report_type"] = report_type
    data.setdefault("credit_union", cfg.get("credit_union", ""))
    data["source"] = "ai"

    try:
        schema = ReportSchema.model_validate(data)
    except ValidationError as exc:
        errs = [
            f"{'.'.join(str(p) for p in e['loc'])}: {e['msg']}"
            for e in exc.errors()
        ]
        return EscalationResult(
            schema=None, used_api=True, model=model,
            errors=["AI-proposed schema failed validation: " + "; ".join(errs)],
        )

    return EscalationResult(
        schema=schema, used_api=True, model=model,
        messages=[f"AI ({model}) proposed a validated schema."],
    )
