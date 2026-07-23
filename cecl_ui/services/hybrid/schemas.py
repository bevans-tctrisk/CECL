"""Typed report-schema contract (Pydantic v2).

These models mirror the report-relevant slice of a ``client_configs/<cu>.yaml``
file. They are the canonical definition of a "report schema" used by the
hybrid router for local validation, AI escalation, and registry storage.

Only the fields that materially affect a generated CECL report are modeled
here — pools, credit grades, column mappings, and the report-format flags.
Everything else in the YAML (folder paths, file patterns, economic data, ...)
is intentionally ignored so schema drift in unrelated areas does not trigger
spurious escalations.
"""
from __future__ import annotations

import hashlib
import json
from typing import Any

from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    field_validator,
    model_validator,
)

# Report types the wizard knows how to emit. Kept in sync with
# cecl_ui/routes/run.py::reports() and generate_report.generate_report().
KNOWN_REPORT_TYPES: tuple[str, ...] = (
    "tct",
    "vizo",
    "vizo_supp",
    "mgmt_adj_napkin",
    "impdet",
)

# Column mappings that a report genuinely cannot be produced without.
REQUIRED_COLUMN_KEYS: tuple[str, ...] = ("member_number", "current_balance")

# Provenance of a stored schema.
SCHEMA_SOURCES: tuple[str, ...] = ("manual", "local", "ai")


class CreditGrade(BaseModel):
    """A single FICO band with its CECL reserve rate."""

    model_config = ConfigDict(extra="ignore")

    label: str
    min_score: int = Field(ge=0, le=1000)
    max_score: int = Field(ge=0, le=1000)
    reserve_rate: float = Field(ge=0.0, le=1.0)

    @model_validator(mode="after")
    def _check_range(self) -> "CreditGrade":
        if self.min_score > self.max_score:
            raise ValueError(
                f"credit grade '{self.label}': min_score "
                f"({self.min_score}) > max_score ({self.max_score})"
            )
        return self


class Pool(BaseModel):
    """A loan pool and its risk/ACL characteristics."""

    model_config = ConfigDict(extra="ignore")

    name: str
    risk_rated: bool = True
    acl_months: int | None = Field(default=None, ge=0)
    excluded: bool = False
    use_default_mgmt_adj: bool = False

    @field_validator("name")
    @classmethod
    def _name_not_blank(cls, v: str) -> str:
        if not (v or "").strip():
            raise ValueError("pool name must not be blank")
        return v


class ReportSchema(BaseModel):
    """The validated, report-relevant schema for one CU + report type."""

    model_config = ConfigDict(extra="ignore")

    credit_union: str
    report_type: str
    pools: list[Pool] = Field(default_factory=list)
    credit_grades: list[CreditGrade] = Field(default_factory=list)
    column_mappings: dict[str, str] = Field(default_factory=dict)
    reports: dict[str, bool] = Field(default_factory=dict)
    default_pool: str | None = None
    source: str = "local"

    @field_validator("report_type")
    @classmethod
    def _known_report_type(cls, v: str) -> str:
        if v not in KNOWN_REPORT_TYPES:
            raise ValueError(
                f"unknown report_type '{v}'; expected one of "
                f"{', '.join(KNOWN_REPORT_TYPES)}"
            )
        return v

    @field_validator("source")
    @classmethod
    def _known_source(cls, v: str) -> str:
        if v not in SCHEMA_SOURCES:
            raise ValueError(
                f"unknown source '{v}'; expected one of "
                f"{', '.join(SCHEMA_SOURCES)}"
            )
        return v

    @model_validator(mode="after")
    def _check_grades_non_overlapping(self) -> "ReportSchema":
        # Grades must not overlap so a FICO score maps to exactly one band.
        bands = sorted(self.credit_grades, key=lambda g: g.min_score)
        for prev, cur in zip(bands, bands[1:]):
            if cur.min_score <= prev.max_score:
                raise ValueError(
                    f"credit grades overlap: '{prev.label}' "
                    f"[{prev.min_score}-{prev.max_score}] and '{cur.label}' "
                    f"[{cur.min_score}-{cur.max_score}]"
                )
        return self

    # ------------------------------------------------------------------
    # Construction / serialization helpers
    # ------------------------------------------------------------------
    @classmethod
    def from_config(
        cls, cfg: dict[str, Any], report_type: str, *, source: str = "local"
    ) -> "ReportSchema":
        """Build a schema from a raw client-config dict.

        Raises ``pydantic.ValidationError`` when the config's report-relevant
        slice is malformed — that error is what triggers AI escalation.
        """
        return cls(
            credit_union=cfg.get("credit_union", ""),
            report_type=report_type,
            pools=cfg.get("pools") or [],
            credit_grades=cfg.get("credit_grades") or [],
            column_mappings=cfg.get("column_mappings") or {},
            reports=cfg.get("reports") or {},
            default_pool=cfg.get("default_pool"),
            source=source,
        )

    def canonical_json(self) -> str:
        """Stable JSON used for storage and change detection."""
        return json.dumps(
            self.model_dump(mode="json"), sort_keys=True, separators=(",", ":")
        )

    def checksum(self) -> str:
        """SHA-256 of the canonical JSON — cheap 'has this changed?' key."""
        return hashlib.sha256(self.canonical_json().encode("utf-8")).hexdigest()
