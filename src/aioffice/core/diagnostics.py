"""Machine-readable diagnostics shared by validation, patches, and exporters."""

from __future__ import annotations

from enum import StrEnum
from typing import Any

from pydantic import BaseModel, ConfigDict, Field


class Severity(StrEnum):
    INFO = "info"
    WARNING = "warning"
    ERROR = "error"
    FATAL = "fatal"


class Diagnostic(BaseModel):
    model_config = ConfigDict(extra="forbid")

    severity: Severity
    code: str
    message: str
    node_ids: list[str] = Field(default_factory=list)
    path: str | None = None
    recoverable: bool = True
    suggested_actions: list[dict[str, Any]] = Field(default_factory=list)


class ValidationResult(BaseModel):
    model_config = ConfigDict(extra="forbid")

    diagnostics: list[Diagnostic] = Field(default_factory=list)

    @property
    def valid(self) -> bool:
        return not any(
            item.severity in (Severity.ERROR, Severity.FATAL) for item in self.diagnostics
        )

    @property
    def errors(self) -> list[Diagnostic]:
        return [
            item
            for item in self.diagnostics
            if item.severity in (Severity.ERROR, Severity.FATAL)
        ]

    @property
    def warnings(self) -> list[Diagnostic]:
        return [item for item in self.diagnostics if item.severity == Severity.WARNING]
