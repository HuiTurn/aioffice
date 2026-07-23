"""Strict persisted models for the local AiOffice workspace."""

from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field, model_validator

from aioffice.spec.models import NodeId


class WorkspaceModel(BaseModel):
    model_config = ConfigDict(extra="forbid")


class ArtifactEntry(WorkspaceModel):
    artifact_id: NodeId
    kind: Literal["document"] = "document"
    format: Literal["docx"] = "docx"
    first_revision: int = Field(default=1, ge=1)
    latest_revision: int = Field(ge=1)
    source_name: str


class WorkspaceIndex(WorkspaceModel):
    workspace_version: Literal["0.1"] = "0.1"
    workspace_id: NodeId
    artifacts: dict[str, ArtifactEntry] = Field(default_factory=dict)


class PatchRecord(WorkspaceModel):
    base_revision: int = Field(ge=1)
    result_revision: int = Field(ge=2)
    idempotency_key: str | None = None
    operations: list[dict[str, Any]]
    changes: list[dict[str, Any]] = Field(default_factory=list)
    diagnostics: list[dict[str, Any]] = Field(default_factory=list)
    fidelity: dict[str, Any] | None = None
    diff: dict[str, Any] | None = None

    @model_validator(mode="after")
    def validate_revision_step(self) -> "PatchRecord":
        if self.result_revision != self.base_revision + 1:
            raise ValueError("A patch record must advance exactly one revision.")
        return self


__all__ = ["ArtifactEntry", "PatchRecord", "WorkspaceIndex"]
