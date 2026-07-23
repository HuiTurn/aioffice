"""Render contracts that distinguish previews from native layout evidence."""

from __future__ import annotations

import hashlib
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field

from aioffice.core.diagnostics import Diagnostic


class RenderOptions(BaseModel):
    model_config = ConfigDict(extra="forbid", allow_inf_nan=False)

    page_view: bool = True
    include_document_metadata: bool = True
    locale: str | None = None
    font_environment_hash: str | None = None
    dpi: int = Field(default=144, ge=72, le=600)
    page_number: int | None = Field(default=None, ge=1)
    timeout_seconds: float = Field(default=60.0, gt=0, le=300)


class RenderResult(BaseModel):
    model_config = ConfigDict(extra="forbid", frozen=True)

    format: Literal["html", "png", "pdf", "svg"]
    media_type: str
    provider: str
    provider_version: str
    fidelity: Literal["approximate", "native"]
    verification_status: Literal["preview_only", "unverified", "verified"]
    content: bytes = Field(repr=False, exclude=True)
    content_size: int = Field(ge=0)
    content_sha256: str
    cache_key: str
    diagnostics: list[Diagnostic] = Field(default_factory=list)
    metadata: dict[str, Any] = Field(default_factory=dict)

    def write(self, target: str | Path) -> Path:
        path = Path(target)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(self.content)
        return path

    def summary(self) -> dict[str, Any]:
        return self.model_dump(mode="json")

    @classmethod
    def create(
        cls,
        *,
        format: Literal["html", "png", "pdf", "svg"],
        media_type: str,
        provider: str,
        provider_version: str,
        fidelity: Literal["approximate", "native"],
        verification_status: Literal["preview_only", "unverified", "verified"],
        content: bytes,
        cache_material: bytes,
        diagnostics: list[Diagnostic] | None = None,
        metadata: dict[str, Any] | None = None,
    ) -> "RenderResult":
        return cls(
            format=format,
            media_type=media_type,
            provider=provider,
            provider_version=provider_version,
            fidelity=fidelity,
            verification_status=verification_status,
            content=content,
            content_size=len(content),
            content_sha256=hashlib.sha256(content).hexdigest(),
            cache_key=hashlib.sha256(cache_material).hexdigest(),
            diagnostics=diagnostics or [],
            metadata=metadata or {},
        )


class VisualComparison(BaseModel):
    model_config = ConfigDict(extra="forbid", allow_inf_nan=False)

    passed: bool
    baseline_size: tuple[int, int]
    candidate_size: tuple[int, int]
    mean_absolute_error: float = Field(ge=0, le=1)
    changed_pixel_ratio: float = Field(ge=0, le=1)
    max_mean_absolute_error: float = Field(ge=0, le=1)
    max_changed_pixel_ratio: float = Field(ge=0, le=1)
    pixel_tolerance: int = Field(ge=0, le=255)
    diagnostics: list[Diagnostic] = Field(default_factory=list)


__all__ = ["RenderOptions", "RenderResult", "VisualComparison"]
