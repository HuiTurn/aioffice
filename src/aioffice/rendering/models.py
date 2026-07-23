"""Render contracts that distinguish previews from native layout evidence."""

from __future__ import annotations

import hashlib
import os
import tempfile
from pathlib import Path
from typing import Any, Literal

from pydantic import BaseModel, ConfigDict, Field

from aioffice.core.diagnostics import Diagnostic
from aioffice.core.errors import RenderingError


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


class PageVisualAnalysis(BaseModel):
    model_config = ConfigDict(extra="forbid", frozen=True, allow_inf_nan=False)

    page_number: int = Field(ge=1)
    width_pixels: int = Field(ge=1)
    height_pixels: int = Field(ge=1)
    background_color: str = Field(pattern=r"^#[0-9A-F]{6}$")
    ink_pixel_ratio: float = Field(ge=0, le=1)
    appears_blank: bool
    content_bbox: tuple[int, int, int, int] | None = None
    whitespace_top_ratio: float = Field(ge=0, le=1)
    whitespace_right_ratio: float = Field(ge=0, le=1)
    whitespace_bottom_ratio: float = Field(ge=0, le=1)
    whitespace_left_ratio: float = Field(ge=0, le=1)
    edge_contact: list[Literal["top", "right", "bottom", "left"]] = Field(
        default_factory=list
    )
    diagnostics: list[Diagnostic] = Field(default_factory=list)


class RenderedPage(BaseModel):
    model_config = ConfigDict(extra="forbid", frozen=True)

    page_number: int = Field(ge=1)
    media_type: Literal["image/png"] = "image/png"
    provider: str
    provider_version: str
    dpi: int = Field(ge=72, le=600)
    width_pixels: int = Field(ge=1)
    height_pixels: int = Field(ge=1)
    content: bytes = Field(repr=False, exclude=True)
    content_size: int = Field(ge=0)
    content_sha256: str
    cache_key: str
    analysis: PageVisualAnalysis | None = None
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
        page_number: int,
        provider: str,
        provider_version: str,
        dpi: int,
        width_pixels: int,
        height_pixels: int,
        content: bytes,
        cache_material: bytes,
        analysis: PageVisualAnalysis | None = None,
        diagnostics: list[Diagnostic] | None = None,
        metadata: dict[str, Any] | None = None,
    ) -> "RenderedPage":
        return cls(
            page_number=page_number,
            provider=provider,
            provider_version=provider_version,
            dpi=dpi,
            width_pixels=width_pixels,
            height_pixels=height_pixels,
            content=content,
            content_size=len(content),
            content_sha256=hashlib.sha256(content).hexdigest(),
            cache_key=hashlib.sha256(cache_material).hexdigest(),
            analysis=analysis,
            diagnostics=diagnostics or [],
            metadata=metadata or {},
        )


class PaginatedRenderResult(BaseModel):
    model_config = ConfigDict(extra="forbid", frozen=True)

    provider: str
    provider_version: str
    fidelity: Literal["native"] = "native"
    verification_status: Literal["unverified", "verified"]
    page_count: int = Field(ge=1)
    pdf: RenderResult
    pages: list[RenderedPage] = Field(min_length=1)
    cache_key: str
    diagnostics: list[Diagnostic] = Field(default_factory=list)
    metadata: dict[str, Any] = Field(default_factory=dict)

    def summary(self) -> dict[str, Any]:
        return self.model_dump(mode="json")

    def write(
        self,
        directory: str | Path,
        *,
        stem: str = "document",
        overwrite: bool = False,
    ) -> dict[str, Path | list[Path]]:
        if (
            not stem
            or stem in {".", ".."}
            or Path(stem).name != stem
        ):
            raise RenderingError(
                "Render output stem must be one filename component."
            )
        root = Path(directory)
        root.mkdir(parents=True, exist_ok=True)
        digits = max(4, len(str(self.page_count)))
        pdf_target = root / f"{stem}.pdf"
        page_targets = [
            root / f"{stem}-page-{page.page_number:0{digits}d}.png"
            for page in self.pages
        ]
        targets = [pdf_target, *page_targets]
        existing = [path for path in targets if path.exists()]
        if existing and not overwrite:
            names = ", ".join(str(path) for path in existing)
            raise RenderingError(
                f"Refusing to overwrite existing render evidence: {names}"
            )

        with tempfile.TemporaryDirectory(
            prefix=".aioffice-render-write-",
            dir=root,
        ) as staging_directory:
            staging = Path(staging_directory)
            staged_pdf = staging / pdf_target.name
            staged_pdf.write_bytes(self.pdf.content)
            staged_pages: list[Path] = []
            for page, target in zip(self.pages, page_targets, strict=True):
                staged_page = staging / target.name
                staged_page.write_bytes(page.content)
                staged_pages.append(staged_page)
            staged_targets = [
                (staged_pdf, pdf_target),
                *zip(staged_pages, page_targets, strict=True),
            ]
            if overwrite:
                for staged_path, target in staged_targets:
                    os.replace(staged_path, target)
            else:
                created: list[Path] = []
                try:
                    for staged_path, target in staged_targets:
                        with target.open("xb") as output:
                            created.append(target)
                            output.write(staged_path.read_bytes())
                except BaseException:
                    for path in created:
                        path.unlink(missing_ok=True)
                    raise
        return {"pdf": pdf_target, "pages": page_targets}


__all__ = [
    "PageVisualAnalysis",
    "PaginatedRenderResult",
    "RenderedPage",
    "RenderOptions",
    "RenderResult",
    "VisualComparison",
]
