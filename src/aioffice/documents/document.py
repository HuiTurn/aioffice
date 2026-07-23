"""Document artifact, validation, export, inspection, and atomic patching."""

from __future__ import annotations

import hashlib
import json
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Literal, Mapping, Sequence, cast

from pydantic import TypeAdapter, ValidationError

from aioffice.core.diagnostics import Diagnostic, Severity, ValidationResult
from aioffice.core.diff import DocumentDiff, compute_document_diff
from aioffice.core.errors import (
    ExportError,
    NativePackageError,
    SecurityError,
    SpecValidationError,
    UnsupportedFormatError,
)
from aioffice.formats.docx import compile_docx, export_docx
from aioffice.formats.docx_images import simple_inline_image_from_ref
from aioffice.formats.docx_import import import_docx
from aioffice.formats.docx_native import apply_docx_operations
from aioffice.formats.html import export_html
from aioffice.formats.markdown import export_markdown, import_markdown
from aioffice.native import (
    FidelityPolicy,
    FidelityReport,
    IdentityManifest,
    MANIFEST_PART_URI,
    NativePackage,
    build_identity_manifest,
    serialize_identity_manifest,
)
from aioffice.operations.text import (
    format_entire_text,
    format_text_range,
    node_plain_text,
    replace_node_text,
    resolve_text_selection,
)
from aioffice.rendering import (
    LIBREOFFICE_PROVIDER,
    PaginatedRenderResult,
    RenderOptions,
    RenderResult,
    libreoffice_render_capabilities,
    render_docx_libreoffice,
    render_docx_pages_libreoffice,
    render_semantic_html,
)
from aioffice.security import SecurityPolicy
from aioffice.spec.models import (
    AiOfficeDocumentSpec,
    AssetRef,
    Block,
    DOCUMENT_SCHEMA_URL,
    BulletList,
    DocumentField,
    Heading,
    ImageBlock,
    ImageInsert,
    ImageUpdate,
    LEGACY_DOCUMENT_SCHEMA_URL,
    LEGACY_SPEC_VERSION,
    Length,
    NamedStyle,
    NativeRef,
    OpaqueBlock,
    OrderedList,
    Paragraph,
    ParagraphStyle,
    SectionLayout,
    SPEC_VERSION,
    Table,
    TableCellFormat,
    TableColumn,
    TableLayout,
    TextStyle,
)
from aioffice.styles import style_catalog, theme_named_styles
from aioffice.themes import get_theme

from .assets import ImageAsset, prepare_image_asset


def _document_fields(spec: AiOfficeDocumentSpec) -> list[DocumentField]:
    blocks = [
        *(
            node
            for node in spec.content
            if isinstance(node, (Heading, Paragraph))
        ),
        *(
            block
            for part in spec.header_footers
            for block in part.content
            if isinstance(block, Paragraph)
        ),
    ]
    return [
        inline
        for block in blocks
        for inline in block.content
        if isinstance(inline, DocumentField)
    ]


def _validation_error_diagnostics(error: ValidationError) -> list[Diagnostic]:
    diagnostics: list[Diagnostic] = []
    for item in error.errors(include_url=False):
        path = ".".join(str(part) for part in item["loc"])
        diagnostics.append(
            Diagnostic(
                severity=Severity.ERROR,
                code="INVALID_SPEC",
                message=item["msg"],
                path=path or None,
                recoverable=True,
                suggested_actions=[{"action": "fix_field", "path": path}],
            )
        )
    return diagnostics


def _parse_spec(value: AiOfficeDocumentSpec | Mapping[str, Any]) -> AiOfficeDocumentSpec:
    if isinstance(value, AiOfficeDocumentSpec):
        return value.model_copy(deep=True)
    payload = deepcopy(dict(value))
    if payload.get("spec_version") == LEGACY_SPEC_VERSION:
        payload["spec_version"] = SPEC_VERSION
        if payload.get("$schema") in {None, LEGACY_DOCUMENT_SCHEMA_URL}:
            payload["$schema"] = DOCUMENT_SCHEMA_URL
    try:
        return AiOfficeDocumentSpec.model_validate(payload)
    except ValidationError as error:
        diagnostics = _validation_error_diagnostics(error)
        summary = "; ".join(f"{item.path or '<root>'}: {item.message}" for item in diagnostics[:3])
        raise SpecValidationError(
            f"Invalid AiOffice Document Spec: {summary}", diagnostics
        ) from error


@dataclass(slots=True)
class PatchResult:
    success: bool
    base_revision: int
    result_revision: int
    dry_run: bool
    document: "Document | None" = None
    changes: list[dict[str, Any]] = field(default_factory=list)
    diagnostics: list[Diagnostic] = field(default_factory=list)
    idempotency_key: str | None = None
    fidelity: FidelityReport | None = None
    diff: DocumentDiff | None = None

    def model_dump(self) -> dict[str, Any]:
        return {
            "success": self.success,
            "base_revision": self.base_revision,
            "result_revision": self.result_revision,
            "dry_run": self.dry_run,
            "changes": deepcopy(self.changes),
            "diagnostics": [item.model_dump(mode="json") for item in self.diagnostics],
            "idempotency_key": self.idempotency_key,
            "fidelity": (
                self.fidelity.model_dump(mode="json") if self.fidelity is not None else None
            ),
            "diff": self.diff.model_dump(mode="json") if self.diff is not None else None,
        }


class _PatchFailure(Exception):
    def __init__(self, diagnostic: Diagnostic) -> None:
        super().__init__(diagnostic.message)
        self.diagnostic = diagnostic


class Document:
    """A validated, logically immutable Document artifact."""

    def __init__(
        self,
        spec: AiOfficeDocumentSpec,
        *,
        native: NativePackage | None = None,
        import_diagnostics: Sequence[Diagnostic] = (),
    ) -> None:
        self._spec = spec.model_copy(deep=True)
        self._native = native.clone() if native is not None else None
        self._import_diagnostics = [
            diagnostic.model_copy(deep=True) for diagnostic in import_diagnostics
        ]

    @classmethod
    def from_spec(cls, spec: AiOfficeDocumentSpec | Mapping[str, Any]) -> "Document":
        return cls(_parse_spec(spec))

    @classmethod
    def from_json(cls, source: str | bytes | Path) -> "Document":
        if isinstance(source, Path):
            raw = source.read_text(encoding="utf-8")
        elif isinstance(source, bytes):
            raw = source.decode("utf-8")
        elif source.lstrip().startswith(("{", "[")):
            raw = source
        else:
            raw = Path(source).read_text(encoding="utf-8")
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError as error:
            diagnostic = Diagnostic(
                severity=Severity.ERROR,
                code="INVALID_SPEC",
                message=f"Invalid JSON at line {error.lineno}, column {error.colno}: {error.msg}",
                path=f"line:{error.lineno}:column:{error.colno}",
            )
            raise SpecValidationError(diagnostic.message, [diagnostic]) from error
        if not isinstance(payload, dict):
            raise SpecValidationError("The root of an AiOffice Document Spec must be an object.")
        return cls.from_spec(payload)

    @classmethod
    def from_markdown(cls, source: str | Path, *, title: str | None = None) -> "Document":
        if isinstance(source, Path):
            text = source.read_text(encoding="utf-8")
        elif "\n" not in source and Path(source).suffix.lower() in {".md", ".markdown"}:
            text = Path(source).read_text(encoding="utf-8")
        else:
            text = source
        return cls(import_markdown(text, title=title))

    @classmethod
    def from_docx(
        cls,
        source: str | Path | bytes,
        *,
        roundtrip: FidelityPolicy | str = FidelityPolicy.PRESERVE_UNKNOWN,
        security_policy: SecurityPolicy | None = None,
        identity_manifest: IdentityManifest | None = None,
    ) -> "Document":
        imported = import_docx(
            source,
            roundtrip=roundtrip,
            security_policy=security_policy,
            identity_manifest=identity_manifest,
        )
        return cls(
            imported.spec,
            native=imported.native,
            import_diagnostics=imported.diagnostics,
        )

    @property
    def id(self) -> str:
        return self._spec.artifact.id

    @property
    def kind(self) -> str:
        return self._spec.artifact.kind

    @property
    def revision(self) -> int:
        return self._spec.artifact.revision

    @property
    def spec_version(self) -> str:
        return self._spec.spec_version

    @property
    def spec(self) -> AiOfficeDocumentSpec:
        return self._spec.model_copy(deep=True)

    @property
    def origin(self) -> str:
        return "native" if self._native is not None else "semantic"

    @property
    def fidelity(self) -> FidelityReport | None:
        return self._native.fidelity_report() if self._native is not None else None

    @property
    def import_diagnostics(self) -> list[Diagnostic]:
        return [diagnostic.model_copy(deep=True) for diagnostic in self._import_diagnostics]

    def to_spec(self) -> dict[str, Any]:
        return self._spec.model_dump(mode="json", by_alias=True, exclude_none=True)

    def to_json(self, *, indent: int = 2) -> str:
        return json.dumps(self.to_spec(), ensure_ascii=False, indent=indent) + "\n"

    def read_image(self, image_id: str) -> ImageAsset:
        """Return verified bytes for one projected native image occurrence."""

        normalized_id = image_id[1:] if image_id.startswith("#") else image_id
        image = next(
            (
                node
                for node in self._spec.content
                if isinstance(node, ImageBlock)
                and node.id == normalized_id
            ),
            None,
        )
        if image is None:
            raise NativePackageError(
                f"No projected image matched #{normalized_id}."
            )
        if self._native is None:
            raise NativePackageError(
                "Image bytes are available only while the native DOCX package "
                "is attached."
            )
        if not isinstance(image.source_ref, NativeRef):
            raise NativePackageError(
                f"Image {image.id!r} has no trusted native source reference."
            )
        native_image = simple_inline_image_from_ref(
            self._native,
            image.source_ref,
        )
        asset_matches = [
            asset
            for asset in self._spec.assets
            if asset.id == image.asset_id
        ]
        if len(asset_matches) != 1:
            raise NativePackageError(
                f"Image {image.id!r} does not resolve to one asset record."
            )
        asset = asset_matches[0]
        if (
            image.asset_id != native_image.asset_id
            or asset.sha256 != native_image.sha256
            or asset.media_type != native_image.media_type
            or (
                asset.size_bytes is not None
                and asset.size_bytes != native_image.size_bytes
            )
        ):
            raise NativePackageError(
                f"Image {image.id!r} metadata no longer matches the native "
                "package."
            )
        payload = self._native.get_part(native_image.part_uri)
        digest = hashlib.sha256(payload).hexdigest()
        if (
            digest != native_image.sha256
            or len(payload) != native_image.size_bytes
        ):
            raise NativePackageError(
                f"Image {image.id!r} failed binary integrity verification."
            )
        return ImageAsset(
            image_id=image.id,
            asset_id=asset.id,
            media_type=native_image.media_type,
            filename=native_image.filename,
            sha256=digest,
            data=payload,
        )

    def image_bytes(self, image_id: str) -> bytes:
        """Return only the verified bytes for a projected image."""

        return self.read_image(image_id).data

    def extract_image(
        self,
        image_id: str,
        target: str | Path,
        *,
        overwrite: bool = False,
    ) -> Path:
        """Write one verified native image without embedding it in JSON."""

        return self.read_image(image_id).write(
            target,
            overwrite=overwrite,
        )

    def replace_image(
        self,
        image_id: str,
        source: bytes | bytearray | memoryview | str | Path | ImageAsset,
        *,
        media_type: str | None = None,
        dry_run: bool = False,
        base_revision: int | None = None,
        idempotency_key: str | None = None,
    ) -> PatchResult:
        """Replace one projected native image through a bounded binary channel."""

        if self._native is None:
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=[
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="UNSUPPORTED_FEATURE",
                        message=(
                            "Image replacement requires an attached native DOCX "
                            "package."
                        ),
                        node_ids=[self.id],
                        suggested_actions=[
                            {"action": "open_native_docx"},
                            {"action": "inspect_capabilities"},
                        ],
                    )
                ],
                idempotency_key=idempotency_key,
            )
        try:
            prepared = prepare_image_asset(
                source,
                media_type=media_type,
                security_policy=self._native.security_policy,
            )
        except (
            NativePackageError,
            OSError,
            SecurityError,
            TypeError,
        ) as error:
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=[
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_ASSET_INPUT",
                        message=str(error),
                        node_ids=[self.id],
                        suggested_actions=[
                            {
                                "action": "use_supported_raster_image",
                                "media_types": [
                                    "image/png",
                                    "image/jpeg",
                                    "image/gif",
                                    "image/bmp",
                                    "image/tiff",
                                ],
                            }
                        ],
                    )
                ],
                idempotency_key=idempotency_key,
            )
        operation = {
            "op": "image.replace",
            "target": image_id,
            "asset": prepared.asset.model_dump(
                mode="json",
                exclude_none=True,
            ),
        }
        return self._apply(
            [operation],
            dry_run=dry_run,
            base_revision=base_revision,
            idempotency_key=idempotency_key,
            image_payloads={
                prepared.asset.id: prepared.data,
            },
        )

    def insert_image_after(
        self,
        target: str,
        source: bytes | bytearray | memoryview | str | Path | ImageAsset,
        *,
        width: Length | Mapping[str, Any],
        height: Length | Mapping[str, Any],
        alt_text: str,
        media_type: str | None = None,
        image_id: str | None = None,
        name: str | None = None,
        title: str | None = None,
        paragraph_style: ParagraphStyle | Mapping[str, Any] | None = None,
        dry_run: bool = False,
        base_revision: int | None = None,
        idempotency_key: str | None = None,
    ) -> PatchResult:
        """Insert one native inline image after a mapped top-level node."""

        if self._native is None:
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=[
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="UNSUPPORTED_FEATURE",
                        message=(
                            "Image insertion requires an attached native DOCX "
                            "package."
                        ),
                        node_ids=[self.id],
                        suggested_actions=[
                            {"action": "open_native_docx"},
                            {"action": "inspect_capabilities"},
                        ],
                    )
                ],
                idempotency_key=idempotency_key,
            )
        try:
            prepared = prepare_image_asset(
                source,
                media_type=media_type,
                security_policy=self._native.security_policy,
            )
        except (
            NativePackageError,
            OSError,
            SecurityError,
            TypeError,
        ) as error:
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=[
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_ASSET_INPUT",
                        message=str(error),
                        node_ids=[self.id],
                        suggested_actions=[
                            {"action": "use_supported_raster_image"},
                        ],
                    )
                ],
                idempotency_key=idempotency_key,
            )
        image_payload: dict[str, Any] = {
            "width": width,
            "height": height,
            "name": name or prepared.asset.filename,
            "alt_text": alt_text,
            "title": title,
            "paragraph_style": paragraph_style,
        }
        if image_id is not None:
            image_payload["id"] = image_id
        try:
            image_insert = ImageInsert.model_validate(image_payload)
        except ValidationError as error:
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=[
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_IMAGE_INSERT",
                        message="Image insertion metadata is invalid.",
                        node_ids=[self.id],
                        suggested_actions=[
                            {
                                "action": "inspect_image_insert_schema",
                                "diagnostics": [
                                    item.model_dump(mode="json")
                                    for item in _validation_error_diagnostics(
                                        error
                                    )
                                ],
                            }
                        ],
                    )
                ],
                idempotency_key=idempotency_key,
            )
        operation = {
            "op": "image.insert_after",
            "target": target,
            "image": image_insert.model_dump(
                mode="json",
                exclude_none=True,
            ),
            "asset": prepared.asset.model_dump(
                mode="json",
                exclude_none=True,
            ),
        }
        return self._apply(
            [operation],
            dry_run=dry_run,
            base_revision=base_revision,
            idempotency_key=idempotency_key,
            image_payloads={
                prepared.asset.id: prepared.data,
            },
        )

    def inspect(self, *, response_format: str = "compact") -> dict[str, Any]:
        if response_format not in {"compact", "expanded", "summary"}:
            raise ValueError("response_format must be compact, expanded, or summary.")
        counts: dict[str, int] = {}
        for node in self._spec.content:
            counts[node.type] = counts.get(node.type, 0) + 1
        result: dict[str, Any] = {
            "artifact_id": self.id,
            "kind": self.kind,
            "revision": self.revision,
            "spec_version": self.spec_version,
            "title": self._spec.metadata.title,
            "origin": self.origin,
            "node_count": len(self._spec.content),
            "node_types": counts,
            "diagnostic_count": len(self._import_diagnostics),
            "style_count": len(style_catalog(self._spec)),
            "section_count": len(self._spec.sections),
            "header_footer_count": len(self._spec.header_footers),
            "field_count": len(_document_fields(self._spec)),
            "image_count": sum(
                isinstance(node, ImageBlock)
                for node in self._spec.content
            ),
            "asset_count": len(self._spec.assets),
        }
        if response_format == "compact":
            style_usage: dict[str, int] = {}
            nodes: list[dict[str, Any]] = []
            for node in self._spec.content:
                compact: dict[str, Any] = {
                    "id": node.id,
                    "type": node.type,
                }
                if isinstance(node, Heading):
                    compact.update(text=node.plain_text, level=node.level)
                elif isinstance(node, Paragraph):
                    compact["text"] = node.plain_text
                elif isinstance(node, (BulletList, OrderedList)):
                    compact.update(item_count=len(node.items), items=node.items[:3])
                elif isinstance(node, Table):
                    compact.update(
                        column_count=len(node.columns),
                        row_count=len(node.rows),
                        regular_grid=node.metadata.get(
                            "regular_grid",
                            all(
                                cell.column_span == 1
                                and cell.row_span == 1
                                for row in node.rows
                                for cell in row.cells
                            ),
                        ),
                        logical_grid=node.metadata.get(
                            "logical_grid",
                            True,
                        ),
                        layout=node.layout.model_dump(
                            mode="json",
                            exclude_none=True,
                        ),
                        columns=[
                            {
                                "id": column.id,
                                "key": column.key,
                                "title": column.title,
                                "data_type": column.data_type,
                                "width": (
                                    column.width.model_dump(mode="json")
                                    if column.width is not None
                                    else None
                                ),
                            }
                            for column in node.columns
                        ],
                        rows=[
                            {
                                "id": row.id,
                                "cells": [
                                    {
                                        "id": cell.id,
                                        "column_key": cell.column_key,
                                        "text": cell.plain_text,
                                        "column_span": cell.column_span,
                                        "row_span": cell.row_span,
                                        "format": cell.format.model_dump(
                                            mode="json",
                                            exclude_none=True,
                                        ),
                                        "content_ids": [
                                            paragraph.id
                                            for paragraph in cell.content
                                        ],
                                    }
                                    for cell in row.cells
                                ],
                                "allow_break_across_pages": (
                                    row.allow_break_across_pages
                                ),
                                "height": (
                                    row.height.model_dump(mode="json")
                                    if row.height is not None
                                    else None
                                ),
                                "height_rule": row.height_rule,
                            }
                            for row in node.rows[:3]
                        ],
                    )
                elif isinstance(node, ImageBlock):
                    asset = next(
                        (
                            candidate
                            for candidate in self._spec.assets
                            if candidate.id == node.asset_id
                        ),
                        None,
                    )
                    compact.update(
                        asset_id=node.asset_id,
                        placement=node.placement,
                        width=node.width.model_dump(mode="json"),
                        height=node.height.model_dump(mode="json"),
                        name=node.name,
                        alt_text=node.alt_text,
                        title=node.title,
                        capabilities=node.capabilities,
                        supported_operations=(
                            [
                                "image.insert_after",
                                "image.replace",
                                "image.update",
                                "paragraph.format",
                                "node.remove",
                            ]
                            if self._native is not None
                            else []
                        ),
                        editable=node.editable,
                        asset=(
                            asset.model_dump(
                                mode="json",
                                exclude_none=True,
                            )
                            if asset is not None
                            else None
                        ),
                    )
                elif isinstance(node, OpaqueBlock):
                    compact.update(
                        summary=node.summary,
                        capabilities=node.capabilities,
                        editable=node.editable,
                    )
                if isinstance(node, (Heading, Paragraph)) and node.style_ref is not None:
                    compact["style_ref"] = node.style_ref
                    style_usage[node.style_ref] = style_usage.get(node.style_ref, 0) + 1
                if isinstance(node, (Heading, Paragraph)):
                    compact["fields"] = [
                        {
                            "id": inline.id,
                            "kind": inline.kind,
                            "number_format": inline.number_format,
                            "cached_result": inline.cached_result,
                            "editable": inline.editable,
                        }
                        for inline in node.content
                        if isinstance(inline, DocumentField)
                    ]
                nodes.append(compact)
            result["nodes"] = nodes
            result["styles"] = [
                {
                    "id": style.id,
                    "name": style.name,
                    "semantic_role": style.semantic_role,
                    "heading_level": style.heading_level,
                    "based_on": style.based_on,
                    "usage_count": style_usage.get(style.id, 0),
                }
                for style in style_catalog(self._spec).values()
                if not style.hidden
            ]
            result["sections"] = [
                {
                    "id": section.id,
                    "start_at": section.start_at,
                    "start_type": section.layout.start_type,
                    "page_size": (
                        section.layout.page_size.model_dump(
                            mode="json",
                            exclude_none=True,
                        )
                        if section.layout.page_size is not None
                        else None
                    ),
                    "column_count": (
                        section.layout.columns.count
                        if section.layout.columns is not None
                        else None
                    ),
                    "header_footer": (
                        section.header_footer.model_dump(
                            mode="json",
                            exclude_none=True,
                        )
                        if section.header_footer is not None
                        else {}
                    ),
                }
                for section in self._spec.sections
            ]
            result["header_footers"] = [
                {
                    "id": part.id,
                    "kind": part.kind,
                    "block_count": len(part.content),
                    "blocks": [
                        {
                            "id": block.id,
                            "type": block.type,
                            **(
                                {
                                    "text": block.plain_text,
                                    "fields": [
                                        {
                                            "id": inline.id,
                                            "kind": inline.kind,
                                            "number_format": inline.number_format,
                                            "cached_result": inline.cached_result,
                                            "editable": inline.editable,
                                        }
                                        for inline in block.content
                                        if isinstance(inline, DocumentField)
                                    ],
                                }
                                if isinstance(block, Paragraph)
                                else {"summary": block.summary}
                            ),
                        }
                        for block in part.content
                    ],
                    "projection_complete": part.metadata.get(
                        "projection_complete",
                        True,
                    ),
                }
                for part in self._spec.header_footers
            ]
        elif response_format == "expanded":
            result["nodes"] = [
                node.model_dump(mode="json", exclude_none=True) for node in self._spec.content
            ]
            result["sections"] = [
                section.model_dump(mode="json", exclude_none=True)
                for section in self._spec.sections
            ]
            result["header_footers"] = [
                part.model_dump(mode="json", exclude_none=True)
                for part in self._spec.header_footers
            ]
            result["assets"] = [
                asset.model_dump(mode="json", exclude_none=True)
                for asset in self._spec.assets
            ]
        return result

    def diff(
        self,
        other: "Document",
        *,
        include_native: bool = False,
    ) -> DocumentDiff:
        """Return a stable semantic diff keyed by persistent node identities."""

        return compute_document_diff(
            self._spec,
            other._spec,
            include_native=include_native,
        )

    def render(
        self,
        *,
        format: str = "html",
        provider: str = "semantic-html",
        options: RenderOptions | Mapping[str, Any] | None = None,
    ) -> RenderResult:
        """Render through an explicit provider with declared layout fidelity."""

        normalized_format = format.lower().lstrip(".")
        active_options = (
            options
            if isinstance(options, RenderOptions)
            else RenderOptions.model_validate(options or {})
        )
        if provider == "semantic-html":
            if normalized_format != "html":
                raise UnsupportedFormatError(
                    "provider='semantic-html' supports only format='html'."
                )
            return render_semantic_html(self._spec, active_options)
        if provider == LIBREOFFICE_PROVIDER:
            if normalized_format not in {"pdf", "png"}:
                raise UnsupportedFormatError(
                    "provider='libreoffice' supports format='pdf' or format='png'."
                )
            return render_docx_libreoffice(
                self.to_bytes("docx"),
                format=cast(Literal["pdf", "png"], normalized_format),
                options=active_options,
            )
        raise UnsupportedFormatError(
            f"Unknown render provider {provider!r}; use 'semantic-html' or "
            f"{LIBREOFFICE_PROVIDER!r}."
        )

    def render_pages(
        self,
        *,
        provider: str = LIBREOFFICE_PROVIDER,
        page_numbers: Sequence[int] | None = None,
        options: RenderOptions | Mapping[str, Any] | None = None,
        analyze: bool = False,
        max_pages: int = 100,
    ) -> PaginatedRenderResult:
        """Render one native PDF and a bounded set of consistent page images."""

        if provider != LIBREOFFICE_PROVIDER:
            raise UnsupportedFormatError(
                "Paginated rendering currently requires provider='libreoffice'."
            )
        active_options = (
            options
            if isinstance(options, RenderOptions)
            else RenderOptions.model_validate(options or {})
        )
        return render_docx_pages_libreoffice(
            self.to_bytes("docx"),
            page_numbers=page_numbers,
            options=active_options,
            analyze=analyze,
            max_pages=max_pages,
        )

    def capabilities(self) -> dict[str, Any]:
        native_render = libreoffice_render_capabilities()
        native_extension = self._spec.extensions.get(
            "dev.aioffice.native",
            {},
        )
        detached_native_projection = (
            self._native is None
            and native_extension.get("authority") == "native"
        )
        operations = [
            "text.replace",
            "paragraph.format",
            "text.format",
            "node.append",
            "node.insert_after",
            "node.move_after",
            "node.move_before",
            "node.remove",
            "node.update",
            "style.apply",
            "style.define",
            "style.format",
            "section.format",
            "field.update",
            "table.format",
            "table.column.format",
            "table.cell.format",
        ]
        if self._native is not None:
            operations = [
                "text.replace",
                "paragraph.format",
                "text.format",
                "node.move_after",
                "node.move_before",
                "node.remove",
                "style.apply",
                "style.define",
                "style.format",
                "section.format",
                "field.update",
                "image.insert_after",
                "table.format",
                "table.column.format",
                "table.cell.format",
            ]
            if any(
                isinstance(node, ImageBlock)
                for node in self._spec.content
            ):
                operations.append("image.replace")
                operations.append("image.update")
        elif detached_native_projection:
            operations.remove("node.move_after")
            operations.remove("node.move_before")
        ambiguous_node_ids = sorted(
            {
                node_id
                for diagnostic in self._import_diagnostics
                if diagnostic.code == "IDENTITY_AMBIGUOUS"
                for node_id in diagnostic.node_ids
            }
        )
        return {
            "artifact_id": self.id,
            "kind": self.kind,
            "origin": self.origin,
            "spec_version": self.spec_version,
            "import_formats": ["json", "markdown", "docx"],
            "export_formats": ["json", "markdown", "html", "docx"],
            "render": {
                "providers": [
                    {
                        "name": "semantic-html",
                        "formats": ["html"],
                        "fidelity": "approximate",
                        "verification_status": "preview_only",
                    },
                    native_render,
                ],
                "native_visual_verification_available": native_render[
                    "available"
                ],
            },
            "operations": operations,
            "assets": {
                "binary_in_json": False,
                "image_projection": "native_metadata_only",
                "projected_image_count": sum(
                    isinstance(node, ImageBlock)
                    for node in self._spec.content
                ),
                "asset_count": len(self._spec.assets),
                "read_api": "Document.read_image(image_id)",
                "bytes_api": "Document.image_bytes(image_id)",
                "extract_api": (
                    "Document.extract_image(image_id, target, overwrite=False)"
                ),
                "native_update_operation": "image.update",
                "native_update_fields": [
                    "width",
                    "height",
                    "alt_text",
                    "title",
                ],
                "clearable_update_fields": [
                    "alt_text",
                    "title",
                ],
                "single_dimension_resize": "preserve_aspect_ratio",
                "two_dimension_resize": "exact",
                "native_geometry_patch": [
                    "wp:inline/wp:extent",
                    "pic:spPr/a:xfrm/a:ext",
                ],
                "native_replace_api": (
                    "Document.replace_image(image_id, source, media_type=None)"
                ),
                "native_replace_operation": "image.replace",
                "binary_write_in_json": False,
                "binary_write_transport": "out_of_band",
                "replacement_media_types": [
                    "image/png",
                    "image/jpeg",
                    "image/gif",
                    "image/bmp",
                    "image/tiff",
                ],
                "replacement_strategy": "occurrence_copy_on_write",
                "replacement_preserves": [
                    "image_occurrence_id",
                    "display_extent",
                    "alternative_text",
                    "title",
                    "unrelated_occurrences",
                    "original_image_part",
                ],
                "native_insert_api": (
                    "Document.insert_image_after(target, source, width=..., "
                    "height=..., alt_text=...)"
                ),
                "native_insert_operation": "image.insert_after",
                "insert_placement": "inline",
                "insert_dimensions": "explicit_width_and_height",
                "insert_alt_text": "required",
                "insert_target": "mapped_top_level_body_node",
                "insert_supports_paragraph_style": True,
                "native_layout_operation": "paragraph.format",
                "native_layout_fields": sorted(
                    ParagraphStyle.model_fields
                ),
                "native_layout_target": "projected_image_id",
                "cli_extract": (
                    "aioffice extract-image INPUT IMAGE_ID -o OUTPUT"
                ),
                "cli_replace": (
                    "aioffice replace-image INPUT IMAGE_ID REPLACEMENT -o OUTPUT"
                ),
                "cli_insert": (
                    "aioffice insert-image-after INPUT TARGET REPLACEMENT "
                    "--width VALUE --width-unit UNIT --height VALUE "
                    "--height-unit UNIT --alt-text TEXT -o OUTPUT"
                ),
                "supported_native_subset": [
                    "one embedded DrawingML picture",
                    "inline placement",
                    "explicit positive extent",
                    "rectangular stretch fill",
                    "no crop, rotation, flip, visible outline, or visual effect",
                    "body paragraph with no other visible content",
                ],
                "opaque_native_cases": [
                    "floating or anchored drawing",
                    "text and drawing in one paragraph",
                    "multiple pictures",
                    "linked or external image",
                    "cropped, transformed, outlined, or effected picture",
                    "picture in table, header, or footer",
                    "VML, OLE, or embedded object",
                ],
                "native_render_is_visual_authority": True,
            },
            "selectors": [
                "#node_id",
                "#image_id",
                "#section_id",
                "#header_footer_block_id",
                "#field_id",
                "#table_id + column id/key",
                "#table_id + cell_id",
            ],
            "formatting": {
                "length_units": ["pt", "in", "cm", "mm", "px"],
                "paragraph_properties": sorted(ParagraphStyle.model_fields),
                "text_properties": sorted(TextStyle.model_fields),
                "text_scopes": ["whole_node", "range", "match"],
                "range": {
                    "indexing": "half_open",
                    "unit": "unicode_codepoint",
                },
                "match": {
                    "mode": "exact",
                    "occurrence_indexing": "one_based",
                },
                "clear_semantics": (
                    "Remove direct formatting so named styles or document defaults can apply."
                ),
                "paragraph_surface_contract": {
                    "background": "solid_srgb_fill",
                    "border_edges": [
                        "top",
                        "right",
                        "bottom",
                        "left",
                    ],
                    "border_styles": [
                        "none",
                        "single",
                        "double",
                        "dotted",
                        "dashed",
                        "thick",
                    ],
                    "native_container": "w:pPr",
                    "native_style_inheritance": True,
                    "unsupported_native_features_preserved": [
                        "between_border",
                        "bar_border",
                        "theme_border_colors",
                        "pattern_shading",
                        "theme_shading",
                    ],
                },
                "named_styles": [
                    {
                        "id": style.id,
                        "name": style.name,
                        "semantic_role": style.semantic_role,
                        "heading_level": style.heading_level,
                        "based_on": style.based_on,
                    }
                    for style in style_catalog(self._spec).values()
                    if not style.hidden
                ],
                "section_properties": sorted(SectionLayout.model_fields),
                "section_contract": {
                    "ordered": True,
                    "first_section_start_at": None,
                    "later_sections_start_at": "existing_content_node_id",
                    "native_patch_scope": "one mapped w:sectPr",
                },
                "header_footer_contract": {
                    "part_model": "shared_reusable_parts",
                    "variants": ["default", "first", "even"],
                    "missing_binding": "inherit_previous_section",
                    "editable_blocks": ["paragraph"],
                    "opaque_native_features": [
                        "drawings",
                        "objects",
                        "tables",
                    ],
                },
                "field_contract": {
                    "kinds": [
                        "page_number",
                        "page_count",
                        "section_number",
                        "section_page_count",
                    ],
                    "native_unknown_kind": "read_only",
                    "cached_result": "non_authoritative",
                    "generated_form": "complex_field",
                    "native_forms": ["complex_field", "simple_field"],
                    "native_patch_scope": "one field instruction",
                },
                "table_contract": {
                    "model": (
                        "semantic_columns_rows_cells_and_logical_spans"
                    ),
                    "table_properties": sorted(TableLayout.model_fields),
                    "column_properties": ["width"],
                    "cell_properties": sorted(
                        TableCellFormat.model_fields
                    ),
                    "native_table_patch_scope": "one w:tbl",
                    "native_column_patch_scope": (
                        "one regular w:gridCol and its one-to-one cells"
                    ),
                    "native_cell_patch_scope": (
                        "one mapped anchor w:tc properties element"
                    ),
                    "border_contract": {
                        "styles": [
                            "none",
                            "single",
                            "double",
                            "dotted",
                            "dashed",
                            "thick",
                        ],
                        "width_range_points": [0.25, 12],
                        "space_range_points": [0, 31],
                        "table_edges": [
                            "top",
                            "right",
                            "bottom",
                            "left",
                            "inside_horizontal",
                            "inside_vertical",
                        ],
                        "cell_edges": [
                            "top",
                            "right",
                            "bottom",
                            "left",
                        ],
                        "direct_cell_precedence": True,
                        "color_modes": ["srgb", "auto"],
                        "unsupported_theme_colors_preserved": True,
                        "clear_semantics": (
                            "Remove known direct border XML so table styles "
                            "or inherited formatting can apply."
                        ),
                        "none_semantics": (
                            "Write an explicit no-border edge that suppresses "
                            "the corresponding inherited or table edge."
                        ),
                    },
                    "tables": [
                        {
                            "id": node.id,
                            "regular_grid": node.metadata.get(
                                "regular_grid",
                                all(
                                    cell.column_span == 1
                                    and cell.row_span == 1
                                    for row in node.rows
                                    for cell in row.cells
                                ),
                            ),
                            "logical_grid": node.metadata.get(
                                "logical_grid",
                                True,
                            ),
                            "column_ids": [
                                column.id for column in node.columns
                            ],
                            "column_keys": [
                                column.key for column in node.columns
                            ],
                            "cell_ids": [
                                cell.id
                                for row in node.rows
                                for cell in row.cells
                            ],
                            "rich_cell_paragraph_ids": [
                                paragraph.id
                                for row in node.rows
                                for cell in row.cells
                                for paragraph in cell.content
                            ],
                            "read_only_cell_ids": [
                                cell.id
                                for row in node.rows
                                for cell in row.cells
                                if cell.metadata.get(
                                    "content_editable"
                                )
                                is False
                            ],
                        }
                        for node in self._spec.content
                        if isinstance(node, Table)
                    ],
                },
            },
            "structural_editing": {
                "available": not detached_native_projection,
                "move_operation": "node.move_after",
                "move_operations": {
                    "after": "node.move_after",
                    "before": "node.move_before",
                },
                "selectors": "stable_top_level_content_ids",
                "position": "after_complete_anchor_range",
                "positions": [
                    "after_complete_anchor_range",
                    "before_complete_anchor_range",
                ],
                "native_scope": "word_document_body_top_level",
                "multi_element_nodes": "move_as_one_contiguous_group",
                "section_policy": "same_section_only",
                "section_start_anchor_movable": False,
                "prepend_to_section": "rebind_section_start_at",
                "native_section_carrier_movable": False,
                "third_party_identity_manifest": (
                    "attach_on_first_successful_structural_edit"
                ),
                "native_authority_requires_attached_package": True,
                "source_mutation": False,
                "supports_dry_run": True,
            },
            "identity": {
                "source": native_extension.get(
                    "identity_source",
                    "semantic_spec" if self._native is None else None,
                ),
                "embedded_on_docx_export": True,
                "ambiguous_node_ids": ambiguous_node_ids,
                "safe_to_commit": not ambiguous_node_ids,
            },
            "roundtrip": (
                {
                    "format": self._native.format_name,
                    "policy": self._native.policy.value,
                    "affected_parts": list(self._native.affected_parts),
                    "noop_exact": not self._native.affected_parts,
                }
                if self._native is not None
                else None
            ),
        }

    def validate(self) -> ValidationResult:
        diagnostics = self.import_diagnostics
        native_projection = self._native is not None or (
            self._spec.extensions.get("dev.aioffice.native", {}).get("authority")
            == "native"
        )
        seen: dict[str, str] = {self.id: "artifact"}
        content_positions = {
            node.id: index for index, node in enumerate(self._spec.content)
        }
        used_section_anchors: set[str] = set()
        previous_section_position = -1
        previous_heading_level: int | None = None
        style_issue_severity = (
            Severity.WARNING if native_projection else Severity.ERROR
        )

        if not self._spec.content:
            diagnostics.append(
                Diagnostic(
                    severity=Severity.WARNING,
                    code="EMPTY_DOCUMENT",
                    message="The document has no content.",
                    node_ids=[self.id],
                    suggested_actions=[{"action": "add_content"}],
                )
            )

        assets_by_id: dict[str, AssetRef] = {}
        for asset_index, asset in enumerate(self._spec.assets):
            asset_path = f"assets.{asset_index}"
            if asset.id in seen:
                diagnostics.append(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=f"Duplicate node ID {asset.id!r}.",
                        node_ids=[asset.id],
                        path=f"{asset_path}.id",
                    )
                )
            else:
                seen[asset.id] = asset_path
            if asset.id in assets_by_id:
                diagnostics.append(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=f"Duplicate asset ID {asset.id!r}.",
                        node_ids=[asset.id],
                        path=f"{asset_path}.id",
                    )
                )
            else:
                assets_by_id[asset.id] = asset

        if get_theme(self._spec.theme.ref) is None:
            diagnostics.append(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="UNSUPPORTED_FEATURE",
                    message=f"Theme {self._spec.theme.ref!r} is not registered.",
                    recoverable=True,
                    suggested_actions=[{"action": "use_theme", "name": "business-clean"}],
                )
            )

        for index, section in enumerate(self._spec.sections):
            if section.id in seen:
                diagnostics.append(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=f"Duplicate node ID {section.id!r}.",
                        node_ids=[section.id],
                        path=f"sections.{index}.id",
                        suggested_actions=[{"action": "assign_unique_id"}],
                    )
                )
            else:
                seen[section.id] = f"sections.{index}"
            if (
                section.revision_added > self.revision
                or section.revision_updated > self.revision
            ):
                diagnostics.append(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            f"Section {section.id!r} references a revision newer "
                            f"than artifact revision {self.revision}."
                        ),
                        node_ids=[section.id],
                        path=f"sections.{index}",
                    )
                )
            if index == 0:
                if section.start_at is not None:
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="INVALID_SECTION_ANCHOR",
                            message="The first document section must have start_at=null.",
                            node_ids=[section.id],
                            path=f"sections.{index}.start_at",
                            suggested_actions=[
                                {
                                    "action": "clear_section_anchor",
                                    "section_id": section.id,
                                }
                            ],
                        )
                    )
            elif section.start_at is None:
                diagnostics.append(
                    Diagnostic(
                        severity=(
                            Severity.WARNING
                            if native_projection
                            else Severity.ERROR
                        ),
                        code="INVALID_SECTION_ANCHOR",
                        message=(
                            f"Section {section.id!r} has no content anchor; this is "
                            "only tolerated for an empty native section."
                        ),
                        node_ids=[section.id],
                        path=f"sections.{index}.start_at",
                        recoverable=True,
                        suggested_actions=[
                            {
                                "action": "set_section_anchor",
                                "section_id": section.id,
                            }
                        ],
                    )
                )
            else:
                position = content_positions.get(section.start_at)
                if position is None:
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="INVALID_SECTION_ANCHOR",
                            message=(
                                f"Section {section.id!r} starts at missing content "
                                f"node {section.start_at!r}."
                            ),
                            node_ids=[section.id, section.start_at],
                            path=f"sections.{index}.start_at",
                            suggested_actions=[{"action": "inspect_nodes"}],
                        )
                    )
                elif section.start_at in used_section_anchors:
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="INVALID_SECTION_ANCHOR",
                            message=(
                                f"Multiple sections start at node "
                                f"{section.start_at!r}."
                            ),
                            node_ids=[section.id, section.start_at],
                            path=f"sections.{index}.start_at",
                        )
                    )
                elif position <= previous_section_position:
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="INVALID_SECTION_ORDER",
                            message="Document sections are not in content order.",
                            node_ids=[section.id],
                            path=f"sections.{index}.start_at",
                            suggested_actions=[{"action": "reorder_sections"}],
                        )
                    )
                else:
                    used_section_anchors.add(section.start_at)
                    previous_section_position = position

            bindings = section.header_footer
            if bindings is not None:
                even_bound = (
                    bindings.header_even is not None
                    or bindings.footer_even is not None
                )
                first_bound = (
                    bindings.header_first is not None
                    or bindings.footer_first is not None
                )
                if even_bound and not (
                    self._spec.settings is not None
                    and self._spec.settings.even_and_odd_headers is True
                ):
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.WARNING,
                            code="HEADER_FOOTER_BINDING_INACTIVE",
                            message=(
                                f"Section {section.id!r} has even-page bindings, "
                                "but even_and_odd_headers is not enabled."
                            ),
                            node_ids=[section.id],
                            path=f"sections.{index}.header_footer",
                            suggested_actions=[
                                {"action": "enable_even_and_odd_headers"}
                            ],
                        )
                    )
                if first_bound and section.layout.different_first_page is not True:
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.WARNING,
                            code="HEADER_FOOTER_BINDING_INACTIVE",
                            message=(
                                f"Section {section.id!r} has first-page bindings, "
                                "but different_first_page is not enabled."
                            ),
                            node_ids=[section.id],
                            path=f"sections.{index}.header_footer",
                            suggested_actions=[
                                {
                                    "action": "set_section_format",
                                    "different_first_page": True,
                                }
                            ],
                        )
                    )

            columns = section.layout.columns
            page_size = section.layout.page_size
            if (
                columns is not None
                and not columns.equal_width
                and page_size is not None
                and section.layout.margin_left is not None
                and section.layout.margin_right is not None
            ):
                width, _ = page_size.dimensions_points()
                printable_width = (
                    width
                    - section.layout.margin_left.to_points()
                    - section.layout.margin_right.to_points()
                    - (
                        section.layout.gutter.to_points()
                        if section.layout.gutter is not None
                        else 0
                    )
                )
                required_width = sum(
                    column.width.to_points()
                    for column in columns.columns
                ) + sum(
                    column.space_after.to_points()
                    for column in columns.columns[:-1]
                )
                if required_width > printable_width + 0.01:
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="SECTION_COLUMNS_OVERFLOW",
                            message=(
                                f"Section {section.id!r} columns require "
                                f"{required_width:.2f}pt but only "
                                f"{printable_width:.2f}pt is printable."
                            ),
                            node_ids=[section.id],
                            path=f"sections.{index}.layout.columns",
                            suggested_actions=[
                                {
                                    "action": "reduce_column_widths_or_margins",
                                    "available_width_pt": printable_width,
                                }
                            ],
                        )
                    )

        named_styles = style_catalog(self._spec)
        for style in self._spec.styles:
            for field_name, referenced_id in (
                ("based_on", style.based_on),
                ("next_style", style.next_style),
            ):
                if referenced_id is not None and referenced_id not in named_styles:
                    diagnostics.append(
                        Diagnostic(
                            severity=style_issue_severity,
                            code="STYLE_NOT_FOUND",
                            message=(
                                f"Named style {style.id!r} references missing "
                                f"{field_name} style {referenced_id!r}."
                            ),
                            path=f"styles.{style.id}.{field_name}",
                            recoverable=True,
                            suggested_actions=[
                                {"action": "define_style", "style_id": referenced_id}
                            ],
                        )
                    )

        reported_cycles: set[frozenset[str]] = set()
        for style_id in named_styles:
            visiting: list[str] = []
            current_id: str | None = style_id
            while current_id is not None and current_id in named_styles:
                if current_id in visiting:
                    cycle = frozenset(visiting[visiting.index(current_id) :])
                    if cycle not in reported_cycles:
                        reported_cycles.add(cycle)
                        diagnostics.append(
                            Diagnostic(
                                severity=style_issue_severity,
                                code="STYLE_INHERITANCE_CYCLE",
                                message=(
                                    "Named style inheritance cycle detected: "
                                    f"{' -> '.join(sorted(cycle))}."
                                ),
                                recoverable=True,
                                suggested_actions=[
                                    {
                                        "action": "clear_style_base",
                                        "style_id": current_id,
                                    }
                                ],
                            )
                        )
                    break
                visiting.append(current_id)
                current_id = named_styles[current_id].based_on

        for part_index, part in enumerate(self._spec.header_footers):
            if part.id in seen:
                diagnostics.append(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=f"Duplicate node ID {part.id!r}.",
                        node_ids=[part.id],
                        path=f"header_footers.{part_index}.id",
                    )
                )
            else:
                seen[part.id] = f"header_footers.{part_index}"
            if (
                part.revision_added > self.revision
                or part.revision_updated > self.revision
            ):
                diagnostics.append(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            f"Header/footer part {part.id!r} references a "
                            "revision newer than the artifact."
                        ),
                        node_ids=[part.id],
                        path=f"header_footers.{part_index}",
                    )
                )
            for block_index, block in enumerate(part.content):
                block_path = (
                    f"header_footers.{part_index}.content.{block_index}"
                )
                if block.id in seen:
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="INVALID_SPEC",
                            message=f"Duplicate node ID {block.id!r}.",
                            node_ids=[block.id],
                            path=f"{block_path}.id",
                        )
                    )
                else:
                    seen[block.id] = block_path
                if (
                    block.revision_added > self.revision
                    or block.revision_updated > self.revision
                ):
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="INVALID_SPEC",
                            message=(
                                f"Header/footer block {block.id!r} references "
                                "a revision newer than the artifact."
                            ),
                            node_ids=[block.id],
                            path=block_path,
                        )
                    )
                if isinstance(block, OpaqueBlock) and not native_projection:
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="UNSUPPORTED_FEATURE",
                            message=(
                                "Opaque header/footer content can only be "
                                "preserved from a native document."
                            ),
                            node_ids=[block.id],
                            path=block_path,
                        )
                    )
                if (
                    isinstance(block, Paragraph)
                    and block.style_ref is not None
                    and block.style_ref not in named_styles
                ):
                    diagnostics.append(
                        Diagnostic(
                            severity=style_issue_severity,
                            code="STYLE_NOT_FOUND",
                            message=(
                                f"Header/footer paragraph {block.id!r} references "
                                f"missing named style {block.style_ref!r}."
                            ),
                            node_ids=[block.id],
                            path=f"{block_path}.style_ref",
                            recoverable=True,
                        )
                    )

        for index, node in enumerate(self._spec.content):
            if node.id in seen:
                diagnostics.append(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=f"Duplicate node ID {node.id!r}.",
                        node_ids=[node.id],
                        path=f"content.{index}.id",
                        suggested_actions=[{"action": "assign_unique_id"}],
                    )
                )
            else:
                seen[node.id] = f"content.{index}"

            if node.revision_added > self.revision or node.revision_updated > self.revision:
                diagnostics.append(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            f"Node {node.id!r} references a revision newer than "
                            f"artifact revision {self.revision}."
                        ),
                        node_ids=[node.id],
                        path=f"content.{index}",
                    )
                )

            if isinstance(node, ImageBlock):
                asset = assets_by_id.get(node.asset_id)
                if asset is None:
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="ASSET_NOT_FOUND",
                            message=(
                                f"Image {node.id!r} references missing asset "
                                f"{node.asset_id!r}."
                            ),
                            node_ids=[node.id, node.asset_id],
                            path=f"content.{index}.asset_id",
                            suggested_actions=[
                                {"action": "inspect_assets"}
                            ],
                        )
                    )
                if not native_projection:
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="UNSUPPORTED_FEATURE",
                            message=(
                                "Image blocks in this release can only be "
                                "projected from an attached native DOCX."
                            ),
                            node_ids=[node.id],
                            path=f"content.{index}",
                        )
                    )
                if node.alt_text is None or not node.alt_text.strip():
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.WARNING,
                            code="IMAGE_ALT_TEXT_MISSING",
                            message=(
                                f"Image {node.id!r} has no native alternative "
                                "text."
                            ),
                            node_ids=[node.id],
                            path=f"content.{index}.alt_text",
                            recoverable=True,
                            suggested_actions=[
                                {
                                    "action": "add_native_image_alt_text",
                                    "image_id": node.id,
                                }
                            ],
                        )
                    )
                if (
                    node.style_ref is not None
                    and node.style_ref not in named_styles
                ):
                    diagnostics.append(
                        Diagnostic(
                            severity=style_issue_severity,
                            code="STYLE_NOT_FOUND",
                            message=(
                                f"Image {node.id!r} references missing "
                                f"paragraph style {node.style_ref!r}."
                            ),
                            node_ids=[node.id],
                            path=f"content.{index}.style_ref",
                            recoverable=True,
                        )
                    )

            if isinstance(node, OpaqueBlock) and not native_projection:
                diagnostics.append(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="UNSUPPORTED_FEATURE",
                        message=(
                            "Opaque body content can only be preserved from "
                            "a native document."
                        ),
                        node_ids=[node.id],
                        path=f"content.{index}",
                    )
                )

            if isinstance(node, Heading):
                if previous_heading_level is not None and node.level > previous_heading_level + 1:
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.WARNING,
                            code="HEADING_LEVEL_JUMP",
                            message=(
                                f"Heading {node.id!r} jumps from level "
                                f"{previous_heading_level} to {node.level}."
                            ),
                            node_ids=[node.id],
                            path=f"content.{index}.level",
                            suggested_actions=[
                                {
                                    "action": "set_heading_level",
                                    "maximum": previous_heading_level + 1,
                                }
                            ],
                        )
                    )
                previous_heading_level = node.level

            if isinstance(node, (Heading, Paragraph)):
                if node.style_ref is not None and node.style_ref not in named_styles:
                    diagnostics.append(
                        Diagnostic(
                            severity=style_issue_severity,
                            code="STYLE_NOT_FOUND",
                            message=(
                                f"Node {node.id!r} references missing named style "
                                f"{node.style_ref!r}."
                            ),
                            node_ids=[node.id],
                            path=f"content.{index}.style_ref",
                            recoverable=True,
                            suggested_actions=[
                                {"action": "define_style", "style_id": node.style_ref},
                                {"action": "clear_style_ref", "node_id": node.id},
                            ],
                        )
                    )
                elif node.style_ref is not None:
                    referenced_style = named_styles[node.style_ref]
                    style_is_heading = (
                        referenced_style.semantic_role == "heading"
                        and referenced_style.heading_level is not None
                        and referenced_style.heading_level <= 6
                    )
                    if style_is_heading != isinstance(node, Heading):
                        diagnostics.append(
                            Diagnostic(
                                severity=style_issue_severity,
                                code="STYLE_SEMANTIC_MISMATCH",
                                message=(
                                    f"Node {node.id!r} type {node.type!r} is incompatible "
                                    f"with named style {node.style_ref!r} role "
                                    f"{referenced_style.semantic_role!r}."
                                ),
                                node_ids=[node.id],
                                path=f"content.{index}.style_ref",
                                recoverable=True,
                                suggested_actions=[
                                    {
                                        "action": "apply_compatible_style",
                                        "style_id": node.style_ref,
                                    }
                                ],
                            )
                        )
                    elif (
                        isinstance(node, Heading)
                        and referenced_style.heading_level != node.level
                    ):
                        diagnostics.append(
                            Diagnostic(
                                severity=style_issue_severity,
                                code="STYLE_SEMANTIC_MISMATCH",
                                message=(
                                    f"Heading {node.id!r} level {node.level} conflicts "
                                    f"with named style {node.style_ref!r} level "
                                    f"{referenced_style.heading_level}."
                                ),
                                node_ids=[node.id],
                                path=f"content.{index}.level",
                                recoverable=True,
                                suggested_actions=[
                                    {
                                        "action": "set_heading_level",
                                        "level": referenced_style.heading_level,
                                    }
                                ],
                            )
                        )
                elif isinstance(node, Heading) and node.style_ref is None:
                    implicit_style = named_styles.get(f"Heading{node.level}")
                    if (
                        implicit_style is None
                        or implicit_style.semantic_role != "heading"
                        or implicit_style.heading_level != node.level
                    ):
                        diagnostics.append(
                            Diagnostic(
                                severity=style_issue_severity,
                                code="STYLE_SEMANTIC_MISMATCH",
                                message=(
                                    f"Heading {node.id!r} has no compatible implicit "
                                    f"Heading{node.level} named style."
                                ),
                                node_ids=[node.id],
                                path=f"content.{index}.style_ref",
                                recoverable=True,
                                suggested_actions=[
                                    {
                                        "action": "apply_compatible_style",
                                        "heading_level": node.level,
                                    }
                                ],
                            )
                        )

            if isinstance(node, Table):
                keys = [column.key for column in node.columns]
                if len(keys) != len(set(keys)):
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="INVALID_SPEC",
                            message=f"Table {node.id!r} has duplicate column keys.",
                            node_ids=[node.id],
                            path=f"content.{index}.columns",
                        )
                    )
                known_keys = set(keys)
                active_section = self._spec.sections[0]
                for candidate_section in self._spec.sections[1:]:
                    anchor_position = (
                        content_positions.get(candidate_section.start_at)
                        if candidate_section.start_at is not None
                        else None
                    )
                    if (
                        anchor_position is not None
                        and anchor_position <= index
                    ):
                        active_section = candidate_section
                page_size = active_section.layout.page_size
                margin_left = active_section.layout.margin_left
                margin_right = active_section.layout.margin_right
                if (
                    page_size is not None
                    and margin_left is not None
                    and margin_right is not None
                ):
                    page_width, _ = page_size.dimensions_points()
                    printable_width = (
                        page_width
                        - margin_left.to_points()
                        - margin_right.to_points()
                        - (
                            active_section.layout.gutter.to_points()
                            if active_section.layout.gutter is not None
                            else 0
                        )
                    )
                    explicit_widths = [
                        column.width.to_points()
                        for column in node.columns
                        if column.width is not None
                    ]
                    table_indent = (
                        node.layout.indent.to_points()
                        if node.layout.indent is not None
                        else 0
                    )
                    if (
                        len(explicit_widths) == len(node.columns)
                        and sum(explicit_widths) + table_indent
                        > printable_width + 0.01
                    ):
                        diagnostics.append(
                            Diagnostic(
                                severity=Severity.WARNING,
                                code="TABLE_WIDTH_OVERFLOW",
                                message=(
                                    f"Table {node.id!r} requests "
                                    f"{sum(explicit_widths) + table_indent:.2f}pt "
                                    f"inside {printable_width:.2f}pt of printable "
                                    "section width."
                                ),
                                node_ids=[
                                    active_section.id,
                                    node.id,
                                    *[
                                        column.id
                                        for column in node.columns
                                    ],
                                ],
                                path=f"content.{index}.columns",
                                recoverable=True,
                                suggested_actions=[
                                    {
                                        "action": "resize_table_columns",
                                        "available_width_pt": printable_width,
                                    }
                                ],
                            )
                        )
                    preferred = node.layout.preferred_width
                    if (
                        preferred is not None
                        and preferred.mode == "exact"
                        and isinstance(preferred.value, Length)
                        and preferred.value.to_points() + table_indent
                        > printable_width + 0.01
                    ):
                        diagnostics.append(
                            Diagnostic(
                                severity=Severity.WARNING,
                                code="TABLE_WIDTH_OVERFLOW",
                                message=(
                                    f"Table {node.id!r} preferred width exceeds "
                                    "the printable section width."
                                ),
                                node_ids=[active_section.id, node.id],
                                path=(
                                    f"content.{index}.layout."
                                    "preferred_width"
                                ),
                                recoverable=True,
                                suggested_actions=[
                                    {
                                        "action": "set_table_width",
                                        "maximum_width_pt": (
                                            printable_width - table_indent
                                        ),
                                    }
                                ],
                            )
                        )
                if (
                    node.layout.algorithm == "fixed"
                    and any(
                        column.width is None
                        for column in node.columns
                    )
                ):
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.WARNING,
                            code="TABLE_COLUMN_WIDTH_INCOMPLETE",
                            message=(
                                f"Fixed-layout table {node.id!r} has columns "
                                "without explicit widths."
                            ),
                            node_ids=[
                                node.id,
                                *[
                                    column.id
                                    for column in node.columns
                                    if column.width is None
                                ],
                            ],
                            path=f"content.{index}.columns",
                            suggested_actions=[
                                {"action": "set_table_column_widths"}
                            ],
                        )
                    )
                for column_index, column in enumerate(node.columns):
                    if column.id in seen:
                        diagnostics.append(
                            Diagnostic(
                                severity=Severity.ERROR,
                                code="INVALID_SPEC",
                                message=f"Duplicate node ID {column.id!r}.",
                                node_ids=[node.id, column.id],
                                path=(
                                    f"content.{index}.columns."
                                    f"{column_index}.id"
                                ),
                                suggested_actions=[
                                    {"action": "assign_unique_id"}
                                ],
                            )
                        )
                    else:
                        seen[column.id] = (
                            f"content.{index}.columns.{column_index}"
                        )
                    if (
                        column.revision_added > self.revision
                        or column.revision_updated > self.revision
                    ):
                        diagnostics.append(
                            Diagnostic(
                                severity=Severity.ERROR,
                                code="INVALID_SPEC",
                                message=(
                                    f"Table column {column.id!r} references "
                                    "a revision newer than the artifact."
                                ),
                                node_ids=[node.id, column.id],
                                path=(
                                    f"content.{index}.columns."
                                    f"{column_index}"
                                ),
                            )
                        )
                grid_owners: dict[tuple[int, int], str] = {}
                column_positions = {
                    column.key: position
                    for position, column in enumerate(node.columns)
                }
                for row_index, row in enumerate(node.rows):
                    if row.id in seen:
                        diagnostics.append(
                            Diagnostic(
                                severity=Severity.ERROR,
                                code="INVALID_SPEC",
                                message=f"Duplicate node ID {row.id!r}.",
                                node_ids=[row.id],
                                path=f"content.{index}.rows.{row_index}.id",
                            )
                        )
                    else:
                        seen[row.id] = f"content.{index}.rows.{row_index}"
                    if (
                        row.revision_added > self.revision
                        or row.revision_updated > self.revision
                    ):
                        diagnostics.append(
                            Diagnostic(
                                severity=Severity.ERROR,
                                code="INVALID_SPEC",
                                message=(
                                    f"Table row {row.id!r} references a "
                                    "revision newer than the artifact."
                                ),
                                node_ids=[node.id, row.id],
                                path=f"content.{index}.rows.{row_index}",
                            )
                        )
                    cell_keys = [
                        cell.column_key
                        for cell in row.cells
                    ]
                    unknown = sorted(set(cell_keys) - known_keys)
                    if unknown:
                        diagnostics.append(
                            Diagnostic(
                                severity=Severity.ERROR,
                                code="INVALID_SPEC",
                                message=(
                                    f"Table row {row.id!r} uses unknown columns: "
                                    f"{', '.join(unknown)}."
                                ),
                                node_ids=[node.id, row.id],
                                path=f"content.{index}.rows.{row_index}.cells",
                            )
                        )
                    for cell_index, cell in enumerate(row.cells):
                        cell_path = (
                            f"content.{index}.rows.{row_index}."
                            f"cells.{cell_index}"
                        )
                        if cell.id in seen:
                            diagnostics.append(
                                Diagnostic(
                                    severity=Severity.ERROR,
                                    code="INVALID_SPEC",
                                    message=(
                                        f"Duplicate node ID {cell.id!r}."
                                    ),
                                    node_ids=[node.id, row.id, cell.id],
                                    path=f"{cell_path}.id",
                                    suggested_actions=[
                                        {"action": "assign_unique_id"}
                                    ],
                                )
                            )
                        else:
                            seen[cell.id] = cell_path
                        if (
                            cell.revision_added > self.revision
                            or cell.revision_updated > self.revision
                        ):
                            diagnostics.append(
                                Diagnostic(
                                    severity=Severity.ERROR,
                                    code="INVALID_SPEC",
                                    message=(
                                        f"Table cell {cell.id!r} references "
                                        "a revision newer than the artifact."
                                    ),
                                    node_ids=[node.id, row.id, cell.id],
                                    path=cell_path,
                                )
                            )
                        start_column = column_positions.get(
                            cell.column_key
                        )
                        if start_column is not None:
                            end_column = (
                                start_column + cell.column_span
                            )
                            end_row = row_index + cell.row_span
                            if (
                                end_column > len(node.columns)
                                or end_row > len(node.rows)
                            ):
                                diagnostics.append(
                                    Diagnostic(
                                        severity=Severity.ERROR,
                                        code="TABLE_CELL_SPAN_INVALID",
                                        message=(
                                            f"Table cell {cell.id!r} spans "
                                            "outside the logical grid."
                                        ),
                                        node_ids=[
                                            node.id,
                                            row.id,
                                            cell.id,
                                        ],
                                        path=cell_path,
                                        suggested_actions=[
                                            {
                                                "action": (
                                                    "reduce_table_cell_span"
                                                )
                                            }
                                        ],
                                    )
                                )
                            else:
                                for covered_row in range(
                                    row_index,
                                    end_row,
                                ):
                                    for covered_column in range(
                                        start_column,
                                        end_column,
                                    ):
                                        coordinate = (
                                            covered_row,
                                            covered_column,
                                        )
                                        owner = grid_owners.get(
                                            coordinate
                                        )
                                        if owner is not None:
                                            diagnostics.append(
                                                Diagnostic(
                                                    severity=Severity.ERROR,
                                                    code=(
                                                        "TABLE_CELL_OVERLAP"
                                                    ),
                                                    message=(
                                                        f"Table cell "
                                                        f"{cell.id!r} overlaps "
                                                        f"{owner!r}."
                                                    ),
                                                    node_ids=[
                                                        node.id,
                                                        cell.id,
                                                        owner,
                                                    ],
                                                    path=cell_path,
                                                )
                                            )
                                        else:
                                            grid_owners[
                                                coordinate
                                            ] = cell.id
                        for paragraph_index, paragraph in enumerate(
                            cell.content
                        ):
                            paragraph_path = (
                                f"{cell_path}.content."
                                f"{paragraph_index}"
                            )
                            if paragraph.id in seen:
                                diagnostics.append(
                                    Diagnostic(
                                        severity=Severity.ERROR,
                                        code="INVALID_SPEC",
                                        message=(
                                            f"Duplicate node ID "
                                            f"{paragraph.id!r}."
                                        ),
                                        node_ids=[
                                            node.id,
                                            cell.id,
                                            paragraph.id,
                                        ],
                                        path=f"{paragraph_path}.id",
                                    )
                                )
                            else:
                                seen[paragraph.id] = paragraph_path
                            if (
                                paragraph.revision_added
                                > self.revision
                                or paragraph.revision_updated
                                > self.revision
                            ):
                                diagnostics.append(
                                    Diagnostic(
                                        severity=Severity.ERROR,
                                        code="INVALID_SPEC",
                                        message=(
                                            f"Cell paragraph "
                                            f"{paragraph.id!r} references "
                                            "a revision newer than the "
                                            "artifact."
                                        ),
                                        node_ids=[
                                            node.id,
                                            cell.id,
                                            paragraph.id,
                                        ],
                                        path=paragraph_path,
                                    )
                                )
                            if (
                                paragraph.style_ref is not None
                                and paragraph.style_ref
                                not in named_styles
                            ):
                                diagnostics.append(
                                    Diagnostic(
                                        severity=style_issue_severity,
                                        code="STYLE_NOT_FOUND",
                                        message=(
                                            f"Table-cell paragraph "
                                            f"{paragraph.id!r} references "
                                            f"missing named style "
                                            f"{paragraph.style_ref!r}."
                                        ),
                                        node_ids=[
                                            node.id,
                                            cell.id,
                                            paragraph.id,
                                        ],
                                        path=(
                                            f"{paragraph_path}."
                                            "style_ref"
                                        ),
                                        recoverable=True,
                                    )
                                )
                            elif paragraph.style_ref is not None:
                                referenced_style = named_styles[
                                    paragraph.style_ref
                                ]
                                if (
                                    referenced_style.semantic_role
                                    == "heading"
                                ):
                                    diagnostics.append(
                                        Diagnostic(
                                            severity=(
                                                style_issue_severity
                                            ),
                                            code=(
                                                "STYLE_SEMANTIC_MISMATCH"
                                            ),
                                            message=(
                                                f"Table-cell paragraph "
                                                f"{paragraph.id!r} cannot "
                                                "use a semantic heading "
                                                f"style "
                                                f"{paragraph.style_ref!r}."
                                            ),
                                            node_ids=[
                                                node.id,
                                                cell.id,
                                                paragraph.id,
                                            ],
                                            path=(
                                                f"{paragraph_path}."
                                                "style_ref"
                                            ),
                                            recoverable=True,
                                        )
                                    )
                uncovered = [
                    (
                        row_position,
                        node.columns[column_position].key,
                    )
                    for row_position in range(len(node.rows))
                    for column_position in range(len(node.columns))
                    if (
                        row_position,
                        column_position,
                    )
                    not in grid_owners
                ]
                if uncovered:
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="TABLE_GRID_INCOMPLETE",
                            message=(
                                f"Table {node.id!r} leaves "
                                f"{len(uncovered)} logical grid positions "
                                "uncovered."
                            ),
                            node_ids=[node.id],
                            path=f"content.{index}.rows",
                            suggested_actions=[
                                {
                                    "action": "fill_table_cells",
                                    "positions": [
                                        {
                                            "row": row_position,
                                            "column_key": column_key,
                                        }
                                        for row_position, column_key in (
                                            uncovered[:20]
                                        )
                                    ],
                                }
                            ],
                        )
                    )

        for field_index, document_field in enumerate(
            _document_fields(self._spec)
        ):
            field_path = f"fields.{field_index}"
            if document_field.id in seen:
                diagnostics.append(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=f"Duplicate node ID {document_field.id!r}.",
                        node_ids=[document_field.id],
                        path=f"{field_path}.id",
                        suggested_actions=[{"action": "assign_unique_id"}],
                    )
                )
            else:
                seen[document_field.id] = field_path
            if (
                document_field.revision_added > self.revision
                or document_field.revision_updated > self.revision
            ):
                diagnostics.append(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            f"Field {document_field.id!r} references a revision "
                            "newer than the artifact."
                        ),
                        node_ids=[document_field.id],
                        path=field_path,
                    )
                )
            if document_field.kind == "native" and not native_projection:
                diagnostics.append(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="UNSUPPORTED_FEATURE",
                        message=(
                            "A native-only field can only be preserved from an "
                            "existing native document."
                        ),
                        node_ids=[document_field.id],
                        path=field_path,
                    )
                )

        if _document_fields(self._spec) and (
            self._spec.settings is not None
            and self._spec.settings.update_fields_on_open is False
        ):
            diagnostics.append(
                Diagnostic(
                    severity=Severity.WARNING,
                    code="FIELD_REFRESH_DISABLED",
                    message=(
                        "The document contains dynamic fields but explicitly "
                        "disables update_fields_on_open; cached results may be stale."
                    ),
                    node_ids=[field.id for field in _document_fields(self._spec)],
                    suggested_actions=[
                        {"action": "enable_update_fields_on_open"}
                    ],
                )
            )
        return ValidationResult(diagnostics=diagnostics)

    def export(self, target: str | Path) -> Path:
        validation = self.validate()
        if not validation.valid:
            summary = "; ".join(item.message for item in validation.errors)
            raise ExportError(f"Document validation failed: {summary}")
        path = Path(target)
        suffix = path.suffix.lower()
        path.parent.mkdir(parents=True, exist_ok=True)
        if suffix == ".json":
            path.write_text(self.to_json(), encoding="utf-8")
        elif suffix in {".md", ".markdown"}:
            path.write_text(export_markdown(self._spec), encoding="utf-8")
        elif suffix in {".html", ".htm"}:
            path.write_text(export_html(self._spec), encoding="utf-8")
        elif suffix == ".docx":
            if self._native is not None:
                self._native.write(path)
            else:
                export_docx(self._spec, path)
        else:
            raise UnsupportedFormatError(
                f"Unsupported export format {suffix or '<none>'!r}; "
                "use .json, .md, .html, or .docx."
            )
        return path

    def to_bytes(self, format: str = "docx") -> bytes:
        normalized = format.lower().lstrip(".")
        if normalized == "docx":
            if self._native is not None:
                return self._native.export_bytes()
            return compile_docx(self._spec)
        if normalized == "json":
            return self.to_json().encode("utf-8")
        if normalized in {"md", "markdown"}:
            return export_markdown(self._spec).encode("utf-8")
        if normalized in {"html", "htm"}:
            return export_html(self._spec).encode("utf-8")
        raise UnsupportedFormatError(
            f"Unsupported byte export format {format!r}; use docx, json, md, or html."
        )

    def synchronize_identity_manifest(self) -> "Document":
        updated = Document(
            self._spec,
            native=self._native,
            import_diagnostics=self._import_diagnostics,
        )
        if updated._native is not None and updated._native.has_part(MANIFEST_PART_URI):
            updated._native.set_part(
                MANIFEST_PART_URI,
                serialize_identity_manifest(build_identity_manifest(updated._spec)),
            )
        return updated

    def apply(
        self,
        operations: Sequence[Mapping[str, Any]],
        *,
        dry_run: bool = False,
        base_revision: int | None = None,
        idempotency_key: str | None = None,
    ) -> PatchResult:
        return self._apply(
            operations,
            dry_run=dry_run,
            base_revision=base_revision,
            idempotency_key=idempotency_key,
            image_payloads={},
        )

    def _apply(
        self,
        operations: Sequence[Mapping[str, Any]],
        *,
        dry_run: bool,
        base_revision: int | None,
        idempotency_key: str | None,
        image_payloads: Mapping[str, bytes],
    ) -> PatchResult:
        expected_revision = self.revision if base_revision is None else base_revision
        if expected_revision != self.revision:
            diagnostic = Diagnostic(
                severity=Severity.ERROR,
                code="REVISION_CONFLICT",
                message=(
                    f"Patch targets revision {expected_revision}, but the document is at "
                    f"revision {self.revision}."
                ),
                node_ids=[self.id],
                suggested_actions=[{"action": "refresh_artifact"}],
            )
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=[diagnostic],
                idempotency_key=idempotency_key,
            )
        if not operations:
            diagnostic = Diagnostic(
                severity=Severity.ERROR,
                code="INVALID_SPEC",
                message="A patch must contain at least one operation.",
                suggested_actions=[{"action": "add_operation"}],
            )
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=[diagnostic],
                idempotency_key=idempotency_key,
            )
        native_image_operations = {
            str(operation.get("op"))
            for operation in operations
        }.intersection(
            {
                "image.insert_after",
                "image.update",
                "image.replace",
            }
        )
        image_ids = {
            node.id
            for node in self._spec.content
            if isinstance(node, ImageBlock)
        }
        image_layout_requested = any(
            operation.get("op") == "paragraph.format"
            and isinstance(operation.get("target"), str)
            and str(operation["target"]).removeprefix("#") in image_ids
            for operation in operations
        )
        if image_layout_requested:
            native_image_operations.add("paragraph.format(image)")
        if self._native is None and native_image_operations:
            diagnostic = Diagnostic(
                severity=Severity.ERROR,
                code="UNSUPPORTED_FEATURE",
                message=(
                    f"{', '.join(sorted(native_image_operations))} requires an "
                    "attached native DOCX package; "
                    "a detached JSON projection cannot safely mutate DrawingML."
                ),
                node_ids=[self.id],
                suggested_actions=[
                    {"action": "open_native_docx"},
                    {"action": "inspect_capabilities"},
                ],
            )
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=[diagnostic],
                idempotency_key=idempotency_key,
            )
        detached_native_move = (
            self._native is None
            and self._spec.extensions.get("dev.aioffice.native", {}).get(
                "authority"
            )
            == "native"
            and any(
                operation.get("op")
                in {"node.move_after", "node.move_before"}
                for operation in operations
            )
        )
        if detached_native_move:
            diagnostic = Diagnostic(
                severity=Severity.ERROR,
                code="UNSUPPORTED_FEATURE",
                message=(
                    "node.move_after and node.move_before require the attached "
                    "native DOCX package for a native-authority projection; "
                    "detached JSON cannot prove or relocate the complete XML "
                    "element range."
                ),
                node_ids=[self.id],
                suggested_actions=[
                    {"action": "open_native_docx"},
                    {"action": "inspect_capabilities"},
                ],
            )
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=[diagnostic],
                idempotency_key=idempotency_key,
            )
        if any(
            operation.get("op")
            in {"image.insert_after", "image.replace"}
            for operation in operations
        ) and not image_payloads:
            diagnostic = Diagnostic(
                severity=Severity.ERROR,
                code="BINARY_ASSET_REQUIRED",
                message=(
                    "Image insertion/replacement binary data cannot be supplied "
                    "through a JSON Patch; use the dedicated Document API or CLI."
                ),
                node_ids=[self.id],
                suggested_actions=[
                    {"action": "use_out_of_band_image_api"},
                ],
            )
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=[diagnostic],
                idempotency_key=idempotency_key,
            )

        payload = self.to_spec()
        next_revision = self.revision + 1
        changes: list[dict[str, Any]] = []
        fidelity: FidelityReport | None = None
        try:
            for operation in operations:
                changes.append(self._apply_operation(payload, dict(operation), next_revision))
            payload["artifact"]["revision"] = next_revision
            updated = Document.from_spec(payload)
            validation = updated.validate()
            if not validation.valid:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="VALIDATION_FAILED",
                        message="The patch result failed document validation.",
                        node_ids=[self.id],
                        suggested_actions=[
                            {
                                "action": "inspect_diagnostics",
                                "diagnostics": [
                                    item.model_dump(mode="json") for item in validation.errors
                                ],
                            }
                        ],
                    )
                )
            if self._native is not None:
                native, fidelity, identity_updates = apply_docx_operations(
                    self._native,
                    self._spec,
                    updated._spec,
                    operations,
                    image_payloads=image_payloads,
                )
                for node in updated._spec.content:
                    if node.id in identity_updates:
                        node.source_ref = identity_updates[node.id]
                for section in updated._spec.sections:
                    if section.id in identity_updates:
                        section.source_ref = identity_updates[section.id]
                for part in updated._spec.header_footers:
                    if part.id in identity_updates:
                        part.source_ref = identity_updates[part.id]
                    for block in part.content:
                        if block.id in identity_updates:
                            block.source_ref = identity_updates[block.id]
                for document_field in _document_fields(updated._spec):
                    if document_field.id in identity_updates:
                        document_field.source_ref = identity_updates[
                            document_field.id
                        ]
                for table in (
                    node
                    for node in updated._spec.content
                    if isinstance(node, Table)
                ):
                    for column in table.columns:
                        if column.id in identity_updates:
                            column.source_ref = identity_updates[
                                column.id
                            ]
                    for row in table.rows:
                        if row.id in identity_updates:
                            row.source_ref = identity_updates[row.id]
                        for cell in row.cells:
                            if cell.id in identity_updates:
                                cell.source_ref = identity_updates[cell.id]
                            for paragraph in cell.content:
                                if paragraph.id in identity_updates:
                                    paragraph.source_ref = (
                                        identity_updates[paragraph.id]
                                    )
                updated._native = native
                updated._import_diagnostics = self.import_diagnostics
        except _PatchFailure as error:
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=[error.diagnostic],
                idempotency_key=idempotency_key,
            )
        except SecurityError as error:
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=[
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="SECURITY_POLICY_VIOLATION",
                        message=str(error),
                        node_ids=[self.id],
                        recoverable=True,
                        suggested_actions=[
                            {"action": "inspect_security_policy"},
                            {"action": "reduce_asset_size"},
                        ],
                    )
                ],
                idempotency_key=idempotency_key,
            )
        except NativePackageError as error:
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=[
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="NATIVE_PATCH_FAILED",
                        message=str(error),
                        node_ids=[self.id],
                        recoverable=True,
                        suggested_actions=[
                            {"action": "use_supported_native_operation"},
                            {"action": "inspect_capabilities"},
                        ],
                    )
                ],
                idempotency_key=idempotency_key,
            )
        except SpecValidationError as error:
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=error.diagnostics,
                idempotency_key=idempotency_key,
            )

        return PatchResult(
            success=True,
            base_revision=self.revision,
            result_revision=next_revision,
            dry_run=dry_run,
            document=updated,
            changes=changes,
            diagnostics=validation.warnings,
            idempotency_key=idempotency_key,
            fidelity=fidelity,
            diff=self.diff(updated),
        )

    @staticmethod
    def _normalize_new_block(candidate: dict[str, Any], next_revision: int) -> dict[str, Any]:
        candidate = deepcopy(candidate)
        candidate.setdefault("revision_added", next_revision)
        candidate.setdefault("revision_updated", next_revision)
        try:
            block = TypeAdapter(Block).validate_python(candidate)
        except ValidationError as error:
            details = _validation_error_diagnostics(error)
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="INVALID_SPEC",
                    message="Patch content is not a valid document block.",
                    suggested_actions=[
                        {
                            "action": "fix_content",
                            "diagnostics": [item.model_dump(mode="json") for item in details],
                        }
                    ],
                )
            ) from error
        return block.model_dump(mode="json", exclude_none=True)

    @staticmethod
    def _target_id(target: Any) -> str:
        if not isinstance(target, str) or not target:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="INVALID_SPEC",
                    message="Patch target must be a node ID or #node_id selector.",
                )
            )
        return target[1:] if target.startswith("#") else target

    @staticmethod
    def _find_node(payload: dict[str, Any], target: Any) -> tuple[int, dict[str, Any]]:
        target_id = Document._target_id(target)
        candidates = [
            *payload["content"],
            *(
                block
                for part in payload.get("header_footers", [])
                for block in part.get("content", [])
            ),
            *(
                paragraph
                for table in payload.get("content", [])
                if table.get("type") == "table"
                for row in table.get("rows", [])
                for cell in row.get("cells", [])
                for paragraph in cell.get("content", [])
            ),
        ]
        matches = [
            (index, node)
            for index, node in enumerate(candidates)
            if node.get("id") == target_id
        ]
        if not matches:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="TARGET_NOT_FOUND",
                    message=f"No node matched #{target_id}.",
                    suggested_actions=[{"action": "inspect_nodes"}],
                )
            )
        if len(matches) > 1:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="AMBIGUOUS_SELECTOR",
                    message=f"Multiple nodes matched #{target_id}.",
                    node_ids=[target_id],
                    suggested_actions=[{"action": "repair_duplicate_ids"}],
                )
            )
        return matches[0]

    @staticmethod
    def _find_content_node(
        payload: dict[str, Any],
        target: Any,
    ) -> tuple[int, dict[str, Any]]:
        target_id = Document._target_id(target)
        matches = [
            (index, node)
            for index, node in enumerate(payload["content"])
            if node.get("id") == target_id
        ]
        if not matches:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="TARGET_NOT_FOUND",
                    message=f"No top-level content node matched #{target_id}.",
                    suggested_actions=[{"action": "inspect_nodes"}],
                )
            )
        if len(matches) > 1:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="AMBIGUOUS_SELECTOR",
                    message=f"Multiple nodes matched #{target_id}.",
                    node_ids=[target_id],
                )
            )
        return matches[0]

    @staticmethod
    def _find_image(
        payload: dict[str, Any],
        target: Any,
    ) -> dict[str, Any]:
        _, node = Document._find_content_node(payload, target)
        if node.get("type") != "image":
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="TARGET_TYPE_MISMATCH",
                    message=(
                        f"Target #{node.get('id')} is not an image block."
                    ),
                    node_ids=[str(node.get("id"))],
                    suggested_actions=[{"action": "inspect_images"}],
                )
            )
        return node

    @staticmethod
    def _find_field(
        payload: dict[str, Any],
        target: Any,
    ) -> dict[str, Any]:
        target_id = Document._target_id(target)
        text_blocks = [
            *(
                node
                for node in payload.get("content", [])
                if node.get("type") in {"heading", "paragraph"}
            ),
            *(
                block
                for part in payload.get("header_footers", [])
                for block in part.get("content", [])
                if block.get("type") == "paragraph"
            ),
        ]
        matches = [
            inline
            for block in text_blocks
            for inline in block.get("content", [])
            if inline.get("type") == "field"
            and inline.get("id") == target_id
        ]
        if not matches:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="TARGET_NOT_FOUND",
                    message=f"No field matched #{target_id}.",
                    suggested_actions=[{"action": "inspect_fields"}],
                )
            )
        if len(matches) > 1:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="AMBIGUOUS_SELECTOR",
                    message=f"Multiple fields matched #{target_id}.",
                    node_ids=[target_id],
                    suggested_actions=[{"action": "repair_duplicate_ids"}],
                )
            )
        return matches[0]

    @staticmethod
    def _find_table(
        payload: dict[str, Any],
        target: Any,
    ) -> dict[str, Any]:
        target_id = Document._target_id(target)
        matches = [
            node
            for node in payload.get("content", [])
            if node.get("type") == "table"
            and node.get("id") == target_id
        ]
        if not matches:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="TARGET_NOT_FOUND",
                    message=f"No table matched #{target_id}.",
                    suggested_actions=[{"action": "inspect_tables"}],
                )
            )
        if len(matches) > 1:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="AMBIGUOUS_SELECTOR",
                    message=f"Multiple tables matched #{target_id}.",
                    node_ids=[target_id],
                )
            )
        return matches[0]

    @staticmethod
    def _find_table_column(
        table: dict[str, Any],
        selector: Any,
    ) -> dict[str, Any]:
        if not isinstance(selector, str) or not selector:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="INVALID_SPEC",
                    message=(
                        "table.column.format requires a non-empty "
                        "column ID or key."
                    ),
                    node_ids=[table["id"]],
                )
            )
        normalized = selector[1:] if selector.startswith("#") else selector
        matches = [
            column
            for column in table.get("columns", [])
            if column.get("id") == normalized
            or column.get("key") == normalized
        ]
        if not matches:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="TARGET_NOT_FOUND",
                    message=(
                        f"No column matched {selector!r} in table "
                        f"{table['id']!r}."
                    ),
                    node_ids=[table["id"]],
                    suggested_actions=[{"action": "inspect_table_columns"}],
                )
            )
        if len(matches) > 1:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="AMBIGUOUS_SELECTOR",
                    message=(
                        f"Column selector {selector!r} is ambiguous in "
                        f"table {table['id']!r}."
                    ),
                    node_ids=[
                        table["id"],
                        *[column["id"] for column in matches],
                    ],
                )
            )
        return matches[0]

    @staticmethod
    def _find_table_cell(
        table: dict[str, Any],
        selector: Any,
    ) -> tuple[dict[str, Any], dict[str, Any]]:
        if not isinstance(selector, str) or not selector:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="INVALID_SPEC",
                    message=(
                        "table.cell.format requires a non-empty cell ID."
                    ),
                    node_ids=[table["id"]],
                )
            )
        normalized = selector[1:] if selector.startswith("#") else selector
        matches = [
            (row, cell)
            for row in table.get("rows", [])
            for cell in row.get("cells", [])
            if cell.get("id") == normalized
        ]
        if not matches:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="TARGET_NOT_FOUND",
                    message=(
                        f"No cell matched {selector!r} in table "
                        f"{table['id']!r}."
                    ),
                    node_ids=[table["id"]],
                    suggested_actions=[
                        {"action": "inspect_table_cells"}
                    ],
                )
            )
        if len(matches) > 1:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="AMBIGUOUS_SELECTOR",
                    message=(
                        f"Cell selector {selector!r} is ambiguous in "
                        f"table {table['id']!r}."
                    ),
                    node_ids=[
                        table["id"],
                        *[cell["id"] for _, cell in matches],
                    ],
                )
            )
        return matches[0]

    @staticmethod
    def _find_section(
        payload: dict[str, Any],
        target: Any,
    ) -> tuple[int, dict[str, Any]]:
        target_id = Document._target_id(target)
        matches = [
            (index, section)
            for index, section in enumerate(payload["sections"])
            if section.get("id") == target_id
        ]
        if not matches:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="TARGET_NOT_FOUND",
                    message=f"No section matched #{target_id}.",
                    suggested_actions=[{"action": "inspect_sections"}],
                )
            )
        if len(matches) > 1:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="AMBIGUOUS_SELECTOR",
                    message=f"Multiple sections matched #{target_id}.",
                    node_ids=[target_id],
                    suggested_actions=[{"action": "repair_duplicate_ids"}],
                )
            )
        return matches[0]

    @staticmethod
    def _apply_operation(
        payload: dict[str, Any], operation: dict[str, Any], next_revision: int
    ) -> dict[str, Any]:
        operation_name = operation.get("op")
        if operation_name == "image.insert_after":
            unexpected = sorted(
                set(operation) - {"op", "target", "image", "asset"}
            )
            index, after_node = Document._find_content_node(
                payload,
                operation.get("target"),
            )
            if unexpected:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "image.insert_after received unknown fields: "
                            f"{', '.join(unexpected)}."
                        ),
                        node_ids=[after_node["id"]],
                    )
                )
            try:
                image_insert = ImageInsert.model_validate(
                    operation.get("image")
                )
                asset = AssetRef.model_validate(
                    operation.get("asset")
                )
            except ValidationError as error:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="image.insert_after has invalid metadata.",
                        node_ids=[after_node["id"]],
                        suggested_actions=[
                            {
                                "action": "inspect_image_insert_schema",
                                "diagnostics": [
                                    item.model_dump(mode="json")
                                    for item in _validation_error_diagnostics(
                                        error
                                    )
                                ],
                            }
                        ],
                    )
                ) from error
            if (
                asset.id != f"asset_{asset.sha256}"
                or asset.filename is None
                or asset.size_bytes is None
                or asset.size_bytes <= 0
            ):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "image.insert_after requires a content-addressed "
                            "asset with a filename and positive byte count."
                        ),
                        node_ids=[after_node["id"]],
                    )
                )
            known_node_ids = {
                str(node.get("id"))
                for node in payload.get("content", [])
            } | {
                str(candidate.get("id"))
                for candidate in payload.get("assets", [])
            }
            if image_insert.id in known_node_ids:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            f"Image ID {image_insert.id!r} already exists."
                        ),
                        node_ids=[image_insert.id],
                        suggested_actions=[
                            {"action": "assign_unique_id"}
                        ],
                    )
                )
            image = ImageBlock(
                id=image_insert.id,
                asset_id=asset.id,
                width=image_insert.width,
                height=image_insert.height,
                name=image_insert.name,
                alt_text=image_insert.alt_text,
                title=image_insert.title,
                paragraph_style=image_insert.paragraph_style,
                revision_added=next_revision,
                revision_updated=next_revision,
            ).model_dump(
                mode="json",
                exclude_none=True,
            )
            replacement_asset = asset.model_dump(
                mode="json",
                exclude_none=True,
            )
            existing_assets = [
                candidate
                for candidate in payload.get("assets", [])
                if candidate.get("id") == asset.id
            ]
            if len(existing_assets) > 1 or (
                existing_assets
                and existing_assets[0] != replacement_asset
            ):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "image.insert_after asset identity conflicts with "
                            "an existing asset record."
                        ),
                        node_ids=[image_insert.id, asset.id],
                    )
                )
            if not existing_assets:
                payload.setdefault("assets", []).append(
                    replacement_asset
                )
            payload["content"].insert(index + 1, image)
            return {
                "operation": "image.insert_after",
                "after": after_node["id"],
                "created_nodes": [image_insert.id],
                "asset_ids": [asset.id],
                "binary_transport": "out_of_band",
                "placement": "inline",
            }

        if operation_name == "image.replace":
            unexpected = sorted(
                set(operation) - {"op", "target", "asset"}
            )
            image = Document._find_image(
                payload,
                operation.get("target"),
            )
            if unexpected:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "image.replace received unknown fields: "
                            f"{', '.join(unexpected)}."
                        ),
                        node_ids=[image["id"]],
                    )
                )
            try:
                replacement = AssetRef.model_validate(
                    operation.get("asset")
                )
            except ValidationError as error:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="image.replace has invalid asset metadata.",
                        node_ids=[image["id"]],
                        suggested_actions=[
                            {
                                "action": "inspect_asset_ref_schema",
                                "diagnostics": [
                                    item.model_dump(mode="json")
                                    for item in _validation_error_diagnostics(
                                        error
                                    )
                                ],
                            }
                        ],
                    )
                ) from error
            if (
                replacement.id != f"asset_{replacement.sha256}"
                or replacement.filename is None
                or replacement.size_bytes is None
                or replacement.size_bytes <= 0
            ):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "image.replace requires a content-addressed asset ID, "
                            "filename, and positive byte count."
                        ),
                        node_ids=[image["id"]],
                    )
                )

            before_image = deepcopy(image)
            before_asset = next(
                (
                    deepcopy(asset)
                    for asset in payload.get("assets", [])
                    if asset.get("id") == image.get("asset_id")
                ),
                None,
            )
            replacement_payload = replacement.model_dump(
                mode="json",
                exclude_none=True,
            )
            existing_replacements = [
                asset
                for asset in payload.get("assets", [])
                if asset.get("id") == replacement.id
            ]
            if len(existing_replacements) > 1 or (
                existing_replacements
                and existing_replacements[0] != replacement_payload
            ):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "image.replace asset identity conflicts with an "
                            "existing asset record."
                        ),
                        node_ids=[image["id"], replacement.id],
                    )
                )

            image["asset_id"] = replacement.id
            image["revision_updated"] = next_revision
            referenced_asset_ids = {
                node.get("asset_id")
                for node in payload.get("content", [])
                if node.get("type") == "image"
            }
            retained_assets = [
                asset
                for asset in payload.get("assets", [])
                if asset.get("id") != replacement.id
                and not (
                    asset.get("id") == before_image.get("asset_id")
                    and asset.get("id") not in referenced_asset_ids
                )
            ]
            retained_assets.append(replacement_payload)
            payload["assets"] = retained_assets
            return {
                "operation": "image.replace",
                "image_ids": [image["id"]],
                "replacement_strategy": "occurrence_copy_on_write",
                "binary_transport": "out_of_band",
                "asset_change": {
                    "before": before_asset,
                    "after": replacement_payload,
                },
                "property_changes": [
                    {
                        "path": "asset_id",
                        "before": before_image.get("asset_id"),
                        "after": replacement.id,
                    }
                ],
            }

        if operation_name == "image.update":
            unexpected = sorted(
                set(operation) - {"op", "target", "set", "clear"}
            )
            image = Document._find_image(
                payload,
                operation.get("target"),
            )
            if unexpected:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "image.update received unknown fields: "
                            f"{', '.join(unexpected)}."
                        ),
                        node_ids=[image["id"]],
                    )
                )
            set_values = operation.get("set", {})
            clear_values = operation.get("clear", [])
            valid_shape = (
                isinstance(set_values, dict)
                and isinstance(clear_values, list)
                and all(
                    isinstance(value, str)
                    for value in clear_values
                )
                and len(clear_values) == len(set(clear_values))
            )
            known_fields = set(ImageUpdate.model_fields)
            unknown = (
                sorted(
                    (set(set_values) | set(clear_values))
                    - known_fields
                )
                if valid_shape
                else []
            )
            overlap = (
                sorted(set(set_values) & set(clear_values))
                if valid_shape
                else []
            )
            invalid_clear = (
                sorted(set(clear_values) - {"alt_text", "title"})
                if valid_shape
                else []
            )
            has_null = (
                any(value is None for value in set_values.values())
                if isinstance(set_values, dict)
                else False
            )
            if (
                not valid_shape
                or not set_values
                and not clear_values
                or unknown
                or overlap
                or invalid_clear
                or has_null
            ):
                detail = (
                    "set must be an object and clear a unique string list"
                    if not valid_shape
                    else "at least one change is required"
                    if not set_values and not clear_values
                    else f"unknown properties: {', '.join(unknown)}"
                    if unknown
                    else (
                        "properties both set and cleared: "
                        f"{', '.join(overlap)}"
                    )
                    if overlap
                    else (
                        "properties cannot be cleared: "
                        f"{', '.join(invalid_clear)}"
                    )
                    if invalid_clear
                    else "set values cannot be null"
                )
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=f"Invalid image.update: {detail}.",
                        node_ids=[image["id"]],
                    )
                )
            try:
                normalized_set = ImageUpdate.model_validate(
                    set_values
                ).model_dump(mode="json", exclude_none=True)
            except ValidationError as error:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="image.update has invalid values.",
                        node_ids=[image["id"]],
                        suggested_actions=[
                            {
                                "action": "inspect_image_update_schema",
                                "diagnostics": [
                                    item.model_dump(mode="json")
                                    for item in _validation_error_diagnostics(
                                        error
                                    )
                                ],
                            }
                        ],
                    )
                ) from error
            before = deepcopy(image)
            candidate = {
                **before,
                **deepcopy(normalized_set),
            }
            for field_name in clear_values:
                candidate.pop(field_name, None)
            geometry_fields = {
                field_name
                for field_name in ("width", "height")
                if field_name in normalized_set
            }
            if geometry_fields == {"width"}:
                old_width = Length.model_validate(before["width"])
                old_height = Length.model_validate(before["height"])
                new_width = Length.model_validate(candidate["width"])
                candidate["height"] = Length(
                    value=round(
                        new_width.to_points()
                        * old_height.to_points()
                        / old_width.to_points(),
                        6,
                    ),
                    unit="pt",
                ).model_dump(mode="json")
                geometry_fields.add("height")
            elif geometry_fields == {"height"}:
                old_width = Length.model_validate(before["width"])
                old_height = Length.model_validate(before["height"])
                new_height = Length.model_validate(candidate["height"])
                candidate["width"] = Length(
                    value=round(
                        new_height.to_points()
                        * old_width.to_points()
                        / old_height.to_points(),
                        6,
                    ),
                    unit="pt",
                ).model_dump(mode="json")
                geometry_fields.add("width")
            candidate["revision_updated"] = next_revision
            try:
                normalized = ImageBlock.model_validate(
                    candidate
                ).model_dump(mode="json", exclude_none=True)
            except ValidationError as error:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="image.update produced an invalid image block.",
                        node_ids=[image["id"]],
                        suggested_actions=[
                            {
                                "action": "inspect_image_block_schema",
                                "diagnostics": [
                                    item.model_dump(mode="json")
                                    for item in _validation_error_diagnostics(
                                        error
                                    )
                                ],
                            }
                        ],
                    )
                ) from error
            image.clear()
            image.update(normalized)
            changed_fields = sorted(
                set(set_values)
                | set(clear_values)
                | geometry_fields
            )
            return {
                "operation": "image.update",
                "image_ids": [image["id"]],
                "resize_mode": (
                    "preserve_aspect_ratio"
                    if len(
                        {
                            field_name
                            for field_name in ("width", "height")
                            if field_name in set_values
                        }
                    )
                    == 1
                    else (
                        "exact"
                        if geometry_fields
                        else None
                    )
                ),
                "property_changes": [
                    {
                        "path": field_name,
                        "before": before.get(field_name),
                        "after": normalized.get(field_name),
                    }
                    for field_name in changed_fields
                    if before.get(field_name)
                    != normalized.get(field_name)
                ],
            }

        if operation_name == "table.format":
            unexpected = sorted(
                set(operation) - {"op", "target", "set", "clear"}
            )
            table = Document._find_table(
                payload,
                operation.get("target"),
            )
            if unexpected:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "table.format received unknown fields: "
                            f"{', '.join(unexpected)}."
                        ),
                        node_ids=[table["id"]],
                    )
                )
            set_values = operation.get("set", {})
            clear_values = operation.get("clear", [])
            if (
                not isinstance(set_values, dict)
                or not isinstance(clear_values, list)
                or any(not isinstance(value, str) for value in clear_values)
                or len(clear_values) != len(set(clear_values))
            ):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "table.format requires an object set and a "
                            "list of unique clear property names."
                        ),
                        node_ids=[table["id"]],
                    )
                )
            known_fields = set(TableLayout.model_fields)
            unknown = sorted(
                (set(set_values) | set(clear_values)) - known_fields
            )
            overlap = sorted(set(set_values) & set(clear_values))
            has_null = any(value is None for value in set_values.values())
            if (
                not set_values
                and not clear_values
                or unknown
                or overlap
                or has_null
            ):
                detail = (
                    "at least one change is required"
                    if not set_values and not clear_values
                    else f"unknown properties: {', '.join(unknown)}"
                    if unknown
                    else (
                        "properties both set and cleared: "
                        f"{', '.join(overlap)}"
                    )
                    if overlap
                    else "set values cannot be null"
                )
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=f"Invalid table.format: {detail}.",
                        node_ids=[table["id"]],
                    )
                )
            before = deepcopy(table.get("layout", {}))
            candidate = {**before, **deepcopy(set_values)}
            for field_name in clear_values:
                candidate.pop(field_name, None)
            try:
                normalized = TableLayout.model_validate(
                    candidate
                ).model_dump(mode="json", exclude_none=True)
            except ValidationError as error:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="table.format has invalid values.",
                        node_ids=[table["id"]],
                        suggested_actions=[
                            {
                                "action": "inspect_table_layout_schema",
                                "diagnostics": [
                                    item.model_dump(mode="json")
                                    for item in _validation_error_diagnostics(
                                        error
                                    )
                                ],
                            }
                        ],
                    )
                ) from error
            table["layout"] = normalized
            table["revision_updated"] = next_revision
            changed_fields = sorted(set(set_values) | set(clear_values))
            return {
                "operation": "table.format",
                "table_ids": [table["id"]],
                "property_changes": [
                    {
                        "path": f"layout.{field_name}",
                        "before": before.get(field_name),
                        "after": normalized.get(field_name),
                    }
                    for field_name in changed_fields
                    if before.get(field_name) != normalized.get(field_name)
                ],
            }

        if operation_name == "table.column.format":
            unexpected = sorted(
                set(operation)
                - {"op", "target", "column", "set", "clear"}
            )
            table = Document._find_table(
                payload,
                operation.get("target"),
            )
            column = Document._find_table_column(
                table,
                operation.get("column"),
            )
            if unexpected:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "table.column.format received unknown fields: "
                            f"{', '.join(unexpected)}."
                        ),
                        node_ids=[table["id"], column["id"]],
                    )
                )
            set_values = operation.get("set", {})
            clear_values = operation.get("clear", [])
            valid_shape = (
                isinstance(set_values, dict)
                and isinstance(clear_values, list)
                and all(
                    isinstance(value, str)
                    for value in clear_values
                )
                and len(clear_values) == len(set(clear_values))
            )
            unknown = (
                sorted(
                    (set(set_values) | set(clear_values)) - {"width"}
                )
                if valid_shape
                else []
            )
            overlap = (
                sorted(set(set_values) & set(clear_values))
                if valid_shape
                else []
            )
            has_null = (
                any(value is None for value in set_values.values())
                if isinstance(set_values, dict)
                else False
            )
            if (
                not valid_shape
                or not set_values
                and not clear_values
                or unknown
                or overlap
                or has_null
            ):
                detail = (
                    "set must be an object and clear a unique string list"
                    if not valid_shape
                    else "at least one change is required"
                    if not set_values and not clear_values
                    else f"unknown properties: {', '.join(unknown)}"
                    if unknown
                    else (
                        "properties both set and cleared: "
                        f"{', '.join(overlap)}"
                    )
                    if overlap
                    else "set values cannot be null"
                )
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "Invalid table.column.format: "
                            f"{detail}."
                        ),
                        node_ids=[table["id"], column["id"]],
                    )
                )
            before = deepcopy(column)
            candidate = {**before, **deepcopy(set_values)}
            for field_name in clear_values:
                candidate.pop(field_name, None)
            candidate["revision_updated"] = next_revision
            try:
                normalized = TableColumn.model_validate(
                    candidate
                ).model_dump(mode="json", exclude_none=True)
            except ValidationError as error:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "table.column.format has invalid values."
                        ),
                        node_ids=[table["id"], column["id"]],
                        suggested_actions=[
                            {
                                "action": "inspect_table_column_schema",
                                "diagnostics": [
                                    item.model_dump(mode="json")
                                    for item in _validation_error_diagnostics(
                                        error
                                    )
                                ],
                            }
                        ],
                    )
                ) from error
            column.clear()
            column.update(normalized)
            table["revision_updated"] = next_revision
            return {
                "operation": "table.column.format",
                "table_ids": [table["id"]],
                "column_ids": [column["id"]],
                "property_changes": [
                    {
                        "path": "width",
                        "before": before.get("width"),
                        "after": normalized.get("width"),
                    }
                ]
                if before.get("width") != normalized.get("width")
                else [],
            }

        if operation_name == "table.cell.format":
            unexpected = sorted(
                set(operation)
                - {"op", "target", "cell", "set", "clear"}
            )
            table = Document._find_table(
                payload,
                operation.get("target"),
            )
            row, cell = Document._find_table_cell(
                table,
                operation.get("cell"),
            )
            if unexpected:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "table.cell.format received unknown fields: "
                            f"{', '.join(unexpected)}."
                        ),
                        node_ids=[table["id"], row["id"], cell["id"]],
                    )
                )
            set_values = operation.get("set", {})
            clear_values = operation.get("clear", [])
            valid_shape = (
                isinstance(set_values, dict)
                and isinstance(clear_values, list)
                and all(
                    isinstance(value, str)
                    for value in clear_values
                )
                and len(clear_values) == len(set(clear_values))
            )
            known_fields = set(TableCellFormat.model_fields)
            unknown = (
                sorted(
                    (set(set_values) | set(clear_values))
                    - known_fields
                )
                if valid_shape
                else []
            )
            overlap = (
                sorted(set(set_values) & set(clear_values))
                if valid_shape
                else []
            )
            has_null = (
                any(value is None for value in set_values.values())
                if isinstance(set_values, dict)
                else False
            )
            if (
                not valid_shape
                or not set_values
                and not clear_values
                or unknown
                or overlap
                or has_null
            ):
                detail = (
                    "set must be an object and clear a unique string list"
                    if not valid_shape
                    else "at least one change is required"
                    if not set_values and not clear_values
                    else f"unknown properties: {', '.join(unknown)}"
                    if unknown
                    else (
                        "properties both set and cleared: "
                        f"{', '.join(overlap)}"
                    )
                    if overlap
                    else "set values cannot be null"
                )
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            f"Invalid table.cell.format: {detail}."
                        ),
                        node_ids=[table["id"], row["id"], cell["id"]],
                    )
                )
            before = deepcopy(cell.get("format", {}))
            candidate = {**before, **deepcopy(set_values)}
            for field_name in clear_values:
                candidate.pop(field_name, None)
            try:
                normalized = TableCellFormat.model_validate(
                    candidate
                ).model_dump(mode="json", exclude_none=True)
            except ValidationError as error:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="table.cell.format has invalid values.",
                        node_ids=[table["id"], row["id"], cell["id"]],
                        suggested_actions=[
                            {
                                "action": (
                                    "inspect_table_cell_format_schema"
                                ),
                                "diagnostics": [
                                    item.model_dump(mode="json")
                                    for item in (
                                        _validation_error_diagnostics(
                                            error
                                        )
                                    )
                                ],
                            }
                        ],
                    )
                ) from error
            cell["format"] = normalized
            cell["revision_updated"] = next_revision
            row["revision_updated"] = next_revision
            table["revision_updated"] = next_revision
            changed_fields = sorted(
                set(set_values) | set(clear_values)
            )
            return {
                "operation": "table.cell.format",
                "table_ids": [table["id"]],
                "row_ids": [row["id"]],
                "cell_ids": [cell["id"]],
                "property_changes": [
                    {
                        "path": f"format.{field_name}",
                        "before": before.get(field_name),
                        "after": normalized.get(field_name),
                    }
                    for field_name in changed_fields
                    if before.get(field_name)
                    != normalized.get(field_name)
                ],
            }

        if operation_name == "field.update":
            unexpected = sorted(
                set(operation) - {"op", "target", "set", "clear"}
            )
            if unexpected:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "field.update received unknown fields: "
                            f"{', '.join(unexpected)}."
                        ),
                    )
                )
            document_field = Document._find_field(
                payload,
                operation.get("target"),
            )
            if document_field.get("editable") is False:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="UNSUPPORTED_FEATURE",
                        message=(
                            f"Native field {document_field['id']!r} is read-only "
                            "because its instruction is not normalized."
                        ),
                        node_ids=[document_field["id"]],
                        suggested_actions=[
                            {"action": "preserve_native_field"}
                        ],
                    )
                )
            set_values = operation.get("set", {})
            clear_values = operation.get("clear", [])
            if (
                not isinstance(set_values, dict)
                or not isinstance(clear_values, list)
                or any(not isinstance(value, str) for value in clear_values)
                or len(clear_values) != len(set(clear_values))
            ):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "field.update requires an object set and a list of "
                            "unique clear property names."
                        ),
                        node_ids=[document_field["id"]],
                    )
                )
            if not set_values and not clear_values:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="field.update requires at least one change.",
                        node_ids=[document_field["id"]],
                    )
                )
            known_fields = {"kind", "number_format"}
            unknown = sorted(
                (set(set_values) | set(clear_values)) - known_fields
            )
            overlap = sorted(set(set_values) & set(clear_values))
            invalid_clear = sorted(set(clear_values) - {"number_format"})
            has_null = any(value is None for value in set_values.values())
            if unknown or overlap or invalid_clear or has_null:
                detail = (
                    f"unknown properties: {', '.join(unknown)}"
                    if unknown
                    else f"properties both set and cleared: {', '.join(overlap)}"
                    if overlap
                    else f"properties cannot be cleared: {', '.join(invalid_clear)}"
                    if invalid_clear
                    else "set values cannot be null"
                )
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=f"Invalid field.update: {detail}.",
                        node_ids=[document_field["id"]],
                    )
                )
            before = deepcopy(document_field)
            candidate = {**before, **deepcopy(set_values)}
            for field_name in clear_values:
                candidate.pop(field_name, None)
            candidate["revision_updated"] = next_revision
            try:
                normalized = DocumentField.model_validate(
                    candidate
                ).model_dump(mode="json", exclude_none=True)
            except ValidationError as error:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="field.update has invalid values.",
                        node_ids=[document_field["id"]],
                        suggested_actions=[
                            {
                                "action": "inspect_field_schema",
                                "diagnostics": [
                                    item.model_dump(mode="json")
                                    for item in _validation_error_diagnostics(
                                        error
                                    )
                                ],
                            }
                        ],
                    )
                ) from error
            document_field.clear()
            document_field.update(normalized)
            changed_fields = sorted(set(set_values) | set(clear_values))
            return {
                "operation": "field.update",
                "field_ids": [document_field["id"]],
                "property_changes": [
                    {
                        "path": field_name,
                        "before": before.get(field_name),
                        "after": normalized.get(field_name),
                    }
                    for field_name in changed_fields
                    if before.get(field_name) != normalized.get(field_name)
                ],
            }

        if operation_name == "section.format":
            unexpected = sorted(set(operation) - {"op", "target", "set", "clear"})
            if unexpected:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "section.format received unknown fields: "
                            f"{', '.join(unexpected)}."
                        ),
                    )
                )
            _, section = Document._find_section(payload, operation.get("target"))
            set_values = operation.get("set", {})
            clear_values = operation.get("clear", [])
            if (
                not isinstance(set_values, dict)
                or not isinstance(clear_values, list)
                or any(not isinstance(value, str) for value in clear_values)
                or len(clear_values) != len(set(clear_values))
            ):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "section.format requires an object set and a list "
                            "of unique clear property names."
                        ),
                        node_ids=[section["id"]],
                    )
                )
            if not set_values and not clear_values:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="section.format requires at least one change.",
                        node_ids=[section["id"]],
                    )
                )
            known_fields = set(SectionLayout.model_fields)
            unknown = sorted((set(set_values) | set(clear_values)) - known_fields)
            overlap = sorted(set(set_values) & set(clear_values))
            has_null = any(value is None for value in set_values.values())
            if unknown or overlap or has_null:
                detail = (
                    f"unknown properties: {', '.join(unknown)}"
                    if unknown
                    else f"properties both set and cleared: {', '.join(overlap)}"
                    if overlap
                    else "set values cannot be null"
                )
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=f"Invalid section.format: {detail}.",
                        node_ids=[section["id"]],
                        suggested_actions=[
                            {
                                "action": "inspect_section_schema",
                                "properties": sorted(known_fields),
                            }
                        ],
                    )
                )
            before = deepcopy(section.get("layout", {}))
            candidate = {**before, **deepcopy(set_values)}
            for field_name in clear_values:
                candidate.pop(field_name, None)
            try:
                normalized = SectionLayout.model_validate(candidate).model_dump(
                    mode="json",
                    exclude_none=True,
                )
            except ValidationError as error:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="section.format has invalid values.",
                        node_ids=[section["id"]],
                        suggested_actions=[
                            {
                                "action": "fix_section_layout",
                                "diagnostics": [
                                    item.model_dump(mode="json")
                                    for item in _validation_error_diagnostics(error)
                                ],
                            }
                        ],
                    )
                ) from error
            section["layout"] = normalized
            section["revision_updated"] = next_revision
            changed_fields = sorted(set(set_values) | set(clear_values))
            return {
                "operation": "section.format",
                "section_ids": [section["id"]],
                "property_changes": [
                    {
                        "path": f"layout.{field_name}",
                        "before": before.get(field_name),
                        "after": normalized.get(field_name),
                    }
                    for field_name in changed_fields
                    if before.get(field_name) != normalized.get(field_name)
                ],
            }

        if operation_name == "style.define":
            unexpected = sorted(set(operation) - {"op", "style"})
            if unexpected:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            "style.define received unknown fields: "
                            f"{', '.join(unexpected)}."
                        ),
                    )
                )
            raw_style = operation.get("style")
            if not isinstance(raw_style, dict):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="style.define requires a named style object in style.",
                    )
                )
            try:
                named_style = NamedStyle.model_validate(raw_style)
            except ValidationError as error:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="style.define contains an invalid named style.",
                        suggested_actions=[
                            {
                                "action": "fix_style",
                                "diagnostics": [
                                    item.model_dump(mode="json")
                                    for item in _validation_error_diagnostics(error)
                                ],
                            }
                        ],
                    )
                ) from error
            known_ids = {
                style.id
                for style in theme_named_styles(payload.get("theme", {}).get("ref", ""))
            } | {str(style.get("id")) for style in payload.get("styles", [])}
            if named_style.id in known_ids:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="STYLE_ALREADY_EXISTS",
                        message=f"Named style {named_style.id!r} already exists.",
                        suggested_actions=[
                            {"action": "format_style", "style_id": named_style.id}
                        ],
                    )
                )
            payload.setdefault("styles", []).append(
                named_style.model_dump(mode="json", exclude_none=True)
            )
            return {
                "operation": "style.define",
                "style_ids": [named_style.id],
            }

        if operation_name == "style.apply":
            unexpected = sorted(set(operation) - {"op", "target", "style_ref"})
            if unexpected:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            f"style.apply received unknown fields: {', '.join(unexpected)}."
                        ),
                    )
                )
            _, node = Document._find_node(payload, operation.get("target"))
            if node["type"] not in {"heading", "paragraph"}:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="UNSUPPORTED_FEATURE",
                        message=f"style.apply does not support node type {node['type']!r}.",
                        node_ids=[node["id"]],
                    )
                )
            if "style_ref" not in operation:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="style.apply requires style_ref; use null to clear it.",
                        node_ids=[node["id"]],
                    )
                )
            style_ref = operation["style_ref"]
            if style_ref is not None and (not isinstance(style_ref, str) or not style_ref):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="style.apply style_ref must be a non-empty string or null.",
                        node_ids=[node["id"]],
                    )
                )
            named_style_catalog = {
                style.id: style
                for style in theme_named_styles(payload.get("theme", {}).get("ref", ""))
            }
            named_style_catalog.update(
                {
                    style.id: style
                    for raw_style in payload.get("styles", [])
                    for style in [NamedStyle.model_validate(raw_style)]
                }
            )
            if style_ref is not None and style_ref not in named_style_catalog:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="STYLE_NOT_FOUND",
                        message=f"Named style {style_ref!r} does not exist.",
                        node_ids=[node["id"]],
                        recoverable=True,
                        suggested_actions=[
                            {"action": "inspect_styles"},
                            {"action": "define_style", "style_id": style_ref},
                        ],
                    )
                )
            is_header_footer_block = any(
                candidate.get("id") == node["id"]
                for part in payload.get("header_footers", [])
                for candidate in part.get("content", [])
            )
            is_table_cell_paragraph = any(
                candidate.get("id") == node["id"]
                for table in payload.get("content", [])
                if table.get("type") == "table"
                for row in table.get("rows", [])
                for cell in row.get("cells", [])
                for candidate in cell.get("content", [])
            )
            if (
                (
                    is_header_footer_block
                    or is_table_cell_paragraph
                )
                and style_ref is not None
                and named_style_catalog[style_ref].semantic_role == "heading"
            ):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="UNSUPPORTED_FEATURE",
                        message=(
                            "Header/footer and table-cell paragraphs cannot "
                            "be promoted to semantic headings."
                        ),
                        node_ids=[node["id"]],
                        suggested_actions=[
                            {"action": "use_paragraph_named_style"}
                        ],
                    )
                )
            before = node.get("style_ref")
            before_type = node["type"]
            before_level = node.get("level")
            if style_ref is None:
                node.pop("style_ref", None)
            else:
                node["style_ref"] = style_ref
                selected_style = named_style_catalog[style_ref]
                if (
                    selected_style.semantic_role == "heading"
                    and selected_style.heading_level is not None
                    and selected_style.heading_level <= 6
                ):
                    node["type"] = "heading"
                    node["level"] = selected_style.heading_level
                    if style_ref.casefold() == (
                        f"Heading{selected_style.heading_level}".casefold()
                    ):
                        node.pop("style_ref", None)
                else:
                    node["type"] = "paragraph"
                    node.pop("level", None)
            node["revision_updated"] = next_revision
            property_changes = [
                {
                    "path": "style_ref",
                    "before": before,
                    "after": node.get("style_ref"),
                }
            ]
            if before_type != node["type"]:
                property_changes.append(
                    {
                        "path": "type",
                        "before": before_type,
                        "after": node["type"],
                    }
                )
            if before_level != node.get("level"):
                property_changes.append(
                    {
                        "path": "level",
                        "before": before_level,
                        "after": node.get("level"),
                    }
                )
            return {
                "operation": "style.apply",
                "node_ids": [node["id"]],
                "property_changes": property_changes,
            }

        if operation_name == "style.format":
            unexpected = sorted(set(operation) - {"op", "target", "paragraph", "text"})
            if unexpected:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            f"style.format received unknown fields: {', '.join(unexpected)}."
                        ),
                    )
                )
            raw_target = operation.get("target")
            if not isinstance(raw_target, str) or not raw_target:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="style.format target must be a named style ID.",
                    )
                )
            available_style_ids = {
                str(style.get("id")) for style in payload.get("styles", [])
            } | {
                style.id
                for style in theme_named_styles(payload.get("theme", {}).get("ref", ""))
            }
            style_id = (
                raw_target[1:]
                if raw_target.startswith("@")
                and raw_target not in available_style_ids
                else raw_target
            )
            style_index = next(
                (
                    index
                    for index, style in enumerate(payload.get("styles", []))
                    if style.get("id") == style_id
                ),
                None,
            )
            if style_index is None:
                theme_style = next(
                    (
                        style
                        for style in theme_named_styles(
                            payload.get("theme", {}).get("ref", "")
                        )
                        if style.id == style_id
                    ),
                    None,
                )
                if theme_style is None:
                    raise _PatchFailure(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="STYLE_NOT_FOUND",
                            message=f"Named style {style_id!r} does not exist.",
                            recoverable=True,
                            suggested_actions=[{"action": "inspect_styles"}],
                        )
                    )
                payload.setdefault("styles", []).append(
                    theme_style.model_dump(mode="json", exclude_none=True)
                )
                style_index = len(payload["styles"]) - 1

            style_payload = deepcopy(payload["styles"][style_index])
            property_changes: list[dict[str, Any]] = []
            changed = False
            for scope_name, model_type, field_name in (
                ("paragraph", ParagraphStyle, "paragraph_style"),
                ("text", TextStyle, "text_style"),
            ):
                scope = operation.get(scope_name)
                if scope is None:
                    continue
                if not isinstance(scope, dict) or set(scope) - {"set", "clear"}:
                    raise _PatchFailure(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="INVALID_SPEC",
                            message=(
                                f"style.format {scope_name} must contain only set and clear."
                            ),
                        )
                    )
                set_values = scope.get("set", {})
                clear_values = scope.get("clear", [])
                if (
                    not isinstance(set_values, dict)
                    or not isinstance(clear_values, list)
                    or any(not isinstance(value, str) for value in clear_values)
                    or len(clear_values) != len(set(clear_values))
                ):
                    raise _PatchFailure(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="INVALID_SPEC",
                            message=(
                                f"style.format {scope_name} requires an object set and "
                                "a list of unique clear property names."
                            ),
                        )
                    )
                if not set_values and not clear_values:
                    raise _PatchFailure(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="INVALID_SPEC",
                            message=(
                                f"style.format {scope_name} requires at least one change."
                            ),
                        )
                    )
                known_fields = set(model_type.model_fields)
                unknown = sorted((set(set_values) | set(clear_values)) - known_fields)
                overlap = sorted(set(set_values) & set(clear_values))
                if unknown or overlap or any(value is None for value in set_values.values()):
                    detail = (
                        f"unknown properties: {', '.join(unknown)}"
                        if unknown
                        else f"properties both set and cleared: {', '.join(overlap)}"
                        if overlap
                        else "set values cannot be null"
                    )
                    raise _PatchFailure(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="INVALID_SPEC",
                            message=f"Invalid style.format {scope_name}: {detail}.",
                            suggested_actions=[
                                {
                                    "action": "inspect_style_schema",
                                    "properties": sorted(known_fields),
                                }
                            ],
                        )
                    )
                before_values = deepcopy(style_payload.get(field_name, {}))
                candidate = {**before_values, **deepcopy(set_values)}
                for property_name in clear_values:
                    candidate.pop(property_name, None)
                try:
                    normalized = model_type.model_validate(candidate).model_dump(
                        mode="json",
                        exclude_none=True,
                    )
                except ValidationError as error:
                    raise _PatchFailure(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="INVALID_SPEC",
                            message=f"style.format {scope_name} has invalid values.",
                            suggested_actions=[
                                {
                                    "action": "fix_style",
                                    "diagnostics": [
                                        item.model_dump(mode="json")
                                        for item in _validation_error_diagnostics(error)
                                    ],
                                }
                            ],
                        )
                    ) from error
                if normalized:
                    style_payload[field_name] = normalized
                else:
                    style_payload.pop(field_name, None)
                for property_name in sorted(set(set_values) | set(clear_values)):
                    if before_values.get(property_name) != normalized.get(property_name):
                        property_changes.append(
                            {
                                "path": f"{field_name}.{property_name}",
                                "before": before_values.get(property_name),
                                "after": normalized.get(property_name),
                            }
                        )
                changed = True
            if not changed:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="style.format requires paragraph or text changes.",
                    )
                )
            try:
                normalized_style = NamedStyle.model_validate(style_payload)
            except ValidationError as error:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="style.format produced an invalid named style.",
                    )
                ) from error
            payload["styles"][style_index] = normalized_style.model_dump(
                mode="json",
                exclude_none=True,
            )
            return {
                "operation": "style.format",
                "style_ids": [style_id],
                "property_changes": property_changes,
            }

        if operation_name == "text.replace":
            _, node = Document._find_node(payload, operation.get("target"))
            unexpected = sorted(
                set(operation) - {"op", "target", "search", "replacement", "replace_all"}
            )
            if unexpected:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(f"text.replace received unknown fields: {', '.join(unexpected)}."),
                        node_ids=[node["id"]],
                    )
                )
            search = operation.get("search")
            replacement = operation.get("replacement")
            replace_all = operation.get("replace_all", False)
            if not isinstance(search, str) or not search:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="text.replace requires a non-empty search string.",
                    )
                )
            if not isinstance(replacement, str):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="text.replace requires a string replacement.",
                    )
                )
            if not isinstance(replace_all, bool):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="text.replace requires replace_all to be a boolean.",
                    )
                )
            if node["type"] not in {"heading", "paragraph"}:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="UNSUPPORTED_FEATURE",
                        message=f"text.replace does not support node type {node['type']!r}.",
                        node_ids=[node["id"]],
                    )
                )
            if any(
                inline.get("type") == "field"
                for inline in node.get("content", [])
            ):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="UNSUPPORTED_FEATURE",
                        message=(
                            "text.replace cannot cross or rewrite dynamic fields; "
                            "target the field with field.update or edit a field-free "
                            "paragraph."
                        ),
                        node_ids=[node["id"]],
                        suggested_actions=[{"action": "inspect_fields"}],
                    )
                )
            count = replace_node_text(
                node,
                search=search,
                replacement=replacement,
                replace_all=bool(replace_all),
            )
            if count == 0:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="TARGET_NOT_FOUND",
                        message=f"Search text {search!r} was not found in node {node['id']!r}.",
                        node_ids=[node["id"]],
                        suggested_actions=[{"action": "inspect_node", "node_id": node["id"]}],
                    )
                )
            node["revision_updated"] = next_revision
            return {
                "operation": "text.replace",
                "node_ids": [node["id"]],
                "replacement_count": count,
            }

        if operation_name in {"paragraph.format", "text.format"}:
            _, node = Document._find_node(payload, operation.get("target"))
            supported_node_types = {"heading", "paragraph"}
            if operation_name == "paragraph.format":
                supported_node_types.add("image")
            if node["type"] not in supported_node_types:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="UNSUPPORTED_FEATURE",
                        message=(f"{operation_name} does not support node type {node['type']!r}."),
                        node_ids=[node["id"]],
                    )
                )
            allowed_keys = {"op", "target", "set", "clear"}
            if operation_name == "text.format":
                allowed_keys.update({"range", "match"})
            unexpected = sorted(set(operation) - allowed_keys)
            if unexpected:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            f"{operation_name} received unknown fields: {', '.join(unexpected)}."
                        ),
                        node_ids=[node["id"]],
                    )
                )
            set_values = operation.get("set", {})
            clear_values = operation.get("clear", [])
            if not isinstance(set_values, dict):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=f"{operation_name} requires set to be an object.",
                    )
                )
            if (
                not isinstance(clear_values, list)
                or any(not isinstance(value, str) for value in clear_values)
                or len(clear_values) != len(set(clear_values))
            ):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            f"{operation_name} requires clear to be a list of "
                            "unique property names."
                        ),
                    )
                )
            if not set_values and not clear_values:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=f"{operation_name} requires at least one set or clear field.",
                    )
                )
            if any(value is None for value in set_values.values()):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            f"{operation_name} uses clear for property removal; "
                            "set values cannot be null."
                        ),
                    )
                )
            overlap = sorted(set(set_values) & set(clear_values))
            if overlap:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            f"{operation_name} cannot set and clear the same fields: "
                            f"{', '.join(overlap)}."
                        ),
                    )
                )

            style_field = (
                "paragraph_style" if operation_name == "paragraph.format" else "text_style"
            )
            style_model = ParagraphStyle if operation_name == "paragraph.format" else TextStyle
            known_fields = set(style_model.model_fields)
            unknown_fields = sorted((set(set_values) | set(clear_values)) - known_fields)
            if unknown_fields:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            f"{operation_name} does not recognize: {', '.join(unknown_fields)}."
                        ),
                        node_ids=[node["id"]],
                        suggested_actions=[
                            {
                                "action": "inspect_style_schema",
                                "properties": sorted(known_fields),
                            }
                        ],
                    )
                )
            selection = None
            plain_text = node_plain_text(node)
            if operation_name == "text.format":
                if (
                    operation.get("range") is not None
                    or operation.get("match") is not None
                ) and any(
                    inline.get("type") == "field"
                    for inline in node.get("content", [])
                ):
                    raise _PatchFailure(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="UNSUPPORTED_FEATURE",
                            message=(
                                "Range or match text formatting is disabled for "
                                "paragraphs containing dynamic fields because "
                                "field results are not authoritative text."
                            ),
                            node_ids=[node["id"]],
                            suggested_actions=[
                                {"action": "format_whole_paragraph_text"}
                            ],
                        )
                    )
                try:
                    selection = resolve_text_selection(
                        plain_text,
                        range_value=operation.get("range"),
                        match_value=operation.get("match"),
                    )
                except (ValidationError, ValueError) as error:
                    raise _PatchFailure(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="INVALID_TEXT_SELECTION",
                            message=str(error),
                            node_ids=[node["id"]],
                            suggested_actions=[
                                {
                                    "action": "inspect_node_text",
                                    "length": len(plain_text),
                                    "unit": "unicode_codepoint",
                                }
                            ],
                        )
                    ) from error
            before_style = deepcopy(node.get(style_field, {}))
            try:
                if selection is not None:
                    normalized_set = TextStyle.model_validate(set_values).model_dump(
                        mode="json",
                        exclude_none=True,
                    )
                    format_text_range(
                        node,
                        selection,
                        set_values=normalized_set,
                        clear_values=clear_values,
                    )
                    normalized = {}
                elif operation_name == "text.format":
                    before_style, normalized = format_entire_text(
                        node,
                        set_values=set_values,
                        clear_values=clear_values,
                    )
                else:
                    candidate = deepcopy(before_style)
                    candidate.update(deepcopy(set_values))
                    for field_name in clear_values:
                        candidate.pop(field_name, None)
                    normalized = style_model.model_validate(candidate).model_dump(
                        mode="json",
                        exclude_none=True,
                    )
            except ValidationError as error:
                details = _validation_error_diagnostics(error)
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=f"{operation_name} contains invalid style values.",
                        node_ids=[node["id"]],
                        suggested_actions=[
                            {
                                "action": "fix_style",
                                "diagnostics": [item.model_dump(mode="json") for item in details],
                            }
                        ],
                    )
                ) from error
            if operation_name == "paragraph.format":
                if normalized:
                    node[style_field] = normalized
                else:
                    node.pop(style_field, None)
            node["revision_updated"] = next_revision
            changed_fields = sorted(set(set_values) | set(clear_values))
            if selection is not None:
                return {
                    "operation": operation_name,
                    "node_ids": [node["id"]],
                    "selection": selection.model_dump(mode="json"),
                    "selected_text": plain_text[selection.start : selection.end],
                    "fields": changed_fields,
                }
            return {
                "operation": operation_name,
                "node_ids": [node["id"]],
                "property_changes": [
                    {
                        "path": f"{style_field}.{field_name}",
                        "before": before_style.get(field_name),
                        "after": normalized.get(field_name),
                    }
                    for field_name in changed_fields
                    if before_style.get(field_name) != normalized.get(field_name)
                ],
            }

        if operation_name == "node.append":
            target = operation.get("target", "$")
            if target not in {"$", payload["artifact"]["id"], f"#{payload['artifact']['id']}"}:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="UNSUPPORTED_FEATURE",
                        message="node.append currently supports only the document root target '$'.",
                    )
                )
            content = operation.get("content")
            if not isinstance(content, dict):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="node.append requires an object in content.",
                    )
                )
            candidate = Document._normalize_new_block(content, next_revision)
            payload["content"].append(candidate)
            return {
                "operation": "node.append",
                "created_nodes": [candidate["id"]],
            }

        if operation_name == "node.insert_after":
            index, node = Document._find_content_node(
                payload,
                operation.get("target"),
            )
            content = operation.get("content")
            if not isinstance(content, dict):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="node.insert_after requires an object in content.",
                    )
                )
            candidate = Document._normalize_new_block(content, next_revision)
            payload["content"].insert(index + 1, candidate)
            return {
                "operation": "node.insert_after",
                "after": node["id"],
                "created_nodes": [candidate["id"]],
            }

        if operation_name in {"node.move_after", "node.move_before"}:
            anchor_field = (
                "after"
                if operation_name == "node.move_after"
                else "before"
            )
            unexpected = sorted(
                set(operation) - {"op", "target", anchor_field}
            )
            target_index, target_node = Document._find_content_node(
                payload,
                operation.get("target"),
            )
            anchor_index, anchor_node = Document._find_content_node(
                payload,
                operation.get(anchor_field),
            )
            if unexpected:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            f"{operation_name} received unknown fields: "
                            f"{', '.join(unexpected)}."
                        ),
                        node_ids=[target_node["id"]],
                    )
                )
            if target_node["id"] == anchor_node["id"]:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            f"{operation_name} target and {anchor_field} "
                            "must be different nodes."
                        ),
                        node_ids=[target_node["id"]],
                    )
                )
            already_positioned = (
                target_index == anchor_index + 1
                if anchor_field == "after"
                else target_index + 1 == anchor_index
            )
            if already_positioned:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="NO_CHANGES",
                        message=(
                            f"Node {target_node['id']!r} is already immediately "
                            f"{anchor_field} {anchor_node['id']!r}."
                        ),
                        node_ids=[
                            target_node["id"],
                            anchor_node["id"],
                        ],
                        suggested_actions=[
                            {"action": "choose_different_anchor"}
                        ],
                    )
                )

            content_positions = {
                str(node["id"]): index
                for index, node in enumerate(payload["content"])
            }
            section_starts: list[tuple[int, int]] = [(0, 0)]
            section_anchor_ids: set[str] = set()
            for section_index, section in enumerate(
                payload.get("sections", [])[1:],
                start=1,
            ):
                start_at = section.get("start_at")
                if not isinstance(start_at, str):
                    continue
                start_position = content_positions.get(start_at)
                if start_position is None:
                    continue
                section_starts.append(
                    (start_position, section_index)
                )
                section_anchor_ids.add(start_at)

            def section_for(position: int) -> int:
                return max(
                    (
                        section_index
                        for start, section_index in section_starts
                        if start <= position
                    ),
                    default=0,
                )

            if target_node["id"] in section_anchor_ids:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="UNSUPPORTED_FEATURE",
                        message=(
                            f"{operation_name} cannot move a node that anchors "
                            "the start of a document section."
                        ),
                        node_ids=[target_node["id"]],
                        suggested_actions=[
                            {"action": "move_section_content_not_anchor"}
                        ],
                    )
                )
            target_section = section_for(target_index)
            anchor_section = section_for(anchor_index)
            if target_section != anchor_section:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="CROSS_SECTION_MOVE_UNSUPPORTED",
                        message=(
                            f"{operation_name} currently preserves section "
                            "semantics by requiring target and anchor to be "
                            "in the same section."
                        ),
                        node_ids=[
                            target_node["id"],
                            anchor_node["id"],
                        ],
                        suggested_actions=[
                            {
                                "action": "choose_anchor_in_same_section",
                                "section_index": target_section,
                            }
                        ],
                    )
                )
            rebind_section_start = (
                anchor_field == "before"
                and anchor_node["id"] in section_anchor_ids
            )

            previous_after = (
                payload["content"][target_index - 1]["id"]
                if target_index
                else None
            )
            moved = payload["content"].pop(target_index)
            new_anchor_index, _ = Document._find_content_node(
                payload,
                anchor_node["id"],
            )
            moved["revision_updated"] = next_revision
            insert_index = (
                new_anchor_index + 1
                if anchor_field == "after"
                else new_anchor_index
            )
            payload["content"].insert(insert_index, moved)
            section_start_change: dict[str, Any] | None = None
            if rebind_section_start:
                section = payload["sections"][anchor_section]
                section_start_change = {
                    "section_id": section["id"],
                    "from": anchor_node["id"],
                    "to": target_node["id"],
                }
                section["start_at"] = target_node["id"]
                section["revision_updated"] = next_revision
            change = {
                "operation": operation_name,
                "moved_nodes": [target_node["id"]],
                "from_after": previous_after,
                "section_index": target_section,
            }
            change[anchor_field] = anchor_node["id"]
            if section_start_change is not None:
                change["section_start_updated"] = section_start_change
            return change

        if operation_name == "node.remove":
            index, node = Document._find_content_node(
                payload,
                operation.get("target"),
            )
            payload["content"].pop(index)
            return {"operation": "node.remove", "removed_nodes": [node["id"]]}

        if operation_name == "node.update":
            _, node = Document._find_content_node(
                payload,
                operation.get("target"),
            )
            changes = operation.get("changes")
            if not isinstance(changes, dict) or not changes:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="node.update requires a non-empty changes object.",
                    )
                )
            immutable = {"id", "type", "revision_added", "revision_updated"}
            forbidden = sorted(immutable.intersection(changes))
            if forbidden:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=f"node.update cannot change: {', '.join(forbidden)}.",
                        node_ids=[node["id"]],
                    )
                )
            node.update(deepcopy(changes))
            node["revision_updated"] = next_revision
            return {
                "operation": "node.update",
                "node_ids": [node["id"]],
                "fields": sorted(changes),
            }

        raise _PatchFailure(
            Diagnostic(
                severity=Severity.ERROR,
                code="UNSUPPORTED_FEATURE",
                message=(
                    f"Unsupported operation {operation_name!r}; AiOffice supports "
                    "text.replace, paragraph.format, text.format, node.append, "
                    "node.insert_after, node.move_after, node.move_before, "
                    "node.remove, "
                    "node.update, style.apply, "
                    "style.define, style.format, section.format, field.update, "
                    "image.insert_after, image.replace, image.update, "
                    "table.format, table.column.format, and table.cell.format."
                ),
                suggested_actions=[{"action": "use_supported_operation"}],
            )
        )

    def __repr__(self) -> str:
        return (
            f"Document(id={self.id!r}, revision={self.revision}, nodes={len(self._spec.content)})"
        )


def open_artifact(
    source: str | Path,
    *,
    roundtrip: FidelityPolicy | str = FidelityPolicy.PRESERVE_UNKNOWN,
    security_policy: SecurityPolicy | None = None,
    identity_manifest: IdentityManifest | None = None,
) -> Document:
    path = Path(source)
    suffix = path.suffix.lower()
    if suffix == ".json":
        return Document.from_json(path)
    if suffix in {".md", ".markdown"}:
        return Document.from_markdown(path)
    if suffix == ".docx":
        return Document.from_docx(
            path,
            roundtrip=roundtrip,
            security_policy=security_policy,
            identity_manifest=identity_manifest,
        )
    if suffix in {".html", ".htm"}:
        raise UnsupportedFormatError(f"Importing {suffix} is planned for a later release.")
    raise UnsupportedFormatError(
        f"Unsupported source format {suffix or '<none>'!r}; use .json, .md, or .docx."
    )


__all__ = ["Document", "PatchResult", "open_artifact"]
