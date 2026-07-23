"""Strict Pydantic models for the AiOffice Document Spec 0.2 draft."""

from __future__ import annotations

from typing import Annotated, Any, Literal

from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    StringConstraints,
    field_validator,
    model_validator,
)

from aioffice._version import __version__
from aioffice.core.ids import new_id

SPEC_VERSION = "0.2-draft.1"
DOCUMENT_SCHEMA_URL = "https://schemas.aioffice.dev/spec/draft/0.2/document.json"
LEGACY_SPEC_VERSION = "1.0"
LEGACY_DOCUMENT_SCHEMA_URL = "https://schemas.aioffice.dev/spec/1.0/document.json"

NodeId = Annotated[str, StringConstraints(pattern=r"^[A-Za-z][A-Za-z0-9_.:-]{0,127}$")]
Mark = Literal[
    "strong",
    "emphasis",
    "underline",
    "strike",
    "code",
    "subscript",
    "superscript",
    "link",
    "highlight",
]


class StrictModel(BaseModel):
    model_config = ConfigDict(extra="forbid", populate_by_name=True)


class ArtifactDescriptor(StrictModel):
    id: NodeId = Field(default_factory=lambda: new_id("doc"))
    kind: Literal["document"] = "document"
    revision: int = Field(default=1, ge=1)


class DocumentMetadata(StrictModel):
    title: str | None = None
    author: str | None = None
    subject: str | None = None
    keywords: list[str] = Field(default_factory=list)
    language: str | None = None
    custom: dict[str, Any] = Field(default_factory=dict)


class ThemeRef(StrictModel):
    ref: str = "business-clean"


class AssetRef(StrictModel):
    id: NodeId = Field(default_factory=lambda: new_id("asset"))
    sha256: Annotated[str, StringConstraints(pattern=r"^[a-fA-F0-9]{64}$")]
    media_type: str
    filename: str | None = None


class NativeRef(StrictModel):
    format: Literal["docx", "xlsx", "pptx"]
    part_uri: Annotated[str, StringConstraints(pattern=r"^/[^\\\x00]*$")]
    native_kind: str
    element_index: int | None = Field(default=None, ge=0)
    element_indices: list[int] = Field(default_factory=list)
    path_hint: str | None = None
    native_id: str | None = None
    fingerprint: Annotated[
        str,
        StringConstraints(pattern=r"^sha256:[a-fA-F0-9]{64}$"),
    ] | None = None

    @model_validator(mode="after")
    def validate_element_indices(self) -> "NativeRef":
        if any(index < 0 for index in self.element_indices):
            raise ValueError("Native element indices cannot be negative.")
        if len(self.element_indices) != len(set(self.element_indices)):
            raise ValueError("Native element indices must be unique.")
        if self.element_indices != sorted(self.element_indices):
            raise ValueError("Native element indices must be sorted.")
        if (
            self.element_index is not None
            and self.element_indices
            and self.element_index != self.element_indices[0]
        ):
            raise ValueError("element_index must equal the first element_indices value.")
        return self


class TextSpan(StrictModel):
    type: Literal["text"] = "text"
    text: str
    marks: list[Mark] = Field(default_factory=list)
    href: str | None = None

    @model_validator(mode="after")
    def validate_link(self) -> "TextSpan":
        if "link" in self.marks and not self.href:
            raise ValueError("A text span marked as a link must include href.")
        if self.href and "link" not in self.marks:
            raise ValueError("A text span with href must include the link mark.")
        return self


class NodeBase(StrictModel):
    id: NodeId = Field(default_factory=lambda: new_id("node"))
    tags: list[str] = Field(default_factory=list)
    metadata: dict[str, Any] = Field(default_factory=dict)
    source_ref: NativeRef | str | None = None
    revision_added: int = Field(default=1, ge=1)
    revision_updated: int = Field(default=1, ge=1)


class Heading(NodeBase):
    id: NodeId = Field(default_factory=lambda: new_id("heading"))
    type: Literal["heading"] = "heading"
    level: int = Field(default=1, ge=1, le=6)
    text: str


class Paragraph(NodeBase):
    id: NodeId = Field(default_factory=lambda: new_id("para"))
    type: Literal["paragraph"] = "paragraph"
    text: str | None = None
    content: list[TextSpan] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_content(self) -> "Paragraph":
        if self.text is None and not self.content:
            raise ValueError("A paragraph must include text or content.")
        if self.text is not None and self.content:
            raise ValueError("A paragraph cannot include both text and content.")
        return self

    @property
    def plain_text(self) -> str:
        if self.text is not None:
            return self.text
        return "".join(span.text for span in self.content)


class ListBase(NodeBase):
    items: list[str] = Field(min_length=1)

    @field_validator("items")
    @classmethod
    def non_empty_items(cls, value: list[str]) -> list[str]:
        if any(not item.strip() for item in value):
            raise ValueError("List items cannot be empty.")
        return value


class BulletList(ListBase):
    id: NodeId = Field(default_factory=lambda: new_id("list"))
    type: Literal["bullet_list"] = "bullet_list"


class OrderedList(ListBase):
    id: NodeId = Field(default_factory=lambda: new_id("list"))
    type: Literal["ordered_list"] = "ordered_list"


class TableColumn(StrictModel):
    key: Annotated[str, StringConstraints(pattern=r"^[A-Za-z_][A-Za-z0-9_.-]*$")]
    title: str
    data_type: Literal["text", "number", "integer", "boolean", "date", "enum"] = "text"


class TableRow(StrictModel):
    id: NodeId = Field(default_factory=lambda: new_id("row"))
    values: dict[str, Any]


class Table(NodeBase):
    id: NodeId = Field(default_factory=lambda: new_id("table"))
    type: Literal["table"] = "table"
    columns: list[TableColumn] = Field(min_length=1)
    rows: list[TableRow] = Field(default_factory=list)


class PageBreak(NodeBase):
    id: NodeId = Field(default_factory=lambda: new_id("break"))
    type: Literal["page_break"] = "page_break"


class OpaqueBlock(NodeBase):
    id: NodeId = Field(default_factory=lambda: new_id("opaque"))
    type: Literal["opaque"] = "opaque"
    summary: str
    capabilities: list[
        Literal["inspect", "move", "delete", "render"]
    ] = Field(default_factory=lambda: ["inspect", "render"])
    editable: Literal[False] = False


Block = Annotated[
    Heading | Paragraph | BulletList | OrderedList | Table | PageBreak | OpaqueBlock,
    Field(discriminator="type"),
]


class AiOfficeDocumentSpec(StrictModel):
    schema_url: Literal[
        "https://schemas.aioffice.dev/spec/draft/0.2/document.json"
    ] = Field(
        default=DOCUMENT_SCHEMA_URL,
        alias="$schema",
    )
    spec_version: Literal["0.2-draft.1"] = SPEC_VERSION
    engine_version: str = __version__
    artifact: ArtifactDescriptor = Field(default_factory=ArtifactDescriptor)
    metadata: DocumentMetadata = Field(default_factory=DocumentMetadata)
    theme: ThemeRef = Field(default_factory=ThemeRef)
    content: list[Block] = Field(default_factory=list)
    assets: list[AssetRef] = Field(default_factory=list)
    extensions: dict[str, dict[str, Any]] = Field(default_factory=dict)

    @field_validator("extensions")
    @classmethod
    def validate_extension_namespaces(
        cls, value: dict[str, dict[str, Any]]
    ) -> dict[str, dict[str, Any]]:
        for namespace in value:
            if "." not in namespace:
                raise ValueError(
                    f"Extension namespace {namespace!r} must be a reverse-domain name."
                )
        return value
