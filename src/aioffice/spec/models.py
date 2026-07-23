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

SPEC_VERSION = "0.2-draft.3"
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
    model_config = ConfigDict(
        extra="forbid",
        populate_by_name=True,
        allow_inf_nan=False,
    )


class Length(StrictModel):
    """An explicit physical or CSS length.

    Unitless layout values are intentionally rejected so an agent cannot
    silently confuse points, inches, pixels, and OOXML twips.
    """

    value: float
    unit: Literal["pt", "in", "cm", "mm", "px"]

    @field_validator("value", mode="before")
    @classmethod
    def reject_boolean_value(cls, value: object) -> object:
        if isinstance(value, bool):
            raise ValueError("Length value cannot be a boolean.")
        return value

    def to_points(self) -> float:
        factors = {
            "pt": 1.0,
            "in": 72.0,
            "cm": 72.0 / 2.54,
            "mm": 72.0 / 25.4,
            "px": 72.0 / 96.0,
        }
        return self.value * factors[self.unit]

    def to_css(self) -> str:
        value = f"{self.value:.6f}".rstrip("0").rstrip(".")
        return f"{value}{self.unit}"


class LineSpacing(StrictModel):
    """Paragraph line-spacing with an unambiguous rule/value pair."""

    rule: Literal["multiple", "exact", "at_least"] = "multiple"
    value: float | Length = 1.0

    @model_validator(mode="after")
    def validate_value_for_rule(self) -> "LineSpacing":
        if self.rule == "multiple":
            if isinstance(self.value, Length):
                raise ValueError("Multiple line spacing requires a numeric multiplier.")
            if self.value <= 0:
                raise ValueError("Line-spacing multiplier must be greater than zero.")
        else:
            if not isinstance(self.value, Length):
                raise ValueError(f"{self.rule} line spacing requires an explicit length.")
            if self.value.to_points() <= 0:
                raise ValueError("Line-spacing length must be greater than zero.")
        return self


HexColor = Annotated[str, StringConstraints(pattern=r"^#[0-9A-Fa-f]{6}$")]


class ParagraphStyle(StrictModel):
    """Direct paragraph formatting, independent of named native styles."""

    alignment: Literal["left", "center", "right", "justify", "distribute"] | None = None
    spacing_before: Length | None = None
    spacing_after: Length | None = None
    line_spacing: LineSpacing | None = None
    indent_left: Length | None = None
    indent_right: Length | None = None
    first_line_indent: Length | None = None
    hanging_indent: Length | None = None
    keep_with_next: bool | None = None
    keep_together: bool | None = None
    page_break_before: bool | None = None
    widow_control: bool | None = None

    @model_validator(mode="after")
    def validate_indentation(self) -> "ParagraphStyle":
        if self.first_line_indent is not None and self.hanging_indent is not None:
            raise ValueError("first_line_indent and hanging_indent are mutually exclusive.")
        non_negative = (
            "spacing_before",
            "spacing_after",
            "indent_left",
            "indent_right",
            "first_line_indent",
            "hanging_indent",
        )
        if any(
            (value := getattr(self, field_name)) is not None and value.to_points() < 0
            for field_name in non_negative
        ):
            raise ValueError("Paragraph spacing and indentation cannot be negative.")
        return self


class TextStyle(StrictModel):
    """Direct character formatting that can be lowered to native run properties."""

    font_family: str | None = Field(default=None, min_length=1, max_length=255)
    font_family_east_asia: str | None = Field(default=None, min_length=1, max_length=255)
    font_size: Length | None = None
    color: HexColor | None = None
    background_color: HexColor | None = None
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strike: bool | None = None
    small_caps: bool | None = None
    all_caps: bool | None = None
    letter_spacing: Length | None = None
    baseline: Literal["normal", "superscript", "subscript"] | None = None

    @field_validator("color", "background_color")
    @classmethod
    def normalize_color(cls, value: str | None) -> str | None:
        return value.upper() if value is not None else None

    @field_validator("font_size")
    @classmethod
    def validate_font_size(cls, value: Length | None) -> Length | None:
        if value is not None and value.to_points() <= 0:
            raise ValueError("font_size must be greater than zero.")
        return value


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
    fingerprint: (
        Annotated[
            str,
            StringConstraints(pattern=r"^sha256:[a-fA-F0-9]{64}$"),
        ]
        | None
    ) = None

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
    style: TextStyle | None = None

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
    text: str | None = None
    content: list[TextSpan] = Field(default_factory=list)
    paragraph_style: ParagraphStyle | None = None
    text_style: TextStyle | None = None

    @model_validator(mode="after")
    def validate_content(self) -> "Heading":
        if self.text is None and not self.content:
            raise ValueError("A heading must include text or content.")
        if self.text is not None and self.content:
            raise ValueError("A heading cannot include both text and content.")
        return self

    @property
    def plain_text(self) -> str:
        if self.text is not None:
            return self.text
        return "".join(span.text for span in self.content)


class Paragraph(NodeBase):
    id: NodeId = Field(default_factory=lambda: new_id("para"))
    type: Literal["paragraph"] = "paragraph"
    text: str | None = None
    content: list[TextSpan] = Field(default_factory=list)
    paragraph_style: ParagraphStyle | None = None
    text_style: TextStyle | None = None

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
    capabilities: list[Literal["inspect", "move", "delete", "render"]] = Field(
        default_factory=lambda: ["inspect", "render"]
    )
    editable: Literal[False] = False


Block = Annotated[
    Heading | Paragraph | BulletList | OrderedList | Table | PageBreak | OpaqueBlock,
    Field(discriminator="type"),
]


class AiOfficeDocumentSpec(StrictModel):
    schema_url: Literal["https://schemas.aioffice.dev/spec/draft/0.2/document.json"] = Field(
        default=DOCUMENT_SCHEMA_URL,
        alias="$schema",
    )
    spec_version: Literal[
        "0.2-draft.1",
        "0.2-draft.2",
        "0.2-draft.3",
    ] = SPEC_VERSION
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
