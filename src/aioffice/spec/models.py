"""Strict Pydantic models for the AiOffice Document Spec 0.2 draft."""

from __future__ import annotations

import hashlib
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

SPEC_VERSION = "0.2-draft.36"
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
StyleId = Annotated[
    str,
    StringConstraints(
        min_length=1,
        max_length=255,
        pattern=r"^[^\x00-\x1F\x7F]+$",
    ),
]
SemanticStyleRole = Literal[
    "body",
    "title",
    "subtitle",
    "heading",
    "quote",
    "caption",
    "code",
    "list",
    "custom",
]


class BorderLine(StrictModel):
    """One explicit border edge with OOXML-compatible physical constraints."""

    style: Literal[
        "none",
        "single",
        "double",
        "dotted",
        "dashed",
        "thick",
    ]
    width: Length | None = None
    color: HexColor | Literal["auto"] = "auto"
    space: Length | None = None

    @field_validator("color")
    @classmethod
    def normalize_color(cls, value: str) -> str:
        return value.upper() if value.startswith("#") else value

    @model_validator(mode="after")
    def validate_border(self) -> "BorderLine":
        if self.style == "none":
            if self.width is not None or self.space is not None:
                raise ValueError(
                    "A none border cannot include width or space."
                )
            return self
        if self.width is None:
            raise ValueError("A visible border requires an explicit width.")
        width_points = self.width.to_points()
        if width_points < 0.25 or width_points > 12:
            raise ValueError(
                "Border width must be between 0.25pt and 12pt."
            )
        if self.space is not None:
            space_points = self.space.to_points()
            if space_points < 0 or space_points > 31:
                raise ValueError(
                    "Border space must be between 0pt and 31pt."
                )
        return self


class ParagraphBorders(StrictModel):
    """Four direct paragraph edges within the supported OOXML subset."""

    top: BorderLine | None = None
    right: BorderLine | None = None
    bottom: BorderLine | None = None
    left: BorderLine | None = None


class ParagraphStyle(StrictModel):
    """Direct paragraph formatting, independent of named native styles."""

    alignment: Literal["left", "center", "right", "justify", "distribute"] | None = None
    background_color: HexColor | None = None
    borders: ParagraphBorders | None = None
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
    outline_level: int | None = Field(default=None, ge=1, le=9, strict=True)

    @field_validator("background_color")
    @classmethod
    def normalize_background_color(
        cls,
        value: str | None,
    ) -> str | None:
        return value.upper() if value is not None else None

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


class DocumentDefaults(StrictModel):
    """Formatting inherited before named styles and direct node formatting."""

    paragraph_style: ParagraphStyle | None = None
    text_style: TextStyle | None = None


class PageSize(StrictModel):
    """A standard paper preset or exact custom physical dimensions."""

    preset: Literal[
        "letter",
        "legal",
        "executive",
        "a3",
        "a4",
        "tabloid",
        "custom",
    ] = "letter"
    orientation: Literal["portrait", "landscape"] = "portrait"
    width: Length | None = None
    height: Length | None = None

    @model_validator(mode="after")
    def validate_dimensions(self) -> "PageSize":
        if self.preset == "custom":
            if self.width is None or self.height is None:
                raise ValueError("Custom page size requires width and height.")
            if self.width.to_points() <= 0 or self.height.to_points() <= 0:
                raise ValueError("Custom page width and height must be positive.")
            if self.width.to_points() > 1584 or self.height.to_points() > 1584:
                raise ValueError("Custom page dimensions cannot exceed 22 inches.")
        elif self.width is not None or self.height is not None:
            raise ValueError("Standard page presets cannot include custom width or height.")
        return self

    def dimensions_points(self) -> tuple[float, float]:
        if self.preset == "custom":
            assert self.width is not None
            assert self.height is not None
            width = self.width.to_points()
            height = self.height.to_points()
        else:
            width, height = {
                "letter": (612.0, 792.0),
                "legal": (612.0, 1008.0),
                "executive": (522.0, 756.0),
                "a3": (841.9, 1190.55),
                "a4": (595.3, 841.9),
                "tabloid": (792.0, 1224.0),
            }[self.preset]
        if self.orientation == "landscape" and width < height:
            return height, width
        if self.orientation == "portrait" and width > height:
            return height, width
        return width, height


class SectionColumn(StrictModel):
    width: Length
    space_after: Length = Field(
        default_factory=lambda: Length(value=36, unit="pt")
    )

    @model_validator(mode="after")
    def validate_column(self) -> "SectionColumn":
        if self.width.to_points() <= 0:
            raise ValueError("Section column width must be positive.")
        if self.space_after.to_points() < 0:
            raise ValueError("Section column spacing cannot be negative.")
        return self


class ColumnLayout(StrictModel):
    count: int = Field(default=1, ge=1, le=45, strict=True)
    equal_width: bool = True
    spacing: Length = Field(default_factory=lambda: Length(value=36, unit="pt"))
    separator: bool = False
    columns: list[SectionColumn] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_columns(self) -> "ColumnLayout":
        if self.spacing.to_points() < 0:
            raise ValueError("Column spacing cannot be negative.")
        if self.equal_width and self.columns:
            raise ValueError("Equal-width columns cannot include explicit column widths.")
        if not self.equal_width and len(self.columns) != self.count:
            raise ValueError(
                "Unequal-width column layout requires one definition per column."
            )
        return self


class SectionLayout(StrictModel):
    """Supported direct properties of one WordprocessingML ``w:sectPr``."""

    start_type: Literal[
        "continuous",
        "next_page",
        "even_page",
        "odd_page",
        "next_column",
    ] | None = None
    page_size: PageSize | None = None
    margin_top: Length | None = None
    margin_right: Length | None = None
    margin_bottom: Length | None = None
    margin_left: Length | None = None
    gutter: Length | None = None
    header_distance: Length | None = None
    footer_distance: Length | None = None
    columns: ColumnLayout | None = None
    vertical_alignment: Literal["top", "center", "both", "bottom"] | None = None
    different_first_page: bool | None = None
    page_number_start: int | None = Field(default=None, ge=0, strict=True)
    page_number_format: Literal[
        "decimal",
        "upper_roman",
        "lower_roman",
        "upper_letter",
        "lower_letter",
    ] | None = None

    @model_validator(mode="after")
    def validate_page_geometry(self) -> "SectionLayout":
        non_negative = (
            "margin_top",
            "margin_right",
            "margin_bottom",
            "margin_left",
            "gutter",
            "header_distance",
            "footer_distance",
        )
        if any(
            (value := getattr(self, field_name)) is not None
            and value.to_points() < 0
            for field_name in non_negative
            if field_name not in {"margin_top", "margin_bottom"}
        ):
            raise ValueError(
                "Horizontal margins, gutter, and header/footer distances "
                "cannot be negative."
            )
        if self.page_size is not None:
            width, height = self.page_size.dimensions_points()
            horizontal = sum(
                value.to_points()
                for value in (self.margin_left, self.margin_right, self.gutter)
                if value is not None
            )
            vertical = sum(
                value.to_points()
                for value in (self.margin_top, self.margin_bottom)
                if value is not None
            )
            if horizontal >= width:
                raise ValueError("Horizontal margins and gutter must leave positive page width.")
            if vertical >= height:
                raise ValueError("Vertical margins must leave positive page height.")
        return self


def _default_section_layout() -> SectionLayout:
    return SectionLayout(
        page_size=PageSize(),
        margin_top=Length(value=72, unit="pt"),
        margin_right=Length(value=72, unit="pt"),
        margin_bottom=Length(value=72, unit="pt"),
        margin_left=Length(value=72, unit="pt"),
        gutter=Length(value=0, unit="pt"),
        header_distance=Length(value=36, unit="pt"),
        footer_distance=Length(value=36, unit="pt"),
        columns=ColumnLayout(),
        vertical_alignment="top",
        different_first_page=False,
    )


class NamedStyle(StrictModel):
    """An AI-addressable paragraph style with an explicit inheritance contract."""

    id: StyleId
    name: str = Field(min_length=1, max_length=255)
    kind: Literal["paragraph"] = "paragraph"
    semantic_role: SemanticStyleRole = "custom"
    heading_level: int | None = Field(default=None, ge=1, le=9, strict=True)
    based_on: StyleId | None = None
    next_style: StyleId | None = None
    paragraph_style: ParagraphStyle | None = None
    text_style: TextStyle | None = None
    quick_style: bool | None = None
    hidden: bool | None = None
    metadata: dict[str, Any] = Field(default_factory=dict)

    @model_validator(mode="after")
    def validate_semantic_role(self) -> "NamedStyle":
        if self.semantic_role == "heading" and self.heading_level is None:
            raise ValueError("A heading named style must include heading_level.")
        if self.semantic_role != "heading" and self.heading_level is not None:
            raise ValueError("heading_level is only valid for semantic_role='heading'.")
        outline_level = (
            self.paragraph_style.outline_level
            if self.paragraph_style is not None
            else None
        )
        if outline_level is not None and (
            self.semantic_role != "heading" or self.heading_level != outline_level
        ):
            raise ValueError(
                "A named style outline_level must match its semantic heading_level."
            )
        if self.based_on == self.id:
            raise ValueError("A named style cannot be based on itself.")
        return self


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
    media_type: Annotated[
        str,
        StringConstraints(
            min_length=3,
            max_length=255,
            pattern=(
                r"^[A-Za-z0-9][A-Za-z0-9!#$&^_.+-]*/"
                r"[A-Za-z0-9][A-Za-z0-9!#$&^_.+-]*$"
            ),
        ),
    ]
    filename: str | None = None
    size_bytes: int | None = Field(default=None, ge=0, strict=True)

    @field_validator("sha256")
    @classmethod
    def normalize_sha256(cls, value: str) -> str:
        return value.lower()

    @field_validator("filename")
    @classmethod
    def validate_filename(cls, value: str | None) -> str | None:
        if (
            value is not None
            and (
                not value
                or value in {".", ".."}
                or "/" in value
                or "\\" in value
                or "\x00" in value
            )
        ):
            raise ValueError("Asset filename must be one safe basename.")
        return value


class NativeRef(StrictModel):
    format: Literal["docx", "xlsx", "pptx"]
    part_uri: Annotated[str, StringConstraints(pattern=r"^/[^\\\x00]*$")]
    native_kind: str
    element_index: int | None = Field(default=None, ge=0)
    element_indices: list[int] = Field(default_factory=list)
    sub_index: int | None = Field(default=None, ge=0)
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
        if self.sub_index is not None and self.element_index is None:
            raise ValueError("sub_index requires an element_index.")
        return self


class HeaderFooterBindings(StrictModel):
    """Explicit section bindings; a missing slot inherits from the previous section."""

    header_default: NodeId | None = None
    header_first: NodeId | None = None
    header_even: NodeId | None = None
    footer_default: NodeId | None = None
    footer_first: NodeId | None = None
    footer_even: NodeId | None = None


class DocumentSettings(StrictModel):
    """Document-wide layout switches that cannot be scoped to one section."""

    even_and_odd_headers: bool | None = None
    update_fields_on_open: bool | None = None


class DocumentSection(StrictModel):
    """A stable semantic section anchored at its first content node."""

    id: NodeId = Field(default_factory=lambda: new_id("section"))
    type: Literal["section"] = "section"
    start_at: NodeId | None = None
    layout: SectionLayout = Field(default_factory=_default_section_layout)
    header_footer: HeaderFooterBindings | None = None
    source_ref: NativeRef | str | None = None
    metadata: dict[str, Any] = Field(default_factory=dict)
    revision_added: int = Field(default=1, ge=1)
    revision_updated: int = Field(default=1, ge=1)


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


class DocumentField(StrictModel):
    """An AI-addressable dynamic Word field with a non-authoritative cached result."""

    id: NodeId = Field(default_factory=lambda: new_id("field"))
    type: Literal["field"] = "field"
    kind: Literal[
        "page_number",
        "page_count",
        "section_number",
        "section_page_count",
        "native",
    ]
    number_format: Literal[
        "decimal",
        "upper_roman",
        "lower_roman",
        "upper_letter",
        "lower_letter",
    ] | None = None
    cached_result: str | None = None
    instruction: str | None = None
    style: TextStyle | None = None
    editable: bool = True
    source_ref: NativeRef | str | None = None
    metadata: dict[str, Any] = Field(default_factory=dict)
    revision_added: int = Field(default=1, ge=1)
    revision_updated: int = Field(default=1, ge=1)

    @model_validator(mode="after")
    def validate_native_field(self) -> "DocumentField":
        if self.kind == "native":
            if not self.instruction:
                raise ValueError("A native field requires its preserved instruction.")
            if self.editable:
                raise ValueError("A native field must be read-only.")
            if self.number_format is not None:
                raise ValueError("A native field cannot claim a normalized number format.")
        elif self.instruction is not None:
            raise ValueError(
                "Normalized fields keep native instructions in metadata, not instruction."
            )
        elif not self.editable:
            raise ValueError("A normalized field must remain editable.")
        return self

    @property
    def display_text(self) -> str:
        if self.cached_result is not None:
            return self.cached_result
        if self.kind == "native":
            return "⟦field⟧"
        return "1"


InlineContent = TextSpan | DocumentField


class Heading(NodeBase):
    id: NodeId = Field(default_factory=lambda: new_id("heading"))
    type: Literal["heading"] = "heading"
    level: int = Field(default=1, ge=1, le=6)
    text: str | None = None
    content: list[InlineContent] = Field(default_factory=list)
    style_ref: StyleId | None = None
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
        return "".join(
            span.text if isinstance(span, TextSpan) else span.display_text
            for span in self.content
        )


class Paragraph(NodeBase):
    id: NodeId = Field(default_factory=lambda: new_id("para"))
    type: Literal["paragraph"] = "paragraph"
    text: str | None = None
    content: list[InlineContent] = Field(default_factory=list)
    style_ref: StyleId | None = None
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
        return "".join(
            span.text if isinstance(span, TextSpan) else span.display_text
            for span in self.content
        )


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


class TableWidth(StrictModel):
    """A preferred table width with explicit auto, percent, or physical units."""

    mode: Literal["auto", "percent", "exact"] = "auto"
    value: float | Length | None = None

    @field_validator("value", mode="before")
    @classmethod
    def reject_boolean_value(cls, value: object) -> object:
        if isinstance(value, bool):
            raise ValueError("Table width cannot be a boolean.")
        return value

    @model_validator(mode="after")
    def validate_mode_value(self) -> "TableWidth":
        if self.mode == "auto":
            if self.value is not None:
                raise ValueError("Auto table width cannot include a value.")
        elif self.mode == "percent":
            if isinstance(self.value, Length) or self.value is None:
                raise ValueError("Percent table width requires a numeric value.")
            if self.value <= 0 or self.value > 100:
                raise ValueError("Percent table width must be greater than 0 and at most 100.")
        else:
            if not isinstance(self.value, Length):
                raise ValueError("Exact table width requires an explicit physical length.")
            if self.value.to_points() <= 0:
                raise ValueError("Exact table width must be greater than zero.")
        return self


class TableBorders(StrictModel):
    """Table perimeter and internal grid edges."""

    top: BorderLine | None = None
    right: BorderLine | None = None
    bottom: BorderLine | None = None
    left: BorderLine | None = None
    inside_horizontal: BorderLine | None = None
    inside_vertical: BorderLine | None = None


class TableCellBorders(StrictModel):
    """Four direct cell edges that override conflicting table borders."""

    top: BorderLine | None = None
    right: BorderLine | None = None
    bottom: BorderLine | None = None
    left: BorderLine | None = None


class TableLayout(StrictModel):
    """Supported table-wide geometry independent of cell data semantics."""

    style_ref: StyleId | None = None
    preferred_width: TableWidth | None = None
    alignment: Literal["left", "center", "right"] | None = None
    algorithm: Literal["autofit", "fixed"] | None = None
    indent: Length | None = None
    cell_spacing: Length | None = None
    cell_margin_top: Length | None = None
    cell_margin_right: Length | None = None
    cell_margin_bottom: Length | None = None
    cell_margin_left: Length | None = None
    borders: TableBorders | None = None
    repeat_header: bool | None = None

    @model_validator(mode="after")
    def validate_lengths(self) -> "TableLayout":
        for field_name in (
            "indent",
            "cell_spacing",
            "cell_margin_top",
            "cell_margin_right",
            "cell_margin_bottom",
            "cell_margin_left",
        ):
            value = getattr(self, field_name)
            if value is not None and value.to_points() < 0:
                raise ValueError(f"{field_name} cannot be negative.")
        return self


class TableColumn(StrictModel):
    id: NodeId = Field(default_factory=lambda: new_id("column"))
    type: Literal["table_column"] = "table_column"
    key: Annotated[str, StringConstraints(pattern=r"^[A-Za-z_][A-Za-z0-9_.-]*$")]
    title: str
    data_type: Literal["text", "number", "integer", "boolean", "date", "enum"] = "text"
    width: Length | None = None
    source_ref: NativeRef | str | None = None
    revision_added: int = Field(default=1, ge=1)
    revision_updated: int = Field(default=1, ge=1)

    @field_validator("width")
    @classmethod
    def validate_width(cls, value: Length | None) -> Length | None:
        if value is not None and value.to_points() <= 0:
            raise ValueError("Table column width must be greater than zero.")
        return value


class TableCellFormat(StrictModel):
    """Cell-local presentation that can be patched without rebuilding content."""

    vertical_alignment: Literal["top", "center", "bottom"] | None = None
    no_wrap: bool | None = None
    fit_text: bool | None = None
    background_color: HexColor | None = None
    borders: TableCellBorders | None = None
    margin_top: Length | None = None
    margin_right: Length | None = None
    margin_bottom: Length | None = None
    margin_left: Length | None = None

    @field_validator("background_color")
    @classmethod
    def normalize_color(cls, value: str | None) -> str | None:
        return value.upper() if value is not None else None

    @model_validator(mode="after")
    def validate_margins(self) -> "TableCellFormat":
        for field_name in (
            "margin_top",
            "margin_right",
            "margin_bottom",
            "margin_left",
        ):
            value = getattr(self, field_name)
            if value is not None and value.to_points() < 0:
                raise ValueError(f"{field_name} cannot be negative.")
        return self


class TableCell(StrictModel):
    """One logical cell anchored to a semantic column key."""

    id: NodeId = Field(default_factory=lambda: new_id("cell"))
    type: Literal["table_cell"] = "table_cell"
    column_key: Annotated[
        str,
        StringConstraints(pattern=r"^[A-Za-z_][A-Za-z0-9_.-]*$"),
    ]
    value: Any = None
    content: list[Paragraph] = Field(default_factory=list)
    column_span: int = Field(default=1, ge=1, le=63, strict=True)
    row_span: int = Field(default=1, ge=1, le=32767, strict=True)
    format: TableCellFormat = Field(default_factory=TableCellFormat)
    source_ref: NativeRef | str | None = None
    metadata: dict[str, Any] = Field(default_factory=dict)
    revision_added: int = Field(default=1, ge=1)
    revision_updated: int = Field(default=1, ge=1)

    @model_validator(mode="after")
    def validate_content(self) -> "TableCell":
        if self.content and self.value is not None:
            raise ValueError(
                "A table cell cannot include both scalar value and rich content."
            )
        if any(
            isinstance(inline, DocumentField)
            for paragraph in self.content
            for inline in paragraph.content
        ):
            raise ValueError(
                "Dynamic fields inside table cells are not supported yet."
            )
        return self

    @property
    def plain_text(self) -> str:
        if self.content:
            return "\n".join(paragraph.plain_text for paragraph in self.content)
        if self.value is None:
            return ""
        if isinstance(self.value, bool):
            return "true" if self.value else "false"
        return str(self.value)


class TableRow(StrictModel):
    id: NodeId = Field(default_factory=lambda: new_id("row"))
    type: Literal["table_row"] = "table_row"
    cells: list[TableCell] = Field(default_factory=list)
    allow_break_across_pages: bool | None = None
    height: Length | None = None
    height_rule: Literal["at_least", "exact"] | None = None
    source_ref: NativeRef | str | None = None
    revision_added: int = Field(default=1, ge=1)
    revision_updated: int = Field(default=1, ge=1)

    @model_validator(mode="before")
    @classmethod
    def migrate_legacy_values(cls, value: object) -> object:
        if not isinstance(value, dict) or "values" not in value:
            return value
        payload = dict(value)
        values = payload.pop("values")
        if "cells" in payload:
            raise ValueError(
                "A table row cannot include both legacy values and cells."
            )
        if not isinstance(values, dict):
            raise ValueError("Legacy table row values must be an object.")
        row_id = payload.get("id")
        if not isinstance(row_id, str):
            row_id = new_id("row")
            payload["id"] = row_id
        payload["cells"] = [
            {
                "id": (
                    "cell_"
                    + hashlib.sha256(
                        f"{row_id}:{key}".encode()
                    ).hexdigest()[:24]
                ),
                "type": "table_cell",
                "column_key": key,
                "value": cell_value,
            }
            for key, cell_value in values.items()
        ]
        return payload

    @model_validator(mode="after")
    def validate_height(self) -> "TableRow":
        if self.height is not None and self.height.to_points() <= 0:
            raise ValueError("Table row height must be greater than zero.")
        if self.height_rule is not None and self.height is None:
            raise ValueError("Table row height_rule requires height.")
        return self


class Table(NodeBase):
    id: NodeId = Field(default_factory=lambda: new_id("table"))
    type: Literal["table"] = "table"
    columns: list[TableColumn] = Field(min_length=1)
    rows: list[TableRow] = Field(default_factory=list)
    layout: TableLayout = Field(default_factory=TableLayout)


class PageBreak(NodeBase):
    id: NodeId = Field(default_factory=lambda: new_id("break"))
    type: Literal["page_break"] = "page_break"


class ImageCrop(StrictModel):
    """Rectangular source crop in percentage points of the original image."""

    left: float = Field(default=0, ge=0, lt=100)
    top: float = Field(default=0, ge=0, lt=100)
    right: float = Field(default=0, ge=0, lt=100)
    bottom: float = Field(default=0, ge=0, lt=100)

    @field_validator("left", "top", "right", "bottom", mode="before")
    @classmethod
    def normalize_percentage(cls, value: object) -> object:
        if isinstance(value, bool):
            raise ValueError("Image crop percentages cannot be booleans.")
        if isinstance(value, (int, float)):
            return round(float(value), 3)
        return value

    @model_validator(mode="after")
    def validate_visible_area(self) -> "ImageCrop":
        if not any(
            (
                self.left,
                self.top,
                self.right,
                self.bottom,
            )
        ):
            raise ValueError(
                "Image crop must remove at least one non-zero edge; "
                "clear crop to show the complete source."
            )
        if self.left + self.right >= 100:
            raise ValueError(
                "Image left and right crop must leave visible source width."
            )
        if self.top + self.bottom >= 100:
            raise ValueError(
                "Image top and bottom crop must leave visible source height."
            )
        return self


class FloatingImageHorizontalPosition(StrictModel):
    """Horizontal position of a floating image relative to a Word layout frame."""

    relative_to: Literal[
        "character",
        "column",
        "inside_margin",
        "left_margin",
        "margin",
        "outside_margin",
        "page",
        "right_margin",
    ]
    offset: Length

    @model_validator(mode="after")
    def validate_offset(self) -> "FloatingImageHorizontalPosition":
        emu = round(self.offset.to_points() * 12_700)
        if emu < -(2**63) or emu > 2**63 - 1:
            raise ValueError(
                "Floating image horizontal offset must fit an OOXML Int64 EMU."
            )
        return self


class FloatingImageVerticalPosition(StrictModel):
    """Vertical position of a floating image relative to a Word layout frame."""

    relative_to: Literal[
        "bottom_margin",
        "inside_margin",
        "line",
        "margin",
        "outside_margin",
        "page",
        "paragraph",
        "top_margin",
    ]
    offset: Length

    @model_validator(mode="after")
    def validate_offset(self) -> "FloatingImageVerticalPosition":
        emu = round(self.offset.to_points() * 12_700)
        if emu < -(2**63) or emu > 2**63 - 1:
            raise ValueError(
                "Floating image vertical offset must fit an OOXML Int64 EMU."
            )
        return self


class FloatingImageTextWrap(StrictModel):
    """Conservative rectangular text wrapping around a floating image."""

    mode: Literal["square"] = "square"
    side: Literal["both_sides", "largest", "left", "right"]
    distance_top: Length
    distance_right: Length
    distance_bottom: Length
    distance_left: Length

    @model_validator(mode="after")
    def validate_distances(self) -> "FloatingImageTextWrap":
        for field_name in (
            "distance_top",
            "distance_right",
            "distance_bottom",
            "distance_left",
        ):
            value = getattr(self, field_name)
            emu = round(value.to_points() * 12_700)
            if emu < 0 or emu > 2**32 - 1:
                raise ValueError(
                    f"Floating image {field_name} must fit a non-negative "
                    "OOXML UInt32 EMU."
                )
        return self


class FloatingImageLayout(StrictModel):
    """Read-only, lossless evidence for one conservative Word floating anchor."""

    horizontal: FloatingImageHorizontalPosition
    vertical: FloatingImageVerticalPosition
    wrap: FloatingImageTextWrap
    relative_height: int = Field(ge=0, le=2**32 - 1, strict=True)
    behind_text: bool = Field(strict=True)
    locked: bool = Field(strict=True)
    layout_in_cell: bool = Field(strict=True)
    allow_overlap: bool = Field(strict=True)


class FloatingImageLayoutUpdate(StrictModel):
    """Selectable fields accepted by ``image.anchor.update``."""

    horizontal: FloatingImageHorizontalPosition | None = None
    vertical: FloatingImageVerticalPosition | None = None
    wrap: FloatingImageTextWrap | None = None
    relative_height: int | None = Field(
        default=None,
        ge=0,
        le=2**32 - 1,
        strict=True,
    )
    behind_text: bool | None = Field(default=None, strict=True)
    locked: bool | None = Field(default=None, strict=True)
    layout_in_cell: bool | None = Field(default=None, strict=True)
    allow_overlap: bool | None = Field(default=None, strict=True)

    @model_validator(mode="after")
    def validate_changes(self) -> "FloatingImageLayoutUpdate":
        if not self.model_fields_set:
            raise ValueError(
                "Floating image layout update requires at least one field."
            )
        return self


class ImageBlock(NodeBase):
    """One AI-addressable image occurrence backed by a native binary asset."""

    id: NodeId = Field(default_factory=lambda: new_id("image"))
    type: Literal["image"] = "image"
    asset_id: NodeId
    placement: Literal["inline", "floating"] = "inline"
    width: Length
    height: Length
    crop: ImageCrop | None = None
    floating: FloatingImageLayout | None = None
    name: str | None = None
    alt_text: str | None = None
    title: str | None = None
    style_ref: StyleId | None = None
    paragraph_style: ParagraphStyle | None = None
    capabilities: list[
        Literal["inspect", "extract", "delete", "render"]
    ] = Field(
        default_factory=lambda: [
            "inspect",
            "extract",
            "delete",
            "render",
        ]
    )
    editable: Literal[False] = False

    @model_validator(mode="after")
    def validate_size(self) -> "ImageBlock":
        if self.width.to_points() <= 0 or self.height.to_points() <= 0:
            raise ValueError("Image width and height must be greater than zero.")
        if (self.placement == "floating") != (self.floating is not None):
            raise ValueError(
                "Floating image placement requires floating layout evidence, "
                "and inline placement forbids it."
            )
        if self.capabilities != [
            "inspect",
            "extract",
            "delete",
            "render",
        ]:
            raise ValueError(
                "A native image must declare its exact supported capabilities."
            )
        return self


class HeaderFooterImageBlock(ImageBlock):
    """One supported native image occurrence inside a reusable page story."""

    capabilities: list[
        Literal["inspect", "extract", "delete", "render"]
    ] = Field(
        default_factory=lambda: [
            "inspect",
            "extract",
            "render",
        ],
        json_schema_extra={
            "prefixItems": [
                {"const": "inspect"},
                {"const": "extract"},
                {"const": "render"},
            ],
            "items": False,
            "minItems": 3,
            "maxItems": 3,
        },
    )

    @model_validator(mode="after")
    def validate_size(self) -> "HeaderFooterImageBlock":
        if self.width.to_points() <= 0 or self.height.to_points() <= 0:
            raise ValueError("Image width and height must be greater than zero.")
        if (self.placement == "floating") != (self.floating is not None):
            raise ValueError(
                "Floating image placement requires floating layout evidence, "
                "and inline placement forbids it."
            )
        if self.capabilities != [
            "inspect",
            "extract",
            "render",
        ]:
            raise ValueError(
                "A header/footer image must declare its exact supported "
                "capabilities."
            )
        return self


class ImageUpdate(StrictModel):
    """Fields accepted by the selective ``image.update`` operation."""

    width: Length | None = None
    height: Length | None = None
    crop: ImageCrop | None = None
    alt_text: str | None = Field(default=None, min_length=1, max_length=4096)
    title: str | None = Field(default=None, min_length=1, max_length=1024)

    @model_validator(mode="after")
    def validate_values(self) -> "ImageUpdate":
        for field_name in ("width", "height"):
            value = getattr(self, field_name)
            if value is None:
                continue
            emu = round(value.to_points() * 12_700)
            if emu <= 0 or emu > 2**63 - 1:
                raise ValueError(
                    f"Image {field_name} must fit a positive OOXML Int64 EMU."
                )
        for field_name in ("alt_text", "title"):
            value = getattr(self, field_name)
            if value is None:
                continue
            if not value.strip():
                raise ValueError(
                    f"Image {field_name} cannot be blank; clear it explicitly."
                )
            if any(
                ord(character) < 0x20
                and character not in {"\t", "\n", "\r"}
                for character in value
            ):
                raise ValueError(
                    f"Image {field_name} contains an invalid XML character."
                )
        return self


class ImageInsert(StrictModel):
    """Fields accepted by the out-of-band native image insertion API."""

    id: NodeId = Field(default_factory=lambda: new_id("image"))
    placement: Literal["inline", "floating"] = "inline"
    floating: FloatingImageLayout | None = None
    width: Length
    height: Length
    name: str | None = Field(default=None, min_length=1, max_length=1024)
    alt_text: str = Field(min_length=1, max_length=4096)
    title: str | None = Field(default=None, min_length=1, max_length=1024)
    paragraph_style: ParagraphStyle | None = None

    @model_validator(mode="after")
    def validate_values(self) -> "ImageInsert":
        if (self.placement == "floating") != (self.floating is not None):
            raise ValueError(
                "Floating image insertion requires floating layout, and "
                "inline insertion forbids it."
            )
        for field_name in ("width", "height"):
            value = getattr(self, field_name)
            emu = round(value.to_points() * 12_700)
            if emu <= 0 or emu > 2**63 - 1:
                raise ValueError(
                    f"Image {field_name} must fit a positive OOXML Int64 EMU."
                )
        for field_name in ("name", "alt_text", "title"):
            value = getattr(self, field_name)
            if value is None:
                continue
            if not value.strip():
                raise ValueError(f"Image {field_name} cannot be blank.")
            if any(
                ord(character) < 0x20
                and character not in {"\t", "\n", "\r"}
                for character in value
            ):
                raise ValueError(
                    f"Image {field_name} contains an invalid XML character."
                )
        return self


class OpaqueBlock(NodeBase):
    id: NodeId = Field(default_factory=lambda: new_id("opaque"))
    type: Literal["opaque"] = "opaque"
    summary: str
    capabilities: list[Literal["inspect", "move", "delete", "render"]] = Field(
        default_factory=lambda: ["inspect", "render"]
    )
    editable: Literal[False] = False


Block = Annotated[
    Heading
    | Paragraph
    | BulletList
    | OrderedList
    | Table
    | PageBreak
    | ImageBlock
    | OpaqueBlock,
    Field(discriminator="type"),
]


HeaderFooterBlock = Annotated[
    Paragraph | HeaderFooterImageBlock | OpaqueBlock,
    Field(discriminator="type"),
]


class HeaderFooterPart(StrictModel):
    """A reusable header or footer part referenced by one or more sections."""

    id: NodeId = Field(default_factory=lambda: new_id("region"))
    type: Literal["header_footer"] = "header_footer"
    kind: Literal["header", "footer"]
    content: list[HeaderFooterBlock] = Field(default_factory=list)
    source_ref: NativeRef | str | None = None
    metadata: dict[str, Any] = Field(default_factory=dict)
    revision_added: int = Field(default=1, ge=1)
    revision_updated: int = Field(default=1, ge=1)


class AiOfficeDocumentSpec(StrictModel):
    schema_url: Literal["https://schemas.aioffice.dev/spec/draft/0.2/document.json"] = Field(
        default=DOCUMENT_SCHEMA_URL,
        alias="$schema",
    )
    spec_version: Literal[
        "0.2-draft.1",
        "0.2-draft.2",
        "0.2-draft.3",
        "0.2-draft.4",
        "0.2-draft.5",
        "0.2-draft.6",
        "0.2-draft.7",
        "0.2-draft.8",
        "0.2-draft.9",
        "0.2-draft.10",
        "0.2-draft.11",
        "0.2-draft.12",
        "0.2-draft.13",
        "0.2-draft.14",
        "0.2-draft.15",
        "0.2-draft.16",
        "0.2-draft.17",
        "0.2-draft.18",
        "0.2-draft.19",
        "0.2-draft.20",
        "0.2-draft.21",
        "0.2-draft.22",
        "0.2-draft.23",
        "0.2-draft.24",
        "0.2-draft.25",
        "0.2-draft.26",
        "0.2-draft.27",
        "0.2-draft.28",
        "0.2-draft.29",
        "0.2-draft.30",
        "0.2-draft.31",
        "0.2-draft.32",
        "0.2-draft.33",
        "0.2-draft.34",
        "0.2-draft.35",
        "0.2-draft.36",
    ] = SPEC_VERSION
    engine_version: str = __version__
    artifact: ArtifactDescriptor = Field(default_factory=ArtifactDescriptor)
    metadata: DocumentMetadata = Field(default_factory=DocumentMetadata)
    theme: ThemeRef = Field(default_factory=ThemeRef)
    defaults: DocumentDefaults = Field(default_factory=DocumentDefaults)
    settings: DocumentSettings | None = None
    styles: list[NamedStyle] = Field(default_factory=list)
    sections: list[DocumentSection] = Field(
        default_factory=lambda: [
            DocumentSection(id="section_default"),
        ],
        min_length=1,
    )
    header_footers: list[HeaderFooterPart] = Field(default_factory=list)
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

    @model_validator(mode="after")
    def validate_named_styles(self) -> "AiOfficeDocumentSpec":
        style_ids = [style.id for style in self.styles]
        if len(style_ids) != len(set(style_ids)):
            raise ValueError("Named style IDs must be unique.")
        section_ids = [section.id for section in self.sections]
        if len(section_ids) != len(set(section_ids)):
            raise ValueError("Document section IDs must be unique.")
        header_footer_ids = [part.id for part in self.header_footers]
        if len(header_footer_ids) != len(set(header_footer_ids)):
            raise ValueError("Header/footer part IDs must be unique.")
        part_kinds = {part.id: part.kind for part in self.header_footers}
        for section in self.sections:
            if section.header_footer is None:
                continue
            for field_name, part_id in section.header_footer.model_dump(
                mode="python"
            ).items():
                if part_id is None:
                    continue
                expected_kind = field_name.split("_", 1)[0]
                if part_id not in part_kinds:
                    raise ValueError(
                        f"Section {section.id!r} references missing "
                        f"header/footer part {part_id!r}."
                    )
                if part_kinds[part_id] != expected_kind:
                    raise ValueError(
                        f"Section {section.id!r} binds {part_id!r} as "
                        f"{expected_kind}, but the part is {part_kinds[part_id]}."
                    )
        return self
