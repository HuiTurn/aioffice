"""Lower semantic operations into minimal mutations of a native DOCX part."""

from __future__ import annotations

from collections.abc import Mapping, Sequence
from copy import deepcopy
import hashlib
from pathlib import PurePosixPath
from typing import Any
from xml.etree import ElementTree as ET

from pydantic import ValidationError

from aioffice.core.errors import NativePackageError
from aioffice.formats.docx import (
    DocxCompileContext,
    append_single_level_numbering_definition,
    compile_list_elements,
    compile_table_element,
    register_table_refs,
)
from aioffice.formats.docx_style import (
    apply_paragraph_mark_text_style,
    apply_paragraph_style,
    apply_text_style,
    patch_paragraph_mark_text_style,
    patch_paragraph_style,
    patch_paragraph_style_ref,
    patch_text_style,
)
from aioffice.formats.docx_named_styles import (
    find_named_style,
    format_named_style,
    upsert_named_style,
)
from aioffice.formats.docx_header_footer import (
    native_ref_for_header_footer_part,
    resolve_relationship_target,
)
from aioffice.formats.docx_fields import (
    FieldStructureError,
    append_complex_field,
    canonical_field_instruction,
    field_match_at,
    native_ref_for_field,
    parse_paragraph_fields,
    patch_field_instruction,
)
from aioffice.formats.docx_images import (
    insert_simple_inline_image_after,
    patch_simple_inline_image,
    replace_simple_inline_image,
    simple_inline_image,
)
from aioffice.formats.docx_section import (
    native_ref_for_section,
    patch_section_layout,
)
from aioffice.formats.docx_tables import (
    native_ref_for_table_cell,
    native_ref_for_table_cell_paragraph,
    native_ref_for_table_column,
    native_ref_for_table_row,
    patch_table_cell_format,
    patch_table_column_width,
    patch_table_layout,
    table_cell_from_ref,
    table_cell_paragraph_from_ref,
)
from aioffice.native import (
    FidelityReport,
    MANIFEST_PART_URI,
    MANIFEST_RELATIONSHIP_TYPE,
    NativePackage,
    build_identity_manifest,
    native_ref_for_part_elements,
    serialize_identity_manifest,
)
from aioffice.native.xml import parse_xml, serialize_xml
from aioffice.operations.text import resolve_text_selection
from aioffice.spec.models import (
    AiOfficeDocumentSpec,
    AssetRef,
    BulletList,
    DocumentField,
    Heading,
    ImageBlock,
    ImageInsert,
    InlineContent,
    NamedStyle,
    NativeRef,
    OrderedList,
    PageBreak,
    Paragraph,
    ParagraphStyle,
    SectionLayout,
    Table,
    TableCell,
    TableCellFormat,
    TableLayout,
    TextSpan,
    TextStyle,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
XML = "http://www.w3.org/XML/1998/namespace"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
HYPERLINK_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)
NUMBERING_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/"
    "relationships/numbering"
)
NUMBERING_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument."
    "wordprocessingml.numbering+xml"
)
RELATIONSHIPS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-package.relationships+xml"
)


def _q(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def _target_id(value: Any) -> str:
    if not isinstance(value, str) or not value:
        raise NativePackageError("Native DOCX patch target must be a node ID.")
    return value[1:] if value.startswith("#") else value


def _relationship_part_uri(source_part: str) -> str:
    source = PurePosixPath(source_part)
    return str(
        source.parent
        / "_rels"
        / f"{source.name}.rels"
    )


def _attach_hyperlink_relationship(
    package: NativePackage,
    *,
    source_part: str,
    target: str,
) -> str:
    relationship_part_uri = _relationship_part_uri(source_part)
    relationship_part_exists = package.has_part(relationship_part_uri)
    relationships = (
        parse_xml(package.get_part(relationship_part_uri))
        if relationship_part_exists
        else ET.Element(_q(REL, "Relationships"))
    )
    if relationships.tag != _q(REL, "Relationships"):
        raise NativePackageError(
            f"DOCX relationship part {relationship_part_uri!r} has "
            "an invalid root."
        )
    relationship_ids = {
        element.get("Id", "")
        for element in relationships.findall(_q(REL, "Relationship"))
    }
    relationship_number = 1
    while (
        f"rIdAiOfficeHyperlink{relationship_number}"
        in relationship_ids
    ):
        relationship_number += 1
    relationship_id = f"rIdAiOfficeHyperlink{relationship_number}"
    ET.SubElement(
        relationships,
        _q(REL, "Relationship"),
        {
            "Id": relationship_id,
            "Type": HYPERLINK_RELATIONSHIP_TYPE,
            "Target": target,
            "TargetMode": "External",
        },
    )
    package.set_part(
        relationship_part_uri,
        serialize_xml(relationships),
        content_type=(
            None
            if relationship_part_exists
            else RELATIONSHIPS_CONTENT_TYPE
        ),
    )
    return relationship_id


def _native_paragraph_anchor(
    container: ET.Element,
    node_id: str,
) -> str:
    existing = {
        value
        for paragraph in container.iter(_q(W, "p"))
        if (value := paragraph.get(_q(W14, "paraId"))) is not None
    }
    return _fresh_native_paragraph_anchor(existing, node_id)


def _fresh_native_paragraph_anchor(
    existing: set[str],
    node_id: str,
) -> str:
    ordinal = 0
    while True:
        candidate = hashlib.sha256(
            f"{node_id}:{ordinal}".encode()
        ).hexdigest()[:8].upper()
        if candidate == "00000000":
            candidate = "00000001"
        if candidate not in existing:
            return candidate
        ordinal += 1


def _numbering_decimal_ids(
    elements: Sequence[ET.Element],
    *,
    attribute_name: str,
    label: str,
) -> set[int]:
    values: list[int] = []
    for element in elements:
        raw_value = element.get(_q(W, attribute_name))
        if (
            raw_value is None
            or not raw_value.isascii()
            or not raw_value.isdecimal()
        ):
            raise NativePackageError(
                f"Native DOCX numbering contains an invalid {label}."
            )
        values.append(int(raw_value))
    if len(values) != len(set(values)):
        raise NativePackageError(
            f"Native DOCX numbering contains duplicate {label} values."
        )
    return set(values)


def _fresh_decimal_id(existing: set[int], *, start: int) -> int:
    candidate = start
    while candidate in existing:
        candidate += 1
    return candidate


def _validate_numbering_child_order(
    numbering: ET.Element,
) -> None:
    phases = {
        _q(W, "numPicBullet"): 0,
        _q(W, "abstractNum"): 1,
        _q(W, "num"): 2,
        _q(W, "numIdMacAtCleanup"): 3,
    }
    last_phase = -1
    cleanup_count = 0
    for child in list(numbering):
        phase = phases.get(child.tag)
        if phase is None:
            continue
        if phase < last_phase:
            raise NativePackageError(
                "Native DOCX numbering children are not in OOXML "
                "schema order."
            )
        last_phase = phase
        if child.tag == _q(W, "numIdMacAtCleanup"):
            cleanup_count += 1
    if cleanup_count > 1:
        raise NativePackageError(
            "Native DOCX numbering contains duplicate "
            "numIdMacAtCleanup elements."
        )


def _ensure_numbering_relationship(
    package: NativePackage,
) -> None:
    source_part = "/word/document.xml"
    relationship_part_uri = _relationship_part_uri(source_part)
    relationship_part_exists = package.has_part(
        relationship_part_uri
    )
    relationships = (
        parse_xml(package.get_part(relationship_part_uri))
        if relationship_part_exists
        else ET.Element(_q(REL, "Relationships"))
    )
    if relationships.tag != _q(REL, "Relationships"):
        raise NativePackageError(
            f"DOCX relationship part {relationship_part_uri!r} has "
            "an invalid root."
        )
    relationship_elements = relationships.findall(
        _q(REL, "Relationship")
    )
    numbering_relationships = [
        relationship
        for relationship in relationship_elements
        if relationship.get("Type") == NUMBERING_RELATIONSHIP_TYPE
    ]
    if len(numbering_relationships) > 1:
        raise NativePackageError(
            "Native DOCX has multiple document numbering "
            "relationships."
        )
    if numbering_relationships:
        relationship = numbering_relationships[0]
        if relationship.get("TargetMode") is not None:
            raise NativePackageError(
                "Native DOCX numbering relationship must be internal."
            )
        if (
            resolve_relationship_target(
                source_part,
                relationship.get("Target", ""),
            )
            != "/word/numbering.xml"
        ):
            raise NativePackageError(
                "Native DOCX numbering relationship targets an "
                "unsupported part."
            )
        return
    relationship_ids = [
        relationship.get("Id", "")
        for relationship in relationship_elements
    ]
    if (
        any(not relationship_id for relationship_id in relationship_ids)
        or len(relationship_ids) != len(set(relationship_ids))
    ):
        raise NativePackageError(
            "Native DOCX relationship IDs are incomplete or ambiguous."
        )
    relationship_number = 1
    while (
        f"rIdAiOfficeNumbering{relationship_number}"
        in relationship_ids
    ):
        relationship_number += 1
    ET.SubElement(
        relationships,
        _q(REL, "Relationship"),
        {
            "Id": f"rIdAiOfficeNumbering{relationship_number}",
            "Type": NUMBERING_RELATIONSHIP_TYPE,
            "Target": "numbering.xml",
        },
    )
    package.set_part(
        relationship_part_uri,
        serialize_xml(relationships),
        content_type=(
            None
            if relationship_part_exists
            else RELATIONSHIPS_CONTENT_TYPE
        ),
    )


def _ensure_numbering_content_type(
    package: NativePackage,
) -> None:
    content_types = parse_xml(
        package.get_part("/[Content_Types].xml")
    )
    if content_types.tag != _q(CT, "Types"):
        raise NativePackageError(
            "DOCX content types part has an invalid root."
        )
    overrides = [
        override
        for override in content_types.findall(_q(CT, "Override"))
        if override.get("PartName") == "/word/numbering.xml"
    ]
    if len(overrides) > 1:
        raise NativePackageError(
            "DOCX content types contain duplicate numbering overrides."
        )
    if overrides:
        if (
            overrides[0].get("ContentType")
            != NUMBERING_CONTENT_TYPE
        ):
            raise NativePackageError(
                "DOCX numbering content type is invalid."
            )
        return
    ET.SubElement(
        content_types,
        _q(CT, "Override"),
        {
            "PartName": "/word/numbering.xml",
            "ContentType": NUMBERING_CONTENT_TYPE,
        },
    )
    package.set_part(
        "/[Content_Types].xml",
        serialize_xml(content_types),
    )


def _append_numbering_definition(
    package: NativePackage,
    *,
    ordered: bool,
) -> int:
    numbering_part_exists = package.has_part(
        "/word/numbering.xml"
    )
    numbering = (
        parse_xml(package.get_part("/word/numbering.xml"))
        if numbering_part_exists
        else ET.Element(_q(W, "numbering"))
    )
    if numbering.tag != _q(W, "numbering"):
        raise NativePackageError(
            "Native DOCX numbering.xml has an invalid root."
        )
    _validate_numbering_child_order(numbering)
    abstract_ids = _numbering_decimal_ids(
        numbering.findall(_q(W, "abstractNum")),
        attribute_name="abstractNumId",
        label="abstractNumId",
    )
    num_ids = _numbering_decimal_ids(
        numbering.findall(_q(W, "num")),
        attribute_name="numId",
        label="numId",
    )
    abstract_id = _fresh_decimal_id(abstract_ids, start=0)
    num_id = _fresh_decimal_id(num_ids, start=1)
    append_single_level_numbering_definition(
        numbering,
        abstract_id=abstract_id,
        num_id=num_id,
        ordered=ordered,
    )
    _validate_numbering_child_order(numbering)
    try:
        numbering_payload = serialize_xml(numbering)
        parse_xml(numbering_payload)
    except (
        NativePackageError,
        UnicodeError,
        ValueError,
    ) as error:
        raise NativePackageError(
            "Native list insertion generated invalid numbering XML."
        ) from error
    _ensure_numbering_relationship(package)
    _ensure_numbering_content_type(package)
    package.set_part(
        "/word/numbering.xml",
        numbering_payload,
        content_type=NUMBERING_CONTENT_TYPE,
    )
    return num_id


def _compile_inserted_list(
    package: NativePackage,
    container: ET.Element,
    block: BulletList | OrderedList,
) -> list[ET.Element]:
    if block.source_ref is not None:
        raise NativePackageError(
            f"Inserted list {block.id!r} cannot claim an existing "
            "native source reference."
        )
    numbering_id = _append_numbering_definition(
        package,
        ordered=isinstance(block, OrderedList),
    )
    reserved_paragraph_ids = {
        para_id
        for paragraph in container.iter(_q(W, "p"))
        if (
            para_id := paragraph.get(_q(W14, "paraId"))
        )
        is not None
    }
    anchors: list[str] = []
    for item_index in range(len(block.items)):
        anchor = _fresh_native_paragraph_anchor(
            reserved_paragraph_ids,
            f"{block.id}:{item_index}",
        )
        anchors.append(anchor)
        reserved_paragraph_ids.add(anchor)
    temporary_container = ET.Element(_q(W, "body"))
    try:
        elements = compile_list_elements(
            temporary_container,
            block,
            DocxCompileContext(),
            numbering_id=numbering_id,
            native_anchors=anchors,
        )
        parse_xml(serialize_xml(temporary_container))
    except (
        NativePackageError,
        UnicodeError,
        ValueError,
    ) as error:
        raise NativePackageError(
            f"Could not compile inserted list {block.id!r} as "
            "valid, safe XML."
        ) from error
    if (
        len(elements) != len(block.items)
        or any(element.tag != _q(W, "p") for element in elements)
    ):
        raise NativePackageError(
            f"Inserted list {block.id!r} did not compile to one "
            "paragraph per item."
        )
    return elements


def _merge_text_style(
    base: TextStyle | None,
    override: TextStyle | None,
) -> TextStyle | None:
    if base is None:
        return override
    if override is None:
        return base
    return TextStyle.model_validate(
        {
            **base.model_dump(mode="json", exclude_none=True),
            **override.model_dump(mode="json", exclude_none=True),
        }
    )


def _append_inserted_text_span(
    package: NativePackage,
    paragraph: ET.Element,
    span: TextSpan,
    *,
    source_part: str,
    default_style: TextStyle | None,
) -> None:
    run_parent = paragraph
    if "link" in span.marks:
        assert span.href is not None
        if span.href.startswith("#"):
            anchor = span.href[1:]
            if not anchor:
                raise NativePackageError(
                    "A native internal hyperlink requires a non-empty anchor."
                )
            run_parent = ET.SubElement(
                paragraph,
                _q(W, "hyperlink"),
                {_q(W, "anchor"): anchor},
            )
        else:
            relationship_id = _attach_hyperlink_relationship(
                package,
                source_part=source_part,
                target=span.href,
            )
            run_parent = ET.SubElement(
                paragraph,
                _q(W, "hyperlink"),
                {_q(R, "id"): relationship_id},
            )
    run = ET.SubElement(run_parent, _q(W, "r"))
    apply_text_style(
        run,
        _merge_text_style(default_style, span.style),
    )
    if span.marks:
        mark_style: dict[str, object] = {}
        if "strong" in span.marks:
            mark_style["bold"] = True
        if "emphasis" in span.marks:
            mark_style["italic"] = True
        if "underline" in span.marks:
            mark_style["underline"] = True
        if "strike" in span.marks:
            mark_style["strike"] = True
        if "subscript" in span.marks:
            mark_style["baseline"] = "subscript"
        if "superscript" in span.marks:
            mark_style["baseline"] = "superscript"
        if "code" in span.marks:
            mark_style["font_family"] = "Consolas"
        if mark_style:
            apply_text_style(
                run,
                TextStyle.model_validate(mark_style),
            )
        properties = run.find(_q(W, "rPr"))
        if properties is None:
            properties = ET.Element(_q(W, "rPr"))
            run.insert(0, properties)
        if "highlight" in span.marks:
            ET.SubElement(
                properties,
                _q(W, "highlight"),
                {_q(W, "val"): "yellow"},
            )
        if "link" in span.marks:
            ET.SubElement(
                properties,
                _q(W, "rStyle"),
                {_q(W, "val"): "Hyperlink"},
            )
    text = ET.SubElement(run, _q(W, "t"))
    if (
        span.text[:1].isspace()
        or span.text[-1:].isspace()
        or "  " in span.text
    ):
        text.set(_q(XML, "space"), "preserve")
    text.text = span.text


def _compile_inserted_body_block(
    package: NativePackage,
    container: ET.Element,
    block: Heading | PageBreak | Paragraph,
    *,
    source_part: str,
    styles_root: ET.Element | None,
) -> tuple[ET.Element, list[DocumentField]]:
    if block.source_ref is not None:
        raise NativePackageError(
            "Inserted content cannot claim an existing native "
            "source reference."
        )
    paragraph = ET.Element(
        _q(W, "p"),
        {
            _q(W14, "paraId"): _native_paragraph_anchor(
                container,
                block.id,
            )
        },
    )
    if isinstance(block, PageBreak):
        run = ET.SubElement(paragraph, _q(W, "r"))
        ET.SubElement(
            run,
            _q(W, "br"),
            {_q(W, "type"): "page"},
        )
        try:
            parse_xml(serialize_xml(paragraph))
        except (
            NativePackageError,
            UnicodeError,
            ValueError,
        ) as error:
            raise NativePackageError(
                "Native page-break insertion generated attributes that "
                "are not valid, safe XML."
            ) from error
        return paragraph, []
    native_style_ref = (
        block.style_ref or f"Heading{block.level}"
        if isinstance(block, Heading)
        else block.style_ref
    )
    if native_style_ref is not None:
        if (
            styles_root is None
            or find_named_style(styles_root, native_style_ref) is None
        ):
            raise NativePackageError(
                f"Native DOCX has no paragraph style "
                f"{native_style_ref!r} required by inserted "
                f"{block.type} {block.id!r}."
            )
    if native_style_ref is not None:
        patch_paragraph_style_ref(
            paragraph,
            native_style_ref,
        )
    apply_paragraph_style(paragraph, block.paragraph_style)
    apply_paragraph_mark_text_style(paragraph, block.text_style)
    content: Sequence[InlineContent] = (
        block.content
        if block.text is None
        else [TextSpan(text=block.text)]
    )
    inserted_fields: list[DocumentField] = []
    for inline in content:
        if isinstance(inline, TextSpan):
            _append_inserted_text_span(
                package,
                paragraph,
                inline,
                source_part=source_part,
                default_style=block.text_style,
            )
            continue
        if inline.kind == "native":
            raise NativePackageError(
                "Native text insertion cannot reconstruct a native-only field "
                "instruction from its semantic projection."
            )
        append_complex_field(
            paragraph,
            inline,
            effective_style=_merge_text_style(
                block.text_style,
                inline.style,
            ),
        )
        inserted_fields.append(inline)
    matches = parse_paragraph_fields(paragraph)
    if len(matches) != len(inserted_fields):
        raise NativePackageError(
            "Inserted field count does not match generated native fields."
        )
    try:
        parse_xml(serialize_xml(paragraph))
    except (NativePackageError, UnicodeError, ValueError) as error:
        raise NativePackageError(
            "Native text insertion generated text or attributes that are not "
            "valid, safe XML."
        ) from error
    return paragraph, inserted_fields


def _compile_inserted_table(
    package: NativePackage,
    container: ET.Element,
    table: Table,
    *,
    source_part: str,
    styles_root: ET.Element | None,
) -> ET.Element:
    components: list[Any] = [
        table,
        *table.columns,
        *table.rows,
        *(
            cell
            for row in table.rows
            for cell in row.cells
        ),
        *(
            paragraph
            for row in table.rows
            for cell in row.cells
            for paragraph in cell.content
        ),
    ]
    forged_ids = [
        component.id
        for component in components
        if component.source_ref is not None
    ]
    if forged_ids:
        raise NativePackageError(
            "Inserted table content cannot claim existing native "
            f"source references: {', '.join(forged_ids)}."
        )
    table_style_ref = table.layout.style_ref or "TableGrid"
    if styles_root is None:
        raise NativePackageError(
            f"Native DOCX has no styles.xml part required by inserted "
            f"table style {table_style_ref!r}."
        )
    table_style_matches = [
        style
        for style in styles_root.findall(_q(W, "style"))
        if style.get(_q(W, "type"), "paragraph") == "table"
        and style.get(_q(W, "styleId")) == table_style_ref
    ]
    if len(table_style_matches) != 1:
        raise NativePackageError(
            f"Native DOCX requires exactly one table style "
            f"{table_style_ref!r} for inserted table {table.id!r}; "
            f"found {len(table_style_matches)}."
        )
    for paragraph in (
        paragraph
        for row in table.rows
        for cell in row.cells
        for paragraph in cell.content
        if paragraph.style_ref is not None
    ):
        assert paragraph.style_ref is not None
        if find_named_style(
            styles_root,
            paragraph.style_ref,
        ) is None:
            raise NativePackageError(
                f"Native DOCX has no paragraph style "
                f"{paragraph.style_ref!r} required by table-cell "
                f"paragraph {paragraph.id!r}."
            )
    context = DocxCompileContext()
    temporary_container = ET.Element(_q(W, "body"))
    try:
        element = compile_table_element(
            temporary_container,
            table,
            context,
        )
    except (ValidationError, ValueError) as error:
        raise NativePackageError(
            f"Could not compile inserted table {table.id!r}: {error}"
        ) from error
    for temporary_id, target in context.hyperlinks:
        hyperlinks = [
            hyperlink
            for hyperlink in element.iter(_q(W, "hyperlink"))
            if hyperlink.get(_q(R, "id")) == temporary_id
        ]
        if len(hyperlinks) != 1:
            raise NativePackageError(
                "Inserted table hyperlink compilation produced "
                "ambiguous relationship evidence."
            )
        hyperlink = hyperlinks[0]
        if target.startswith("#"):
            anchor = target[1:]
            if not anchor:
                raise NativePackageError(
                    "A native internal hyperlink requires a non-empty "
                    "anchor."
                )
            hyperlink.attrib.pop(_q(R, "id"), None)
            hyperlink.set(_q(W, "anchor"), anchor)
        else:
            hyperlink.set(
                _q(R, "id"),
                _attach_hyperlink_relationship(
                    package,
                    source_part=source_part,
                    target=target,
                ),
            )
    preliminary_refs: dict[str, NativeRef] = {}
    try:
        register_table_refs(
            preliminary_refs,
            element,
            0,
            table,
        )
    except ValueError as error:
        raise NativePackageError(
            f"Could not map inserted table {table.id!r}: {error}"
        ) from error
    reserved_paragraph_ids = {
        para_id
        for paragraph in container.iter(_q(W, "p"))
        if (
            para_id := paragraph.get(_q(W14, "paraId"))
        )
        is not None
    }
    for paragraph in (
        paragraph
        for row in table.rows
        for cell in row.cells
        for paragraph in cell.content
    ):
        paragraph_ref = preliminary_refs.get(paragraph.id)
        if paragraph_ref is None:
            raise NativePackageError(
                f"Inserted table paragraph {paragraph.id!r} has no "
                "compiled native reference."
            )
        try:
            _, native_paragraph = table_cell_paragraph_from_ref(
                element,
                paragraph_ref,
            )
        except ValueError as error:
            raise NativePackageError(
                f"Could not resolve inserted table paragraph "
                f"{paragraph.id!r}: {error}"
            ) from error
        native_paragraph.set(
            _q(W14, "paraId"),
            _fresh_native_paragraph_anchor(
                reserved_paragraph_ids,
                paragraph.id,
            ),
        )
        reserved_paragraph_ids.add(
            native_paragraph.get(_q(W14, "paraId"), "")
        )
    try:
        parse_xml(serialize_xml(element))
    except (
        NativePackageError,
        UnicodeError,
        ValueError,
    ) as error:
        raise NativePackageError(
            "Native table insertion generated content that is not "
            "valid, safe XML."
        ) from error
    return element


def _assign_inserted_table_refs(
    table: Table,
    refs: Mapping[str, NativeRef],
) -> None:
    components = [
        *table.columns,
        *table.rows,
        *(
            cell
            for row in table.rows
            for cell in row.cells
        ),
        *(
            paragraph
            for row in table.rows
            for cell in row.cells
            for paragraph in cell.content
        ),
    ]
    missing_ids = [
        component.id
        for component in components
        if component.id not in refs
    ]
    if missing_ids:
        raise NativePackageError(
            "Inserted table identity mapping is incomplete: "
            f"{', '.join(missing_ids)}."
        )
    for component in components:
        component.source_ref = refs[component.id]


def _synchronize_inserted_table_ids(
    table: Table,
    result_table: Table,
) -> None:
    """Reuse IDs assigned by semantic normalization for native lowering."""

    if table.id != result_table.id:
        raise NativePackageError(
            "Inserted table semantic and native root IDs do not match."
        )
    if (
        len(table.columns) != len(result_table.columns)
        or len(table.rows) != len(result_table.rows)
    ):
        raise NativePackageError(
            f"Inserted table {table.id!r} changed structural shape "
            "between semantic normalization and native lowering."
        )
    for source_column, result_column in zip(
        table.columns,
        result_table.columns,
        strict=True,
    ):
        if source_column.key != result_column.key:
            raise NativePackageError(
                f"Inserted table {table.id!r} changed column ordering "
                "between semantic normalization and native lowering."
            )
        source_column.id = result_column.id
    for source_row, result_row in zip(
        table.rows,
        result_table.rows,
        strict=True,
    ):
        if len(source_row.cells) != len(result_row.cells):
            raise NativePackageError(
                f"Inserted table row {source_row.id!r} changed cell "
                "shape between semantic normalization and native lowering."
            )
        source_row.id = result_row.id
        for source_cell, result_cell in zip(
            source_row.cells,
            result_row.cells,
            strict=True,
        ):
            if (
                source_cell.column_key != result_cell.column_key
                or source_cell.column_span != result_cell.column_span
                or source_cell.row_span != result_cell.row_span
                or len(source_cell.content)
                != len(result_cell.content)
            ):
                raise NativePackageError(
                    f"Inserted table row {source_row.id!r} changed "
                    "cell structure between semantic normalization and "
                    "native lowering."
                )
            source_cell.id = result_cell.id
            for source_paragraph, result_paragraph in zip(
                source_cell.content,
                result_cell.content,
                strict=True,
            ):
                source_paragraph.id = result_paragraph.id


def _ensure_identity_manifest_parts(package: NativePackage) -> None:
    """Attach the AiOffice identity part to a third-party OPC package."""

    root_relationships = parse_xml(package.get_part("/_rels/.rels"))
    manifest_relationships = [
        relationship
        for relationship in root_relationships.findall(
            _q(REL, "Relationship")
        )
        if relationship.get("Type") == MANIFEST_RELATIONSHIP_TYPE
    ]
    if len(manifest_relationships) > 1:
        raise NativePackageError(
            "Root relationships contain duplicate AiOffice manifest links."
        )
    if manifest_relationships:
        if manifest_relationships[0].get("TargetMode") is not None:
            raise NativePackageError(
                "AiOffice manifest relationship must be internal."
            )
        target = manifest_relationships[0].get("Target", "")
        if (
            resolve_relationship_target("/", target)
            != MANIFEST_PART_URI
        ):
            raise NativePackageError(
                "Existing AiOffice manifest relationship targets another part."
            )
    else:
        relationship_ids = {
            relationship.get("Id", "")
            for relationship in root_relationships.findall(
                _q(REL, "Relationship")
            )
        }
        relationship_number = 1
        while (
            f"rIdAiOfficeManifest{relationship_number}"
            in relationship_ids
        ):
            relationship_number += 1
        ET.SubElement(
            root_relationships,
            _q(REL, "Relationship"),
            {
                "Id": f"rIdAiOfficeManifest{relationship_number}",
                "Type": MANIFEST_RELATIONSHIP_TYPE,
                "Target": MANIFEST_PART_URI.lstrip("/"),
            },
        )
        package.set_part(
            "/_rels/.rels",
            serialize_xml(root_relationships),
        )

    content_types = parse_xml(
        package.get_part("/[Content_Types].xml")
    )
    overrides = [
        override
        for override in content_types.findall(_q(CT, "Override"))
        if override.get("PartName") == MANIFEST_PART_URI
    ]
    if len(overrides) > 1:
        raise NativePackageError(
            "Content types contain duplicate AiOffice manifest overrides."
        )
    if overrides:
        if overrides[0].get("ContentType") != "application/xml":
            raise NativePackageError(
                "AiOffice manifest content type is not application/xml."
            )
    else:
        manifest_part = next(
            (
                part
                for part in package.parts
                if part.uri == MANIFEST_PART_URI
            ),
            None,
        )
        if (
            package.has_part(MANIFEST_PART_URI)
            and manifest_part is not None
            and manifest_part.content_type == "application/xml"
        ):
            return
        ET.SubElement(
            content_types,
            _q(CT, "Override"),
            {
                "PartName": MANIFEST_PART_URI,
                "ContentType": "application/xml",
            },
        )
        package.set_part(
            "/[Content_Types].xml",
            serialize_xml(content_types),
        )


def _find_source_ref(spec: AiOfficeDocumentSpec, target: Any) -> NativeRef:
    target_id = _target_id(target)
    matches = [
        node
        for node in [
            *spec.content,
            *(
                block
                for part in spec.header_footers
                for block in part.content
            ),
            *(
                paragraph
                for table in spec.content
                if isinstance(table, Table)
                for row in table.rows
                for cell in row.cells
                for paragraph in cell.content
            ),
        ]
        if node.id == target_id
    ]
    if len(matches) != 1:
        raise NativePackageError(
            f"Native DOCX target #{target_id} matched {len(matches)} semantic nodes."
        )
    source_ref = matches[0].source_ref
    if not isinstance(source_ref, NativeRef) or source_ref.format != "docx":
        raise NativePackageError(
            f"Semantic node {target_id!r} has no editable DOCX source reference."
        )
    if source_ref.element_index is None and not source_ref.element_indices:
        raise NativePackageError(
            f"Semantic node {target_id!r} has no indexed DOCX element mapping."
        )
    return source_ref


def _fields(spec: AiOfficeDocumentSpec) -> list[DocumentField]:
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


def _find_field(
    spec: AiOfficeDocumentSpec,
    target: Any,
) -> tuple[DocumentField, NativeRef]:
    target_id = _target_id(target)
    matches = [field for field in _fields(spec) if field.id == target_id]
    if len(matches) != 1:
        raise NativePackageError(
            f"Native DOCX field target #{target_id} matched {len(matches)} fields."
        )
    document_field = matches[0]
    source_ref = document_field.source_ref
    if (
        not isinstance(source_ref, NativeRef)
        or source_ref.format != "docx"
        or source_ref.element_index is None
        or source_ref.sub_index is None
    ):
        raise NativePackageError(
            f"Field {target_id!r} has no editable DOCX inline source reference."
        )
    return document_field, source_ref


def _find_image(
    spec: AiOfficeDocumentSpec,
    target: Any,
) -> tuple[ImageBlock, NativeRef]:
    target_id = _target_id(target)
    matches = [
        node
        for node in spec.content
        if isinstance(node, ImageBlock) and node.id == target_id
    ]
    if len(matches) != 1:
        raise NativePackageError(
            f"Native DOCX image target #{target_id} matched "
            f"{len(matches)} images."
        )
    image = matches[0]
    source_ref = image.source_ref
    if (
        not isinstance(source_ref, NativeRef)
        or source_ref.format != "docx"
        or source_ref.native_kind != "w:p"
        or source_ref.element_index is None
    ):
        raise NativePackageError(
            f"Image {target_id!r} has no editable DOCX paragraph reference."
        )
    return image, source_ref


def _find_table_column(
    table: Table,
    selector: Any,
) -> tuple[int, str]:
    if not isinstance(selector, str) or not selector:
        raise NativePackageError(
            "table.column.format requires a non-empty column ID or key."
        )
    normalized = selector[1:] if selector.startswith("#") else selector
    matches = [
        (index, column.id)
        for index, column in enumerate(table.columns)
        if column.id == normalized or column.key == normalized
    ]
    if len(matches) != 1:
        raise NativePackageError(
            f"Column selector {selector!r} matched {len(matches)} columns "
            f"in table {table.id!r}."
        )
    return matches[0]


def _find_table_cell(
    table: Table,
    selector: Any,
) -> TableCell:
    if not isinstance(selector, str) or not selector:
        raise NativePackageError(
            "table.cell.format requires a non-empty cell ID."
        )
    normalized = selector[1:] if selector.startswith("#") else selector
    matches = [
        cell
        for row in table.rows
        for cell in row.cells
        if cell.id == normalized
    ]
    if len(matches) != 1:
        raise NativePackageError(
            f"Cell selector {selector!r} matched {len(matches)} cells "
            f"in table {table.id!r}."
        )
    return matches[0]


def _find_section_source_ref(
    spec: AiOfficeDocumentSpec,
    target: Any,
) -> tuple[str, NativeRef]:
    target_id = _target_id(target)
    matches = [section for section in spec.sections if section.id == target_id]
    if len(matches) != 1:
        raise NativePackageError(
            f"Native DOCX target #{target_id} matched {len(matches)} semantic sections."
        )
    source_ref = matches[0].source_ref
    if not isinstance(source_ref, NativeRef) or source_ref.format != "docx":
        raise NativePackageError(
            f"Semantic section {target_id!r} has no editable DOCX source reference."
        )
    if (
        source_ref.part_uri != "/word/document.xml"
        or source_ref.element_index is None
        or source_ref.native_kind
        not in {"w:sectPr-body", "w:sectPr-paragraph"}
    ):
        raise NativePackageError(
            f"Semantic section {target_id!r} is not mapped to a DOCX section."
        )
    return target_id, source_ref


def _section_element(
    body_elements: Sequence[ET.Element],
    source_ref: NativeRef,
) -> tuple[ET.Element, ET.Element, str]:
    assert source_ref.element_index is not None
    if source_ref.element_index >= len(body_elements):
        raise NativePackageError("DOCX section source reference points outside w:body.")
    container = body_elements[source_ref.element_index]
    if source_ref.native_kind == "w:sectPr-body":
        if container.tag != _q(W, "sectPr"):
            raise NativePackageError("DOCX body section source reference is stale.")
        return container, container, "body"
    if container.tag != _q(W, "p"):
        raise NativePackageError("DOCX paragraph section source reference is stale.")
    section = container.find(f"./{_q(W, 'pPr')}/{_q(W, 'sectPr')}")
    if section is None:
        raise NativePackageError("DOCX paragraph no longer contains its mapped section.")
    return section, container, "paragraph"


def _source_indices(source_ref: NativeRef) -> list[int]:
    if source_ref.element_indices:
        return list(source_ref.element_indices)
    if source_ref.element_index is not None:
        return [source_ref.element_index]
    raise NativePackageError("DOCX source reference has no native element indices.")


def _occurrences(text: str, search: str, replace_all: bool) -> list[tuple[int, int]]:
    found: list[tuple[int, int]] = []
    start = 0
    while True:
        index = text.find(search, start)
        if index < 0:
            break
        found.append((index, index + len(search)))
        if not replace_all:
            break
        start = index + len(search)
    return found


def _replace_text_nodes(
    element: ET.Element,
    search: str,
    replacement: str,
    *,
    replace_all: bool,
) -> int:
    text_nodes = list(element.iter(_q(W, "t")))
    values = [node.text or "" for node in text_nodes]
    joined = "".join(values)
    matches = _occurrences(joined, search, replace_all)
    if not matches:
        raise NativePackageError(f"Search text {search!r} was not found in native DOCX node.")

    spans: list[tuple[int, int]] = []
    cursor = 0
    for value in values:
        spans.append((cursor, cursor + len(value)))
        cursor += len(value)

    def locate(position: int) -> tuple[int, int]:
        for node_index, (start, end) in enumerate(spans):
            if start <= position < end:
                return node_index, position - start
        raise NativePackageError("Could not map semantic text to native DOCX runs.")

    for start, end in reversed(matches):
        start_node_index, start_offset = locate(start)
        end_node_index, end_offset_last = locate(end - 1)
        end_offset = end_offset_last + 1
        if start_node_index == end_node_index:
            node = text_nodes[start_node_index]
            current = node.text or ""
            node.text = current[:start_offset] + replacement + current[end_offset:]
        else:
            start_node = text_nodes[start_node_index]
            end_node = text_nodes[end_node_index]
            start_node.text = (start_node.text or "")[:start_offset] + replacement
            for middle_index in range(start_node_index + 1, end_node_index):
                text_nodes[middle_index].text = ""
            end_node.text = (end_node.text or "")[end_offset:]

    for node in text_nodes:
        value = node.text or ""
        xml_space = _q(XML, "space")
        if value[:1].isspace() or value[-1:].isspace() or "  " in value:
            node.set(xml_space, "preserve")
        else:
            node.attrib.pop(xml_space, None)
    return len(matches)


def _set_xml_space(text: ET.Element) -> None:
    value = text.text or ""
    xml_space = _q(XML, "space")
    if value[:1].isspace() or value[-1:].isspace() or "  " in value:
        text.set(xml_space, "preserve")
    else:
        text.attrib.pop(xml_space, None)


def _run_text(run: ET.Element) -> str:
    return "".join(node.text or "" for node in run.iter(_q(W, "t")))


def _clone_run_segment(
    run: ET.Element,
    text_children: Sequence[ET.Element],
    *,
    preserve_tail: bool,
) -> ET.Element:
    clone = ET.Element(run.tag, dict(run.attrib))
    clone.text = run.text
    properties = run.find(_q(W, "rPr"))
    if properties is not None:
        clone.append(deepcopy(properties))
    clone.extend(text_children)
    clone.tail = run.tail if preserve_tail else None
    return clone


def _split_and_format_run(
    run: ET.Element,
    *,
    run_start: int,
    selection_start: int,
    selection_end: int,
    style: TextStyle,
    fields: set[str],
    parent_map: Mapping[ET.Element, ET.Element],
) -> int:
    children = list(run)
    unsupported = [
        child.tag.rsplit("}", 1)[-1]
        for child in children
        if child.tag not in {_q(W, "rPr"), _q(W, "t")}
    ]
    if unsupported:
        raise NativePackageError(
            "A partial text range crosses a complex native run containing "
            f"{', '.join(sorted(set(unsupported)))}; refusing to duplicate or "
            "drop unknown inline content."
        )
    text_nodes = [child for child in children if child.tag == _q(W, "t")]
    segments: list[tuple[bool, list[ET.Element]]] = []
    cursor = run_start
    for text_node in text_nodes:
        value = text_node.text or ""
        node_start = cursor
        node_end = cursor + len(value)
        cursor = node_end
        cuts = {
            0,
            len(value),
            max(0, min(len(value), selection_start - node_start)),
            max(0, min(len(value), selection_end - node_start)),
        }
        ordered = sorted(cuts)
        for left, right in zip(ordered, ordered[1:]):
            if left == right:
                continue
            piece_start = node_start + left
            piece_end = node_start + right
            selected = (
                piece_start >= selection_start
                and piece_end <= selection_end
                and piece_start < piece_end
            )
            cloned_text = deepcopy(text_node)
            cloned_text.text = value[left:right]
            cloned_text.tail = text_node.tail if right == len(value) else None
            _set_xml_space(cloned_text)
            if segments and segments[-1][0] == selected:
                segments[-1][1].append(cloned_text)
            else:
                segments.append((selected, [cloned_text]))
    if not segments:
        raise NativePackageError("Could not split the selected native DOCX text run.")
    try:
        parent = parent_map[run]
    except KeyError as error:
        raise NativePackageError("Could not locate the native run parent.") from error
    index = list(parent).index(run)
    parent.remove(run)
    selected_runs = 0
    for offset, (selected, text_children) in enumerate(segments):
        clone = _clone_run_segment(
            run,
            text_children,
            preserve_tail=offset == len(segments) - 1,
        )
        if selected:
            patch_text_style(clone, style, fields)
            selected_runs += 1
        parent.insert(index + offset, clone)
    return selected_runs


def _format_text_range(
    paragraph: ET.Element,
    *,
    start: int,
    end: int,
    style: TextStyle,
    fields: set[str],
) -> int:
    runs = list(paragraph.iter(_q(W, "r")))
    parent_map = {child: parent for parent in paragraph.iter() for child in list(parent)}
    cursor = 0
    selected_runs = 0
    for run in runs:
        value = _run_text(run)
        run_start = cursor
        run_end = cursor + len(value)
        cursor = run_end
        if not value or run_end <= start or run_start >= end:
            continue
        if start <= run_start and run_end <= end:
            patch_text_style(run, style, fields)
            selected_runs += 1
        else:
            selected_runs += _split_and_format_run(
                run,
                run_start=run_start,
                selection_start=start,
                selection_end=end,
                style=style,
                fields=fields,
                parent_map=parent_map,
            )
    if selected_runs == 0:
        raise NativePackageError("The selected text range mapped to no editable DOCX runs.")
    return selected_runs


def apply_docx_operations(
    package: NativePackage,
    spec: AiOfficeDocumentSpec,
    result_spec: AiOfficeDocumentSpec,
    operations: Sequence[Mapping[str, Any]],
    *,
    changes: Sequence[Mapping[str, Any]],
    image_payloads: Mapping[str, bytes] | None = None,
) -> tuple[NativePackage, FidelityReport, dict[str, NativeRef]]:
    supported = {
        "text.replace",
        "paragraph.format",
        "text.format",
        "node.append",
        "node.insert_after",
        "node.insert_before",
        "node.remove",
        "style.apply",
        "style.define",
        "style.format",
        "section.insert_before",
        "section.format",
        "field.update",
        "image.insert_after",
        "image.replace",
        "image.update",
        "node.move_after",
        "node.move_before",
        "table.format",
        "table.column.format",
        "table.cell.format",
    }
    unsupported = sorted({str(operation.get("op")) for operation in operations} - supported)
    if unsupported:
        raise NativePackageError(
            "Imported DOCX V0.2 currently supports native lowering for "
            "text.replace, paragraph.format, text.format, "
            "node.append, node.insert_after, node.insert_before, node.move_after, "
            "node.move_before, node.remove, "
            "style.apply, style.define, style.format, "
            "section.insert_before, section.format, and "
            "field.update, image.insert_after, image.replace, image.update, "
            "table.format, "
            "table.column.format, and "
            "table.cell.format; "
            f"unsupported: {', '.join(unsupported)}."
        )
    if len(changes) != len(operations):
        raise NativePackageError(
            "Native DOCX lowering requires one semantic change record "
            "for every Patch operation."
        )

    updated = package.clone()
    root = parse_xml(updated.get_part("/word/document.xml"))
    body = root.find(_q(W, "body"))
    if body is None:
        raise NativePackageError("DOCX main document part has no w:body.")
    style_operations = {
        "style.apply",
        "style.define",
        "style.format",
    }.intersection(str(operation.get("op")) for operation in operations)
    has_body_insert = any(
        operation.get("op")
        in {"node.append", "node.insert_after", "node.insert_before"}
        for operation in operations
    )
    styles_root: ET.Element | None = None
    styles_changed = False
    changed_xml_parts: set[str] = set()
    if style_operations:
        if not updated.has_part("/word/styles.xml"):
            raise NativePackageError(
                "Native DOCX has no /word/styles.xml part for named-style editing."
            )
        styles_root = parse_xml(updated.get_part("/word/styles.xml"))
    elif has_body_insert and updated.has_part("/word/styles.xml"):
        styles_root = parse_xml(updated.get_part("/word/styles.xml"))
    original_elements = list(body)
    part_roots: dict[str, ET.Element] = {
        "/word/document.xml": root,
    }
    part_containers: dict[str, ET.Element] = {
        "/word/document.xml": body,
    }

    def elements_for_ref(
        source_ref: NativeRef,
    ) -> tuple[ET.Element, list[ET.Element]]:
        part_uri = source_ref.part_uri
        if source_ref.native_kind in {"w:tc", "w:tc/w:p"}:
            if (
                part_uri != "/word/document.xml"
                or source_ref.element_index is None
            ):
                raise NativePackageError(
                    "Table-cell references must point into document.xml."
                )
            body_elements = list(body)
            if source_ref.element_index >= len(body_elements):
                raise NativePackageError(
                    "Table-cell reference points outside w:body."
                )
            table = body_elements[source_ref.element_index]
            if table.tag != _q(W, "tbl"):
                raise NativePackageError(
                    "Table-cell reference no longer points to w:tbl."
                )
            try:
                if source_ref.native_kind == "w:tc":
                    return table, [
                        table_cell_from_ref(table, source_ref)
                    ]
                cell, paragraph = table_cell_paragraph_from_ref(
                    table,
                    source_ref,
                )
                return cell, [paragraph]
            except ValueError as error:
                raise NativePackageError(str(error)) from error
        container = part_containers.get(part_uri)
        if container is None:
            if not updated.has_part(part_uri):
                raise NativePackageError(
                    f"DOCX source part {part_uri!r} no longer exists."
                )
            part_root = parse_xml(updated.get_part(part_uri))
            if part_root.tag not in {_q(W, "hdr"), _q(W, "ftr")}:
                raise NativePackageError(
                    f"DOCX source part {part_uri!r} is not an editable "
                    "header/footer part."
                )
            part_roots[part_uri] = part_root
            part_containers[part_uri] = part_root
            container = part_root
        indices = _source_indices(source_ref)
        elements = list(container)
        if any(index >= len(elements) for index in indices):
            raise NativePackageError(
                f"DOCX source reference points outside {part_uri!r}."
            )
        return container, [elements[index] for index in indices]

    source_elements: dict[str, tuple[list[ET.Element], NativeRef]] = {}
    semantic_nodes = [
        *spec.content,
        *(
            block
            for part in spec.header_footers
            for block in part.content
        ),
    ]
    for node in semantic_nodes:
        source_ref = node.source_ref
        if not isinstance(source_ref, NativeRef):
            continue
        _, mapped_elements = elements_for_ref(source_ref)
        source_elements[node.id] = (mapped_elements, source_ref)
    source_fields: dict[str, tuple[ET.Element, NativeRef]] = {}
    for document_field in _fields(spec):
        source_ref = document_field.source_ref
        if not isinstance(source_ref, NativeRef):
            continue
        _, mapped_elements = elements_for_ref(source_ref)
        if len(mapped_elements) != 1 or mapped_elements[0].tag != _q(W, "p"):
            raise NativePackageError(
                f"Field {document_field.id!r} is not mapped to one w:p element."
            )
        source_fields[document_field.id] = (
            mapped_elements[0],
            source_ref,
        )
    source_tables: dict[str, tuple[Table, ET.Element, NativeRef]] = {}
    for table in (
        node
        for node in spec.content
        if isinstance(node, Table)
    ):
        source_ref = table.source_ref
        if not isinstance(source_ref, NativeRef):
            continue
        _, mapped_elements = elements_for_ref(source_ref)
        if (
            len(mapped_elements) != 1
            or mapped_elements[0].tag != _q(W, "tbl")
        ):
            raise NativePackageError(
                f"Table {table.id!r} is not mapped to one w:tbl element."
            )
        source_tables[table.id] = (
            table,
            mapped_elements[0],
            source_ref,
        )

    def live_table_for_target(
        target: Any,
    ) -> tuple[Table, ET.Element, NativeRef]:
        target_id = _target_id(target)
        mapped = source_tables.get(target_id)
        if mapped is None:
            raise NativePackageError(
                f"No mapped native table matched {target_id!r}."
            )
        source_table, table_element, source_ref = mapped
        container = part_containers.get(source_ref.part_uri)
        if (
            source_ref.part_uri != "/word/document.xml"
            or container is not body
            or table_element not in list(body)
        ):
            raise NativePackageError(
                f"Table {target_id!r} is no longer a mapped top-level "
                "document body table."
            )
        return source_table, table_element, source_ref

    source_sections: dict[
        str,
        tuple[ET.Element, ET.Element, str],
    ] = {}
    for section in spec.sections:
        source_ref = section.source_ref
        if (
            isinstance(source_ref, NativeRef)
            and source_ref.part_uri == "/word/document.xml"
            and source_ref.native_kind
            in {"w:sectPr-body", "w:sectPr-paragraph"}
        ):
            source_sections[section.id] = _section_element(
                original_elements,
                source_ref,
            )
    section_indices = {
        section.id: index
        for index, section in enumerate(spec.sections)
    }
    native_section_starts = {
        section.id: section.start_at
        for section in spec.sections
    }

    def synchronize_section_start(
        change: Mapping[str, Any],
        *,
        operation_name: str,
        anchor_id: str,
        new_start_id: str,
        new_start_elements: Sequence[ET.Element],
        can_rebind: bool,
    ) -> None:
        section_start_update = change.get(
            "section_start_updated"
        )
        if not can_rebind:
            if section_start_update is not None:
                raise NativePackageError(
                    f"{operation_name} cannot carry section-start "
                    "rebind evidence."
                )
            return
        matching_section_ids = [
            section_id
            for section_id, start_at in native_section_starts.items()
            if start_at == anchor_id
        ]
        if section_start_update is None:
            if matching_section_ids:
                raise NativePackageError(
                    f"{operation_name} is missing the semantic "
                    "section-start rebind required by its anchor."
                )
            return
        if (
            not isinstance(section_start_update, Mapping)
            or len(matching_section_ids) != 1
        ):
            raise NativePackageError(
                f"{operation_name} has invalid or ambiguous "
                "section-start evidence."
            )
        section_id = matching_section_ids[0]
        section_index = section_indices.get(section_id)
        if (
            section_start_update.get("section_id") != section_id
            or section_start_update.get("from") != anchor_id
            or section_start_update.get("to") != new_start_id
            or change.get("section_index") != section_index
            or section_index is None
            or section_index <= 0
        ):
            raise NativePackageError(
                f"{operation_name} section-start evidence does not "
                "match the semantic section model."
            )
        previous_section_id = next(
            (
                candidate_id
                for candidate_id, candidate_index in (
                    section_indices.items()
                )
                if candidate_index == section_index - 1
            ),
            None,
        )
        if previous_section_id is None:
            raise NativePackageError(
                f"{operation_name} cannot resolve the section preceding "
                "its target."
            )
        previous_boundary = source_sections.get(
            previous_section_id
        )
        if previous_boundary is None:
            raise NativePackageError(
                f"{operation_name} cannot prove the native boundary "
                "preceding the target section."
            )
        _, boundary_container, _ = previous_boundary
        current_body_elements = list(body)
        if (
            boundary_container not in current_body_elements
            or not new_start_elements
            or any(
                element not in current_body_elements
                for element in new_start_elements
            )
            or current_body_elements.index(boundary_container)
            >= min(
                current_body_elements.index(element)
                for element in new_start_elements
            )
        ):
            raise NativePackageError(
                f"{operation_name} result is not positioned after "
                "its proven native section boundary."
            )
        native_section_starts[section_id] = new_start_id

    inserted_images: dict[str, ET.Element] = {}
    inserted_nodes: set[str] = set()
    inserted_sections: set[str] = set()
    inserted_fields: dict[str, tuple[ET.Element, int]] = {}
    moved_nodes: set[str] = set()
    removed_nodes: set[str] = set()

    for operation, change in zip(operations, changes, strict=True):
        operation_name = operation.get("op")
        if operation_name == "style.define":
            assert styles_root is not None
            try:
                named_style = NamedStyle.model_validate(operation.get("style"))
            except ValidationError as error:
                raise NativePackageError(
                    f"Could not lower style.define values: {error}"
                ) from error
            if find_named_style(styles_root, named_style.id) is not None:
                raise NativePackageError(
                    f"Native DOCX paragraph style {named_style.id!r} already exists."
                )
            upsert_named_style(styles_root, named_style, custom_style=True)
            styles_changed = True
            continue
        if operation_name == "style.format":
            assert styles_root is not None
            paragraph_scope = operation.get("paragraph", {})
            text_scope = operation.get("text", {})
            paragraph_set = paragraph_scope.get("set", {})
            paragraph_fields = set(paragraph_set) | set(
                paragraph_scope.get("clear", [])
            )
            text_set = text_scope.get("set", {})
            text_fields = set(text_set) | set(text_scope.get("clear", []))
            try:
                paragraph_style = ParagraphStyle.model_validate(paragraph_set)
                text_style = TextStyle.model_validate(text_set)
            except ValidationError as error:
                raise NativePackageError(
                    f"Could not lower style.format values: {error}"
                ) from error
            raw_target = operation.get("target")
            if not isinstance(raw_target, str) or not raw_target:
                raise NativePackageError("style.format requires a named style ID.")
            style_id = raw_target
            if (
                find_named_style(styles_root, style_id) is None
                and raw_target.startswith("@")
            ):
                style_id = raw_target[1:]
            format_named_style(
                styles_root,
                style_id,
                paragraph_style=paragraph_style,
                paragraph_fields=paragraph_fields,
                text_style=text_style,
                text_fields=text_fields,
            )
            styles_changed = True
            continue
        if operation_name == "section.insert_before":
            target_id = _target_id(operation.get("target"))
            mapped_target = source_elements.get(target_id)
            if mapped_target is None:
                raise NativePackageError(
                    "section.insert_before target has no mapped native "
                    "body elements."
                )
            target_elements, target_ref = mapped_target
            if (
                target_ref.part_uri != "/word/document.xml"
                or part_containers.get(target_ref.part_uri) is not body
            ):
                raise NativePackageError(
                    "section.insert_before requires a top-level document "
                    "body target."
                )
            current_body_elements = list(body)
            if (
                not target_elements
                or any(
                    element not in current_body_elements
                    for element in target_elements
                )
            ):
                raise NativePackageError(
                    "section.insert_before target is no longer in the "
                    "document body."
                )
            target_indices = [
                current_body_elements.index(element)
                for element in target_elements
            ]
            if target_indices != list(
                range(
                    target_indices[0],
                    target_indices[0] + len(target_indices),
                )
            ):
                raise NativePackageError(
                    "section.insert_before target native range is not "
                    "contiguous."
                )

            created_sections = change.get("created_sections")
            split_section_id = change.get("split_section_id")
            new_section_index = change.get("section_index")
            if (
                change.get("operation") != operation_name
                or not isinstance(created_sections, list)
                or len(created_sections) != 1
                or not isinstance(created_sections[0], str)
                or not isinstance(split_section_id, str)
                or not isinstance(new_section_index, int)
            ):
                raise NativePackageError(
                    "section.insert_before requires one trusted semantic "
                    "created-section record."
                )
            new_section_id = created_sections[0]
            split_section_index = section_indices.get(
                split_section_id
            )
            split_boundary = source_sections.get(split_section_id)
            if (
                split_section_index is None
                or new_section_index != split_section_index + 1
                or split_boundary is None
                or change.get("start_at") != target_id
                or new_section_id in source_sections
                or new_section_id in native_section_starts
            ):
                raise NativePackageError(
                    "section.insert_before semantic section evidence does "
                    "not match the native section model."
                )
            if target_id in native_section_starts.values():
                raise NativePackageError(
                    "section.insert_before refuses to create an empty "
                    "section before an existing section start."
                )
            raw_section = operation.get("section")
            raw_layout = (
                raw_section.get("layout", {})
                if isinstance(raw_section, Mapping)
                else None
            )
            if not isinstance(raw_layout, Mapping):
                raise NativePackageError(
                    "section.insert_before section.layout must be an "
                    "object."
                )
            layout_fields = set(raw_layout) | {"start_type"}
            if change.get("layout_fields") != sorted(layout_fields):
                raise NativePackageError(
                    "section.insert_before layout evidence does not match "
                    "the requested fields."
                )
            result_section = next(
                (
                    candidate
                    for candidate in result_spec.sections
                    if candidate.id == new_section_id
                ),
                None,
            )
            result_split_section = next(
                (
                    candidate
                    for candidate in result_spec.sections
                    if candidate.id == split_section_id
                ),
                None,
            )
            if (
                result_section is None
                or result_split_section is None
                or result_section.header_footer
                != result_split_section.header_footer
                or change.get("header_footer_inherited") is not True
            ):
                raise NativePackageError(
                    "section.insert_before result does not preserve its "
                    "semantic anchor and header/footer inheritance."
                )

            (
                split_section_element,
                split_boundary_container,
                split_container_kind,
            ) = split_boundary
            if (
                split_boundary_container not in current_body_elements
                or current_body_elements.index(
                    split_boundary_container
                )
                < target_indices[-1]
            ):
                raise NativePackageError(
                    "section.insert_before target is not contained by its "
                    "proven native section boundary."
                )
            if (
                split_section_element.find(
                    _q(W, "sectPrChange")
                )
                is not None
            ):
                raise NativePackageError(
                    "section.insert_before refuses tracked section "
                    "properties."
                )
            if split_section_index > 0:
                previous_section_id = next(
                    (
                        candidate_id
                        for candidate_id, candidate_index in (
                            section_indices.items()
                        )
                        if candidate_index == split_section_index - 1
                    ),
                    None,
                )
                previous_boundary = (
                    source_sections.get(previous_section_id)
                    if previous_section_id is not None
                    else None
                )
                if (
                    previous_boundary is None
                    or previous_boundary[1]
                    not in current_body_elements
                    or current_body_elements.index(
                        previous_boundary[1]
                    )
                    >= target_indices[0]
                ):
                    raise NativePackageError(
                        "section.insert_before cannot prove content before "
                        "the target inside its containing section."
                    )

            reserved_paragraph_ids = {
                para_id
                for paragraph in body.iter(_q(W, "p"))
                if (
                    para_id := paragraph.get(_q(W14, "paraId"))
                )
                is not None
            }
            carrier = ET.Element(
                _q(W, "p"),
                {
                    _q(W14, "paraId"): (
                        _fresh_native_paragraph_anchor(
                            reserved_paragraph_ids,
                            f"{new_section_id}:boundary",
                        )
                    )
                },
            )
            carrier_properties = ET.SubElement(
                carrier,
                _q(W, "pPr"),
            )
            copied_boundary = deepcopy(split_section_element)
            copied_boundary.tail = None
            carrier_properties.append(copied_boundary)
            body.insert(target_indices[0], carrier)
            patch_section_layout(
                split_section_element,
                result_section.layout,
                layout_fields,
            )

            source_sections[split_section_id] = (
                copied_boundary,
                carrier,
                "paragraph",
            )
            source_sections[new_section_id] = (
                split_section_element,
                split_boundary_container,
                split_container_kind,
            )
            for section_id, section_index in list(
                section_indices.items()
            ):
                if section_index >= new_section_index:
                    section_indices[section_id] = section_index + 1
            section_indices[new_section_id] = new_section_index
            native_section_starts[new_section_id] = target_id
            inserted_sections.add(new_section_id)
            changed_xml_parts.add("/word/document.xml")
            continue
        if operation_name == "section.format":
            target_id = _target_id(operation.get("target"))
            mapped_section = source_sections.get(target_id)
            if mapped_section is None:
                raise NativePackageError(
                    f"Semantic section {target_id!r} has no editable "
                    "DOCX section boundary."
                )
            section, _, _ = mapped_section
            fields = set(operation.get("set", {})) | set(
                operation.get("clear", [])
            )
            try:
                SectionLayout.model_validate(operation.get("set", {}))
            except ValidationError as error:
                raise NativePackageError(
                    f"Could not lower section.format values: {error}"
                ) from error
            result_section = next(
                (
                    candidate
                    for candidate in result_spec.sections
                    if candidate.id == target_id
                ),
                None,
            )
            if result_section is None:
                raise NativePackageError(
                    f"Patch result no longer contains section "
                    f"{target_id!r}."
                )
            patch_section_layout(
                section,
                result_section.layout,
                fields,
            )
            changed_xml_parts.add("/word/document.xml")
            continue
        if operation_name == "field.update":
            source_field, source_ref = _find_field(
                spec,
                operation.get("target"),
            )
            _, mapped_elements = elements_for_ref(source_ref)
            if len(mapped_elements) != 1 or mapped_elements[0].tag != _q(W, "p"):
                raise NativePackageError(
                    "field.update requires a native reference to one w:p element."
                )
            result_field = next(
                (
                    candidate
                    for candidate in _fields(result_spec)
                    if candidate.id == source_field.id
                ),
                None,
            )
            if result_field is None:
                raise NativePackageError(
                    f"Patch result no longer contains field {source_field.id!r}."
                )
            assert source_ref.sub_index is not None
            try:
                match = field_match_at(
                    mapped_elements[0],
                    source_ref.sub_index,
                )
                patch_field_instruction(match, result_field)
                result_field.metadata.update(
                    {
                        "native_form": match.form,
                        "native_instruction": canonical_field_instruction(
                            result_field
                        ),
                        "dirty": True,
                    }
                )
            except FieldStructureError as error:
                raise NativePackageError(
                    f"Could not patch field {source_field.id!r}: {error}"
                ) from error
            changed_xml_parts.add(source_ref.part_uri)
            continue
        if operation_name == "image.update":
            source_image, source_ref = _find_image(
                spec,
                operation.get("target"),
            )
            _, mapped_elements = elements_for_ref(source_ref)
            if (
                len(mapped_elements) != 1
                or mapped_elements[0].tag != _q(W, "p")
            ):
                raise NativePackageError(
                    "image.update requires a native reference to one w:p."
                )
            result_image = next(
                (
                    candidate
                    for candidate in result_spec.content
                    if isinstance(candidate, ImageBlock)
                    and candidate.id == source_image.id
                ),
                None,
            )
            if result_image is None:
                raise NativePackageError(
                    f"Patch result no longer contains image "
                    f"{source_image.id!r}."
                )
            fields = set(operation.get("set", {})) | set(
                operation.get("clear", [])
            )
            patch_simple_inline_image(
                updated,
                mapped_elements[0],
                source_part=source_ref.part_uri,
                result=result_image,
                fields=fields,
            )
            changed_xml_parts.add(source_ref.part_uri)
            continue
        if operation_name == "image.insert_after":
            source_ref = _find_source_ref(
                spec,
                operation.get("target"),
            )
            container, mapped_elements = elements_for_ref(source_ref)
            try:
                image_insert = ImageInsert.model_validate(
                    operation.get("image")
                )
                asset = AssetRef.model_validate(operation.get("asset"))
            except ValidationError as error:
                raise NativePackageError(
                    f"Could not lower image.insert_after metadata: {error}"
                ) from error
            payload = (
                image_payloads.get(asset.id)
                if image_payloads is not None
                else None
            )
            if payload is None:
                raise NativePackageError(
                    "image.insert_after requires a verified out-of-band "
                    "binary payload."
                )
            result_image = next(
                (
                    candidate
                    for candidate in result_spec.content
                    if isinstance(candidate, ImageBlock)
                    and candidate.id == image_insert.id
                ),
                None,
            )
            if (
                result_image is None
                or result_image.asset_id != asset.id
            ):
                raise NativePackageError(
                    "image.insert_after result does not contain its new image."
                )
            inserted = insert_simple_inline_image_after(
                updated,
                container,
                mapped_elements,
                source_part=source_ref.part_uri,
                image=image_insert,
                asset=asset,
                payload=payload,
            )
            inserted_images[image_insert.id] = inserted
            changed_xml_parts.add(source_ref.part_uri)
            continue
        if operation_name == "image.replace":
            source_image, source_ref = _find_image(
                spec,
                operation.get("target"),
            )
            _, mapped_elements = elements_for_ref(source_ref)
            if (
                len(mapped_elements) != 1
                or mapped_elements[0].tag != _q(W, "p")
            ):
                raise NativePackageError(
                    "image.replace requires a native reference to one w:p."
                )
            result_image = next(
                (
                    candidate
                    for candidate in result_spec.content
                    if isinstance(candidate, ImageBlock)
                    and candidate.id == source_image.id
                ),
                None,
            )
            if result_image is None:
                raise NativePackageError(
                    f"Patch result no longer contains image "
                    f"{source_image.id!r}."
                )
            try:
                asset = AssetRef.model_validate(operation.get("asset"))
            except ValidationError as error:
                raise NativePackageError(
                    f"Could not lower image.replace asset metadata: {error}"
                ) from error
            payload = (
                image_payloads.get(asset.id)
                if image_payloads is not None
                else None
            )
            if payload is None:
                raise NativePackageError(
                    "image.replace requires a verified out-of-band binary payload."
                )
            if result_image.asset_id != asset.id:
                raise NativePackageError(
                    "image.replace result does not reference its replacement asset."
                )
            replace_simple_inline_image(
                updated,
                mapped_elements[0],
                source_part=source_ref.part_uri,
                asset=asset,
                payload=payload,
            )
            changed_xml_parts.add(source_ref.part_uri)
            continue
        if operation_name in {
            "node.append",
            "node.insert_after",
            "node.insert_before",
        }:
            current_elements = list(body)
            target_id: str | None = None
            if operation_name == "node.append":
                section_property_indices = [
                    index
                    for index, element in enumerate(current_elements)
                    if element.tag == _q(W, "sectPr")
                ]
                if (
                    len(section_property_indices) > 1
                    or (
                        section_property_indices
                        and section_property_indices[0]
                        != len(current_elements) - 1
                    )
                ):
                    raise NativePackageError(
                        "node.append requires the optional body-level "
                        "w:sectPr to be the final and only direct section "
                        "properties element."
                    )
                insert_index = (
                    section_property_indices[0]
                    if section_property_indices
                    else len(current_elements)
                )
            else:
                target_id = _target_id(operation.get("target"))
                mapped_source = source_elements.get(target_id)
                if mapped_source is None:
                    source_ref = _find_source_ref(
                        spec,
                        operation.get("target"),
                    )
                    container, mapped_elements = elements_for_ref(
                        source_ref
                    )
                else:
                    mapped_elements, source_ref = mapped_source
                    container = part_containers[source_ref.part_uri]
                if (
                    source_ref.part_uri != "/word/document.xml"
                    or container is not body
                    or not mapped_elements
                ):
                    raise NativePackageError(
                        f"{operation_name} requires a mapped top-level "
                        "document body anchor."
                    )
                if any(
                    element not in current_elements
                    for element in mapped_elements
                ):
                    raise NativePackageError(
                        f"{operation_name} anchor is no longer in the "
                        "document body."
                    )
                anchor_indices = [
                    current_elements.index(element)
                    for element in mapped_elements
                ]
                if anchor_indices != list(
                    range(
                        anchor_indices[0],
                        anchor_indices[0] + len(anchor_indices),
                    )
                ):
                    raise NativePackageError(
                        f"{operation_name} requires the anchor's complete "
                        "native range to remain contiguous."
                    )
                if (
                    operation_name == "node.insert_after"
                    and any(
                        element.tag == _q(W, "sectPr")
                        or element.find(
                            f".//{_q(W, 'sectPr')}"
                        )
                        is not None
                        for element in mapped_elements
                    )
                ):
                    raise NativePackageError(
                        "node.insert_after refuses an anchor that carries a "
                        "native section boundary."
                    )
                insert_index = (
                    max(anchor_indices) + 1
                    if operation_name == "node.insert_after"
                    else min(anchor_indices)
                )
            created_nodes = change.get("created_nodes")
            if (
                change.get("operation") != operation_name
                or not isinstance(created_nodes, list)
                or len(created_nodes) != 1
                or not isinstance(created_nodes[0], str)
            ):
                raise NativePackageError(
                    f"{operation_name} requires one trusted semantic "
                    "created-node record."
                )
            created_id = created_nodes[0]
            content = operation.get("content")
            if not isinstance(content, Mapping):
                raise NativePackageError(
                    f"{operation_name} content must be an object."
                )
            candidate_payload = deepcopy(dict(content))
            supplied_id = candidate_payload.get("id")
            if supplied_id is not None and supplied_id != created_id:
                raise NativePackageError(
                    f"{operation_name} semantic and native created IDs "
                    "do not match."
                )
            candidate_payload["id"] = created_id
            try:
                if candidate_payload.get("type") == "paragraph":
                    candidate: (
                        BulletList
                        | Heading
                        | OrderedList
                        | PageBreak
                        | Paragraph
                        | Table
                    ) = Paragraph.model_validate(candidate_payload)
                elif candidate_payload.get("type") == "heading":
                    candidate = Heading.model_validate(candidate_payload)
                elif candidate_payload.get("type") == "page_break":
                    candidate = PageBreak.model_validate(
                        candidate_payload
                    )
                elif candidate_payload.get("type") == "bullet_list":
                    candidate = BulletList.model_validate(
                        candidate_payload
                    )
                elif candidate_payload.get("type") == "ordered_list":
                    candidate = OrderedList.model_validate(
                        candidate_payload
                    )
                elif candidate_payload.get("type") == "table":
                    candidate = Table.model_validate(
                        candidate_payload
                    )
                else:
                    raise NativePackageError(
                        f"Imported DOCX {operation_name} currently "
                        "supports only paragraph, heading, page_break, "
                        "bullet_list, ordered_list, and table content."
                    )
            except ValidationError as error:
                raise NativePackageError(
                    f"Could not lower {operation_name} content: {error}"
                ) from error
            new_fields: list[DocumentField] = []
            if isinstance(candidate, Table):
                matching_result_tables = [
                    result_node
                    for result_node in result_spec.content
                    if isinstance(result_node, Table)
                    and result_node.id == created_id
                ]
                if len(matching_result_tables) > 1:
                    raise NativePackageError(
                        f"Inserted table {created_id!r} is ambiguous in "
                        "the semantic patch result."
                    )
                if matching_result_tables:
                    _synchronize_inserted_table_ids(
                        candidate,
                        matching_result_tables[0],
                    )
                inserted_element = _compile_inserted_table(
                    updated,
                    body,
                    candidate,
                    source_part="/word/document.xml",
                    styles_root=styles_root,
                )
                inserted_elements = [inserted_element]
            elif isinstance(
                candidate,
                (BulletList, OrderedList),
            ):
                inserted_elements = _compile_inserted_list(
                    updated,
                    body,
                    candidate,
                )
            else:
                (
                    inserted_element,
                    new_fields,
                ) = _compile_inserted_body_block(
                    updated,
                    body,
                    candidate,
                    source_part="/word/document.xml",
                    styles_root=styles_root,
                )
                inserted_elements = [inserted_element]
            for offset, inserted_element in enumerate(
                inserted_elements
            ):
                body.insert(insert_index + offset, inserted_element)
            if target_id is not None:
                synchronize_section_start(
                    change,
                    operation_name=str(operation_name),
                    anchor_id=target_id,
                    new_start_id=created_id,
                    new_start_elements=inserted_elements,
                    can_rebind=(
                        operation_name == "node.insert_before"
                    ),
                )
            temporary_ref = native_ref_for_part_elements(
                inserted_elements,
                list(
                    range(
                        insert_index,
                        insert_index + len(inserted_elements),
                    )
                ),
                part_uri="/word/document.xml",
                native_kind=(
                    "w:tbl"
                    if isinstance(candidate, Table)
                    else (
                        "w:p-group"
                        if isinstance(
                            candidate,
                            (BulletList, OrderedList),
                        )
                        else (
                            "w:page-break"
                            if isinstance(candidate, PageBreak)
                            else "w:p"
                        )
                    )
                ),
                root_path="/w:document/w:body",
                native_id=inserted_elements[0].get(
                    _q(W14, "paraId")
                ),
            )
            if isinstance(candidate, Table):
                inserted_element = inserted_elements[0]
                candidate.source_ref = temporary_ref
                component_refs: dict[str, NativeRef] = {}
                try:
                    register_table_refs(
                        component_refs,
                        inserted_element,
                        insert_index,
                        candidate,
                    )
                except ValueError as error:
                    raise NativePackageError(
                        f"Could not map inserted table "
                        f"{candidate.id!r}: {error}"
                    ) from error
                _assign_inserted_table_refs(
                    candidate,
                    component_refs,
                )
                source_tables[created_id] = (
                    candidate,
                    inserted_element,
                    temporary_ref,
                )
                for row in candidate.rows:
                    for cell in row.cells:
                        for paragraph in cell.content:
                            paragraph_ref = paragraph.source_ref
                            assert isinstance(
                                paragraph_ref,
                                NativeRef,
                            )
                            try:
                                _, native_paragraph = (
                                    table_cell_paragraph_from_ref(
                                        inserted_element,
                                        paragraph_ref,
                                    )
                                )
                            except ValueError as error:
                                raise NativePackageError(
                                    "Could not register inserted table "
                                    f"paragraph {paragraph.id!r}: {error}"
                                ) from error
                            source_elements[paragraph.id] = (
                                [native_paragraph],
                                paragraph_ref,
                            )
            source_elements[created_id] = (
                inserted_elements,
                temporary_ref,
            )
            for field_ordinal, document_field in enumerate(
                new_fields
            ):
                inserted_fields[document_field.id] = (
                    inserted_elements[0],
                    field_ordinal,
                )
            inserted_nodes.add(created_id)
            changed_xml_parts.add("/word/document.xml")
            continue
        if operation_name == "table.format":
            (
                source_table,
                table_element,
                source_ref,
            ) = live_table_for_target(
                operation.get("target"),
            )
            if table_element.tag != _q(W, "tbl"):
                raise NativePackageError(
                    "table.format requires a native reference to one w:tbl."
                )
            result_table = next(
                (
                    candidate
                    for candidate in result_spec.content
                    if isinstance(candidate, Table)
                    and candidate.id == source_table.id
                ),
                None,
            )
            if result_table is None:
                raise NativePackageError(
                    f"Patch result no longer contains table "
                    f"{source_table.id!r}."
                )
            fields = set(operation.get("set", {})) | set(
                operation.get("clear", [])
            )
            try:
                TableLayout.model_validate(operation.get("set", {}))
                patch_table_layout(
                    table_element,
                    result_table.layout,
                    fields,
                )
            except (ValidationError, ValueError) as error:
                raise NativePackageError(
                    f"Could not lower table.format values: {error}"
                ) from error
            changed_xml_parts.add(source_ref.part_uri)
            continue
        if operation_name == "table.column.format":
            (
                source_table,
                table_element,
                source_ref,
            ) = live_table_for_target(
                operation.get("target"),
            )
            column_index, column_id = _find_table_column(
                source_table,
                operation.get("column"),
            )
            if table_element.tag != _q(W, "tbl"):
                raise NativePackageError(
                    "table.column.format requires a native reference "
                    "to one w:tbl."
                )
            result_table = next(
                (
                    candidate
                    for candidate in result_spec.content
                    if isinstance(candidate, Table)
                    and candidate.id == source_table.id
                ),
                None,
            )
            result_column = (
                next(
                    (
                        column
                        for column in result_table.columns
                        if column.id == column_id
                    ),
                    None,
                )
                if result_table is not None
                else None
            )
            if result_column is None:
                raise NativePackageError(
                    f"Patch result no longer contains table column "
                    f"{column_id!r}."
                )
            try:
                patch_table_column_width(
                    table_element,
                    column_index,
                    result_column.width,
                )
            except ValueError as error:
                raise NativePackageError(
                    f"Could not lower table.column.format: {error}"
                ) from error
            changed_xml_parts.add(source_ref.part_uri)
            continue
        if operation_name == "table.cell.format":
            (
                source_table,
                table_element,
                source_ref,
            ) = live_table_for_target(
                operation.get("target"),
            )
            source_cell = _find_table_cell(
                source_table,
                operation.get("cell"),
            )
            if not isinstance(source_cell.source_ref, NativeRef):
                raise NativePackageError(
                    f"Table cell {source_cell.id!r} has no editable "
                    "DOCX source reference."
                )
            if table_element.tag != _q(W, "tbl"):
                raise NativePackageError(
                    "table.cell.format requires a native reference "
                    "to one w:tbl."
                )
            result_table = next(
                (
                    candidate
                    for candidate in result_spec.content
                    if isinstance(candidate, Table)
                    and candidate.id == source_table.id
                ),
                None,
            )
            result_cell = (
                next(
                    (
                        cell
                        for row in result_table.rows
                        for cell in row.cells
                        if cell.id == source_cell.id
                    ),
                    None,
                )
                if result_table is not None
                else None
            )
            if result_cell is None:
                raise NativePackageError(
                    f"Patch result no longer contains table cell "
                    f"{source_cell.id!r}."
                )
            fields = set(operation.get("set", {})) | set(
                operation.get("clear", [])
            )
            try:
                TableCellFormat.model_validate(
                    operation.get("set", {})
                )
                native_cell = table_cell_from_ref(
                    table_element,
                    source_cell.source_ref,
                )
                patch_table_cell_format(
                    native_cell,
                    result_cell.format,
                    fields,
                )
            except (ValidationError, ValueError) as error:
                raise NativePackageError(
                    f"Could not lower table.cell.format: {error}"
                ) from error
            changed_xml_parts.add(source_ref.part_uri)
            continue

        target_id = _target_id(operation.get("target"))
        mapped_source = source_elements.get(target_id)
        if mapped_source is None:
            source_ref = _find_source_ref(
                spec,
                operation.get("target"),
            )
            container, elements = elements_for_ref(source_ref)
        else:
            elements, source_ref = mapped_source
            container = part_containers[source_ref.part_uri]
        if operation_name in {"node.move_after", "node.move_before"}:
            anchor_field = (
                "after"
                if operation_name == "node.move_after"
                else "before"
            )
            anchor_id = _target_id(operation.get(anchor_field))
            anchor_source = source_elements.get(anchor_id)
            if anchor_source is None:
                raise NativePackageError(
                    f"{operation_name} anchor has no mapped native elements."
                )
            anchor_elements, anchor_ref = anchor_source
            if (
                source_ref.part_uri != "/word/document.xml"
                or anchor_ref.part_uri != source_ref.part_uri
                or container is not body
                or part_containers[anchor_ref.part_uri] is not body
            ):
                raise NativePackageError(
                    f"{operation_name} requires target and anchor to be "
                    "top-level document body nodes."
                )
            current_elements = list(body)
            if (
                not elements
                or not anchor_elements
                or any(
                    element not in current_elements
                    for element in [*elements, *anchor_elements]
                )
            ):
                raise NativePackageError(
                    f"{operation_name} target or anchor is no longer in "
                    "the document body."
                )
            if set(map(id, elements)).intersection(
                map(id, anchor_elements)
            ):
                raise NativePackageError(
                    f"{operation_name} target and anchor native ranges overlap."
                )

            def contiguous_indices(
                group: list[ET.Element],
            ) -> list[int]:
                indices = [
                    current_elements.index(element)
                    for element in group
                ]
                if indices != list(
                    range(indices[0], indices[0] + len(indices))
                ):
                    raise NativePackageError(
                        f"{operation_name} requires each mapped native range "
                        "to remain contiguous."
                    )
                return indices

            contiguous_indices(elements)
            contiguous_indices(anchor_elements)
            if any(
                element.tag == _q(W, "sectPr")
                or element.find(f".//{_q(W, 'sectPr')}") is not None
                for element in [*elements, *anchor_elements]
            ):
                raise NativePackageError(
                    f"{operation_name} refuses target or anchor elements "
                    "that carry a native section boundary."
                )
            for element in elements:
                body.remove(element)
            remaining = list(body)
            anchor_indices = [
                remaining.index(element)
                for element in anchor_elements
            ]
            insert_index = (
                max(anchor_indices) + 1
                if anchor_field == "after"
                else min(anchor_indices)
            )
            for offset, element in enumerate(elements):
                body.insert(insert_index + offset, element)
            synchronize_section_start(
                change,
                operation_name=str(operation_name),
                anchor_id=anchor_id,
                new_start_id=target_id,
                new_start_elements=elements,
                can_rebind=(
                    operation_name == "node.move_before"
                ),
            )
            moved_nodes.add(target_id)
            changed_xml_parts.add("/word/document.xml")
            continue
        if operation_name == "style.apply":
            if len(elements) != 1 or elements[0].tag != _q(W, "p"):
                raise NativePackageError(
                    "style.apply requires a native reference to one w:p element."
                )
            style_ref = operation.get("style_ref")
            native_style_ref = style_ref
            if native_style_ref is None:
                target_id = _target_id(operation.get("target"))
                result_node = next(
                    (node for node in result_spec.content if node.id == target_id),
                    None,
                )
                if isinstance(result_node, Heading):
                    native_style_ref = f"Heading{result_node.level}"
            if native_style_ref is not None:
                if not isinstance(native_style_ref, str) or not native_style_ref:
                    raise NativePackageError(
                        "style.apply style_ref must be a non-empty string or null."
                    )
                assert styles_root is not None
                if find_named_style(styles_root, native_style_ref) is None:
                    raise NativePackageError(
                        f"Native DOCX has no paragraph style {native_style_ref!r}."
                    )
            patch_paragraph_style_ref(elements[0], native_style_ref)
            changed_xml_parts.add(source_ref.part_uri)
        elif operation_name == "text.replace":
            if len(elements) != 1:
                raise NativePackageError(
                    "text.replace requires a native reference to exactly one element."
                )
            search = operation.get("search")
            replacement = operation.get("replacement")
            if not isinstance(search, str) or not search or not isinstance(replacement, str):
                raise NativePackageError(
                    "text.replace requires a non-empty search and string replacement."
                )
            _replace_text_nodes(
                elements[0],
                search,
                replacement,
                replace_all=bool(operation.get("replace_all", False)),
            )
            changed_xml_parts.add(source_ref.part_uri)
        elif operation_name == "paragraph.format":
            if len(elements) != 1 or elements[0].tag != _q(W, "p"):
                raise NativePackageError(
                    "paragraph.format requires a native reference to one w:p element."
                )
            target_is_image = any(
                isinstance(node, ImageBlock)
                and node.id == target_id
                for node in spec.content
            )
            original_image = (
                simple_inline_image(
                    updated,
                    elements[0],
                    source_part=source_ref.part_uri,
                )
                if target_is_image
                else None
            )
            if target_is_image and original_image is None:
                raise NativePackageError(
                    "paragraph.format image target no longer matches its "
                    "conservative native projection."
                )
            fields = set(operation.get("set", {})) | set(operation.get("clear", []))
            try:
                style = ParagraphStyle.model_validate(operation.get("set", {}))
            except ValidationError as error:
                raise NativePackageError(
                    f"Could not lower paragraph.format values: {error}"
                ) from error
            patch_paragraph_style(elements[0], style, fields)
            if target_is_image and simple_inline_image(
                updated,
                elements[0],
                source_part=source_ref.part_uri,
            ) != original_image:
                raise NativePackageError(
                    "paragraph.format changed the projected native picture."
                )
            changed_xml_parts.add(source_ref.part_uri)
        elif operation_name == "text.format":
            if len(elements) != 1 or elements[0].tag != _q(W, "p"):
                raise NativePackageError(
                    "text.format requires a native reference to one w:p element."
                )
            fields = set(operation.get("set", {})) | set(operation.get("clear", []))
            try:
                style = TextStyle.model_validate(operation.get("set", {}))
            except ValidationError as error:
                raise NativePackageError(f"Could not lower text.format values: {error}") from error
            text = "".join(node.text or "" for node in elements[0].iter(_q(W, "t")))
            try:
                selection = resolve_text_selection(
                    text,
                    range_value=operation.get("range"),
                    match_value=operation.get("match"),
                )
            except (ValidationError, ValueError) as error:
                raise NativePackageError(
                    f"Could not resolve native text.format selection: {error}"
                ) from error
            if selection is None:
                runs = list(elements[0].iter(_q(W, "r")))
                patch_paragraph_mark_text_style(elements[0], style, fields)
                for run in runs:
                    patch_text_style(run, style, fields)
            else:
                _format_text_range(
                    elements[0],
                    start=selection.start,
                    end=selection.end,
                    style=style,
                    fields=fields,
                )
            changed_xml_parts.add(source_ref.part_uri)
        elif operation_name == "node.remove":
            if any(
                element.tag == _q(W, "sectPr")
                or element.find(f".//{_q(W, 'sectPr')}") is not None
                for element in elements
            ):
                raise NativePackageError(
                    "node.remove refuses elements that carry a native "
                    "section boundary."
                )
            for element in elements:
                if element not in list(container):
                    raise NativePackageError("DOCX node has already been removed by this patch.")
                container.remove(element)
            removed_nodes.add(target_id)
            changed_xml_parts.add(source_ref.part_uri)

    result_section_starts = {
        section.id: section.start_at
        for section in result_spec.sections
    }
    if result_section_starts != native_section_starts:
        raise NativePackageError(
            "Native structural operations did not reproduce the semantic "
            "section-start model."
        )

    current_indices = {id(element): index for index, element in enumerate(list(body))}
    identity_updates: dict[str, NativeRef] = {}
    for node_id, (elements, original_ref) in source_elements.items():
        current_container = part_containers[original_ref.part_uri]
        part_indices = {
            id(element): index
            for index, element in enumerate(list(current_container))
        }
        indices = [part_indices.get(id(element)) for element in elements]
        if any(index is None for index in indices):
            continue
        normalized_indices = [index for index in indices if index is not None]
        root_path = (
            "/w:document/w:body"
            if original_ref.part_uri == "/word/document.xml"
            else "/w:hdr"
            if part_roots[original_ref.part_uri].tag == _q(W, "hdr")
            else "/w:ftr"
        )
        identity_updates[node_id] = native_ref_for_part_elements(
            elements,
            normalized_indices,
            part_uri=original_ref.part_uri,
            native_kind=original_ref.native_kind,
            root_path=root_path,
            native_id=original_ref.native_id,
        )
    for image_id, paragraph in inserted_images.items():
        paragraph_index = current_indices.get(id(paragraph))
        if paragraph_index is None:
            raise NativePackageError(
                f"Inserted image {image_id!r} is no longer in the document body."
            )
        identity_updates[image_id] = native_ref_for_part_elements(
            [paragraph],
            [paragraph_index],
            part_uri="/word/document.xml",
            native_kind="w:p",
            root_path="/w:document/w:body",
            native_id=paragraph.get(_q(W14, "paraId")),
        )
    for field_id, (paragraph, original_ref) in source_fields.items():
        if original_ref.sub_index is None:
            continue
        current_container = part_containers[original_ref.part_uri]
        paragraph_index = next(
            (
                index
                for index, element in enumerate(list(current_container))
                if element is paragraph
            ),
            None,
        )
        if paragraph_index is None:
            continue
        try:
            match = field_match_at(paragraph, original_ref.sub_index)
        except FieldStructureError:
            continue
        root_path = (
            "/w:document/w:body"
            if original_ref.part_uri == "/word/document.xml"
            else "/w:hdr"
            if part_roots[original_ref.part_uri].tag == _q(W, "hdr")
            else "/w:ftr"
        )
        identity_updates[field_id] = native_ref_for_field(
            paragraph,
            paragraph_index,
            match,
            part_uri=original_ref.part_uri,
            root_path=root_path,
        )
    for field_id, (paragraph, field_ordinal) in inserted_fields.items():
        paragraph_index = current_indices.get(id(paragraph))
        if paragraph_index is None:
            continue
        try:
            match = field_match_at(paragraph, field_ordinal)
        except FieldStructureError:
            continue
        identity_updates[field_id] = native_ref_for_field(
            paragraph,
            paragraph_index,
            match,
            part_uri="/word/document.xml",
            root_path="/w:document/w:body",
        )
    for table_id, (source_table, table_element, source_ref) in (
        source_tables.items()
    ):
        current_container = part_containers[source_ref.part_uri]
        table_index = next(
            (
                index
                for index, element in enumerate(list(current_container))
                if element is table_element
            ),
            None,
        )
        if table_index is None:
            continue
        result_table = next(
            (
                candidate
                for candidate in result_spec.content
                if isinstance(candidate, Table)
                and candidate.id == table_id
            ),
            None,
        )
        if result_table is None:
            continue
        result_column_ids = {
            column.id for column in result_table.columns
        }
        for column_index, column in enumerate(source_table.columns):
            if (
                column.id not in result_column_ids
                or not isinstance(column.source_ref, NativeRef)
            ):
                continue
            try:
                identity_updates[column.id] = (
                    native_ref_for_table_column(
                        table_element,
                        table_index,
                        column_index,
                    )
                )
            except ValueError:
                continue
        result_row_ids = {row.id for row in result_table.rows}
        for row_index, row in enumerate(source_table.rows, start=1):
            if (
                row.id not in result_row_ids
                or not isinstance(row.source_ref, NativeRef)
            ):
                continue
            try:
                identity_updates[row.id] = native_ref_for_table_row(
                    table_element,
                    table_index,
                    row_index,
                )
            except ValueError:
                continue
        result_cell_ids = {
            cell.id
            for row in result_table.rows
            for cell in row.cells
        }
        result_paragraph_ids = {
            paragraph.id
            for row in result_table.rows
            for cell in row.cells
            for paragraph in cell.content
        }
        native_rows = table_element.findall(_q(W, "tr"))
        for source_row in source_table.rows:
            for source_cell in source_row.cells:
                cell_ref = source_cell.source_ref
                if (
                    source_cell.id not in result_cell_ids
                    or not isinstance(cell_ref, NativeRef)
                ):
                    continue
                try:
                    native_cell = table_cell_from_ref(
                        table_element,
                        cell_ref,
                    )
                except ValueError:
                    continue
                coordinates = next(
                    (
                        (native_row_index, native_cell_index)
                        for native_row_index, native_row in enumerate(
                            native_rows
                        )
                        for native_cell_index, candidate in enumerate(
                            native_row.findall(_q(W, "tc"))
                        )
                        if candidate is native_cell
                    ),
                    None,
                )
                if coordinates is None:
                    continue
                native_row_index, native_cell_index = coordinates
                identity_updates[source_cell.id] = (
                    native_ref_for_table_cell(
                        table_element,
                        table_index,
                        native_row_index,
                        native_cell_index,
                    )
                )
                for source_paragraph in source_cell.content:
                    paragraph_ref = source_paragraph.source_ref
                    if (
                        source_paragraph.id
                        not in result_paragraph_ids
                        or not isinstance(paragraph_ref, NativeRef)
                    ):
                        continue
                    try:
                        paragraph_cell, native_paragraph = (
                            table_cell_paragraph_from_ref(
                                table_element,
                                paragraph_ref,
                            )
                        )
                    except ValueError:
                        continue
                    if paragraph_cell is not native_cell:
                        continue
                    paragraphs = native_cell.findall(_q(W, "p"))
                    paragraph_index = next(
                        (
                            paragraph_index
                            for paragraph_index, candidate in enumerate(
                                paragraphs
                            )
                            if candidate is native_paragraph
                        ),
                        None,
                    )
                    if paragraph_index is None:
                        continue
                    identity_updates[source_paragraph.id] = (
                        native_ref_for_table_cell_paragraph(
                            table_element,
                            table_index,
                            native_row_index,
                            native_cell_index,
                            paragraph_index,
                        )
                    )
    for section_id, (section, container, container_kind) in source_sections.items():
        index = current_indices.get(id(container))
        if index is None:
            continue
        identity_updates[section_id] = native_ref_for_section(
            section,
            index,
            container=container_kind,
        )
    for part in spec.header_footers:
        source_ref = part.source_ref
        if not isinstance(source_ref, NativeRef):
            continue
        part_root = part_roots.get(source_ref.part_uri)
        if part_root is None:
            part_root = parse_xml(updated.get_part(source_ref.part_uri))
            part_roots[source_ref.part_uri] = part_root
        identity_updates[part.id] = native_ref_for_header_footer_part(
            part_root,
            source_ref.part_uri,
            kind=part.kind,
        )

    for part_uri in sorted(changed_xml_parts):
        updated.set_part(part_uri, serialize_xml(part_roots[part_uri]))
    if styles_changed:
        assert styles_root is not None
        updated.set_part("/word/styles.xml", serialize_xml(styles_root))
    structural_identity_required = bool(
        inserted_images
        or inserted_nodes
        or inserted_sections
        or moved_nodes
        or removed_nodes
    )
    if structural_identity_required:
        _ensure_identity_manifest_parts(updated)
    if (
        updated.has_part(MANIFEST_PART_URI)
        or structural_identity_required
    ):
        manifest_spec = result_spec.model_copy(deep=True)
        for node in manifest_spec.content:
            if node.id in identity_updates:
                node.source_ref = identity_updates[node.id]
        for section in manifest_spec.sections:
            if section.id in identity_updates:
                section.source_ref = identity_updates[section.id]
        for part in manifest_spec.header_footers:
            if part.id in identity_updates:
                part.source_ref = identity_updates[part.id]
            for block in part.content:
                if block.id in identity_updates:
                    block.source_ref = identity_updates[block.id]
        for document_field in _fields(manifest_spec):
            if document_field.id in identity_updates:
                document_field.source_ref = identity_updates[
                    document_field.id
                ]
        for table in (
            node
            for node in manifest_spec.content
            if isinstance(node, Table)
        ):
            for column in table.columns:
                if column.id in identity_updates:
                    column.source_ref = identity_updates[column.id]
            for row in table.rows:
                if row.id in identity_updates:
                    row.source_ref = identity_updates[row.id]
                for cell in row.cells:
                    if cell.id in identity_updates:
                        cell.source_ref = identity_updates[cell.id]
                    for paragraph in cell.content:
                        if paragraph.id in identity_updates:
                            paragraph.source_ref = identity_updates[
                                paragraph.id
                            ]
        updated.set_part(
            MANIFEST_PART_URI,
            serialize_identity_manifest(build_identity_manifest(manifest_spec)),
            content_type="application/xml",
        )
    return updated, updated.fidelity_report(), identity_updates


__all__ = ["apply_docx_operations"]
