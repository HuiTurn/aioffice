"""Small, dependency-free WordprocessingML compiler for Document IR."""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Any, Mapping, Sequence
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from aioffice.native.identity import (
    MANIFEST_PART_URI,
    MANIFEST_RELATIONSHIP_TYPE,
    build_identity_manifest,
    native_ref_for_elements,
    native_ref_for_part_elements,
    serialize_identity_manifest,
)
from aioffice.formats.docx_header_footer import (
    FOOTER_CONTENT_TYPE,
    FOOTER_RELATIONSHIP_TYPE,
    HEADER_CONTENT_TYPE,
    HEADER_RELATIONSHIP_TYPE,
    SETTINGS_CONTENT_TYPE,
    SETTINGS_RELATIONSHIP_TYPE,
    apply_header_footer_bindings,
    native_ref_for_header_footer_part,
    settings_xml,
)
from aioffice.formats.docx_fields import (
    append_complex_field,
    native_ref_for_field,
    parse_paragraph_fields,
)
from aioffice.formats.docx_style import (
    apply_paragraph_mark_text_style,
    apply_paragraph_style,
    apply_text_style,
)
from aioffice.formats.docx_named_styles import upsert_named_style
from aioffice.formats.docx_section import (
    apply_section_layout,
    native_ref_for_section,
)
from aioffice.formats.docx_tables import (
    apply_table_cell_format,
    apply_table_layout,
    apply_table_row,
    native_ref_for_table_cell,
    native_ref_for_table_cell_paragraph,
    native_ref_for_table_column,
    native_ref_for_table_row,
)
from aioffice.spec.models import (
    AiOfficeDocumentSpec,
    BulletList,
    DocumentField,
    Heading,
    HeaderFooterPart,
    ImageBlock,
    InlineContent,
    OpaqueBlock,
    OrderedList,
    PageBreak,
    Paragraph,
    ParagraphStyle,
    Table,
    TableCell,
    TableWidth,
    TextSpan,
    TextStyle,
    NativeRef,
)
from aioffice.styles import (
    resolve_document_defaults,
    style_catalog,
    theme_named_styles,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
CP = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
DC = "http://purl.org/dc/elements/1.1/"
DCTERMS = "http://purl.org/dc/terms/"
DCTERMS_XSI = "http://www.w3.org/2001/XMLSchema-instance"
EP = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

ET.register_namespace("w", W)
ET.register_namespace("w14", W14)
ET.register_namespace("mc", MC)
ET.register_namespace("r", R)
ET.register_namespace("cp", CP)
ET.register_namespace("dc", DC)
ET.register_namespace("dcterms", DCTERMS)
ET.register_namespace("xsi", DCTERMS_XSI)
ET.register_namespace("vt", VT)


def _q(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def _xml(element: ET.Element) -> bytes:
    return ET.tostring(element, encoding="utf-8", xml_declaration=True)


def _child(parent: ET.Element, name: str, **attributes: str) -> ET.Element:
    return ET.SubElement(
        parent, _q(W, name), {_q(W, key): value for key, value in attributes.items()}
    )


def _relationship(
    parent: ET.Element, rel_id: str, rel_type: str, target: str, **attrs: str
) -> None:
    values = {"Id": rel_id, "Type": rel_type, "Target": target, **attrs}
    ET.SubElement(parent, "Relationship", values)


def _string_value(value: Any) -> str:
    if value is None:
        return ""
    if isinstance(value, bool):
        return "true" if value else "false"
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False, separators=(",", ":"))
    return str(value)


def _native_anchor(node_id: str, ordinal: int = 0) -> str:
    value = hashlib.sha256(f"{node_id}:{ordinal}".encode()).hexdigest()[:8].upper()
    return value if value != "00000000" else "00000001"


@dataclass
class DocxCompileContext:
    hyperlinks: list[tuple[str, str]] = field(default_factory=list)

    def add_hyperlink(self, target: str) -> str:
        rel_id = f"rId{len(self.hyperlinks) + 3}"
        self.hyperlinks.append((rel_id, target))
        return rel_id


@dataclass
class _CompiledHeaderFooters:
    parts: dict[str, bytes] = field(default_factory=dict)
    content_types: dict[str, str] = field(default_factory=dict)
    document_relationships: list[tuple[str, str, str]] = field(
        default_factory=list
    )
    relationship_ids: dict[str, str] = field(default_factory=dict)
    refs: dict[str, NativeRef] = field(default_factory=dict)


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


def _add_run(
    parent: ET.Element,
    span: TextSpan,
    context: DocxCompileContext,
    *,
    default_style: TextStyle | None = None,
) -> None:
    run_parent = parent
    if "link" in span.marks and span.href:
        run_parent = ET.SubElement(
            parent,
            _q(W, "hyperlink"),
            {_q(R, "id"): context.add_hyperlink(span.href)},
        )
    run = _child(run_parent, "r")
    apply_text_style(run, _merge_text_style(default_style, span.style))
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
            apply_text_style(run, TextStyle.model_validate(mark_style))
        properties = run.find(_q(W, "rPr"))
        if properties is None:
            properties = ET.Element(_q(W, "rPr"))
            run.insert(0, properties)
        if "highlight" in span.marks:
            _child(properties, "highlight", val="yellow")
        if "link" in span.marks:
            _child(properties, "rStyle", val="Hyperlink")
    text = _child(run, "t")
    if span.text[:1].isspace() or span.text[-1:].isspace() or "  " in span.text:
        text.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
    text.text = span.text


def _add_inline(
    parent: ET.Element,
    inline: InlineContent,
    context: DocxCompileContext,
    *,
    default_style: TextStyle | None = None,
) -> None:
    if isinstance(inline, TextSpan):
        _add_run(
            parent,
            inline,
            context,
            default_style=default_style,
        )
        return
    append_complex_field(
        parent,
        inline,
        effective_style=_merge_text_style(default_style, inline.style),
    )


def _add_paragraph(
    body: ET.Element,
    spans: Sequence[InlineContent],
    context: DocxCompileContext,
    *,
    style: str | None = None,
    numbering_id: int | None = None,
    native_anchor: str | None = None,
    paragraph_style: ParagraphStyle | None = None,
    text_style: TextStyle | None = None,
) -> ET.Element:
    attributes = {_q(W14, "paraId"): native_anchor} if native_anchor else {}
    paragraph = ET.SubElement(body, _q(W, "p"), attributes)
    if style or numbering_id is not None:
        properties = _child(paragraph, "pPr")
        if style:
            _child(properties, "pStyle", val=style)
        if numbering_id is not None:
            numbering = _child(properties, "numPr")
            _child(numbering, "ilvl", val="0")
            _child(numbering, "numId", val=str(numbering_id))
    apply_paragraph_style(paragraph, paragraph_style)
    apply_paragraph_mark_text_style(paragraph, text_style)
    for span in spans:
        _add_inline(
            paragraph,
            span,
            context,
            default_style=text_style,
        )
    return paragraph


def _register_field_refs(
    refs: dict[str, NativeRef],
    paragraph: ET.Element,
    paragraph_index: int,
    spans: Sequence[InlineContent],
    *,
    part_uri: str,
    root_path: str,
) -> None:
    fields = [span for span in spans if isinstance(span, DocumentField)]
    if not fields:
        return
    matches = parse_paragraph_fields(paragraph)
    if len(matches) != len(fields):
        raise ValueError("Generated DOCX field count does not match semantic fields.")
    for document_field, match in zip(fields, matches):
        refs[document_field.id] = native_ref_for_field(
            paragraph,
            paragraph_index,
            match,
            part_uri=part_uri,
            root_path=root_path,
        )


def compile_list_elements(
    body: ET.Element,
    block: BulletList | OrderedList,
    context: DocxCompileContext,
    *,
    numbering_id: int,
    native_anchors: Sequence[str] | None = None,
) -> list[ET.Element]:
    anchors = (
        list(native_anchors)
        if native_anchors is not None
        else [
            _native_anchor(block.id, item_index)
            for item_index in range(len(block.items))
        ]
    )
    if len(anchors) != len(block.items):
        raise ValueError(
            "List item count does not match native paragraph anchors."
        )
    return [
        _add_paragraph(
            body,
            [TextSpan(text=item)],
            context,
            numbering_id=numbering_id,
            native_anchor=anchors[item_index],
        )
        for item_index, item in enumerate(block.items)
    ]


def compile_table_element(
    body: ET.Element,
    table: Table,
    context: DocxCompileContext,
) -> ET.Element:
    element = _child(body, "tbl")
    grid = _child(element, "tblGrid")
    native_widths = [
        (
            str(round(column.width.to_points() * 20))
            if column.width is not None
            else "2400"
        )
        for column in table.columns
    ]
    for width in native_widths:
        _child(grid, "gridCol", w=width)

    column_indices = {
        column.key: index
        for index, column in enumerate(table.columns)
    }

    def add_header_row() -> None:
        row = _child(element, "tr")
        for column_index, column in enumerate(table.columns):
            cell = _child(row, "tc")
            cell_properties = _child(cell, "tcPr")
            _child(
                cell_properties,
                "tcW",
                w=native_widths[column_index],
                type="dxa",
            )
            span = TextSpan(text=column.title, marks=["strong"])
            _add_paragraph(cell, [span], context)

    placements = _semantic_table_cell_placements(table)
    add_header_row()
    for row_index, row_placements in enumerate(placements):
        row = _child(element, "tr")
        apply_table_row(row, table.rows[row_index])
        for table_cell, continuation in row_placements:
            start_column = column_indices[table_cell.column_key]
            cell = _child(row, "tc")
            cell_properties = _child(cell, "tcPr")
            cell_width = sum(
                int(native_widths[index])
                for index in range(
                    start_column,
                    start_column + table_cell.column_span,
                )
            )
            _child(
                cell_properties,
                "tcW",
                w=str(cell_width),
                type="dxa",
            )
            if table_cell.column_span > 1:
                _child(
                    cell_properties,
                    "gridSpan",
                    val=str(table_cell.column_span),
                )
            if continuation:
                _child(cell_properties, "vMerge", val="continue")
                _add_paragraph(cell, [], context)
                continue
            if table_cell.row_span > 1:
                _child(cell_properties, "vMerge", val="restart")
            apply_table_cell_format(cell, table_cell.format)
            if table_cell.content:
                for paragraph in table_cell.content:
                    spans = (
                        paragraph.content
                        if paragraph.text is None
                        else [TextSpan(text=paragraph.text)]
                    )
                    _add_paragraph(
                        cell,
                        spans,
                        context,
                        style=paragraph.style_ref,
                        native_anchor=_native_anchor(paragraph.id),
                        paragraph_style=paragraph.paragraph_style,
                        text_style=paragraph.text_style,
                    )
            else:
                _add_paragraph(
                    cell,
                    [TextSpan(text=table_cell.plain_text)],
                    context,
                )
    effective_layout = table.layout.model_copy(
        update={
            "style_ref": table.layout.style_ref or "TableGrid",
            "preferred_width": (
                table.layout.preferred_width
                or TableWidth(mode="auto")
            ),
            "repeat_header": (
                True
                if table.layout.repeat_header is None
                else table.layout.repeat_header
            ),
        }
    )
    apply_table_layout(element, effective_layout)
    return element


def _semantic_table_cell_placements(
    table: Table,
) -> list[list[tuple[TableCell, bool]]]:
    column_indices = {
        column.key: index
        for index, column in enumerate(table.columns)
    }
    active: dict[int, tuple[TableCell, int]] = {}
    result: list[list[tuple[TableCell, bool]]] = []
    for row_index, row in enumerate(table.rows):
        anchors: dict[int, TableCell] = {}
        for cell in row.cells:
            try:
                start_column = column_indices[cell.column_key]
            except KeyError as error:
                raise ValueError(
                    f"Table cell {cell.id!r} references unknown column "
                    f"{cell.column_key!r}."
                ) from error
            if start_column in anchors:
                raise ValueError(
                    f"Table row {row.id!r} contains multiple cells anchored "
                    f"to {cell.column_key!r}."
                )
            anchors[start_column] = cell

        row_result: list[tuple[TableCell, bool]] = []
        next_active: dict[int, tuple[TableCell, int]] = {}
        column_index = 0
        while column_index < len(table.columns):
            if column_index in active:
                cell, remaining = active[column_index]
                if column_index in anchors:
                    raise ValueError(
                        f"Table row {row.id!r} overlaps row-spanning cell "
                        f"{cell.id!r}."
                    )
                row_result.append((cell, True))
                if remaining > 1:
                    next_active[column_index] = (cell, remaining - 1)
                column_index += cell.column_span
                continue
            cell = anchors.get(column_index)
            if cell is None:
                raise ValueError(
                    f"Table row {row.id!r} leaves logical column "
                    f"{table.columns[column_index].key!r} uncovered."
                )
            if column_index + cell.column_span > len(table.columns):
                raise ValueError(
                    f"Table cell {cell.id!r} spans beyond the table grid."
                )
            row_result.append((cell, False))
            if cell.row_span > 1:
                if row_index + cell.row_span > len(table.rows):
                    raise ValueError(
                        f"Table cell {cell.id!r} spans beyond the final row."
                    )
                next_active[column_index] = (
                    cell,
                    cell.row_span - 1,
                )
            column_index += cell.column_span
        unexpected = sorted(set(anchors) - {
            column_indices[cell.column_key]
            for cell, continuation in row_result
            if not continuation
        })
        if unexpected:
            raise ValueError(
                f"Table row {row.id!r} contains overlapping cell anchors."
            )
        result.append(row_result)
        active = next_active
    return result


def register_table_refs(
    refs: dict[str, NativeRef],
    element: ET.Element,
    table_index: int,
    table: Table,
) -> None:
    for column_index, column in enumerate(table.columns):
        refs[column.id] = native_ref_for_table_column(
            element,
            table_index,
            column_index,
        )
    for row_index, row in enumerate(table.rows, start=1):
        refs[row.id] = native_ref_for_table_row(
            element,
            table_index,
            row_index,
        )
    placements = _semantic_table_cell_placements(table)
    for semantic_row_index, row_placements in enumerate(placements):
        physical_row_index = semantic_row_index + 1
        physical_cell_index = 0
        for table_cell, continuation in row_placements:
            if not continuation:
                refs[table_cell.id] = native_ref_for_table_cell(
                    element,
                    table_index,
                    physical_row_index,
                    physical_cell_index,
                )
                for paragraph_index, paragraph in enumerate(
                    table_cell.content
                ):
                    refs[paragraph.id] = (
                        native_ref_for_table_cell_paragraph(
                            element,
                            table_index,
                            physical_row_index,
                            physical_cell_index,
                            paragraph_index,
                        )
                    )
            physical_cell_index += 1


def _document_xml(
    spec: AiOfficeDocumentSpec,
    context: DocxCompileContext,
    header_footer_relationship_ids: dict[str, str],
) -> tuple[bytes, dict[str, NativeRef]]:
    root = ET.Element(
        _q(W, "document"),
        {_q(MC, "Ignorable"): "w14"},
    )
    body = _child(root, "body")
    refs: dict[str, NativeRef] = {}
    sections = list(spec.sections)
    section_index = 0
    active_section = sections[0]
    section_starts = {
        section.start_at: index
        for index, section in enumerate(sections[1:], start=1)
        if section.start_at is not None
    }
    for block in spec.content:
        starting_section_index = section_starts.get(block.id)
        if starting_section_index is not None:
            if starting_section_index != section_index + 1:
                raise ValueError(
                    "Document sections must be ordered by their content anchors."
                )
            carrier = _child(body, "p")
            carrier_properties = _child(carrier, "pPr")
            carrier_section = _child(carrier_properties, "sectPr")
            apply_header_footer_bindings(
                carrier_section,
                active_section.header_footer,
                header_footer_relationship_ids,
            )
            apply_section_layout(carrier_section, active_section.layout)
            refs[active_section.id] = native_ref_for_section(
                carrier_section,
                len(body) - 1,
                container="paragraph",
            )
            section_index = starting_section_index
            active_section = sections[section_index]
        if isinstance(block, Heading):
            spans = block.content if block.text is None else [TextSpan(text=block.text)]
            element = _add_paragraph(
                body,
                spans,
                context,
                style=block.style_ref or f"Heading{block.level}",
                native_anchor=_native_anchor(block.id),
                paragraph_style=block.paragraph_style,
                text_style=block.text_style,
            )
            index = len(body) - 1
            refs[block.id] = native_ref_for_elements(
                [element],
                [index],
                native_kind="w:p",
                native_id=element.attrib.get(_q(W14, "paraId")),
            )
            _register_field_refs(
                refs,
                element,
                index,
                spans,
                part_uri="/word/document.xml",
                root_path="/w:document/w:body",
            )
        elif isinstance(block, Paragraph):
            spans = block.content if block.text is None else [TextSpan(text=block.text)]
            element = _add_paragraph(
                body,
                spans,
                context,
                style=block.style_ref,
                native_anchor=_native_anchor(block.id),
                paragraph_style=block.paragraph_style,
                text_style=block.text_style,
            )
            index = len(body) - 1
            refs[block.id] = native_ref_for_elements(
                [element],
                [index],
                native_kind="w:p",
                native_id=element.attrib.get(_q(W14, "paraId")),
            )
            _register_field_refs(
                refs,
                element,
                index,
                spans,
                part_uri="/word/document.xml",
                root_path="/w:document/w:body",
            )
        elif isinstance(block, (BulletList, OrderedList)):
            elements = compile_list_elements(
                body,
                block,
                context,
                numbering_id=(
                    1 if isinstance(block, BulletList) else 2
                ),
            )
            indices = list(range(len(body) - len(elements), len(body)))
            refs[block.id] = native_ref_for_elements(
                elements,
                indices,
                native_kind="w:p-group",
                native_id=elements[0].attrib.get(_q(W14, "paraId")),
            )
        elif isinstance(block, Table):
            element = compile_table_element(body, block, context)
            index = len(body) - 1
            refs[block.id] = native_ref_for_elements(
                [element],
                [index],
                native_kind="w:tbl",
            )
            register_table_refs(refs, element, index, block)
        elif isinstance(block, ImageBlock):
            raise ValueError(
                "Semantic DOCX generation cannot compile a native-only "
                "image block; preserve or attach its source DOCX package."
            )
        elif isinstance(block, OpaqueBlock):
            raise ValueError(
                "Semantic DOCX generation cannot compile opaque body content."
            )
        elif isinstance(block, PageBreak):
            paragraph = ET.SubElement(
                body,
                _q(W, "p"),
                {_q(W14, "paraId"): _native_anchor(block.id)},
            )
            run = _child(paragraph, "r")
            _child(run, "br", type="page")
            index = len(body) - 1
            refs[block.id] = native_ref_for_elements(
                [paragraph],
                [index],
                native_kind="w:page-break",
                native_id=paragraph.attrib.get(_q(W14, "paraId")),
            )

    if section_index + 1 != len(sections):
        raise ValueError("Every document section after the first requires a valid start_at.")
    section = _child(body, "sectPr")
    apply_header_footer_bindings(
        section,
        active_section.header_footer,
        header_footer_relationship_ids,
    )
    apply_section_layout(section, active_section.layout)
    refs[active_section.id] = native_ref_for_section(
        section,
        len(body) - 1,
        container="body",
    )
    return _xml(root), refs


def compile_header_footer_part(
    part: HeaderFooterPart,
    *,
    part_uri: str,
    native_anchors: Mapping[str, str] | None = None,
) -> tuple[bytes, dict[str, NativeRef], bytes | None]:
    """Compile one reusable header/footer part and its local relationships."""

    root_name = "hdr" if part.kind == "header" else "ftr"
    root = ET.Element(
        _q(W, root_name),
        {_q(MC, "Ignorable"): "w14"},
    )
    context = DocxCompileContext()
    refs: dict[str, NativeRef] = {
        part.id: native_ref_for_header_footer_part(
            root,
            part_uri,
            kind=part.kind,
        )
    }
    if not part.content:
        _add_paragraph(root, [], context)
    for block in part.content:
        if isinstance(block, (ImageBlock, OpaqueBlock)):
            raise ValueError(
                "Semantic DOCX generation cannot compile native image or "
                "opaque header/footer content."
            )
        spans = block.content if block.text is None else [TextSpan(text=block.text)]
        element = _add_paragraph(
            root,
            spans,
            context,
            style=block.style_ref,
            native_anchor=(
                native_anchors.get(
                    block.id,
                    _native_anchor(block.id),
                )
                if native_anchors is not None
                else _native_anchor(block.id)
            ),
            paragraph_style=block.paragraph_style,
            text_style=block.text_style,
        )
        index = len(root) - 1
        refs[block.id] = native_ref_for_part_elements(
            [element],
            [index],
            part_uri=part_uri,
            native_kind="w:p",
            root_path=f"/w:{root_name}",
            native_id=element.attrib.get(_q(W14, "paraId")),
        )
        _register_field_refs(
            refs,
            element,
            index,
            spans,
            part_uri=part_uri,
            root_path=f"/w:{root_name}",
        )
    refs[part.id] = native_ref_for_header_footer_part(
        root,
        part_uri,
        kind=part.kind,
    )
    relationships = (
        _hyperlink_relationships_xml(context) if context.hyperlinks else None
    )
    return _xml(root), refs, relationships


def _compile_header_footers(spec: AiOfficeDocumentSpec) -> _CompiledHeaderFooters:
    compiled = _CompiledHeaderFooters()
    counters = {"header": 0, "footer": 0}
    for index, part in enumerate(spec.header_footers, start=1):
        counters[part.kind] += 1
        number = counters[part.kind]
        filename = f"{part.kind}{number}.xml"
        part_uri = f"/word/{filename}"
        relationship_id = f"rIdHeaderFooter{index}"
        relationship_type = (
            HEADER_RELATIONSHIP_TYPE
            if part.kind == "header"
            else FOOTER_RELATIONSHIP_TYPE
        )
        content_type = (
            HEADER_CONTENT_TYPE
            if part.kind == "header"
            else FOOTER_CONTENT_TYPE
        )
        payload, refs, relationships = compile_header_footer_part(
            part,
            part_uri=part_uri,
        )
        compiled.parts[part_uri.lstrip("/")] = payload
        compiled.content_types[part_uri] = content_type
        compiled.document_relationships.append(
            (relationship_id, relationship_type, filename)
        )
        compiled.relationship_ids[part.id] = relationship_id
        compiled.refs.update(refs)
        if relationships is not None:
            compiled.parts[
                f"word/_rels/{filename}.rels"
            ] = relationships
    return compiled


def _styles_xml(spec: AiOfficeDocumentSpec) -> bytes:
    root = ET.Element(_q(W, "styles"))
    defaults = _child(root, "docDefaults")
    run_default = _child(defaults, "rPrDefault")
    paragraph_default = _child(defaults, "pPrDefault")
    resolved_defaults = resolve_document_defaults(spec)
    apply_text_style(run_default, resolved_defaults.text_style)
    apply_paragraph_style(paragraph_default, resolved_defaults.paragraph_style)

    theme_style_ids = {style.id for style in theme_named_styles(spec.theme.ref)}
    for named_style in style_catalog(spec).values():
        element = upsert_named_style(
            root,
            named_style,
            custom_style=named_style.id not in theme_style_ids,
        )
        if named_style.id == "Normal":
            element.set(_q(W, "default"), "1")

    hyperlink = _child(root, "style", type="character", styleId="Hyperlink")
    _child(hyperlink, "name", val="Hyperlink")
    hyperlink_run = _child(hyperlink, "rPr")
    _child(hyperlink_run, "color", val="0563C1")
    _child(hyperlink_run, "u", val="single")

    table_grid = _child(root, "style", type="table", styleId="TableGrid")
    _child(table_grid, "name", val="Table Grid")
    _child(table_grid, "basedOn", val="TableNormal")
    borders = _child(_child(table_grid, "tblPr"), "tblBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        _child(borders, edge, val="single", sz="4", space="0", color="B8C2CC")
    return _xml(root)


def append_single_level_numbering_definition(
    root: ET.Element,
    *,
    abstract_id: int,
    num_id: int,
    ordered: bool,
) -> None:
    format_name = "decimal" if ordered else "bullet"
    text_value = "%1." if ordered else "\uf0b7"
    abstract = ET.Element(
        _q(W, "abstractNum"),
        {_q(W, "abstractNumId"): str(abstract_id)},
    )
    _child(abstract, "multiLevelType", val="singleLevel")
    level = _child(abstract, "lvl", ilvl="0")
    _child(level, "start", val="1")
    _child(level, "numFmt", val=format_name)
    _child(level, "lvlText", val=text_value)
    _child(level, "lvlJc", val="left")
    paragraph_properties = _child(level, "pPr")
    _child(paragraph_properties, "tabs").append(
        ET.Element(
            _q(W, "tab"),
            {
                _q(W, "val"): "num",
                _q(W, "pos"): "720",
            },
        )
    )
    _child(paragraph_properties, "ind", left="720", hanging="360")
    if not ordered:
        run_properties = _child(level, "rPr")
        _child(
            run_properties,
            "rFonts",
            ascii="Symbol",
            hAnsi="Symbol",
            hint="default",
        )
    abstract_insert_index = next(
        (
            index
            for index, child in enumerate(root)
            if child.tag
            in {
                _q(W, "num"),
                _q(W, "numIdMacAtCleanup"),
            }
        ),
        len(root),
    )
    root.insert(abstract_insert_index, abstract)
    number = ET.Element(
        _q(W, "num"),
        {_q(W, "numId"): str(num_id)},
    )
    _child(number, "abstractNumId", val=str(abstract_id))
    cleanup_index = next(
        (
            index
            for index, child in enumerate(root)
            if child.tag == _q(W, "numIdMacAtCleanup")
        ),
        len(root),
    )
    root.insert(cleanup_index, number)


def _numbering_xml() -> bytes:
    root = ET.Element(_q(W, "numbering"))
    append_single_level_numbering_definition(
        root,
        abstract_id=0,
        num_id=1,
        ordered=False,
    )
    append_single_level_numbering_definition(
        root,
        abstract_id=1,
        num_id=2,
        ordered=True,
    )
    return _xml(root)


def _content_types_xml(
    extra_overrides: dict[str, str] | None = None,
) -> bytes:
    root = ET.Element("Types", {"xmlns": CT})
    ET.SubElement(
        root,
        "Default",
        Extension="rels",
        ContentType="application/vnd.openxmlformats-package.relationships+xml",
    )
    ET.SubElement(root, "Default", Extension="xml", ContentType="application/xml")
    overrides = {
        "/word/document.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        "/word/styles.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        "/word/numbering.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        "/docProps/core.xml": "application/vnd.openxmlformats-package.core-properties+xml",
        "/docProps/app.xml": "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    }
    overrides.update(extra_overrides or {})
    for part_name, content_type in overrides.items():
        ET.SubElement(root, "Override", PartName=part_name, ContentType=content_type)
    return _xml(root)


def _root_relationships_xml() -> bytes:
    root = ET.Element("Relationships", {"xmlns": REL})
    _relationship(
        root,
        "rId1",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument",
        "word/document.xml",
    )
    _relationship(
        root,
        "rId2",
        "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties",
        "docProps/core.xml",
    )
    _relationship(
        root,
        "rId3",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties",
        "docProps/app.xml",
    )
    _relationship(
        root,
        "rId4",
        MANIFEST_RELATIONSHIP_TYPE,
        MANIFEST_PART_URI.lstrip("/"),
    )
    return _xml(root)


def _hyperlink_relationships_xml(
    context: DocxCompileContext,
) -> bytes:
    root = ET.Element("Relationships", {"xmlns": REL})
    for rel_id, target in context.hyperlinks:
        _relationship(
            root,
            rel_id,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            target,
            TargetMode="External",
        )
    return _xml(root)


def _document_relationships_xml(
    context: DocxCompileContext,
    header_footer_relationships: list[tuple[str, str, str]],
    *,
    has_settings: bool,
) -> bytes:
    root = ET.Element("Relationships", {"xmlns": REL})
    _relationship(
        root,
        "rId1",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles",
        "styles.xml",
    )
    _relationship(
        root,
        "rId2",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering",
        "numbering.xml",
    )
    if has_settings:
        _relationship(
            root,
            "rIdSettings",
            SETTINGS_RELATIONSHIP_TYPE,
            "settings.xml",
        )
    for rel_id, rel_type, target in header_footer_relationships:
        _relationship(root, rel_id, rel_type, target)
    for rel_id, target in context.hyperlinks:
        _relationship(
            root,
            rel_id,
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            target,
            TargetMode="External",
        )
    return _xml(root)


def _core_properties_xml(spec: AiOfficeDocumentSpec) -> bytes:
    root = ET.Element(_q(CP, "coreProperties"))
    if spec.metadata.title:
        ET.SubElement(root, _q(DC, "title")).text = spec.metadata.title
    if spec.metadata.author:
        ET.SubElement(root, _q(DC, "creator")).text = spec.metadata.author
        ET.SubElement(root, _q(CP, "lastModifiedBy")).text = spec.metadata.author
    if spec.metadata.subject:
        ET.SubElement(root, _q(DC, "subject")).text = spec.metadata.subject
    if spec.metadata.keywords:
        ET.SubElement(root, _q(CP, "keywords")).text = ", ".join(spec.metadata.keywords)
    ET.SubElement(root, _q(CP, "revision")).text = str(spec.artifact.revision)
    return _xml(root)


def _app_properties_xml() -> bytes:
    root = ET.Element(_q(EP, "Properties"))
    ET.SubElement(root, _q(EP, "Application")).text = "AiOffice"
    ET.SubElement(root, _q(EP, "AppVersion")).text = "0.2"
    return _xml(root)


def _write_part(archive: ZipFile, path: str, data: bytes) -> None:
    info = ZipInfo(path, date_time=(1980, 1, 1, 0, 0, 0))
    info.compress_type = ZIP_DEFLATED
    info.external_attr = 0o600 << 16
    archive.writestr(info, data)


def compile_docx(spec: AiOfficeDocumentSpec) -> bytes:
    """Compile a validated document into a deterministic OOXML package."""

    context = DocxCompileContext()
    header_footers = _compile_header_footers(spec)
    document_xml, refs = _document_xml(
        spec,
        context,
        header_footers.relationship_ids,
    )
    refs.update(header_footers.refs)
    identity_manifest = build_identity_manifest(spec, refs=refs)
    has_fields = any(
        isinstance(inline, DocumentField)
        for block in [
            *spec.content,
            *(block for part in spec.header_footers for block in part.content),
        ]
        if isinstance(block, (Heading, Paragraph))
        for inline in block.content
    )
    even_and_odd_headers = (
        spec.settings.even_and_odd_headers
        if spec.settings is not None
        else None
    )
    update_fields_on_open = (
        spec.settings.update_fields_on_open
        if spec.settings is not None
        and spec.settings.update_fields_on_open is not None
        else True
        if has_fields
        else None
    )
    settings_payload = (
        settings_xml(
            even_and_odd_headers=even_and_odd_headers,
            update_fields_on_open=update_fields_on_open,
        )
        if even_and_odd_headers is not None
        or update_fields_on_open is not None
        else None
    )
    content_type_overrides = dict(header_footers.content_types)
    if settings_payload is not None:
        content_type_overrides["/word/settings.xml"] = SETTINGS_CONTENT_TYPE
    stream = BytesIO()
    with ZipFile(stream, mode="w") as archive:
        parts = {
            "[Content_Types].xml": _content_types_xml(content_type_overrides),
            "_rels/.rels": _root_relationships_xml(),
            MANIFEST_PART_URI.lstrip("/"): serialize_identity_manifest(identity_manifest),
            "docProps/app.xml": _app_properties_xml(),
            "docProps/core.xml": _core_properties_xml(spec),
            "word/document.xml": document_xml,
            "word/_rels/document.xml.rels": _document_relationships_xml(
                context,
                header_footers.document_relationships,
                has_settings=settings_payload is not None,
            ),
            "word/numbering.xml": _numbering_xml(),
            "word/styles.xml": _styles_xml(spec),
            **header_footers.parts,
        }
        if settings_payload is not None:
            parts["word/settings.xml"] = settings_payload
        for path in sorted(parts):
            _write_part(archive, path, parts[path])
    return stream.getvalue()


def export_docx(spec: AiOfficeDocumentSpec, target: str | Path) -> Path:
    path = Path(target)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(compile_docx(spec))
    return path


__all__ = ["compile_docx", "export_docx"]
