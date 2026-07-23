"""Small, dependency-free WordprocessingML compiler for Document IR."""

from __future__ import annotations

import hashlib
import json
from dataclasses import dataclass, field
from io import BytesIO
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from aioffice.native.identity import (
    MANIFEST_PART_URI,
    MANIFEST_RELATIONSHIP_TYPE,
    build_identity_manifest,
    native_ref_for_elements,
    serialize_identity_manifest,
)
from aioffice.formats.docx_style import (
    apply_paragraph_mark_text_style,
    apply_paragraph_style,
    apply_text_style,
)
from aioffice.spec.models import (
    AiOfficeDocumentSpec,
    BulletList,
    Heading,
    OrderedList,
    PageBreak,
    Paragraph,
    ParagraphStyle,
    Table,
    TextSpan,
    TextStyle,
    NativeRef,
)
from aioffice.themes import get_theme

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
    ET.SubElement(parent, _q(REL, "Relationship"), values)


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
class _DocxContext:
    hyperlinks: list[tuple[str, str]] = field(default_factory=list)

    def add_hyperlink(self, target: str) -> str:
        rel_id = f"rId{len(self.hyperlinks) + 3}"
        self.hyperlinks.append((rel_id, target))
        return rel_id


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
    context: _DocxContext,
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


def _add_paragraph(
    body: ET.Element,
    spans: list[TextSpan],
    context: _DocxContext,
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
        _add_run(paragraph, span, context, default_style=text_style)
    return paragraph


def _add_table(body: ET.Element, table: Table, context: _DocxContext) -> ET.Element:
    element = _child(body, "tbl")
    properties = _child(element, "tblPr")
    _child(properties, "tblStyle", val="TableGrid")
    _child(properties, "tblW", w="0", type="auto")
    grid = _child(element, "tblGrid")
    for _ in table.columns:
        _child(grid, "gridCol", w="2400")

    def add_row(values: list[str], *, header: bool = False) -> None:
        row = _child(element, "tr")
        if header:
            row_properties = _child(row, "trPr")
            _child(row_properties, "tblHeader")
        for value in values:
            cell = _child(row, "tc")
            cell_properties = _child(cell, "tcPr")
            _child(cell_properties, "tcW", w="2400", type="dxa")
            span = TextSpan(text=value, marks=["strong"] if header else [])
            _add_paragraph(cell, [span], context)

    add_row([column.title for column in table.columns], header=True)
    for row in table.rows:
        add_row([_string_value(row.values.get(column.key)) for column in table.columns])
    return element


def _document_xml(
    spec: AiOfficeDocumentSpec,
    context: _DocxContext,
) -> tuple[bytes, dict[str, NativeRef]]:
    root = ET.Element(
        _q(W, "document"),
        {_q(MC, "Ignorable"): "w14"},
    )
    body = _child(root, "body")
    refs: dict[str, NativeRef] = {}
    for block in spec.content:
        if isinstance(block, Heading):
            element = _add_paragraph(
                body,
                [TextSpan(text=block.text)],
                context,
                style=f"Heading{block.level}",
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
        elif isinstance(block, Paragraph):
            spans = block.content if block.text is None else [TextSpan(text=block.text)]
            element = _add_paragraph(
                body,
                spans,
                context,
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
        elif isinstance(block, BulletList):
            elements = [
                _add_paragraph(
                    body,
                    [TextSpan(text=item)],
                    context,
                    numbering_id=1,
                    native_anchor=_native_anchor(block.id, item_index),
                )
                for item_index, item in enumerate(block.items)
            ]
            indices = list(range(len(body) - len(elements), len(body)))
            refs[block.id] = native_ref_for_elements(
                elements,
                indices,
                native_kind="w:p-group",
                native_id=elements[0].attrib.get(_q(W14, "paraId")),
            )
        elif isinstance(block, OrderedList):
            elements = [
                _add_paragraph(
                    body,
                    [TextSpan(text=item)],
                    context,
                    numbering_id=2,
                    native_anchor=_native_anchor(block.id, item_index),
                )
                for item_index, item in enumerate(block.items)
            ]
            indices = list(range(len(body) - len(elements), len(body)))
            refs[block.id] = native_ref_for_elements(
                elements,
                indices,
                native_kind="w:p-group",
                native_id=elements[0].attrib.get(_q(W14, "paraId")),
            )
        elif isinstance(block, Table):
            element = _add_table(body, block, context)
            index = len(body) - 1
            refs[block.id] = native_ref_for_elements(
                [element],
                [index],
                native_kind="w:tbl",
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

    section = _child(body, "sectPr")
    _child(section, "pgSz", w="12240", h="15840")
    _child(section, "pgMar", top="1440", right="1440", bottom="1440", left="1440")
    return _xml(root), refs


def _styles_xml(spec: AiOfficeDocumentSpec) -> bytes:
    theme = get_theme(spec.theme.ref) or get_theme("business-clean") or {}
    tokens = theme.get("tokens", {})
    body_latin = tokens.get("font.body.latin", "Aptos")
    body_east_asia = tokens.get("font.body.east_asia", "Microsoft YaHei")
    heading_latin = tokens.get("font.heading.latin", "Aptos Display")
    heading_east_asia = tokens.get("font.heading.east_asia", body_east_asia)
    primary = tokens.get("color.primary", "#1F4E78").lstrip("#")

    root = ET.Element(_q(W, "styles"))
    defaults = _child(root, "docDefaults")
    run_default = _child(defaults, "rPrDefault")
    run_properties = _child(run_default, "rPr")
    _child(
        run_properties,
        "rFonts",
        ascii=body_latin,
        hAnsi=body_latin,
        eastAsia=body_east_asia,
    )
    _child(run_properties, "sz", val="22")
    _child(run_properties, "szCs", val="22")
    paragraph_default = _child(defaults, "pPrDefault")
    paragraph_properties = _child(paragraph_default, "pPr")
    _child(paragraph_properties, "spacing", after="160", line="276", lineRule="auto")

    normal = _child(root, "style", type="paragraph", default="1", styleId="Normal")
    _child(normal, "name", val="Normal")
    _child(normal, "qFormat")

    sizes = {1: 36, 2: 32, 3: 28, 4: 24, 5: 22, 6: 20}
    for level in range(1, 7):
        style = _child(root, "style", type="paragraph", styleId=f"Heading{level}")
        _child(style, "name", val=f"heading {level}")
        _child(style, "basedOn", val="Normal")
        _child(style, "next", val="Normal")
        _child(style, "qFormat")
        properties = _child(style, "pPr")
        _child(properties, "keepNext")
        _child(properties, "spacing", before=str(240 if level <= 2 else 160), after="120")
        _child(properties, "outlineLvl", val=str(level - 1))
        run_properties = _child(style, "rPr")
        _child(
            run_properties,
            "rFonts",
            ascii=heading_latin,
            hAnsi=heading_latin,
            eastAsia=heading_east_asia,
        )
        _child(run_properties, "b")
        _child(run_properties, "color", val=primary)
        _child(run_properties, "sz", val=str(sizes[level]))
        _child(run_properties, "szCs", val=str(sizes[level]))

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


def _numbering_xml() -> bytes:
    root = ET.Element(_q(W, "numbering"))
    for abstract_id, format_name, text_value in ((0, "bullet", "•"), (1, "decimal", "%1.")):
        abstract = _child(root, "abstractNum", abstractNumId=str(abstract_id))
        _child(abstract, "multiLevelType", val="singleLevel")
        level = _child(abstract, "lvl", ilvl="0")
        _child(level, "start", val="1")
        _child(level, "numFmt", val=format_name)
        _child(level, "lvlText", val=text_value)
        _child(level, "lvlJc", val="left")
        paragraph_properties = _child(level, "pPr")
        _child(paragraph_properties, "tabs").append(
            ET.Element(_q(W, "tab"), {_q(W, "val"): "num", _q(W, "pos"): "720"})
        )
        _child(paragraph_properties, "ind", left="720", hanging="360")
    for num_id, abstract_id in ((1, 0), (2, 1)):
        number = _child(root, "num", numId=str(num_id))
        _child(number, "abstractNumId", val=str(abstract_id))
    return _xml(root)


def _content_types_xml() -> bytes:
    root = ET.Element(_q(CT, "Types"))
    ET.SubElement(
        root,
        _q(CT, "Default"),
        Extension="rels",
        ContentType="application/vnd.openxmlformats-package.relationships+xml",
    )
    ET.SubElement(root, _q(CT, "Default"), Extension="xml", ContentType="application/xml")
    overrides = {
        "/word/document.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
        "/word/styles.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
        "/word/numbering.xml": "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
        "/docProps/core.xml": "application/vnd.openxmlformats-package.core-properties+xml",
        "/docProps/app.xml": "application/vnd.openxmlformats-officedocument.extended-properties+xml",
    }
    for part_name, content_type in overrides.items():
        ET.SubElement(root, _q(CT, "Override"), PartName=part_name, ContentType=content_type)
    return _xml(root)


def _root_relationships_xml() -> bytes:
    root = ET.Element(_q(REL, "Relationships"))
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


def _document_relationships_xml(context: _DocxContext) -> bytes:
    root = ET.Element(_q(REL, "Relationships"))
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

    context = _DocxContext()
    document_xml, refs = _document_xml(spec, context)
    identity_manifest = build_identity_manifest(spec, refs=refs)
    stream = BytesIO()
    with ZipFile(stream, mode="w") as archive:
        parts = {
            "[Content_Types].xml": _content_types_xml(),
            "_rels/.rels": _root_relationships_xml(),
            MANIFEST_PART_URI.lstrip("/"): serialize_identity_manifest(identity_manifest),
            "docProps/app.xml": _app_properties_xml(),
            "docProps/core.xml": _core_properties_xml(spec),
            "word/document.xml": document_xml,
            "word/_rels/document.xml.rels": _document_relationships_xml(context),
            "word/numbering.xml": _numbering_xml(),
            "word/styles.xml": _styles_xml(spec),
        }
        for path in sorted(parts):
            _write_part(archive, path, parts[path])
    return stream.getvalue()


def export_docx(spec: AiOfficeDocumentSpec, target: str | Path) -> Path:
    path = Path(target)
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(compile_docx(spec))
    return path


__all__ = ["compile_docx", "export_docx"]
