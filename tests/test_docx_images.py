from __future__ import annotations

import base64
import copy
import hashlib
import io
import json
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from aioffice.cli.main import main
from aioffice.core.errors import NativePackageError
from aioffice.documents import Document, DocumentBuilder
from aioffice.native import MANIFEST_RELATIONSHIP_TYPE
from aioffice.native.xml import parse_xml, serialize_xml
from aioffice.security import SecurityPolicy
from aioffice.workspace import Workspace

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
VML = "urn:schemas-microsoft-com:vml"
OFFICE = "urn:schemas-microsoft-com:office:office"
WORD_VML = "urn:schemas-microsoft-com:office:word"
WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
IMAGE_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)

PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADUlEQVR42mP8"
    "z8BQDwAFgAIB/75jfwAAAABJRU5ErkJggg=="
)
JPEG = base64.b64decode(
    "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQE"
    "BQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/"
    "2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQU"
    "FBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wAARCAABAAIDASIAAhEBAxEB/8QA"
    "HwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFB"
    "AQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFh"
    "cYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd"
    "4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJ"
    "ytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/8QAHwEAAwEBAQEBAQEBA"
    "QAAAAAAAAECAwQFBgcICQoL/8QAtREAAgECBAQDBAcFBAQAAQJ3AAECAxEEBS"
    "ExBhJBUQdhcRMiMoEIFEKRobHBCSMzUvAVYnLRChYkNOEl8RcYGRomJygpKjU2"
    "Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6goOEhYaHiImKkp"
    "OUlZaXmJmaoqOkpaanqKmqsrO0tba3uLm6wsPExcbHyMnK0tPU1dbX2Nna4uPk"
    "5ebn6Onq8vP09fb3+Pn6/9oADAMBAAIRAxEAPwDweiiiv61P5iP/2Q=="
)


def _q(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def _rewrite_package(
    source: bytes,
    *,
    replacements: dict[str, bytes],
    additions: dict[str, bytes],
    deletions: set[str] | None = None,
) -> bytes:
    removed = deletions or set()
    output = io.BytesIO()
    with (
        ZipFile(io.BytesIO(source)) as before,
        ZipFile(output, "w", compression=ZIP_DEFLATED) as after,
    ):
        for info in before.infolist():
            if info.filename in removed:
                continue
            after.writestr(
                copy.copy(info),
                replacements.get(info.filename, before.read(info.filename)),
            )
        for name, payload in additions.items():
            after.writestr(name, payload)
    return output.getvalue()


def _serialize_with_wps(root: ET.Element) -> bytes:
    payload = serialize_xml(root)
    if b"xmlns:wps=" in payload:
        return payload
    declaration_end = payload.find(b"?>")
    root_end = payload.find(b">", declaration_end + 2)
    assert root_end > declaration_end
    return (
        payload[:root_end]
        + f' xmlns:wps="{WPS}"'.encode()
        + payload[root_end:]
    )


def _image_document(
    *,
    preceding_text: str | None = None,
    mixed_text: str | None = None,
    anchored: bool = False,
    aligned: bool = False,
    percentage_position: bool = False,
    relative_size: bool = False,
    wrap_mode: str = "square",
    cropped: bool = False,
    transform_attributes: dict[str, str] | None = None,
    outlined: bool = False,
    opacity_amount: str | None = None,
    shadowed: bool = False,
    alt_text: str | None = "A compact expert workflow diagram",
) -> bytes:
    if (aligned or percentage_position or relative_size) and not anchored:
        raise ValueError(
            "Alternative positioning requires a floating anchor."
        )
    if aligned and percentage_position:
        raise ValueError(
            "Alignment and percentage positioning are mutually exclusive."
        )
    if wrap_mode not in {
        "square",
        "none",
        "top_and_bottom",
        "tight",
        "through",
    }:
        raise ValueError("Unsupported test wrap mode.")
    if wrap_mode != "square" and not anchored:
        raise ValueError("Explicit floating wrap requires an anchor.")
    builder = DocumentBuilder()
    if preceding_text is not None:
        builder.paragraph(preceding_text, id="before")
    source = builder.paragraph("", id="picture_host").build().to_bytes("docx")

    with ZipFile(io.BytesIO(source)) as archive:
        document = parse_xml(archive.read("word/document.xml"))
        relationships = parse_xml(
            archive.read("word/_rels/document.xml.rels")
        )
        content_types = parse_xml(archive.read("[Content_Types].xml"))

    body = document.find(_q(W, "body"))
    assert body is not None
    paragraphs = body.findall(_q(W, "p"))
    paragraph = paragraphs[-1]
    paragraph_properties = paragraph.find(_q(W, "pPr"))
    paragraph_attributes = dict(paragraph.attrib)
    paragraph.clear()
    paragraph.attrib.update(paragraph_attributes)
    if paragraph_properties is not None:
        paragraph.append(paragraph_properties)
    run = ET.SubElement(paragraph, _q(W, "r"))
    if mixed_text is not None:
        text = ET.SubElement(run, _q(W, "t"))
        text.text = mixed_text
    drawing = ET.SubElement(run, _q(W, "drawing"))
    placement_attributes = (
        {
            "distT": "12700",
            "distB": "25400",
            "distL": "38100",
            "distR": "50800",
            "simplePos": "0",
            "relativeHeight": "1026",
            "behindDoc": "0",
            "locked": "1",
            "layoutInCell": "1",
            "allowOverlap": "1",
            _q(WP14, "anchorId"): "A1B2C3D4",
            _q(WP14, "editId"): "E5F60718",
        }
        if anchored
        else {}
    )
    placement = ET.SubElement(
        drawing,
        _q(WP, "anchor" if anchored else "inline"),
        placement_attributes,
    )
    if anchored:
        ET.SubElement(
            placement,
            _q(WP, "simplePos"),
            {"x": "0", "y": "0"},
        )
        horizontal = ET.SubElement(
            placement,
            _q(WP, "positionH"),
            {"relativeFrom": "margin" if aligned else "column"},
        )
        ET.SubElement(
            horizontal,
            (
                _q(WP14, "pctPosHOffset")
                if percentage_position
                else _q(WP, "align" if aligned else "posOffset")
            ),
        ).text = (
            "37500"
            if percentage_position
            else "center"
            if aligned
            else "457200"
        )
        vertical = ET.SubElement(
            placement,
            _q(WP, "positionV"),
            {"relativeFrom": "page" if aligned else "paragraph"},
        )
        ET.SubElement(
            vertical,
            (
                _q(WP14, "pctPosVOffset")
                if percentage_position
                else _q(WP, "align" if aligned else "posOffset")
            ),
        ).text = (
            "-12500"
            if percentage_position
            else "bottom"
            if aligned
            else "5080"
        )
    ET.SubElement(
        placement,
        _q(WP, "extent"),
        {"cx": "1828800", "cy": "914400"},
    )
    ET.SubElement(
        placement,
        _q(WP, "effectExtent"),
        {"l": "0", "t": "0", "r": "0", "b": "0"},
    )
    if anchored:
        wrap_tag = {
            "square": "wrapSquare",
            "none": "wrapNone",
            "top_and_bottom": "wrapTopAndBottom",
            "tight": "wrapTight",
            "through": "wrapThrough",
        }[wrap_mode]
        wrap = ET.SubElement(
            placement,
            _q(WP, wrap_tag),
            (
                {"wrapText": "bothSides"}
                if wrap_mode in {"square", "tight", "through"}
                else {}
            ),
        )
        if wrap_mode in {"tight", "through"}:
            polygon = ET.SubElement(
                wrap,
                _q(WP, "wrapPolygon"),
                {"edited": "1"},
            )
            ET.SubElement(
                polygon,
                _q(WP, "start"),
                {"x": "0", "y": "0"},
            )
            for x, y in (
                (0, 21600),
                (21600, 21600),
                (21600, 0),
            ):
                ET.SubElement(
                    polygon,
                    _q(WP, "lineTo"),
                    {"x": str(x), "y": str(y)},
                )
    document_properties = {
        "id": "7",
        "name": "Expert diagram",
        "title": "Workflow",
    }
    if alt_text is not None:
        document_properties["descr"] = alt_text
    ET.SubElement(
        placement,
        _q(WP, "docPr"),
        document_properties,
    )
    frame_properties = ET.SubElement(
        placement,
        _q(WP, "cNvGraphicFramePr"),
    )
    ET.SubElement(
        frame_properties,
        _q(A, "graphicFrameLocks"),
        {"noChangeAspect": "1"},
    )
    graphic = ET.SubElement(placement, _q(A, "graphic"))
    graphic_data = ET.SubElement(
        graphic,
        _q(A, "graphicData"),
        {"uri": PIC},
    )
    picture = ET.SubElement(graphic_data, _q(PIC, "pic"))
    non_visual = ET.SubElement(picture, _q(PIC, "nvPicPr"))
    ET.SubElement(
        non_visual,
        _q(PIC, "cNvPr"),
        {"id": "0", "name": "image1.png"},
    )
    ET.SubElement(non_visual, _q(PIC, "cNvPicPr"))
    blip_fill = ET.SubElement(picture, _q(PIC, "blipFill"))
    blip = ET.SubElement(
        blip_fill,
        _q(A, "blip"),
        {_q(R, "embed"): "rIdImage1"},
    )
    if opacity_amount is not None:
        ET.SubElement(
            blip,
            _q(A, "alphaModFix"),
            {"amt": opacity_amount},
        )
    if cropped:
        ET.SubElement(
            blip_fill,
            _q(A, "srcRect"),
            {
                "l": "1000",
                "t": "2000",
                "r": "3000",
                "b": "4000",
            },
        )
    stretch = ET.SubElement(blip_fill, _q(A, "stretch"))
    ET.SubElement(stretch, _q(A, "fillRect"))
    shape = ET.SubElement(picture, _q(PIC, "spPr"))
    transform = ET.SubElement(
        shape,
        _q(A, "xfrm"),
        transform_attributes or {},
    )
    ET.SubElement(transform, _q(A, "off"), {"x": "0", "y": "0"})
    ET.SubElement(
        transform,
        _q(A, "ext"),
        {"cx": "1828800", "cy": "914400"},
    )
    geometry = ET.SubElement(
        shape,
        _q(A, "prstGeom"),
        {"prst": "rect"},
    )
    ET.SubElement(geometry, _q(A, "avLst"))
    if outlined:
        outline = ET.SubElement(
            shape,
            _q(A, "ln"),
            {
                "w": "25400",
                "cap": "flat",
                "cmpd": "sng",
                "algn": "ctr",
            },
        )
        solid_fill = ET.SubElement(outline, _q(A, "solidFill"))
        ET.SubElement(
            solid_fill,
            _q(A, "srgbClr"),
            {"val": "CC0000"},
        )
        ET.SubElement(
            outline,
            _q(A, "prstDash"),
            {"val": "dashDot"},
        )
        ET.SubElement(outline, _q(A, "round"))
    if shadowed:
        effect_list = ET.SubElement(shape, _q(A, "effectLst"))
        outer_shadow = ET.SubElement(
            effect_list,
            _q(A, "outerShdw"),
            {
                "blurRad": "127000",
                "dist": "38100",
                "dir": "2700000",
                "algn": "ctr",
                "rotWithShape": "0",
            },
        )
        color = ET.SubElement(
            outer_shadow,
            _q(A, "srgbClr"),
            {"val": "123456"},
        )
        ET.SubElement(
            color,
            _q(A, "alpha"),
            {"val": "33333"},
        )
    if relative_size:
        relative_width = ET.SubElement(
            placement,
            _q(WP14, "sizeRelH"),
            {"relativeFrom": "margin"},
        )
        ET.SubElement(
            relative_width,
            _q(WP14, "pctWidth"),
        ).text = "50000"
        relative_height = ET.SubElement(
            placement,
            _q(WP14, "sizeRelV"),
            {"relativeFrom": "page"},
        )
        ET.SubElement(
            relative_height,
            _q(WP14, "pctHeight"),
        ).text = "25000"

    ET.SubElement(
        relationships,
        _q(REL, "Relationship"),
        {
            "Id": "rIdImage1",
            "Type": IMAGE_RELATIONSHIP_TYPE,
            "Target": "media/image1.png",
        },
    )
    if not any(
        child.attrib.get("Extension", "").casefold() == "png"
        for child in content_types
    ):
        ET.SubElement(
            content_types,
            _q(CT, "Default"),
            {"Extension": "png", "ContentType": "image/png"},
        )

    ET.register_namespace("wp", WP)
    ET.register_namespace("a", A)
    ET.register_namespace("pic", PIC)
    document_payload = serialize_xml(document)
    ET.register_namespace("", REL)
    relationships_payload = serialize_xml(relationships)
    ET.register_namespace("", CT)
    content_types_payload = serialize_xml(content_types)
    ET.register_namespace("rel", REL)
    ET.register_namespace("ct", CT)
    return _rewrite_package(
        source,
        replacements={
            "word/document.xml": document_payload,
            "word/_rels/document.xml.rels": relationships_payload,
            "[Content_Types].xml": content_types_payload,
        },
        additions={"word/media/image1.png": PNG},
    )


def _header_image_document(
    *,
    kind: str = "header",
    cropped: bool = False,
    anchored: bool = False,
    aligned: bool = False,
    percentage_position: bool = False,
    relative_size: bool = False,
    wrap_mode: str = "square",
    transform_attributes: dict[str, str] | None = None,
    outlined: bool = False,
    opacity_amount: str | None = None,
    shadowed: bool = False,
) -> bytes:
    assert kind in {"header", "footer"}
    body_image = _image_document(
        cropped=cropped,
        anchored=anchored,
        aligned=aligned,
        percentage_position=percentage_position,
        relative_size=relative_size,
        wrap_mode=wrap_mode,
        transform_attributes=transform_attributes,
        outlined=outlined,
        opacity_amount=opacity_amount,
        shadowed=shadowed,
    )
    with ZipFile(io.BytesIO(body_image)) as archive:
        image_document = parse_xml(
            archive.read("word/document.xml")
        )
    image_body = image_document.find(_q(W, "body"))
    assert image_body is not None
    image_paragraph = next(
        paragraph
        for paragraph in image_body.findall(_q(W, "p"))
        if paragraph.find(f".//{_q(W, 'drawing')}") is not None
    )

    source = (
        DocumentBuilder(
            header_footers=[
                {
                    "id": f"logo_{kind}",
                    "kind": kind,
                    "content": [
                        {
                            "id": "logo_placeholder",
                            "type": "paragraph",
                            "text": "",
                        }
                    ],
                }
            ],
            sections=[
                {
                    "id": "logo_section",
                    "header_footer": {
                        f"{kind}_default": f"logo_{kind}",
                    },
                }
            ],
        )
        .paragraph("Body", id="logo_body")
        .build()
        .to_bytes("docx")
    )
    part_name = f"word/{kind}1.xml"
    relationship_name = f"word/_rels/{kind}1.xml.rels"
    with ZipFile(io.BytesIO(source)) as archive:
        header = parse_xml(archive.read(part_name))
        content_types = parse_xml(
            archive.read("[Content_Types].xml")
        )
    for child in list(header):
        header.remove(child)
    header.append(copy.deepcopy(image_paragraph))
    if not any(
        child.get("Extension", "").casefold() == "png"
        for child in content_types
    ):
        ET.SubElement(
            content_types,
            _q(CT, "Default"),
            {
                "Extension": "png",
                "ContentType": "image/png",
            },
        )
    relationships = ET.Element(_q(REL, "Relationships"))
    ET.SubElement(
        relationships,
        _q(REL, "Relationship"),
        {
            "Id": "rIdImage1",
            "Type": IMAGE_RELATIONSHIP_TYPE,
            "Target": "media/image1.png",
        },
    )
    return _rewrite_package(
        source,
        replacements={
            part_name: serialize_xml(header),
            "[Content_Types].xml": serialize_xml(content_types),
        },
        additions={
            relationship_name: serialize_xml(
                relationships
            ),
            "word/media/image1.png": PNG,
        },
    )


def _alternate_content_image_document(
    source: bytes,
    *,
    part_name: str = "word/document.xml",
    fallback_matches_choice: bool = True,
    empty_stretch: bool = False,
    signed_anchor_id: bool = False,
) -> bytes:
    relationship_name = (
        f"{part_name.rsplit('/', 1)[0]}/_rels/"
        f"{part_name.rsplit('/', 1)[1]}.rels"
    )
    with ZipFile(io.BytesIO(source)) as archive:
        root = parse_xml(archive.read(part_name))
        relationships = parse_xml(archive.read(relationship_name))
        content_types = parse_xml(archive.read("[Content_Types].xml"))

    drawing = root.find(f".//{_q(W, 'drawing')}")
    assert drawing is not None
    anchor = drawing.find(f"./{_q(WP, 'anchor')}")
    if anchor is not None:
        if signed_anchor_id:
            anchor.set(_q(WP14, "anchorId"), "-5E4D3C2C")
        for attribute in ("distT", "distB", "distL", "distR"):
            anchor.set(attribute, "0")
    if empty_stretch:
        stretch = drawing.find(f".//{_q(A, 'stretch')}")
        assert stretch is not None
        stretch.clear()
    run = next(
        element
        for element in root.iter(_q(W, "r"))
        if drawing in list(element)
    )
    run.remove(drawing)
    alternate = ET.SubElement(run, _q(MC, "AlternateContent"))
    choice = ET.SubElement(
        alternate,
        _q(MC, "Choice"),
        {"Requires": "wps"},
    )
    choice.append(drawing)
    fallback = ET.SubElement(alternate, _q(MC, "Fallback"))
    picture = ET.SubElement(fallback, _q(W, "pict"))
    shape_type = ET.SubElement(
        picture,
        _q(VML, "shapetype"),
        {
            "id": "_x0000_t75",
            "coordsize": "21600,21600",
            _q(OFFICE, "spt"): "75",
            _q(OFFICE, "preferrelative"): "t",
            "path": "m@4@5l@4@11@9@11@9@5xe",
            "filled": "f",
            "stroked": "f",
        },
    )
    ET.SubElement(
        shape_type,
        _q(VML, "stroke"),
        {"joinstyle": "miter"},
    )
    formulas = ET.SubElement(shape_type, _q(VML, "formulas"))
    for formula in (
        "if lineDrawn pixelLineWidth 0",
        "sum @0 1 0",
        "sum 0 0 @1",
        "prod @2 1 2",
        "prod @3 21600 pixelWidth",
        "prod @3 21600 pixelHeight",
        "sum @0 0 1",
        "prod @6 1 2",
        "prod @7 21600 pixelWidth",
        "sum @8 21600 0",
        "prod @7 21600 pixelHeight",
        "sum @10 21600 0",
    ):
        ET.SubElement(formulas, _q(VML, "f"), {"eqn": formula})
    ET.SubElement(
        shape_type,
        _q(VML, "path"),
        {
            _q(OFFICE, "extrusionok"): "f",
            "gradientshapeok": "t",
            _q(OFFICE, "connecttype"): "rect",
        },
    )
    ET.SubElement(
        shape_type,
        _q(OFFICE, "lock"),
        {
            _q(VML, "ext"): "edit",
            "aspectratio": "t",
        },
    )
    fallback_relationship_id = (
        "rIdImage1"
        if fallback_matches_choice
        else "rIdFallbackImage"
    )
    shape_attributes = {
        "id": "_x0000_i1025",
        "stroked": "f",
        _q(OFFICE, "allowincell"): (
            "t"
            if anchor is None
            or anchor.get("layoutInCell") == "1"
            else "f"
        ),
        "style": (
            (
                "position:absolute;margin-left:36pt;margin-top:0.4pt;"
            )
            if anchor is not None
            else ""
        )
        + (
            "width:144pt;height:72pt;"
            "mso-wrap-style:none;v-text-anchor:middle"
        ),
        "type": "#_x0000_t75",
    }
    if anchor is not None:
        anchor_id = anchor.get(_q(WP14, "anchorId"))
        assert anchor_id is not None
        shape_attributes[_q(WP14, "anchorId")] = anchor_id
    shape = ET.SubElement(
        picture,
        _q(VML, "shape"),
        shape_attributes,
    )
    ET.SubElement(
        shape,
        _q(VML, "imagedata"),
        {
            _q(R, "id"): fallback_relationship_id,
            _q(OFFICE, "detectmouseclick"): "t",
        },
    )
    if anchor is not None:
        ET.SubElement(
            shape,
            _q(WORD_VML, "wrap"),
            {"type": "square"},
        )
    additions: dict[str, bytes] = {}
    if not fallback_matches_choice:
        ET.SubElement(
            relationships,
            _q(REL, "Relationship"),
            {
                "Id": fallback_relationship_id,
                "Type": IMAGE_RELATIONSHIP_TYPE,
                "Target": "media/fallback.jpg",
            },
        )
        additions["word/media/fallback.jpg"] = JPEG
        if not any(
            child.get("Extension", "").casefold() in {"jpg", "jpeg"}
            for child in content_types
        ):
            ET.SubElement(
                content_types,
                _q(CT, "Default"),
                {
                    "Extension": "jpg",
                    "ContentType": "image/jpeg",
                },
            )

    ET.register_namespace("mc", MC)
    ET.register_namespace("v", VML)
    ET.register_namespace("o", OFFICE)
    part_payload = _serialize_with_wps(root)
    return _rewrite_package(
        source,
        replacements={
            part_name: part_payload,
            relationship_name: serialize_xml(relationships),
            "[Content_Types].xml": serialize_xml(content_types),
        },
        additions=additions,
    )


class DocxImageTests(unittest.TestCase):
    def test_cli_exposes_strict_image_and_asset_schemas(self) -> None:
        for kind, required_properties in (
            (
                "image-block",
                {"asset_id", "placement", "width", "height", "editable"},
            ),
            (
                "header-footer-image-block",
                {"asset_id", "placement", "width", "height", "editable"},
            ),
            (
                "image-crop",
                {"left", "top", "right", "bottom"},
            ),
            (
                "image-effect-extent",
                {"left", "top", "right", "bottom"},
            ),
            (
                "image-transform",
                {
                    "rotation_degrees_clockwise",
                    "flip_horizontal",
                    "flip_vertical",
                },
            ),
            (
                "image-outline",
                {"width", "color", "dash"},
            ),
            (
                "image-shadow",
                {
                    "color",
                    "opacity",
                    "blur_radius",
                    "distance",
                    "direction_degrees_clockwise",
                    "alignment",
                    "rotate_with_shape",
                    "effect_extent",
                },
            ),
            (
                "image-alternate-content",
                {
                    "choice_requires_prefix",
                    "choice_requires_namespace",
                    "fallback_kind",
                    "fallback_placement",
                    "synchronized_update_fields",
                    "fallback_asset_matches_choice",
                },
            ),
            (
                "floating-image-effect-extent",
                {"left", "top", "right", "bottom"},
            ),
            (
                "floating-image-horizontal-position",
                {
                    "relative_to",
                    "offset",
                    "alignment",
                    "percentage_offset",
                },
            ),
            (
                "floating-image-text-distances",
                {"top", "right", "bottom", "left"},
            ),
            (
                "floating-image-vertical-position",
                {
                    "relative_to",
                    "offset",
                    "alignment",
                    "percentage_offset",
                },
            ),
            (
                "floating-image-text-wrap",
                {
                    "mode",
                    "side",
                    "distances",
                    "effect_extent",
                    "polygon",
                },
            ),
            (
                "floating-image-wrap-point",
                {"x", "y"},
            ),
            (
                "floating-image-wrap-polygon",
                {"edited", "start", "line_to"},
            ),
            (
                "floating-image-layout",
                {
                    "horizontal",
                    "vertical",
                    "anchor_distances",
                    "anchor_effect_extent",
                    "wrap",
                    "relative_size",
                    "relative_height",
                    "behind_text",
                    "locked",
                    "layout_in_cell",
                    "allow_overlap",
                },
            ),
            (
                "floating-image-layout-update",
                {
                    "horizontal",
                    "vertical",
                    "anchor_distances",
                    "anchor_effect_extent",
                    "wrap",
                    "relative_size",
                    "relative_height",
                    "behind_text",
                    "locked",
                    "layout_in_cell",
                    "allow_overlap",
                },
            ),
            (
                "floating-image-relative-width",
                {"relative_to", "percentage"},
            ),
            (
                "floating-image-relative-height",
                {"relative_to", "percentage"},
            ),
            (
                "floating-image-relative-size",
                {"width", "height"},
            ),
            (
                "asset-ref",
                {"id", "sha256", "media_type", "size_bytes"},
            ),
            (
                "image-insert",
                {
                    "id",
                    "placement",
                    "floating",
                    "width",
                    "height",
                    "transform",
                    "outline",
                    "opacity",
                    "shadow",
                    "alt_text",
                    "paragraph_style",
                },
            ),
            (
                "image-update",
                {
                    "width",
                    "height",
                    "crop",
                    "transform",
                    "outline",
                    "opacity",
                    "shadow",
                    "alt_text",
                    "title",
                },
            ),
        ):
            stdout = io.StringIO()
            with self.subTest(kind=kind), redirect_stdout(stdout):
                self.assertEqual(
                    main(["schema", "--kind", kind]),
                    0,
                )
            schema = json.loads(stdout.getvalue())
            self.assertFalse(schema["additionalProperties"])
            self.assertTrue(
                required_properties.issubset(schema["properties"])
            )
            if kind in {
                "floating-image-horizontal-position",
                "floating-image-vertical-position",
            }:
                self.assertEqual(
                    {
                        tuple(branch["required"])
                        for branch in schema["oneOf"]
                    },
                    {
                        ("offset",),
                        ("alignment",),
                        ("percentage_offset",),
                    },
                )
                for branch in schema["oneOf"]:
                    mode = branch["required"][0]
                    self.assertEqual(
                        branch["properties"][mode],
                        {"not": {"type": "null"}},
                    )
            if kind == "floating-image-text-wrap":
                branches = {
                    branch["properties"]["mode"]["const"]: branch
                    for branch in schema["oneOf"]
                }
                self.assertEqual(
                    set(branches),
                    {
                        "square",
                        "none",
                        "top_and_bottom",
                        "tight",
                        "through",
                    },
                )
                self.assertIn("side", branches["square"]["required"])
                self.assertEqual(
                    branches["square"]["properties"]["side"],
                    {"not": {"type": "null"}},
                )
                self.assertIn("mode", branches["none"]["required"])
                self.assertEqual(
                    branches["none"]["allOf"],
                    [
                        {"not": {"required": ["side"]}},
                        {"not": {"required": ["distances"]}},
                        {"not": {"required": ["effect_extent"]}},
                        {"not": {"required": ["polygon"]}},
                    ],
                )
                self.assertIn(
                    "mode",
                    branches["top_and_bottom"]["required"],
                )
                self.assertEqual(
                    branches["top_and_bottom"]["allOf"],
                    [
                        {"not": {"required": ["side"]}},
                        {"not": {"required": ["polygon"]}},
                    ],
                )
                self.assertEqual(
                    branches["top_and_bottom"]["properties"][
                        "distances"
                    ],
                    {
                        "allOf": [
                            {"not": {"required": ["left"]}},
                            {"not": {"required": ["right"]}},
                        ]
                    },
                )
                for mode in ("tight", "through"):
                    self.assertEqual(
                        set(branches[mode]["required"]),
                        {"mode", "side", "polygon"},
                    )
                    self.assertEqual(
                        branches[mode]["not"],
                        {"required": ["effect_extent"]},
                    )
                    self.assertEqual(
                        branches[mode]["properties"]["distances"],
                        {
                            "allOf": [
                                {"not": {"required": ["top"]}},
                                {"not": {"required": ["bottom"]}},
                            ]
                        },
                    )
            if kind == "floating-image-relative-size":
                self.assertEqual(
                    schema["anyOf"],
                    [
                        {
                            "required": ["width"],
                            "properties": {
                                "width": {
                                    "not": {"type": "null"},
                                }
                            },
                        },
                        {
                            "required": ["height"],
                            "properties": {
                                "height": {
                                    "not": {"type": "null"},
                                }
                            },
                        },
                    ],
                )
            if kind == "header-footer-image-block":
                capabilities = schema["properties"]["capabilities"]
                self.assertEqual(
                    capabilities["prefixItems"],
                    [
                        {"const": "inspect"},
                        {"const": "extract"},
                        {"const": "render"},
                    ],
                )
                self.assertFalse(capabilities["items"])
                self.assertEqual(capabilities["minItems"], 3)
                self.assertEqual(capabilities["maxItems"], 3)

    def test_projects_metadata_and_reads_verified_native_bytes(self) -> None:
        source = _image_document()
        document = Document.from_docx(source)
        spec = document.to_spec()

        self.assertEqual(spec["spec_version"], "0.2-draft.49")
        self.assertEqual(len(spec["content"]), 1)
        image = spec["content"][0]
        self.assertEqual(image["type"], "image")
        self.assertEqual(image["placement"], "inline")
        self.assertEqual(image["width"], {"value": 144.0, "unit": "pt"})
        self.assertEqual(image["height"], {"value": 72.0, "unit": "pt"})
        self.assertEqual(
            image["alt_text"],
            "A compact expert workflow diagram",
        )
        self.assertFalse(image["editable"])
        self.assertEqual(
            image["capabilities"],
            ["inspect", "extract", "delete", "render"],
        )

        self.assertEqual(len(spec["assets"]), 1)
        asset = spec["assets"][0]
        self.assertEqual(asset["id"], image["asset_id"])
        self.assertEqual(asset["media_type"], "image/png")
        self.assertEqual(asset["size_bytes"], len(PNG))
        self.assertEqual(len(asset["sha256"]), 64)

        extracted = document.read_image(image["id"])
        self.assertEqual(extracted.data, PNG)
        self.assertEqual(extracted.sha256, asset["sha256"])
        self.assertEqual(document.image_bytes(f"#{image['id']}"), PNG)
        self.assertEqual(document.to_bytes("docx"), source)

        inspected = document.inspect()
        self.assertEqual(inspected["image_count"], 1)
        self.assertEqual(inspected["asset_count"], 1)
        self.assertEqual(
            inspected["nodes"][0]["asset"]["media_type"],
            "image/png",
        )
        self.assertEqual(
            inspected["nodes"][0]["supported_operations"],
            [
                "image.insert_after",
                "image.replace",
                "image.update",
                "paragraph.format",
                "node.remove",
            ],
        )
        document_capabilities = document.capabilities()
        self.assertIn("image.update", document_capabilities["operations"])
        self.assertIn("image.replace", document_capabilities["operations"])
        self.assertIn(
            "image.insert_after",
            document_capabilities["operations"],
        )
        capabilities = document_capabilities["assets"]
        self.assertFalse(capabilities["binary_in_json"])
        self.assertFalse(capabilities["binary_write_in_json"])
        self.assertEqual(
            capabilities["binary_write_transport"],
            "out_of_band",
        )
        self.assertEqual(
            capabilities["replacement_strategy"],
            "occurrence_copy_on_write",
        )
        self.assertEqual(
            capabilities["insert_dimensions"],
            "explicit_width_and_height",
        )
        self.assertEqual(
            capabilities["insert_alt_text"],
            "required",
        )
        self.assertEqual(
            capabilities["native_layout_operation"],
            "paragraph.format",
        )
        self.assertIn(
            "alignment",
            capabilities["native_layout_fields"],
        )
        self.assertEqual(
            capabilities["native_update_fields"],
            [
                "width",
                "height",
                "crop",
                "transform",
                "outline",
                "opacity",
                "shadow",
                "alt_text",
                "title",
            ],
        )
        self.assertEqual(
            capabilities["clearable_update_fields"],
            [
                "crop",
                "transform",
                "outline",
                "opacity",
                "shadow",
                "alt_text",
                "title",
            ],
        )
        self.assertEqual(
            capabilities["single_dimension_resize"],
            "preserve_aspect_ratio",
        )
        self.assertTrue(capabilities["native_render_is_visual_authority"])

        html = document.to_bytes("html").decode()
        self.assertIn("native-image-placeholder", html)
        self.assertIn(image["asset_id"], html)
        self.assertNotIn(base64.b64encode(PNG).decode(), html)
        markdown = document.to_bytes("markdown").decode()
        self.assertIn(f"aioffice-asset:{image['asset_id']}", markdown)

        detached = Document.from_spec(spec)
        self.assertTrue(detached.validate().valid)
        with self.assertRaises(NativePackageError):
            detached.read_image(image["id"])

    def test_extract_image_refuses_overwrite_and_cli_reports_integrity(self) -> None:
        source = _image_document()
        document = Document.from_docx(source)
        image_id = document.to_spec()["content"][0]["id"]
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "image.png"
            self.assertEqual(
                document.extract_image(image_id, output),
                output,
            )
            self.assertEqual(output.read_bytes(), PNG)
            with self.assertRaises(FileExistsError):
                document.extract_image(image_id, output)

            input_path = Path(directory) / "source.docx"
            cli_output = Path(directory) / "cli.png"
            input_path.write_bytes(source)
            stdout = io.StringIO()
            with redirect_stdout(stdout):
                result = main(
                    [
                        "extract-image",
                        str(input_path),
                        image_id,
                        "-o",
                        str(cli_output),
                    ]
                )
            self.assertEqual(result, 0)
            report = json.loads(stdout.getvalue())
            self.assertEqual(report["image_id"], image_id)
            self.assertEqual(report["size_bytes"], len(PNG))
            self.assertEqual(cli_output.read_bytes(), PNG)

    def test_missing_native_alt_text_is_an_explicit_warning(self) -> None:
        document = Document.from_docx(
            _image_document(alt_text=None)
        )
        self.assertTrue(document.validate().valid)
        self.assertTrue(
            any(
                diagnostic.code == "IMAGE_ALT_TEXT_MISSING"
                for diagnostic in document.validate().warnings
            )
        )

    def test_repeated_image_occurrences_share_one_content_asset(self) -> None:
        source = _image_document()
        with ZipFile(io.BytesIO(source)) as archive:
            document_root = parse_xml(
                archive.read("word/document.xml")
            )
        body = document_root.find(_q(W, "body"))
        assert body is not None
        image_paragraph = body.find(_q(W, "p"))
        assert image_paragraph is not None
        repeated = copy.deepcopy(image_paragraph)
        repeated.attrib[_q(W14, "paraId")] = "ABCDEF12"
        body.insert(len(body) - 1, repeated)
        source = _rewrite_package(
            source,
            replacements={
                "word/document.xml": serialize_xml(document_root),
            },
            additions={},
        )

        document = Document.from_docx(source)
        images = [
            node
            for node in document.to_spec()["content"]
            if node["type"] == "image"
        ]
        self.assertEqual(len(images), 2)
        self.assertEqual(len(document.to_spec()["assets"]), 1)
        self.assertEqual(images[0]["asset_id"], images[1]["asset_id"])
        self.assertEqual(document.image_bytes(images[0]["id"]), PNG)
        self.assertEqual(document.image_bytes(images[1]["id"]), PNG)

    def test_mixed_image_remains_opaque(self) -> None:
        source = _image_document(mixed_text="Caption")
        document = Document.from_docx(source)
        spec = document.to_spec()
        self.assertEqual(spec["content"][0]["type"], "opaque")
        self.assertEqual(spec["assets"], [])
        self.assertEqual(document.to_bytes("docx"), source)
        mixed = Document.from_docx(
            _image_document(mixed_text="Caption")
        ).to_spec()["content"][0]
        self.assertIn("with text", mixed["summary"])

    def test_strict_inline_alternate_content_projects_resizes_and_replaces(
        self,
    ) -> None:
        source = _alternate_content_image_document(
            _image_document(shadowed=True),
            empty_stretch=True,
        )
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        expected_alternate = {
            "choice_requires_prefix": "wps",
            "choice_requires_namespace": WPS,
            "fallback_kind": "vml_picture",
            "fallback_placement": "inline",
            "synchronized_update_fields": ["width", "height"],
            "fallback_asset_matches_choice": True,
        }
        self.assertEqual(image["alternate_content"], expected_alternate)
        self.assertEqual(document.image_bytes(image["id"]), PNG)
        self.assertEqual(document.to_bytes("docx"), source)

        inspected = document.inspect()["nodes"][0]
        self.assertEqual(
            inspected["alternate_content"],
            expected_alternate,
        )
        self.assertEqual(
            inspected["native_update_fields"],
            ["width", "height"],
        )
        self.assertIn("image.replace", inspected["supported_operations"])
        html = document.to_bytes("html").decode()
        self.assertIn('data-aioffice-alternate-content="true"', html)
        self.assertIn(
            'data-aioffice-fallback-asset-matches-choice="true"',
            html,
        )
        capabilities = document.capabilities()["assets"]
        self.assertEqual(
            capabilities["image_alternate_content_schema"],
            "image-alternate-content",
        )
        self.assertEqual(
            capabilities[
                "image_alternate_content_synchronized_update_fields"
            ],
            ["width", "height"],
        )
        self.assertEqual(
            capabilities[
                "image_alternate_content_header_footer_clone"
            ],
            "supported_when_strictly_projected",
        )
        formatted = document.apply(
            [
                {
                    "op": "paragraph.format",
                    "target": image["id"],
                    "set": {"alignment": "center"},
                }
            ]
        )
        self.assertTrue(formatted.success, formatted.model_dump())
        assert formatted.document is not None
        formatted_reopened = Document.from_docx(
            formatted.document.to_bytes("docx")
        )
        self.assertEqual(
            formatted_reopened.to_spec()["content"][0][
                "alternate_content"
            ],
            expected_alternate,
        )

        refused = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {"opacity": 50},
                }
            ]
        )
        self.assertFalse(refused.success)
        self.assertEqual(
            refused.diagnostics[0].code,
            "INVALID_SPEC",
        )
        self.assertIn(
            "alternate-content wrapper",
            refused.diagnostics[0].message,
        )
        self.assertEqual(document.to_bytes("docx"), source)

        resized = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {
                        "width": {"value": 288, "unit": "pt"},
                    },
                }
            ]
        )
        self.assertTrue(resized.success, resized.model_dump())
        assert resized.document is not None
        resized_image = resized.document.to_spec()["content"][0]
        self.assertEqual(
            resized_image["width"],
            {"value": 288.0, "unit": "pt"},
        )
        self.assertEqual(
            resized_image["height"],
            {"value": 144.0, "unit": "pt"},
        )
        with ZipFile(
            io.BytesIO(resized.document.to_bytes("docx"))
        ) as package:
            root = parse_xml(package.read("word/document.xml"))
        extent = root.find(f".//{_q(WP, 'extent')}")
        fallback_shape = root.find(f".//{_q(VML, 'shape')}")
        assert extent is not None
        assert fallback_shape is not None
        self.assertEqual(
            extent.attrib,
            {"cx": "3657600", "cy": "1828800"},
        )
        self.assertEqual(
            fallback_shape.get("style"),
            (
                "width:288pt;height:144pt;"
                "mso-wrap-style:none;v-text-anchor:middle"
            ),
        )

        replaced = resized.document.replace_image(
            image["id"],
            JPEG,
            media_type="image/jpeg",
        )
        self.assertTrue(replaced.success, replaced.model_dump())
        assert replaced.document is not None
        self.assertEqual(replaced.document.image_bytes(image["id"]), JPEG)
        replaced_image = replaced.document.to_spec()["content"][0]
        self.assertEqual(
            replaced_image["alternate_content"],
            expected_alternate,
        )
        with ZipFile(
            io.BytesIO(replaced.document.to_bytes("docx"))
        ) as package:
            root = parse_xml(package.read("word/document.xml"))
            relationships = parse_xml(
                package.read("word/_rels/document.xml.rels")
            )
        blip = root.find(f".//{_q(A, 'blip')}")
        image_data = root.find(f".//{_q(VML, 'imagedata')}")
        assert blip is not None
        assert image_data is not None
        replacement_id = blip.get(_q(R, "embed"))
        self.assertEqual(image_data.get(_q(R, "id")), replacement_id)
        relationship = next(
            item
            for item in relationships.findall(
                _q(REL, "Relationship")
            )
            if item.get("Id") == replacement_id
        )
        self.assertTrue(
            relationship.get("Target", "").endswith(".jpg")
        )

    def test_alternate_content_with_distinct_fallback_refuses_replace(
        self,
    ) -> None:
        source = _alternate_content_image_document(
            _image_document(shadowed=True),
            fallback_matches_choice=False,
        )
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        self.assertFalse(
            image["alternate_content"][
                "fallback_asset_matches_choice"
            ]
        )
        self.assertEqual(document.image_bytes(image["id"]), PNG)
        self.assertEqual(len(document.to_spec()["assets"]), 1)
        inspected = document.inspect()["nodes"][0]
        self.assertNotIn(
            "image.replace",
            inspected["supported_operations"],
        )
        self.assertNotIn(
            "image.replace",
            document.capabilities()["operations"],
        )
        result = document.replace_image(
            image["id"],
            JPEG,
            media_type="image/jpeg",
        )
        self.assertFalse(result.success)
        self.assertEqual(
            result.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )
        self.assertEqual(document.to_bytes("docx"), source)

    def test_header_alternate_content_projects_and_resizes_both_branches(
        self,
    ) -> None:
        source = _alternate_content_image_document(
            _header_image_document(),
            part_name="word/header1.xml",
        )
        document = Document.from_docx(source)
        header = document.to_spec()["header_footers"][0]
        image = header["content"][0]
        self.assertTrue(
            image["alternate_content"][
                "fallback_asset_matches_choice"
            ]
        )
        inspected = document.inspect()["header_footers"][0]["blocks"][0]
        self.assertEqual(
            inspected["native_update_fields"],
            ["width", "height"],
        )
        result = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {
                        "height": {"value": 108, "unit": "pt"},
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        with ZipFile(
            io.BytesIO(result.document.to_bytes("docx"))
        ) as package:
            root = parse_xml(package.read("word/header1.xml"))
        extent = root.find(f".//{_q(WP, 'extent')}")
        fallback_shape = root.find(f".//{_q(VML, 'shape')}")
        assert extent is not None
        assert fallback_shape is not None
        self.assertEqual(
            extent.attrib,
            {"cx": "2743200", "cy": "1371600"},
        )
        self.assertEqual(
            fallback_shape.get("style"),
            (
                "width:216pt;height:108pt;"
                "mso-wrap-style:none;v-text-anchor:middle"
            ),
        )

    def test_offset_floating_alternate_content_projects_without_anchor_edit(
        self,
    ) -> None:
        source = _alternate_content_image_document(
            _image_document(anchored=True, shadowed=True),
            signed_anchor_id=True,
        )
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        self.assertEqual(image["placement"], "floating")
        self.assertEqual(
            image["alternate_content"]["fallback_placement"],
            "floating_offset",
        )
        self.assertTrue(
            image["alternate_content"][
                "fallback_asset_matches_choice"
            ]
        )
        self.assertEqual(
            image["floating"]["horizontal"],
            {
                "relative_to": "column",
                "offset": {"value": 36.0, "unit": "pt"},
            },
        )
        self.assertEqual(
            image["floating"]["vertical"],
            {
                "relative_to": "paragraph",
                "offset": {"value": 0.4, "unit": "pt"},
            },
        )
        inspected = document.inspect()["nodes"][0]
        self.assertNotIn(
            "image.anchor.update",
            inspected["supported_operations"],
        )
        self.assertNotIn(
            "image.anchor.update",
            document.capabilities()["operations"],
        )
        self.assertIn(
            'data-aioffice-fallback-placement="floating_offset"',
            document.to_bytes("html").decode(),
        )
        refused = document.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": image["id"],
                    "set": {"relative_height": 12},
                }
            ]
        )
        self.assertFalse(refused.success)
        self.assertEqual(
            refused.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )
        self.assertEqual(document.to_bytes("docx"), source)

        resized = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {
                        "width": {"value": 216, "unit": "pt"},
                    },
                }
            ]
        )
        self.assertTrue(resized.success, resized.model_dump())
        assert resized.document is not None
        with ZipFile(
            io.BytesIO(resized.document.to_bytes("docx"))
        ) as package:
            root = parse_xml(package.read("word/document.xml"))
        fallback_shape = root.find(f".//{_q(VML, 'shape')}")
        assert fallback_shape is not None
        self.assertEqual(
            fallback_shape.get("style"),
            (
                "position:absolute;margin-left:36pt;margin-top:0.4pt;"
                "width:216pt;height:108pt;"
                "mso-wrap-style:none;v-text-anchor:middle"
            ),
        )
        reopened = Document.from_docx(
            resized.document.to_bytes("docx")
        )
        reopened_image = reopened.to_spec()["content"][0]
        self.assertEqual(reopened_image["placement"], "floating")
        self.assertEqual(
            reopened_image["alternate_content"][
                "fallback_placement"
            ],
            "floating_offset",
        )
        replaced = reopened.replace_image(
            image["id"],
            JPEG,
            media_type="image/jpeg",
        )
        self.assertTrue(replaced.success, replaced.model_dump())
        assert replaced.document is not None
        self.assertEqual(replaced.document.image_bytes(image["id"]), JPEG)

        header_source = _alternate_content_image_document(
            _header_image_document(anchored=True),
            part_name="word/header1.xml",
        )
        header_document = Document.from_docx(header_source)
        header_image = header_document.to_spec()["header_footers"][0][
            "content"
        ][0]
        self.assertEqual(header_image["placement"], "floating")
        self.assertEqual(
            header_image["alternate_content"]["fallback_placement"],
            "floating_offset",
        )
        self.assertNotIn(
            "image.anchor.update",
            header_document.inspect()["header_footers"][0]["blocks"][0][
                "supported_operations"
            ],
        )

    def test_header_clone_rebases_supported_alternate_content_images(
        self,
    ) -> None:
        for anchored in (False, True):
            for fallback_matches_choice in (False, True):
                source = _alternate_content_image_document(
                    _header_image_document(anchored=anchored),
                    part_name="word/header1.xml",
                    fallback_matches_choice=fallback_matches_choice,
                    signed_anchor_id=anchored,
                )
                collision_shape_id: str | None = None
                if not anchored and fallback_matches_choice:
                    collision_shape_id = "shape_" + hashlib.sha256(
                        (
                            "compatibility_clone:vml-shape:"
                            "0:0:0"
                        ).encode()
                    ).hexdigest()[:8].upper()
                    with ZipFile(io.BytesIO(source)) as archive:
                        collision_root = parse_xml(
                            archive.read("word/header1.xml")
                        )
                    collision_shape = collision_root.find(
                        f".//{_q(VML, 'shape')}"
                    )
                    assert collision_shape is not None
                    collision_shape.set("id", collision_shape_id)
                    source = _rewrite_package(
                        source,
                        replacements={
                            "word/header1.xml": _serialize_with_wps(
                                collision_root
                            ),
                        },
                        additions={},
                    )
                document = Document.from_docx(source)
                source_part = document.to_spec()["header_footers"][0]
                source_image = source_part["content"][0]
                result = document.apply(
                    [
                        {
                            "op": "header_footer.clone",
                            "target": source_part["id"],
                            "part": {"id": "compatibility_clone"},
                        }
                    ]
                )
                with self.subTest(
                    anchored=anchored,
                    fallback_matches_choice=fallback_matches_choice,
                ):
                    self.assertTrue(result.success, result.model_dump())
                    assert result.document is not None
                    parts = {
                        part["id"]: part
                        for part in result.document.to_spec()[
                            "header_footers"
                        ]
                    }
                    cloned_image = parts["compatibility_clone"][
                        "content"
                    ][0]
                    self.assertNotEqual(
                        cloned_image["id"],
                        source_image["id"],
                    )
                    self.assertEqual(
                        cloned_image["asset_id"],
                        source_image["asset_id"],
                    )
                    self.assertEqual(
                        cloned_image["alternate_content"],
                        source_image["alternate_content"],
                    )
                    self.assertEqual(
                        result.document.image_bytes(cloned_image["id"]),
                        PNG,
                    )

                    output = result.document.to_bytes("docx")
                    with (
                        ZipFile(io.BytesIO(source)) as before,
                        ZipFile(io.BytesIO(output)) as after,
                    ):
                        self.assertEqual(
                            after.read("word/header1.xml"),
                            before.read("word/header1.xml"),
                        )
                        clone_payload = after.read("word/header2.xml")
                        self.assertIn(
                            f'xmlns:wps="{WPS}"'.encode(),
                            clone_payload,
                        )
                        self.assertEqual(
                            after.read(
                                "word/_rels/header2.xml.rels"
                            ),
                            before.read(
                                "word/_rels/header1.xml.rels"
                            ),
                        )
                        source_root = parse_xml(
                            after.read("word/header1.xml")
                        )
                        clone_root = parse_xml(clone_payload)
                    source_shape = source_root.find(
                        f".//{_q(VML, 'shape')}"
                    )
                    clone_shape = clone_root.find(
                        f".//{_q(VML, 'shape')}"
                    )
                    assert source_shape is not None
                    assert clone_shape is not None
                    self.assertNotEqual(
                        source_shape.get("id"),
                        clone_shape.get("id"),
                    )
                    self.assertTrue(
                        clone_shape.get("id", "").startswith("shape_")
                    )
                    if collision_shape_id is not None:
                        self.assertNotEqual(
                            clone_shape.get("id"),
                            collision_shape_id,
                        )
                    source_drawing_ids = {
                        element.get("id")
                        for element in source_root.iter()
                        if element.tag
                        in {
                            _q(WP, "docPr"),
                            _q(A, "cNvPr"),
                            _q(PIC, "cNvPr"),
                        }
                    }
                    clone_drawing_ids = {
                        element.get("id")
                        for element in clone_root.iter()
                        if element.tag
                        in {
                            _q(WP, "docPr"),
                            _q(A, "cNvPr"),
                            _q(PIC, "cNvPr"),
                        }
                    }
                    self.assertTrue(
                        source_drawing_ids.isdisjoint(
                            clone_drawing_ids
                        )
                    )
                    if anchored:
                        source_anchor = source_root.find(
                            f".//{_q(WP, 'anchor')}"
                        )
                        clone_anchor = clone_root.find(
                            f".//{_q(WP, 'anchor')}"
                        )
                        assert source_anchor is not None
                        assert clone_anchor is not None
                        source_anchor_id = source_anchor.get(
                            _q(WP14, "anchorId")
                        )
                        clone_anchor_id = clone_anchor.get(
                            _q(WP14, "anchorId")
                        )
                        self.assertNotEqual(
                            source_anchor_id,
                            clone_anchor_id,
                        )
                        self.assertEqual(
                            source_shape.get(_q(WP14, "anchorId")),
                            source_anchor_id,
                        )
                        self.assertEqual(
                            clone_shape.get(_q(WP14, "anchorId")),
                            clone_anchor_id,
                        )

                    resized = result.document.apply(
                        [
                            {
                                "op": "image.update",
                                "target": cloned_image["id"],
                                "set": {
                                    "width": {
                                        "value": 180,
                                        "unit": "pt",
                                    }
                                },
                            }
                        ]
                    )
                    self.assertTrue(
                        resized.success,
                        resized.model_dump(),
                    )
                    assert resized.document is not None
                    resized_parts = {
                        part["id"]: part
                        for part in resized.document.to_spec()[
                            "header_footers"
                        ]
                    }
                    self.assertEqual(
                        resized_parts[source_part["id"]]["content"][0][
                            "width"
                        ],
                        {"value": 144.0, "unit": "pt"},
                    )
                    self.assertEqual(
                        resized_parts["compatibility_clone"]["content"][0][
                            "width"
                        ],
                        {"value": 180.0, "unit": "pt"},
                    )
                    if fallback_matches_choice:
                        replaced = resized.document.replace_image(
                            cloned_image["id"],
                            JPEG,
                            media_type="image/jpeg",
                        )
                        self.assertTrue(
                            replaced.success,
                            replaced.model_dump(),
                        )
                        assert replaced.document is not None
                        self.assertEqual(
                            replaced.document.image_bytes(
                                source_image["id"]
                            ),
                            PNG,
                        )
                        self.assertEqual(
                            replaced.document.image_bytes(
                                cloned_image["id"]
                            ),
                            JPEG,
                        )
                    else:
                        refused = resized.document.replace_image(
                            cloned_image["id"],
                            JPEG,
                            media_type="image/jpeg",
                        )
                        self.assertFalse(refused.success)
                        self.assertEqual(
                            refused.diagnostics[0].code,
                            "UNSUPPORTED_FEATURE",
                        )

        malformed_source = _alternate_content_image_document(
            _header_image_document(),
            part_name="word/header1.xml",
        )
        with ZipFile(io.BytesIO(malformed_source)) as archive:
            malformed_root = parse_xml(
                archive.read("word/header1.xml")
            )
        malformed_shape = malformed_root.find(
            f".//{_q(VML, 'shape')}"
        )
        assert malformed_shape is not None
        malformed_shape.set("future", "unsafe")
        malformed_source = _rewrite_package(
            malformed_source,
            replacements={
                "word/header1.xml": _serialize_with_wps(
                    malformed_root
                ),
            },
            additions={},
        )
        malformed_document = Document.from_docx(malformed_source)
        malformed_part = malformed_document.to_spec()[
            "header_footers"
        ][0]
        self.assertEqual(
            malformed_part["content"][0]["type"],
            "opaque",
        )
        refused_clone = malformed_document.apply(
            [
                {
                    "op": "header_footer.clone",
                    "target": malformed_part["id"],
                    "part": {"id": "unsafe_compatibility_clone"},
                }
            ]
        )
        self.assertFalse(refused_clone.success)
        self.assertEqual(
            refused_clone.diagnostics[0].code,
            "NATIVE_PATCH_FAILED",
        )
        self.assertIn(
            "VML",
            refused_clone.diagnostics[0].message,
        )

    def test_unrecognized_alternate_content_remains_opaque(self) -> None:
        mutations = {
            "missing_fallback": lambda alternate: alternate.remove(
                alternate.find(_q(MC, "Fallback"))
            ),
            "wrong_requires": lambda alternate: alternate.find(
                _q(MC, "Choice")
            ).set("Requires", "future"),
            "size_mismatch": lambda alternate: alternate.find(
                f".//{_q(VML, 'shape')}"
            ).set(
                "style",
                (
                    "width:145pt;height:72pt;"
                    "mso-wrap-style:none;v-text-anchor:middle"
                ),
            ),
            "extra_fallback_shape": lambda alternate: alternate.find(
                f"./{_q(MC, 'Fallback')}/{_q(W, 'pict')}"
            ).append(ET.Element(_q(VML, "shape"))),
        }
        for name, mutate in mutations.items():
            source = _alternate_content_image_document(_image_document())
            with ZipFile(io.BytesIO(source)) as archive:
                root = parse_xml(archive.read("word/document.xml"))
            alternate = root.find(f".//{_q(MC, 'AlternateContent')}")
            assert alternate is not None
            mutate(alternate)
            mutated = _rewrite_package(
                source,
                replacements={
                    "word/document.xml": _serialize_with_wps(root),
                },
                additions={},
            )
            with self.subTest(name=name):
                document = Document.from_docx(mutated)
                self.assertEqual(
                    document.to_spec()["content"][0]["type"],
                    "opaque",
                )
                self.assertEqual(document.to_spec()["assets"], [])

        source = _alternate_content_image_document(_image_document())
        with ZipFile(io.BytesIO(source)) as archive:
            document_payload = archive.read("word/document.xml")
        choice_tag = b'<mc:Choice Requires="wps">'
        self.assertIn(choice_tag, document_payload)
        rebound_payload = document_payload.replace(
            choice_tag,
            (
                b'<mc:Choice xmlns:wps="urn:aioffice:test:rebound" '
                b'Requires="wps">'
            ),
            1,
        )
        rebound = _rewrite_package(
            source,
            replacements={"word/document.xml": rebound_payload},
            additions={},
        )
        self.assertEqual(
            Document.from_docx(rebound).to_spec()["content"][0]["type"],
            "opaque",
        )

        missing_fallback_asset = _alternate_content_image_document(
            _image_document(),
            fallback_matches_choice=False,
        )
        missing_fallback_asset = _rewrite_package(
            missing_fallback_asset,
            replacements={},
            additions={},
            deletions={"word/media/fallback.jpg"},
        )
        self.assertEqual(
            Document.from_docx(
                missing_fallback_asset
            ).to_spec()["content"][0]["type"],
            "opaque",
        )

        anchored = _alternate_content_image_document(
            _image_document(anchored=True, aligned=True)
        )
        document = Document.from_docx(anchored)
        self.assertEqual(
            document.to_spec()["content"][0]["type"],
            "opaque",
        )

        floating_mutations = {
            "anchor_id_mismatch": lambda root: root.find(
                f".//{_q(VML, 'shape')}"
            ).set(_q(WP14, "anchorId"), "DEADBEEF"),
            "fallback_position_mismatch": lambda root: root.find(
                f".//{_q(VML, 'shape')}"
            ).set(
                "style",
                (
                    "position:absolute;margin-left:37pt;margin-top:0.4pt;"
                    "width:144pt;height:72pt;"
                    "mso-wrap-style:none;v-text-anchor:middle"
                ),
            ),
            "fallback_wrap_mismatch": lambda root: root.find(
                f".//{_q(WORD_VML, 'wrap')}"
            ).set("type", "tight"),
            "unsynchronized_anchor_distance": lambda root: root.find(
                f".//{_q(WP, 'anchor')}"
            ).set("distL", "12700"),
        }
        for name, mutate in floating_mutations.items():
            source = _alternate_content_image_document(
                _image_document(anchored=True)
            )
            with ZipFile(io.BytesIO(source)) as archive:
                root = parse_xml(archive.read("word/document.xml"))
            mutate(root)
            mutated = _rewrite_package(
                source,
                replacements={
                    "word/document.xml": _serialize_with_wps(root),
                },
                additions={},
            )
            with self.subTest(name=name):
                document = Document.from_docx(mutated)
                self.assertEqual(
                    document.to_spec()["content"][0]["type"],
                    "opaque",
                )

    def test_floating_square_wrap_projects_and_preserves_native_layout(
        self,
    ) -> None:
        source = _image_document(anchored=True, cropped=True)
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        self.assertEqual(image["type"], "image")
        self.assertEqual(image["placement"], "floating")
        expected_layout = {
            "horizontal": {
                "relative_to": "column",
                "offset": {"value": 36.0, "unit": "pt"},
            },
            "vertical": {
                "relative_to": "paragraph",
                "offset": {"value": 0.4, "unit": "pt"},
            },
            "anchor_distances": {
                "top": {"value": 1.0, "unit": "pt"},
                "right": {"value": 4.0, "unit": "pt"},
                "bottom": {"value": 2.0, "unit": "pt"},
                "left": {"value": 3.0, "unit": "pt"},
            },
            "anchor_effect_extent": {
                "left": {"value": 0.0, "unit": "pt"},
                "top": {"value": 0.0, "unit": "pt"},
                "right": {"value": 0.0, "unit": "pt"},
                "bottom": {"value": 0.0, "unit": "pt"},
            },
            "wrap": {
                "mode": "square",
                "side": "both_sides",
            },
            "relative_height": 1026,
            "behind_text": False,
            "locked": True,
            "layout_in_cell": True,
            "allow_overlap": True,
        }
        self.assertEqual(image["floating"], expected_layout)
        self.assertEqual(document.to_bytes("docx"), source)
        self.assertEqual(document.image_bytes(image["id"]), PNG)
        inspected = document.inspect()["nodes"][0]
        self.assertEqual(inspected["placement"], "floating")
        self.assertEqual(inspected["floating"], expected_layout)
        self.assertIn(
            "image.anchor.update",
            inspected["supported_operations"],
        )
        capabilities = document.capabilities()
        self.assertIn(
            "image.anchor.update",
            capabilities["operations"],
        )
        self.assertTrue(
            capabilities["assets"]["floating_layout_editable"]
        )
        self.assertEqual(
            capabilities["assets"]["floating_layout_update_operation"],
            "image.anchor.update",
        )
        self.assertIn(
            'data-aioffice-placement="floating"',
            document.to_bytes("html").decode(),
        )

        result = document.apply(
            [
                {
                    "op": "image.update",
                    "target": f"#{image['id']}",
                    "set": {
                        "width": {"value": 3, "unit": "in"},
                        "crop": {"left": 12.5, "right": 12.5},
                        "alt_text": "Floating expert workflow",
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.diagnostics)
        assert result.document is not None
        updated = result.document.to_spec()["content"][0]
        self.assertEqual(updated["floating"], expected_layout)
        self.assertEqual(updated["placement"], "floating")
        self.assertEqual(result.document.image_bytes(image["id"]), PNG)

        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(result.document.to_bytes("docx"))) as after,
        ):
            self.assertEqual(
                before.read("word/media/image1.png"),
                after.read("word/media/image1.png"),
            )
            self.assertEqual(
                before.read("word/_rels/document.xml.rels"),
                after.read("word/_rels/document.xml.rels"),
            )
            updated_root = parse_xml(after.read("word/document.xml"))
        anchor = updated_root.find(f".//{_q(WP, 'anchor')}")
        assert anchor is not None
        self.assertEqual(
            anchor.find(f"./{_q(WP, 'positionH')}/{_q(WP, 'posOffset')}").text,
            "457200",
        )
        self.assertEqual(
            anchor.find(f"./{_q(WP, 'positionV')}/{_q(WP, 'posOffset')}").text,
            "5080",
        )
        self.assertEqual(
            anchor.find(f"./{_q(WP, 'wrapSquare')}").attrib,
            {"wrapText": "bothSides"},
        )

        reopened = Document.from_docx(result.document.to_bytes("docx"))
        reopened_image = reopened.to_spec()["content"][0]
        self.assertEqual(reopened_image["floating"], expected_layout)
        self.assertEqual(reopened_image["crop"]["left"], 12.5)

        replaced = result.document.replace_image(image["id"], JPEG)
        self.assertTrue(replaced.success, replaced.diagnostics)
        assert replaced.document is not None
        replaced_image = replaced.document.to_spec()["content"][0]
        self.assertEqual(replaced_image["floating"], expected_layout)
        self.assertEqual(replaced_image["crop"], updated["crop"])
        self.assertEqual(replaced.document.image_bytes(image["id"]), JPEG)

    def test_none_and_top_bottom_wrap_project_preserve_and_roundtrip(
        self,
    ) -> None:
        native_tags = {
            "none": "wrapNone",
            "top_and_bottom": "wrapTopAndBottom",
        }
        for mode, native_tag in native_tags.items():
            with self.subTest(mode=mode):
                source = _image_document(
                    anchored=True,
                    wrap_mode=mode,
                    cropped=True,
                )
                document = Document.from_docx(source)
                image = document.to_spec()["content"][0]
                self.assertEqual(image["type"], "image")
                self.assertEqual(image["placement"], "floating")
                self.assertEqual(
                    image["floating"]["wrap"],
                    {"mode": mode},
                )
                self.assertEqual(
                    image["floating"]["anchor_distances"],
                    {
                        "top": {"value": 1.0, "unit": "pt"},
                        "right": {"value": 4.0, "unit": "pt"},
                        "bottom": {"value": 2.0, "unit": "pt"},
                        "left": {"value": 3.0, "unit": "pt"},
                    },
                )
                self.assertNotIn("side", image["floating"]["wrap"])
                self.assertEqual(document.to_bytes("docx"), source)
                self.assertEqual(document.image_bytes(image["id"]), PNG)

                updated = document.apply(
                    [
                        {
                            "op": "image.update",
                            "target": image["id"],
                            "set": {
                                "alt_text": (
                                    f"Preserved {mode} floating image"
                                )
                            },
                        }
                    ]
                )
                self.assertTrue(updated.success, updated.model_dump())
                assert updated.document is not None
                output = updated.document.to_bytes("docx")
                with ZipFile(io.BytesIO(output)) as package:
                    root = parse_xml(package.read("word/document.xml"))
                wrap = root.find(f".//{_q(WP, native_tag)}")
                assert wrap is not None
                self.assertFalse(wrap.attrib)
                self.assertFalse(len(wrap))
                reopened = Document.from_docx(output)
                reopened_image = reopened.to_spec()["content"][0]
                self.assertEqual(
                    reopened_image["floating"],
                    updated.document.to_spec()["content"][0]["floating"],
                )
                self.assertEqual(reopened.image_bytes(image["id"]), PNG)

    def test_tight_and_through_polygons_project_preserve_and_update(
        self,
    ) -> None:
        native_tags = {
            "tight": "wrapTight",
            "through": "wrapThrough",
        }
        original_polygon = {
            "edited": True,
            "start": {"x": 0, "y": 0},
            "line_to": [
                {"x": 0, "y": 21600},
                {"x": 21600, "y": 21600},
                {"x": 21600, "y": 0},
            ],
        }
        updated_polygon = {
            "edited": False,
            "start": {"x": -10800, "y": 0},
            "line_to": [
                {"x": -10800, "y": 10800},
                {"x": 21600, "y": 10800},
                {"x": 21600, "y": 0},
                {"x": -10800, "y": 0},
            ],
        }
        for mode, native_tag in native_tags.items():
            with self.subTest(mode=mode):
                source = _image_document(
                    anchored=True,
                    wrap_mode=mode,
                    cropped=True,
                )
                document = Document.from_docx(source)
                image = document.to_spec()["content"][0]
                self.assertEqual(
                    image["floating"]["wrap"],
                    {
                        "mode": mode,
                        "side": "both_sides",
                        "polygon": original_polygon,
                    },
                )
                self.assertEqual(document.to_bytes("docx"), source)
                self.assertEqual(document.image_bytes(image["id"]), PNG)

                replacement_mode = (
                    "through" if mode == "tight" else "tight"
                )
                requested_wrap = {
                    "mode": replacement_mode,
                    "side": "largest",
                    "distances": {
                        "left": {"value": 3, "unit": "pt"},
                        "right": {"value": 4, "unit": "pt"},
                    },
                    "polygon": updated_polygon,
                }
                result = document.apply(
                    [
                        {
                            "op": "image.anchor.update",
                            "target": image["id"],
                            "set": {"wrap": requested_wrap},
                        }
                    ]
                )
                self.assertTrue(result.success, result.model_dump())
                assert result.document is not None
                updated_image = result.document.to_spec()["content"][0]
                self.assertEqual(
                    updated_image["floating"]["wrap"],
                    requested_wrap,
                )
                output = result.document.to_bytes("docx")
                with ZipFile(io.BytesIO(output)) as package:
                    root = parse_xml(package.read("word/document.xml"))
                self.assertIsNone(
                    root.find(f".//{_q(WP, native_tag)}")
                )
                native_wrap = root.find(
                    f".//{_q(WP, native_tags[replacement_mode])}"
                )
                assert native_wrap is not None
                self.assertEqual(
                    native_wrap.attrib,
                    {
                        "wrapText": "largest",
                        "distR": "50800",
                        "distL": "38100",
                    },
                )
                polygon = native_wrap.find(
                    f"./{_q(WP, 'wrapPolygon')}"
                )
                assert polygon is not None
                self.assertEqual(polygon.attrib, {"edited": "0"})
                self.assertEqual(
                    [
                        (
                            child.tag,
                            child.get("x"),
                            child.get("y"),
                        )
                        for child in polygon
                    ],
                    [
                        (_q(WP, "start"), "-10800", "0"),
                        (_q(WP, "lineTo"), "-10800", "10800"),
                        (_q(WP, "lineTo"), "21600", "10800"),
                        (_q(WP, "lineTo"), "21600", "0"),
                        (_q(WP, "lineTo"), "-10800", "0"),
                    ],
                )
                reopened = Document.from_docx(output)
                self.assertEqual(
                    reopened.to_spec()["content"][0]["floating"]["wrap"],
                    requested_wrap,
                )
                self.assertEqual(reopened.image_bytes(image["id"]), PNG)

    def test_layered_anchor_and_wrap_geometry_updates_and_clears(
        self,
    ) -> None:
        source = _image_document(anchored=True, cropped=True)
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
        anchor = root.find(f".//{_q(WP, 'anchor')}")
        assert anchor is not None
        anchor_effect = anchor.find(f"./{_q(WP, 'effectExtent')}")
        wrap = anchor.find(f"./{_q(WP, 'wrapSquare')}")
        assert anchor_effect is not None
        assert wrap is not None
        anchor_effect.attrib.update(
            {
                "l": "-6350",
                "t": "12700",
                "r": "19050",
                "b": "25400",
            }
        )
        wrap.attrib.update(
            {
                "distT": "63500",
                "distR": "76200",
                "distB": "88900",
                "distL": "101600",
            }
        )
        ET.SubElement(
            wrap,
            _q(WP, "effectExtent"),
            {
                "l": "114300",
                "t": "127000",
                "r": "139700",
                "b": "152400",
            },
        )
        layered = _rewrite_package(
            source,
            replacements={
                "word/document.xml": serialize_xml(root),
            },
            additions={},
        )

        document = Document.from_docx(layered)
        image = document.to_spec()["content"][0]
        floating = image["floating"]
        self.assertEqual(
            floating["anchor_distances"],
            {
                "top": {"value": 1.0, "unit": "pt"},
                "right": {"value": 4.0, "unit": "pt"},
                "bottom": {"value": 2.0, "unit": "pt"},
                "left": {"value": 3.0, "unit": "pt"},
            },
        )
        self.assertEqual(
            floating["anchor_effect_extent"],
            {
                "left": {"value": -0.5, "unit": "pt"},
                "top": {"value": 1.0, "unit": "pt"},
                "right": {"value": 1.5, "unit": "pt"},
                "bottom": {"value": 2.0, "unit": "pt"},
            },
        )
        self.assertEqual(
            floating["wrap"],
            {
                "mode": "square",
                "side": "both_sides",
                "distances": {
                    "top": {"value": 5.0, "unit": "pt"},
                    "right": {"value": 6.0, "unit": "pt"},
                    "bottom": {"value": 7.0, "unit": "pt"},
                    "left": {"value": 8.0, "unit": "pt"},
                },
                "effect_extent": {
                    "left": {"value": 9.0, "unit": "pt"},
                    "top": {"value": 10.0, "unit": "pt"},
                    "right": {"value": 11.0, "unit": "pt"},
                    "bottom": {"value": 12.0, "unit": "pt"},
                },
            },
        )
        self.assertEqual(document.to_bytes("docx"), layered)

        requested = {
            "anchor_distances": {
                "top": {"value": 2.5, "unit": "pt"},
            },
            "anchor_effect_extent": {
                "left": {"value": -1, "unit": "pt"},
                "top": {"value": 2, "unit": "pt"},
                "right": {"value": 3, "unit": "pt"},
                "bottom": {"value": 4, "unit": "pt"},
            },
            "wrap": {
                "mode": "top_and_bottom",
                "distances": {
                    "top": {"value": 6, "unit": "pt"},
                    "bottom": {"value": 7, "unit": "pt"},
                },
                "effect_extent": {
                    "left": {"value": -2, "unit": "pt"},
                    "top": {"value": 1, "unit": "pt"},
                    "right": {"value": 2, "unit": "pt"},
                    "bottom": {"value": 3, "unit": "pt"},
                },
            },
        }
        updated = document.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": image["id"],
                    "set": requested,
                }
            ]
        )
        self.assertTrue(updated.success, updated.model_dump())
        assert updated.document is not None
        updated_layout = updated.document.to_spec()["content"][0][
            "floating"
        ]
        for field_name, value in requested.items():
            self.assertEqual(updated_layout[field_name], value)
        output = updated.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as package:
            updated_root = parse_xml(package.read("word/document.xml"))
            self.assertEqual(
                package.read("word/media/image1.png"),
                PNG,
            )
        updated_anchor = updated_root.find(f".//{_q(WP, 'anchor')}")
        assert updated_anchor is not None
        self.assertEqual(
            {
                name: updated_anchor.get(name)
                for name in ("distT", "distR", "distB", "distL")
            },
            {
                "distT": "31750",
                "distR": None,
                "distB": None,
                "distL": None,
            },
        )
        updated_anchor_effect = updated_anchor.find(
            f"./{_q(WP, 'effectExtent')}"
        )
        assert updated_anchor_effect is not None
        self.assertEqual(
            updated_anchor_effect.attrib,
            {
                "l": "-12700",
                "t": "25400",
                "r": "38100",
                "b": "50800",
            },
        )
        updated_wrap = updated_anchor.find(
            f"./{_q(WP, 'wrapTopAndBottom')}"
        )
        assert updated_wrap is not None
        self.assertEqual(
            updated_wrap.attrib,
            {"distT": "76200", "distB": "88900"},
        )
        updated_wrap_effect = updated_wrap.find(
            f"./{_q(WP, 'effectExtent')}"
        )
        assert updated_wrap_effect is not None
        self.assertEqual(
            updated_wrap_effect.attrib,
            {
                "l": "-25400",
                "t": "12700",
                "r": "25400",
                "b": "38100",
            },
        )
        reopened = Document.from_docx(output)
        self.assertEqual(
            reopened.to_spec()["content"][0]["floating"],
            updated_layout,
        )

        cleared = reopened.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": image["id"],
                    "clear": [
                        "anchor_distances",
                        "anchor_effect_extent",
                    ],
                }
            ]
        )
        self.assertTrue(cleared.success, cleared.model_dump())
        assert cleared.document is not None
        cleared_layout = cleared.document.to_spec()["content"][0][
            "floating"
        ]
        self.assertNotIn("anchor_distances", cleared_layout)
        self.assertNotIn("anchor_effect_extent", cleared_layout)
        cleared_bytes = cleared.document.to_bytes("docx")
        with ZipFile(io.BytesIO(cleared_bytes)) as package:
            cleared_root = parse_xml(package.read("word/document.xml"))
        cleared_anchor = cleared_root.find(f".//{_q(WP, 'anchor')}")
        assert cleared_anchor is not None
        for attribute_name in ("distT", "distR", "distB", "distL"):
            self.assertNotIn(attribute_name, cleared_anchor.attrib)
        self.assertIsNone(
            cleared_anchor.find(f"./{_q(WP, 'effectExtent')}")
        )
        self.assertEqual(
            Document.from_docx(cleared_bytes).to_spec()["content"][0][
                "floating"
            ],
            cleared_layout,
        )

    def test_floating_wrap_mode_switch_replaces_exact_native_child(
        self,
    ) -> None:
        source = _image_document(
            anchored=True,
            wrap_mode="none",
        )
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        distances = {
            "distance_top": {"value": 5, "unit": "pt"},
            "distance_right": {"value": 6, "unit": "pt"},
            "distance_bottom": {"value": 7, "unit": "pt"},
            "distance_left": {"value": 8, "unit": "pt"},
        }

        top_bottom = document.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": image["id"],
                    "set": {
                        "wrap": {
                            "mode": "top_and_bottom",
                            **distances,
                        }
                    },
                }
            ]
        )
        self.assertTrue(top_bottom.success, top_bottom.model_dump())
        assert top_bottom.document is not None
        top_bottom_bytes = top_bottom.document.to_bytes("docx")
        with ZipFile(io.BytesIO(top_bottom_bytes)) as package:
            root = parse_xml(package.read("word/document.xml"))
        anchor = root.find(f".//{_q(WP, 'anchor')}")
        assert anchor is not None
        self.assertIsNone(anchor.find(f"./{_q(WP, 'wrapNone')}"))
        native_top_bottom = anchor.find(
            f"./{_q(WP, 'wrapTopAndBottom')}"
        )
        assert native_top_bottom is not None
        self.assertFalse(native_top_bottom.attrib)
        self.assertEqual(
            {
                name: anchor.get(attribute)
                for name, attribute in (
                    ("top", "distT"),
                    ("right", "distR"),
                    ("bottom", "distB"),
                    ("left", "distL"),
                )
            },
            {
                "top": "63500",
                "right": "76200",
                "bottom": "88900",
                "left": "101600",
            },
        )

        square = Document.from_docx(top_bottom_bytes).apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": image["id"],
                    "set": {
                        "wrap": {
                            "mode": "square",
                            "side": "largest",
                            **distances,
                        }
                    },
                }
            ]
        )
        self.assertTrue(square.success, square.model_dump())
        assert square.document is not None
        square_bytes = square.document.to_bytes("docx")
        with ZipFile(io.BytesIO(square_bytes)) as package:
            root = parse_xml(package.read("word/document.xml"))
        anchor = root.find(f".//{_q(WP, 'anchor')}")
        assert anchor is not None
        self.assertIsNone(
            anchor.find(f"./{_q(WP, 'wrapTopAndBottom')}")
        )
        native_square = anchor.find(f"./{_q(WP, 'wrapSquare')}")
        assert native_square is not None
        self.assertEqual(native_square.attrib, {"wrapText": "largest"})
        reopened = Document.from_docx(square_bytes)
        reopened_layout = reopened.to_spec()["content"][0]["floating"]
        self.assertEqual(
            reopened_layout["wrap"],
            {
                "mode": "square",
                "side": "largest",
            },
        )
        self.assertEqual(
            reopened_layout["anchor_distances"],
            {
                field_name.removeprefix("distance_"): {
                    "value": float(value["value"]),
                    "unit": "pt",
                }
                for field_name, value in distances.items()
            },
        )
        self.assertEqual(reopened.image_bytes(image["id"]), PNG)

    def test_floating_alignment_positions_project_switch_and_roundtrip(
        self,
    ) -> None:
        source = _image_document(
            anchored=True,
            aligned=True,
            cropped=True,
        )
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        self.assertEqual(
            image["floating"]["horizontal"],
            {
                "relative_to": "margin",
                "alignment": "center",
            },
        )
        self.assertEqual(
            image["floating"]["vertical"],
            {
                "relative_to": "page",
                "alignment": "bottom",
            },
        )
        self.assertEqual(document.to_bytes("docx"), source)
        self.assertEqual(
            document.inspect()["nodes"][0]["floating"],
            image["floating"],
        )
        capabilities = document.capabilities()["assets"]
        self.assertEqual(
            capabilities["projected_placements"],
            [
                "inline",
                (
                    "floating_offset_alignment_or_percentage_"
                    "supported_wrap"
                ),
            ],
        )
        self.assertEqual(
            capabilities["floating_position_modes"],
            ["offset", "alignment", "percentage_offset"],
        )
        self.assertEqual(
            capabilities["floating_horizontal_alignments"],
            ["left", "right", "center", "inside", "outside"],
        )
        self.assertEqual(
            capabilities["floating_vertical_alignments"],
            ["top", "bottom", "center", "inside", "outside"],
        )
        self.assertEqual(
            capabilities["floating_wrap_modes"],
            [
                "square",
                "none",
                "top_and_bottom",
                "tight",
                "through",
            ],
        )
        self.assertEqual(
            capabilities["floating_square_wrap_sides"],
            ["both_sides", "largest", "left", "right"],
        )
        self.assertEqual(
            capabilities["floating_polygon_wrap_modes"],
            ["tight", "through"],
        )
        self.assertEqual(
            capabilities["floating_polygon_wrap_sides"],
            ["both_sides", "largest", "left", "right"],
        )
        self.assertEqual(
            capabilities["floating_polygon_schema"],
            "floating-image-wrap-polygon",
        )
        self.assertEqual(
            capabilities["floating_polygon_point_schema"],
            "floating-image-wrap-point",
        )
        self.assertEqual(
            capabilities["floating_polygon_line_to_bounds"],
            {"minimum": 2, "maximum": 4096},
        )
        self.assertEqual(
            capabilities["floating_wrap_distance_authority"],
            "native_anchor_and_wrap_element_attributes_are_separate",
        )
        self.assertEqual(
            capabilities["floating_effect_extent_authority"],
            "wrap_child_overrides_anchor_for_square_and_top_bottom",
        )
        self.assertEqual(
            capabilities["floating_layout_clearable_fields"],
            [
                "anchor_distances",
                "anchor_effect_extent",
                "relative_size",
            ],
        )

        switched = document.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": image["id"],
                    "set": {
                        "horizontal": {
                            "relative_to": "page",
                            "offset": {"value": 48, "unit": "pt"},
                        },
                        "vertical": {
                            "relative_to": "margin",
                            "alignment": "top",
                        },
                    },
                }
            ]
        )
        self.assertTrue(switched.success, switched.model_dump())
        assert switched.document is not None
        switched_image = switched.document.to_spec()["content"][0]
        self.assertEqual(
            switched_image["floating"]["horizontal"],
            {
                "relative_to": "page",
                "offset": {"value": 48.0, "unit": "pt"},
            },
        )
        self.assertEqual(
            switched_image["floating"]["vertical"],
            {
                "relative_to": "margin",
                "alignment": "top",
            },
        )
        switched_bytes = switched.document.to_bytes("docx")
        with ZipFile(io.BytesIO(switched_bytes)) as package:
            root = parse_xml(package.read("word/document.xml"))
        horizontal = root.find(f".//{_q(WP, 'positionH')}")
        vertical = root.find(f".//{_q(WP, 'positionV')}")
        assert horizontal is not None
        assert vertical is not None
        self.assertEqual(
            [child.tag for child in horizontal],
            [_q(WP, "posOffset")],
        )
        self.assertEqual(horizontal[0].text, "609600")
        self.assertEqual(
            [child.tag for child in vertical],
            [_q(WP, "align")],
        )
        self.assertEqual(vertical[0].text, "top")

        reopened = Document.from_docx(switched_bytes)
        restored = reopened.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": image["id"],
                    "set": {
                        "horizontal": {
                            "relative_to": "outside_margin",
                            "alignment": "outside",
                        },
                        "vertical": {
                            "relative_to": "paragraph",
                            "offset": {"value": -12, "unit": "pt"},
                        },
                    },
                }
            ]
        )
        self.assertTrue(restored.success, restored.model_dump())
        assert restored.document is not None
        restored_bytes = restored.document.to_bytes("docx")
        with ZipFile(io.BytesIO(restored_bytes)) as package:
            root = parse_xml(package.read("word/document.xml"))
        horizontal = root.find(f".//{_q(WP, 'positionH')}")
        vertical = root.find(f".//{_q(WP, 'positionV')}")
        assert horizontal is not None
        assert vertical is not None
        self.assertEqual(horizontal.get("relativeFrom"), "outsideMargin")
        self.assertEqual(horizontal[0].tag, _q(WP, "align"))
        self.assertEqual(horizontal[0].text, "outside")
        self.assertEqual(vertical.get("relativeFrom"), "paragraph")
        self.assertEqual(vertical[0].tag, _q(WP, "posOffset"))
        self.assertEqual(vertical[0].text, "-152400")
        roundtripped = Document.from_docx(restored_bytes)
        self.assertEqual(
            roundtripped.to_spec()["content"][0]["floating"],
            restored.document.to_spec()["content"][0]["floating"],
        )
        self.assertEqual(
            roundtripped.image_bytes(image["id"]),
            PNG,
        )

    def test_floating_percentage_positions_project_switch_and_roundtrip(
        self,
    ) -> None:
        source = _image_document(
            anchored=True,
            percentage_position=True,
            cropped=True,
        )
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        self.assertEqual(
            image["floating"]["horizontal"],
            {
                "relative_to": "column",
                "percentage_offset": 37.5,
            },
        )
        self.assertEqual(
            image["floating"]["vertical"],
            {
                "relative_to": "paragraph",
                "percentage_offset": -12.5,
            },
        )
        self.assertEqual(document.to_bytes("docx"), source)
        capabilities = document.capabilities()["assets"]
        self.assertEqual(
            capabilities["floating_percentage_offset_unit"],
            "percentage_points",
        )
        self.assertEqual(
            capabilities["floating_percentage_offset_precision"],
            0.001,
        )
        self.assertEqual(
            capabilities["floating_percentage_offset_native_type"],
            "signed_int32_st_percentage",
        )
        self.assertEqual(
            capabilities["floating_percentage_offset_bounds"],
            {
                "minimum": -2147483.648,
                "maximum": 2147483.647,
            },
        )

        updated = document.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": image["id"],
                    "set": {
                        "horizontal": {
                            "relative_to": "page",
                            "percentage_offset": 62.3454,
                        },
                        "vertical": {
                            "relative_to": "margin",
                            "alignment": "bottom",
                        },
                    },
                }
            ]
        )
        self.assertTrue(updated.success, updated.model_dump())
        assert updated.document is not None
        updated_bytes = updated.document.to_bytes("docx")
        with ZipFile(io.BytesIO(updated_bytes)) as package:
            native_xml = package.read("word/document.xml")
            root = parse_xml(native_xml)
        horizontal = root.find(f".//{_q(WP, 'positionH')}")
        vertical = root.find(f".//{_q(WP, 'positionV')}")
        assert horizontal is not None
        assert vertical is not None
        self.assertEqual(horizontal.get("relativeFrom"), "page")
        self.assertEqual(
            [child.tag for child in horizontal],
            [_q(WP14, "pctPosHOffset")],
        )
        self.assertEqual(horizontal[0].text, "62345")
        self.assertEqual(vertical.get("relativeFrom"), "margin")
        self.assertEqual(vertical[0].tag, _q(WP, "align"))
        self.assertEqual(vertical[0].text, "bottom")
        self.assertIn(
            "wp14",
            (root.get(_q(MC, "Ignorable")) or "").split(),
        )
        self.assertIn(b"xmlns:wp14=", native_xml)

        reopened = Document.from_docx(updated_bytes)
        self.assertEqual(
            reopened.to_spec()["content"][0]["floating"]["horizontal"],
            {
                "relative_to": "page",
                "percentage_offset": 62.345,
            },
        )
        self.assertEqual(reopened.to_bytes("docx"), updated_bytes)

        metadata_updated = reopened.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {
                        "alt_text": "Percentage-positioned expert diagram",
                    },
                }
            ]
        )
        self.assertTrue(
            metadata_updated.success,
            metadata_updated.model_dump(),
        )
        assert metadata_updated.document is not None
        self.assertEqual(
            metadata_updated.document.to_spec()["content"][0][
                "floating"
            ],
            reopened.to_spec()["content"][0]["floating"],
        )
        with ZipFile(
            io.BytesIO(metadata_updated.document.to_bytes("docx"))
        ) as package:
            metadata_root = parse_xml(
                package.read("word/document.xml")
            )
        metadata_horizontal = metadata_root.find(
            f".//{_q(WP, 'positionH')}"
        )
        assert metadata_horizontal is not None
        self.assertEqual(
            ET.tostring(metadata_horizontal),
            ET.tostring(horizontal),
        )

    def test_floating_relative_size_projects_updates_and_clears(self) -> None:
        source = _image_document(
            anchored=True,
            relative_size=True,
            cropped=True,
        )
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        self.assertEqual(
            image["floating"]["relative_size"],
            {
                "width": {
                    "relative_to": "margin",
                    "percentage": 50.0,
                },
                "height": {
                    "relative_to": "page",
                    "percentage": 25.0,
                },
            },
        )
        self.assertEqual(image["width"], {"value": 144.0, "unit": "pt"})
        self.assertEqual(image["height"], {"value": 72.0, "unit": "pt"})
        self.assertEqual(document.to_bytes("docx"), source)

        capabilities = document.capabilities()["assets"]
        self.assertIn(
            "relative_size",
            capabilities["floating_layout_update_fields"],
        )
        self.assertIn(
            "relative_size",
            capabilities["floating_layout_clearable_fields"],
        )
        self.assertEqual(
            capabilities["floating_relative_size_axes"],
            ["width", "height"],
        )
        self.assertEqual(
            capabilities["floating_relative_size_unit"],
            "percentage_points",
        )
        self.assertEqual(
            capabilities["floating_relative_size_precision"],
            0.001,
        )
        self.assertEqual(
            capabilities["floating_relative_size_bounds"],
            {"minimum": 0, "maximum": 2147483.647},
        )
        self.assertIn(
            "absolute_native_fallback_extent",
            capabilities["floating_relative_size_authority"],
        )

        updated = document.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": image["id"],
                    "set": {
                        "relative_size": {
                            "width": {
                                "relative_to": "right_margin",
                                "percentage": 62.3454,
                            }
                        }
                    },
                },
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {
                        "width": {"value": 180, "unit": "pt"},
                        "height": {"value": 90, "unit": "pt"},
                    },
                },
            ]
        )
        self.assertTrue(updated.success, updated.model_dump())
        assert updated.document is not None
        updated_image = updated.document.to_spec()["content"][0]
        self.assertEqual(
            updated_image["floating"]["relative_size"],
            {
                "width": {
                    "relative_to": "right_margin",
                    "percentage": 62.345,
                }
            },
        )
        self.assertEqual(
            updated_image["width"],
            {"value": 180.0, "unit": "pt"},
        )
        self.assertEqual(
            updated_image["height"],
            {"value": 90.0, "unit": "pt"},
        )
        output = updated.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as package:
            native_xml = package.read("word/document.xml")
            root = parse_xml(native_xml)
        anchor = root.find(f".//{_q(WP, 'anchor')}")
        assert anchor is not None
        self.assertEqual(anchor[-1].tag, _q(WP14, "sizeRelH"))
        self.assertIsNone(anchor.find(f"./{_q(WP14, 'sizeRelV')}"))
        relative_width = anchor.find(f"./{_q(WP14, 'sizeRelH')}")
        assert relative_width is not None
        self.assertEqual(
            relative_width.attrib,
            {"relativeFrom": "rightMargin"},
        )
        self.assertEqual(relative_width[0].tag, _q(WP14, "pctWidth"))
        self.assertEqual(relative_width[0].text, "62345")
        extent = anchor.find(f"./{_q(WP, 'extent')}")
        assert extent is not None
        self.assertEqual(
            (extent.get("cx"), extent.get("cy")),
            ("2286000", "1143000"),
        )
        self.assertIn(
            "wp14",
            (root.get(_q(MC, "Ignorable")) or "").split(),
        )
        self.assertIn(b"xmlns:wp14=", native_xml)

        reopened = Document.from_docx(output)
        self.assertEqual(
            reopened.to_spec()["content"][0]["floating"][
                "relative_size"
            ],
            updated_image["floating"]["relative_size"],
        )
        cleared = reopened.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": image["id"],
                    "clear": ["relative_size"],
                }
            ]
        )
        self.assertTrue(cleared.success, cleared.model_dump())
        assert cleared.document is not None
        self.assertNotIn(
            "relative_size",
            cleared.document.to_spec()["content"][0]["floating"],
        )
        with ZipFile(
            io.BytesIO(cleared.document.to_bytes("docx"))
        ) as package:
            cleared_root = parse_xml(package.read("word/document.xml"))
        self.assertIsNone(
            cleared_root.find(f".//{_q(WP14, 'sizeRelH')}")
        )
        self.assertIsNone(
            cleared_root.find(f".//{_q(WP14, 'sizeRelV')}")
        )

    def test_libreoffice_neutral_picture_normalization_is_editable(
        self,
    ) -> None:
        source = _image_document(
            anchored=True,
            aligned=True,
            cropped=True,
        )
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
        shape = root.find(f".//{_q(PIC, 'spPr')}")
        assert shape is not None
        shape.set("bwMode", "auto")
        ET.SubElement(shape, _q(A, "noFill"))
        normalized = _rewrite_package(
            source,
            replacements={
                "word/document.xml": serialize_xml(root),
            },
            additions={},
        )
        document = Document.from_docx(normalized)
        image = document.to_spec()["content"][0]
        self.assertEqual(image["type"], "image")
        self.assertEqual(
            image["floating"]["horizontal"]["alignment"],
            "center",
        )
        self.assertEqual(document.to_bytes("docx"), normalized)

        updated = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {
                        "alt_text": (
                            "LibreOffice-normalized aligned picture"
                        )
                    },
                }
            ]
        )
        self.assertTrue(updated.success, updated.model_dump())
        assert updated.document is not None
        output = updated.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as package:
            root = parse_xml(package.read("word/document.xml"))
        shape = root.find(f".//{_q(PIC, 'spPr')}")
        assert shape is not None
        self.assertEqual(shape.attrib, {"bwMode": "auto"})
        no_fills = shape.findall(f"./{_q(A, 'noFill')}")
        self.assertEqual(len(no_fills), 1)
        self.assertFalse(no_fills[0].attrib)
        self.assertFalse(len(no_fills[0]))
        reopened = Document.from_docx(output)
        self.assertEqual(
            reopened.to_spec()["content"][0]["floating"],
            image["floating"],
        )

    def test_floating_anchor_update_is_selective_and_roundtrips(
        self,
    ) -> None:
        source = _image_document(anchored=True, cropped=True)
        document = Document.from_docx(source)
        before_image = document.to_spec()["content"][0]
        requested = {
            "horizontal": {
                "relative_to": "page",
                "offset": {"value": 72, "unit": "pt"},
            },
            "vertical": {
                "relative_to": "margin",
                "offset": {"value": -18, "unit": "pt"},
            },
            "anchor_distances": {
                "top": {"value": 5, "unit": "pt"},
                "right": {"value": 6, "unit": "pt"},
                "bottom": {"value": 7, "unit": "pt"},
                "left": {"value": 8, "unit": "pt"},
            },
            "anchor_effect_extent": {
                "left": {"value": -0.5, "unit": "pt"},
                "top": {"value": 1, "unit": "pt"},
                "right": {"value": 1.5, "unit": "pt"},
                "bottom": {"value": 2, "unit": "pt"},
            },
            "wrap": {
                "mode": "square",
                "side": "right",
            },
            "relative_height": 2048,
            "behind_text": True,
            "locked": False,
            "layout_in_cell": False,
            "allow_overlap": False,
        }
        result = document.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": f"#{before_image['id']}",
                    "set": requested,
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(result.changes[0]["operation"], "image.anchor.update")
        self.assertEqual(
            {
                change["path"]
                for change in result.changes[0]["property_changes"]
            },
            {f"floating.{field_name}" for field_name in requested},
        )
        assert result.document is not None
        after_image = result.document.to_spec()["content"][0]
        self.assertEqual(after_image["floating"], requested)
        for field_name in (
            "id",
            "asset_id",
            "placement",
            "width",
            "height",
            "crop",
            "transform",
            "name",
            "alt_text",
            "title",
            "capabilities",
            "editable",
        ):
            self.assertEqual(
                after_image.get(field_name),
                before_image.get(field_name),
            )
        self.assertEqual(
            result.document.image_bytes(before_image["id"]),
            PNG,
        )

        output = result.document.to_bytes("docx")
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(output)) as after,
        ):
            self.assertEqual(
                before.read("word/_rels/document.xml.rels"),
                after.read("word/_rels/document.xml.rels"),
            )
            self.assertEqual(
                before.read("word/media/image1.png"),
                after.read("word/media/image1.png"),
            )
            before_root = parse_xml(before.read("word/document.xml"))
            after_root = parse_xml(after.read("word/document.xml"))
        expected_anchor = copy.deepcopy(
            before_root.find(f".//{_q(WP, 'anchor')}")
        )
        after_anchor = after_root.find(f".//{_q(WP, 'anchor')}")
        assert expected_anchor is not None
        assert after_anchor is not None
        expected_horizontal = expected_anchor.find(
            f"./{_q(WP, 'positionH')}"
        )
        expected_vertical = expected_anchor.find(
            f"./{_q(WP, 'positionV')}"
        )
        expected_wrap = expected_anchor.find(
            f"./{_q(WP, 'wrapSquare')}"
        )
        assert expected_horizontal is not None
        assert expected_vertical is not None
        assert expected_wrap is not None
        expected_horizontal.set("relativeFrom", "page")
        expected_horizontal.find(
            f"./{_q(WP, 'posOffset')}"
        ).text = "914400"
        expected_vertical.set("relativeFrom", "margin")
        expected_vertical.find(
            f"./{_q(WP, 'posOffset')}"
        ).text = "-228600"
        expected_wrap.set("wrapText", "right")
        expected_effect_extent = expected_anchor.find(
            f"./{_q(WP, 'effectExtent')}"
        )
        assert expected_effect_extent is not None
        expected_effect_extent.attrib.update(
            {
                "l": "-6350",
                "t": "12700",
                "r": "19050",
                "b": "25400",
            }
        )
        expected_anchor.attrib.update(
            {
                "distT": "63500",
                "distR": "76200",
                "distB": "88900",
                "distL": "101600",
                "relativeHeight": "2048",
                "behindDoc": "1",
                "locked": "0",
                "layoutInCell": "0",
                "allowOverlap": "0",
            }
        )
        self.assertEqual(
            ET.tostring(after_anchor),
            ET.tostring(expected_anchor),
        )
        self.assertEqual(
            after_anchor.get(_q(WP14, "anchorId")),
            "A1B2C3D4",
        )
        self.assertEqual(
            after_anchor.get(_q(WP14, "editId")),
            "E5F60718",
        )

        reopened = Document.from_docx(output)
        reopened_image = reopened.to_spec()["content"][0]
        self.assertEqual(reopened_image["floating"], requested)
        self.assertEqual(reopened.image_bytes(before_image["id"]), PNG)
        self.assertEqual(reopened.to_bytes("docx"), output)

    def test_floating_anchor_update_rejects_unsafe_requests_atomically(
        self,
    ) -> None:
        inline_source = _image_document()
        inline_document = Document.from_docx(inline_source)
        inline_image = inline_document.to_spec()["content"][0]
        inline_result = inline_document.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": inline_image["id"],
                    "set": {"behind_text": True},
                }
            ]
        )
        self.assertFalse(inline_result.success)
        self.assertEqual(
            inline_result.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )
        self.assertEqual(inline_document.to_bytes("docx"), inline_source)

        source = _image_document(anchored=True)
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        invalid_operations = (
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {},
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {"unknown": True},
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {"behind_text": None},
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {"behind_text": "true"},
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {"relative_height": True},
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "horizontal": {
                        "offset": {"value": 12, "unit": "pt"},
                    }
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "horizontal": {
                        "relative_to": "page",
                    }
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "horizontal": {
                        "relative_to": "page",
                        "alignment": None,
                    }
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "horizontal": {
                        "relative_to": "page",
                        "offset": {"value": 12, "unit": "pt"},
                        "alignment": "center",
                    }
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "horizontal": {
                        "relative_to": "page",
                        "alignment": "center",
                        "percentage_offset": 50,
                    }
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "horizontal": {
                        "relative_to": "page",
                        "percentage_offset": True,
                    }
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "horizontal": {
                        "relative_to": "page",
                        "percentage_offset": 2147483.648,
                    }
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "relative_size": {},
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "relative_size": {
                        "width": {
                            "relative_to": "margin",
                            "percentage": True,
                        }
                    },
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "relative_size": {
                        "height": {
                            "relative_to": "page",
                            "percentage": "50",
                        }
                    },
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "relative_size": {
                        "width": {
                            "relative_to": "top_margin",
                            "percentage": 50,
                        }
                    },
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "relative_size": {
                        "height": {
                            "relative_to": "page",
                            "percentage": 2147483.648,
                        }
                    },
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "vertical": {
                        "relative_to": "page",
                        "alignment": "middle",
                    }
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "wrap": {
                        "mode": "square",
                        "side": "left",
                        "distances": {
                            "top": {
                                "value": -1,
                                "unit": "pt",
                            },
                        },
                    }
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "anchor_effect_extent": {
                        "left": {
                            "value": 3_000_000_000,
                            "unit": "pt",
                        },
                        "top": {"value": 0, "unit": "pt"},
                        "right": {"value": 0, "unit": "pt"},
                        "bottom": {"value": 0, "unit": "pt"},
                    }
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "wrap": {
                        "mode": "square",
                    }
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "wrap": {
                        "mode": "none",
                        "side": "both_sides",
                    }
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "wrap": {
                        "mode": "top_and_bottom",
                        "distances": {
                            "right": {
                                "value": 0,
                                "unit": "pt",
                            },
                        },
                    }
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "wrap": {
                        "mode": "none",
                        "effect_extent": {
                            "left": {"value": 0, "unit": "pt"},
                            "top": {"value": 0, "unit": "pt"},
                            "right": {"value": 0, "unit": "pt"},
                            "bottom": {"value": 0, "unit": "pt"},
                        },
                    }
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {"behind_text": True},
                "clear": ["behind_text"],
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {"anchor_distances": {}},
                "clear": ["anchor_distances"],
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "clear": [
                    "anchor_effect_extent",
                    "anchor_effect_extent",
                ],
            },
        )
        for operation in invalid_operations:
            with self.subTest(operation=operation):
                result = document.apply([operation])
                self.assertFalse(result.success)
                self.assertEqual(
                    result.diagnostics[0].code,
                    "INVALID_SPEC",
                )
                self.assertIsNone(result.document)
                self.assertEqual(document.to_bytes("docx"), source)

        detached = Document.from_spec(document.to_spec())
        detached_result = detached.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": image["id"],
                    "set": {"behind_text": True},
                }
            ]
        )
        self.assertFalse(detached_result.success)
        self.assertEqual(
            detached_result.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )

    def test_floating_anchor_and_picture_updates_compose_in_any_order(
        self,
    ) -> None:
        source = _image_document(anchored=True)
        for anchor_first in (True, False):
            document = Document.from_docx(source)
            image = document.to_spec()["content"][0]
            anchor_operation = {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "horizontal": {
                        "relative_to": "page",
                        "offset": {"value": 48, "unit": "pt"},
                    },
                    "locked": False,
                },
            }
            picture_operation = {
                "op": "image.update",
                "target": image["id"],
                "set": {
                    "crop": {"left": 10, "right": 10},
                    "alt_text": "Composed floating picture",
                },
            }
            operations = (
                [anchor_operation, picture_operation]
                if anchor_first
                else [picture_operation, anchor_operation]
            )
            with self.subTest(anchor_first=anchor_first):
                result = document.apply(operations)
                self.assertTrue(result.success, result.model_dump())
                assert result.document is not None
                updated = result.document.to_spec()["content"][0]
                self.assertEqual(
                    updated["floating"]["horizontal"],
                    {
                        "relative_to": "page",
                        "offset": {"value": 48.0, "unit": "pt"},
                    },
                )
                self.assertFalse(updated["floating"]["locked"])
                self.assertEqual(updated["crop"]["left"], 10.0)
                self.assertEqual(
                    updated["alt_text"],
                    "Composed floating picture",
                )
                reopened = Document.from_docx(
                    result.document.to_bytes("docx")
                )
                self.assertEqual(
                    reopened.to_spec()["content"][0]["floating"],
                    updated["floating"],
                )
                self.assertEqual(
                    reopened.to_spec()["content"][0]["crop"],
                    updated["crop"],
                )

    def test_floating_anchor_update_compares_physical_units_by_emu(
        self,
    ) -> None:
        document = Document.from_docx(_image_document(anchored=True))
        image = document.to_spec()["content"][0]
        result = document.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": image["id"],
                    "set": {
                        "horizontal": {
                            "relative_to": "page",
                            "offset": {"value": 1, "unit": "in"},
                        }
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        self.assertEqual(
            result.document.to_spec()["content"][0]["floating"][
                "horizontal"
            ],
            {
                "relative_to": "page",
                "offset": {"value": 1.0, "unit": "in"},
            },
        )
        self.assertEqual(
            result.document.read_image(image["id"]).data,
            PNG,
        )
        output = result.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as package:
            root = parse_xml(package.read("word/document.xml"))
        offset = root.find(
            f".//{_q(WP, 'positionH')}/{_q(WP, 'posOffset')}"
        )
        assert offset is not None
        self.assertEqual(offset.text, "914400")
        reopened = Document.from_docx(output)
        self.assertEqual(
            reopened.to_spec()["content"][0]["floating"]["horizontal"],
            {
                "relative_to": "page",
                "offset": {"value": 72.0, "unit": "pt"},
            },
        )

    def test_unsupported_floating_anchor_variants_remain_opaque(self) -> None:
        cases = (
            "duplicate_position_mode",
            "invalid_alignment_position",
            "percentage_attribute",
            "percentage_non_integer",
            "percentage_out_of_range",
            "percentage_wrong_axis",
            "relative_size_attribute",
            "relative_size_unknown_frame",
            "relative_size_wrong_child",
            "relative_size_non_integer",
            "relative_size_negative",
            "relative_size_out_of_range",
            "relative_size_duplicate_axis",
            "simple_position",
            "tight_wrap",
            "tight_top_distance",
            "tight_effect_extent",
            "tight_invalid_coordinate",
            "tight_invalid_point_order",
            "wrap_none_attribute",
            "wrap_text_content",
            "malformed_wrap_effect_extent",
            "duplicate_wrap",
            "nondefault_bw_mode",
            "duplicate_no_fill",
            "negative_distance",
            "malformed_extension_id",
            "missing_required_attribute",
        )
        for case in cases:
            source = _image_document(
                anchored=True,
                wrap_mode=(
                    "tight"
                    if case.startswith("tight_")
                    and case != "tight_wrap"
                    else "square"
                ),
                relative_size=case.startswith("relative_size_"),
            )
            with ZipFile(io.BytesIO(source)) as archive:
                root = parse_xml(archive.read("word/document.xml"))
            anchor = root.find(f".//{_q(WP, 'anchor')}")
            assert anchor is not None
            if case == "duplicate_position_mode":
                position = anchor.find(f"./{_q(WP, 'positionH')}")
                assert position is not None
                ET.SubElement(position, _q(WP, "align")).text = "center"
            elif case == "invalid_alignment_position":
                position = anchor.find(f"./{_q(WP, 'positionH')}")
                assert position is not None
                position.clear()
                position.attrib["relativeFrom"] = "margin"
                ET.SubElement(position, _q(WP, "align")).text = "middle"
            elif case.startswith("percentage_"):
                position = anchor.find(f"./{_q(WP, 'positionH')}")
                assert position is not None
                position.clear()
                position.attrib["relativeFrom"] = "margin"
                percentage = ET.SubElement(
                    position,
                    _q(
                        WP14,
                        (
                            "pctPosVOffset"
                            if case == "percentage_wrong_axis"
                            else "pctPosHOffset"
                        ),
                    ),
                )
                percentage.text = (
                    "12.5"
                    if case == "percentage_non_integer"
                    else "2147483648"
                    if case == "percentage_out_of_range"
                    else "12500"
                )
                if case == "percentage_attribute":
                    percentage.set("unexpected", "1")
            elif case.startswith("relative_size_"):
                relative_width = anchor.find(
                    f"./{_q(WP14, 'sizeRelH')}"
                )
                assert relative_width is not None
                percentage = relative_width.find(
                    f"./{_q(WP14, 'pctWidth')}"
                )
                assert percentage is not None
                if case == "relative_size_attribute":
                    relative_width.set("unexpected", "1")
                elif case == "relative_size_unknown_frame":
                    relative_width.set("relativeFrom", "column")
                elif case == "relative_size_wrong_child":
                    percentage.tag = _q(WP14, "pctHeight")
                elif case == "relative_size_non_integer":
                    percentage.text = "12.5"
                elif case == "relative_size_negative":
                    percentage.text = "-1"
                elif case == "relative_size_out_of_range":
                    percentage.text = "2147483648"
                else:
                    duplicate = copy.deepcopy(relative_width)
                    anchor.append(duplicate)
            elif case == "simple_position":
                anchor.attrib["simplePos"] = "1"
            elif case == "tight_wrap":
                wrap = anchor.find(f"./{_q(WP, 'wrapSquare')}")
                assert wrap is not None
                wrap.tag = _q(WP, "wrapTight")
            elif case == "tight_top_distance":
                wrap = anchor.find(f"./{_q(WP, 'wrapTight')}")
                assert wrap is not None
                wrap.set("distT", "12700")
            elif case == "tight_effect_extent":
                wrap = anchor.find(f"./{_q(WP, 'wrapTight')}")
                assert wrap is not None
                ET.SubElement(
                    wrap,
                    _q(WP, "effectExtent"),
                    {"l": "0", "t": "0", "r": "0", "b": "0"},
                )
            elif case == "tight_invalid_coordinate":
                point = anchor.find(
                    f"./{_q(WP, 'wrapTight')}/"
                    f"{_q(WP, 'wrapPolygon')}/{_q(WP, 'start')}"
                )
                assert point is not None
                point.set("x", "27273042316901")
            elif case == "tight_invalid_point_order":
                point = anchor.find(
                    f"./{_q(WP, 'wrapTight')}/"
                    f"{_q(WP, 'wrapPolygon')}/{_q(WP, 'start')}"
                )
                assert point is not None
                point.tag = _q(WP, "lineTo")
            elif case == "wrap_none_attribute":
                wrap = anchor.find(f"./{_q(WP, 'wrapSquare')}")
                assert wrap is not None
                wrap.tag = _q(WP, "wrapNone")
            elif case == "wrap_text_content":
                wrap = anchor.find(f"./{_q(WP, 'wrapSquare')}")
                assert wrap is not None
                wrap.tag = _q(WP, "wrapNone")
                wrap.attrib.clear()
                wrap.text = "unexpected"
            elif case == "malformed_wrap_effect_extent":
                wrap = anchor.find(f"./{_q(WP, 'wrapSquare')}")
                assert wrap is not None
                wrap.tag = _q(WP, "wrapTopAndBottom")
                wrap.attrib.clear()
                ET.SubElement(
                    wrap,
                    _q(WP, "effectExtent"),
                    {"l": "0", "t": "0", "r": "0"},
                )
            elif case == "duplicate_wrap":
                ET.SubElement(anchor, _q(WP, "wrapNone"))
            elif case == "nondefault_bw_mode":
                shape = anchor.find(f".//{_q(PIC, 'spPr')}")
                assert shape is not None
                shape.set("bwMode", "gray")
            elif case == "duplicate_no_fill":
                shape = anchor.find(f".//{_q(PIC, 'spPr')}")
                assert shape is not None
                ET.SubElement(shape, _q(A, "noFill"))
                ET.SubElement(shape, _q(A, "noFill"))
            elif case == "negative_distance":
                anchor.attrib["distL"] = "-1"
            elif case == "malformed_extension_id":
                anchor.attrib[_q(WP14, "anchorId")] = "not-hex"
            else:
                anchor.attrib.pop("allowOverlap")
            mutated = _rewrite_package(
                source,
                replacements={
                    "word/document.xml": serialize_xml(root),
                },
                additions={},
            )
            with self.subTest(case=case):
                document = Document.from_docx(mutated)
                spec = document.to_spec()
                self.assertEqual(spec["content"][0]["type"], "opaque")
                self.assertEqual(spec["assets"], [])
                self.assertEqual(document.to_bytes("docx"), mutated)

    def test_picture_transform_projects_updates_clears_and_preserves(
        self,
    ) -> None:
        source = _image_document(
            transform_attributes={
                "rot": "-5400000",
                "flipH": "1",
                "flipV": "false",
            },
        )
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        self.assertEqual(
            image["transform"],
            {
                "rotation_degrees_clockwise": 270.0,
                "flip_horizontal": True,
                "flip_vertical": False,
            },
        )
        self.assertEqual(
            document.inspect()["nodes"][0]["transform"],
            image["transform"],
        )
        self.assertEqual(document.to_bytes("docx"), source)
        self.assertEqual(document.image_bytes(image["id"]), PNG)

        capabilities = document.capabilities()["assets"]
        self.assertIn("transform", capabilities["native_update_fields"])
        self.assertIn("transform", capabilities["clearable_update_fields"])
        self.assertEqual(
            capabilities["image_transform_schema"],
            "image-transform",
        )
        self.assertEqual(
            capabilities["image_rotation_native_units_per_degree"],
            60_000,
        )
        self.assertEqual(
            capabilities["image_rotation_range"],
            {
                "minimum_inclusive": 0,
                "maximum_exclusive": 360,
            },
        )
        html = document.to_bytes("html").decode()
        self.assertIn(
            'data-aioffice-rotation-degrees-clockwise="270"',
            html,
        )
        self.assertIn(
            'data-aioffice-flip-horizontal="true"',
            html,
        )
        self.assertIn(
            'data-aioffice-flip-vertical="false"',
            html,
        )

        metadata_update = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {"alt_text": "Rotated expert workflow"},
                }
            ]
        )
        self.assertTrue(metadata_update.success, metadata_update.model_dump())
        assert metadata_update.document is not None
        metadata_output = metadata_update.document.to_bytes("docx")
        with ZipFile(io.BytesIO(metadata_output)) as package:
            root = parse_xml(package.read("word/document.xml"))
        native_transform = root.find(
            f".//{_q(PIC, 'spPr')}/{_q(A, 'xfrm')}"
        )
        assert native_transform is not None
        self.assertEqual(
            native_transform.attrib,
            {
                "rot": "-5400000",
                "flipH": "1",
                "flipV": "false",
            },
        )

        transformed = metadata_update.document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {
                        "transform": {
                            "rotation_degrees_clockwise": 90.123456,
                            "flip_vertical": True,
                        }
                    },
                }
            ]
        )
        self.assertTrue(transformed.success, transformed.model_dump())
        self.assertEqual(
            transformed.changes[0]["property_changes"],
            [
                {
                    "path": "transform",
                    "before": {
                        "rotation_degrees_clockwise": 270.0,
                        "flip_horizontal": True,
                        "flip_vertical": False,
                    },
                    "after": {
                        "rotation_degrees_clockwise": 90.12345,
                        "flip_horizontal": False,
                        "flip_vertical": True,
                    },
                }
            ],
        )
        assert transformed.document is not None
        transformed_output = transformed.document.to_bytes("docx")
        with ZipFile(io.BytesIO(transformed_output)) as package:
            transformed_root = parse_xml(
                package.read("word/document.xml")
            )
        native_transform = transformed_root.find(
            f".//{_q(PIC, 'spPr')}/{_q(A, 'xfrm')}"
        )
        assert native_transform is not None
        self.assertEqual(
            native_transform.attrib,
            {"rot": "5407407", "flipV": "1"},
        )
        self.assertEqual(
            Document.from_docx(transformed_output).to_spec()["content"][0][
                "transform"
            ],
            {
                "rotation_degrees_clockwise": 90.12345,
                "flip_horizontal": False,
                "flip_vertical": True,
            },
        )

        replaced = transformed.document.replace_image(image["id"], JPEG)
        self.assertTrue(replaced.success, replaced.model_dump())
        assert replaced.document is not None
        self.assertEqual(
            replaced.document.to_spec()["content"][0]["transform"],
            {
                "rotation_degrees_clockwise": 90.12345,
                "flip_horizontal": False,
                "flip_vertical": True,
            },
        )

        cleared = replaced.document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "clear": ["transform"],
                }
            ]
        )
        self.assertTrue(cleared.success, cleared.model_dump())
        assert cleared.document is not None
        cleared_output = cleared.document.to_bytes("docx")
        with ZipFile(io.BytesIO(cleared_output)) as package:
            cleared_root = parse_xml(package.read("word/document.xml"))
        native_transform = cleared_root.find(
            f".//{_q(PIC, 'spPr')}/{_q(A, 'xfrm')}"
        )
        assert native_transform is not None
        self.assertEqual(native_transform.attrib, {})
        self.assertEqual(
            [child.tag for child in native_transform],
            [_q(A, "off"), _q(A, "ext")],
        )
        self.assertNotIn(
            "transform",
            Document.from_docx(cleared_output).to_spec()["content"][0],
        )

    def test_floating_picture_transform_survives_anchor_update(
        self,
    ) -> None:
        document = Document.from_docx(
            _image_document(
                anchored=True,
                outlined=True,
                opacity_amount="81250",
                transform_attributes={
                    "rot": "21660000",
                    "flipV": "true",
                },
            )
        )
        image = document.to_spec()["content"][0]
        self.assertEqual(
            image["transform"],
            {
                "rotation_degrees_clockwise": 1.0,
                "flip_horizontal": False,
                "flip_vertical": True,
            },
        )
        result = document.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": image["id"],
                    "set": {"relative_height": 4096},
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        self.assertEqual(
            result.document.to_spec()["content"][0]["transform"],
            image["transform"],
        )
        self.assertEqual(
            result.document.to_spec()["content"][0]["outline"],
            image["outline"],
        )
        self.assertEqual(
            result.document.to_spec()["content"][0]["opacity"],
            image["opacity"],
        )
        with ZipFile(io.BytesIO(result.document.to_bytes("docx"))) as package:
            root = parse_xml(package.read("word/document.xml"))
        native_transform = root.find(
            f".//{_q(PIC, 'spPr')}/{_q(A, 'xfrm')}"
        )
        assert native_transform is not None
        self.assertEqual(
            native_transform.attrib,
            {"rot": "21660000", "flipV": "true"},
        )

    def test_picture_outline_projects_updates_clears_and_preserves(
        self,
    ) -> None:
        source = _image_document(outlined=True)
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        expected_outline = {
            "width": {"value": 2.0, "unit": "pt"},
            "color": "#CC0000",
            "dash": "dash_dot",
        }
        self.assertEqual(image["outline"], expected_outline)
        self.assertEqual(
            document.inspect()["nodes"][0]["outline"],
            expected_outline,
        )
        self.assertEqual(document.to_bytes("docx"), source)
        html = document.to_bytes("html").decode()
        self.assertIn('data-aioffice-outline-width-pt="2"', html)
        self.assertIn('data-aioffice-outline-color="#CC0000"', html)
        self.assertIn('data-aioffice-outline-dash="dash_dot"', html)

        capabilities = document.capabilities()["assets"]
        self.assertIn("outline", capabilities["native_update_fields"])
        self.assertIn("outline", capabilities["clearable_update_fields"])
        self.assertEqual(
            capabilities["image_outline_native_width_bounds"],
            {
                "minimum_inclusive": 1,
                "maximum_inclusive": 20_116_800,
            },
        )

        metadata_update = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {"alt_text": "Outlined expert workflow"},
                }
            ]
        )
        self.assertTrue(metadata_update.success, metadata_update.model_dump())
        assert metadata_update.document is not None
        metadata_output = metadata_update.document.to_bytes("docx")
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(metadata_output)) as after,
        ):
            before_root = parse_xml(before.read("word/document.xml"))
            after_root = parse_xml(after.read("word/document.xml"))
        before_outline = before_root.find(
            f".//{_q(PIC, 'spPr')}/{_q(A, 'ln')}"
        )
        after_outline = after_root.find(
            f".//{_q(PIC, 'spPr')}/{_q(A, 'ln')}"
        )
        assert before_outline is not None
        assert after_outline is not None
        self.assertEqual(
            serialize_xml(after_outline),
            serialize_xml(before_outline),
        )

        updated = metadata_update.document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {
                        "outline": {
                            "width": {"value": 1.23456, "unit": "pt"},
                            "color": "#00aaff",
                            "dash": "large_dash_dot_dot",
                        }
                    },
                }
            ]
        )
        self.assertTrue(updated.success, updated.model_dump())
        self.assertEqual(
            updated.changes[0]["property_changes"],
            [
                {
                    "path": "outline",
                    "before": expected_outline,
                    "after": {
                        "width": {"value": 1.234567, "unit": "pt"},
                        "color": "#00AAFF",
                        "dash": "large_dash_dot_dot",
                    },
                }
            ],
        )
        assert updated.document is not None
        updated_output = updated.document.to_bytes("docx")
        with ZipFile(io.BytesIO(updated_output)) as package:
            root = parse_xml(package.read("word/document.xml"))
        native_outline = root.find(
            f".//{_q(PIC, 'spPr')}/{_q(A, 'ln')}"
        )
        assert native_outline is not None
        self.assertEqual(
            native_outline.attrib,
            {
                "w": "15679",
                "cap": "flat",
                "cmpd": "sng",
                "algn": "ctr",
            },
        )
        self.assertEqual(
            native_outline.find(f"./{_q(A, 'solidFill')}/{_q(A, 'srgbClr')}").attrib,
            {"val": "00AAFF"},
        )
        self.assertEqual(
            native_outline.find(f"./{_q(A, 'prstDash')}").attrib,
            {"val": "lgDashDotDot"},
        )
        reopened = Document.from_docx(updated_output)
        self.assertEqual(
            reopened.to_spec()["content"][0]["outline"],
            {
                "width": {"value": 1.234567, "unit": "pt"},
                "color": "#00AAFF",
                "dash": "large_dash_dot_dot",
            },
        )

        replaced = reopened.replace_image(image["id"], JPEG)
        self.assertTrue(replaced.success, replaced.model_dump())
        assert replaced.document is not None
        self.assertEqual(
            replaced.document.to_spec()["content"][0]["outline"],
            reopened.to_spec()["content"][0]["outline"],
        )

        cleared = replaced.document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "clear": ["outline"],
                }
            ]
        )
        self.assertTrue(cleared.success, cleared.model_dump())
        assert cleared.document is not None
        cleared_output = cleared.document.to_bytes("docx")
        with ZipFile(io.BytesIO(cleared_output)) as package:
            root = parse_xml(package.read("word/document.xml"))
        self.assertIsNone(
            root.find(f".//{_q(PIC, 'spPr')}/{_q(A, 'ln')}")
        )
        self.assertNotIn(
            "outline",
            Document.from_docx(cleared_output).to_spec()["content"][0],
        )

    def test_picture_outline_projects_all_preset_dashes(self) -> None:
        mappings = {
            "solid": "solid",
            "dot": "dot",
            "sysDot": "system_dot",
            "dash": "dash",
            "sysDash": "system_dash",
            "lgDash": "large_dash",
            "dashDot": "dash_dot",
            "sysDashDot": "system_dash_dot",
            "lgDashDot": "large_dash_dot",
            "lgDashDotDot": "large_dash_dot_dot",
            "sysDashDotDot": "system_dash_dot_dot",
        }
        for native_dash, semantic_dash in mappings.items():
            source = _image_document(outlined=True)
            with ZipFile(io.BytesIO(source)) as package:
                root = parse_xml(package.read("word/document.xml"))
            preset = root.find(f".//{_q(A, 'prstDash')}")
            assert preset is not None
            preset.set("val", native_dash)
            mutated = _rewrite_package(
                source,
                replacements={
                    "word/document.xml": serialize_xml(root),
                },
                additions={},
            )
            with self.subTest(native_dash=native_dash):
                image = Document.from_docx(mutated).to_spec()["content"][0]
                self.assertEqual(image["outline"]["dash"], semantic_dash)

    def test_picture_opacity_projects_updates_clears_and_preserves(
        self,
    ) -> None:
        source = _image_document(opacity_amount="33333")
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        self.assertEqual(image["opacity"], 33.333)
        self.assertEqual(document.inspect()["nodes"][0]["opacity"], 33.333)
        self.assertEqual(document.to_bytes("docx"), source)
        asset = next(
            candidate
            for candidate in document.to_spec()["assets"]
            if candidate["id"] == image["asset_id"]
        )
        self.assertEqual(
            document.read_image(image["id"]).sha256,
            asset["sha256"],
        )
        capabilities = document.capabilities()["assets"]
        self.assertIn("opacity", capabilities["native_update_fields"])
        self.assertIn("opacity", capabilities["clearable_update_fields"])
        self.assertEqual(
            capabilities["image_opacity_unit"],
            "percentage_points",
        )
        self.assertEqual(capabilities["image_opacity_precision"], 0.001)
        self.assertEqual(
            capabilities["image_opacity_range"],
            {
                "minimum_inclusive": 0,
                "maximum_exclusive": 100,
            },
        )
        html = document.to_bytes("html").decode()
        self.assertIn('data-aioffice-opacity="33.333"', html)

        unrelated = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {"alt_text": "Opacity preserved"},
                }
            ]
        )
        self.assertTrue(unrelated.success, unrelated.model_dump())
        assert unrelated.document is not None
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(
                io.BytesIO(unrelated.document.to_bytes("docx"))
            ) as after,
        ):
            before_root = parse_xml(before.read("word/document.xml"))
            after_root = parse_xml(after.read("word/document.xml"))
        before_opacity = before_root.find(f".//{_q(A, 'alphaModFix')}")
        after_opacity = after_root.find(f".//{_q(A, 'alphaModFix')}")
        assert before_opacity is not None
        assert after_opacity is not None
        self.assertEqual(
            serialize_xml(after_opacity),
            serialize_xml(before_opacity),
        )

        updated = unrelated.document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {"opacity": 62.3456},
                }
            ]
        )
        self.assertTrue(updated.success, updated.model_dump())
        self.assertEqual(
            updated.changes[0]["property_changes"],
            [
                {
                    "path": "opacity",
                    "before": 33.333,
                    "after": 62.346,
                }
            ],
        )
        assert updated.document is not None
        updated_output = updated.document.to_bytes("docx")
        with ZipFile(io.BytesIO(updated_output)) as package:
            root = parse_xml(package.read("word/document.xml"))
            self.assertEqual(package.read("word/media/image1.png"), PNG)
        native_opacity = root.find(f".//{_q(A, 'alphaModFix')}")
        assert native_opacity is not None
        self.assertEqual(native_opacity.attrib, {"amt": "62346"})
        reopened = Document.from_docx(updated_output)
        self.assertEqual(
            reopened.to_spec()["content"][0]["opacity"],
            62.346,
        )

        replaced = reopened.replace_image(
            image["id"],
            JPEG,
            media_type="image/jpeg",
        )
        self.assertTrue(replaced.success, replaced.model_dump())
        assert replaced.document is not None
        self.assertEqual(
            replaced.document.to_spec()["content"][0]["opacity"],
            62.346,
        )

        cleared = replaced.document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "clear": ["opacity"],
                }
            ]
        )
        self.assertTrue(cleared.success, cleared.model_dump())
        assert cleared.document is not None
        cleared_output = cleared.document.to_bytes("docx")
        with ZipFile(io.BytesIO(cleared_output)) as package:
            root = parse_xml(package.read("word/document.xml"))
        self.assertIsNone(root.find(f".//{_q(A, 'alphaModFix')}"))
        self.assertNotIn(
            "opacity",
            Document.from_docx(cleared_output).to_spec()["content"][0],
        )

    def test_identity_opacity_is_preserved_without_semantic_effect(
        self,
    ) -> None:
        source = _image_document(opacity_amount="100000")
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        self.assertNotIn("opacity", image)
        self.assertEqual(document.to_bytes("docx"), source)
        result = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {"title": "Identity opacity preserved"},
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        with ZipFile(
            io.BytesIO(result.document.to_bytes("docx"))
        ) as package:
            root = parse_xml(package.read("word/document.xml"))
        opacity = root.find(f".//{_q(A, 'alphaModFix')}")
        assert opacity is not None
        self.assertEqual(opacity.attrib, {"amt": "100000"})

    def test_picture_outer_shadow_projects_updates_clears_and_preserves(
        self,
    ) -> None:
        source = _image_document(shadowed=True)
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        self.assertEqual(
            image["shadow"],
            {
                "color": "#123456",
                "opacity": 33.333,
                "blur_radius": {"value": 10.0, "unit": "pt"},
                "distance": {"value": 3.0, "unit": "pt"},
                "direction_degrees_clockwise": 45.0,
                "alignment": "center",
                "rotate_with_shape": False,
            },
        )
        self.assertEqual(
            document.inspect()["nodes"][0]["shadow"],
            image["shadow"],
        )
        self.assertEqual(document.to_bytes("docx"), source)
        capabilities = document.capabilities()["assets"]
        self.assertIn("shadow", capabilities["native_update_fields"])
        self.assertIn("shadow", capabilities["clearable_update_fields"])
        self.assertEqual(
            capabilities["image_shadow_schema"],
            "image-shadow",
        )
        self.assertEqual(
            capabilities["image_shadow_native_length_bounds"],
            {
                "minimum_inclusive": 0,
                "maximum_inclusive": 2_147_483_647,
            },
        )
        html = document.to_bytes("html").decode()
        self.assertIn(
            'data-aioffice-shadow-color="#123456"',
            html,
        )
        self.assertIn(
            'data-aioffice-shadow-opacity="33.333"',
            html,
        )
        self.assertIn(
            'data-aioffice-shadow-direction-degrees-clockwise="45"',
            html,
        )

        unrelated = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {"title": "Shadow preserved"},
                }
            ]
        )
        self.assertTrue(unrelated.success, unrelated.model_dump())
        assert unrelated.document is not None
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(
                io.BytesIO(unrelated.document.to_bytes("docx"))
            ) as after,
        ):
            before_root = parse_xml(before.read("word/document.xml"))
            after_root = parse_xml(after.read("word/document.xml"))
        before_effects = before_root.find(f".//{_q(A, 'effectLst')}")
        after_effects = after_root.find(f".//{_q(A, 'effectLst')}")
        assert before_effects is not None
        assert after_effects is not None
        self.assertEqual(
            serialize_xml(after_effects),
            serialize_xml(before_effects),
        )

        updated = unrelated.document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {
                        "shadow": {
                            "color": "#abcdef",
                            "opacity": 62.3456,
                            "blur_radius": {
                                "value": 5.12345,
                                "unit": "pt",
                            },
                            "distance": {
                                "value": 2.5,
                                "unit": "pt",
                            },
                            "direction_degrees_clockwise": 315.123456,
                            "alignment": "bottom_right",
                            "rotate_with_shape": True,
                        }
                    },
                }
            ]
        )
        self.assertTrue(updated.success, updated.model_dump())
        assert updated.document is not None
        updated_shadow = updated.document.to_spec()["content"][0]["shadow"]
        self.assertEqual(updated_shadow["color"], "#ABCDEF")
        self.assertEqual(updated_shadow["opacity"], 62.346)
        self.assertEqual(
            updated_shadow["blur_radius"],
            {"value": 5.123465, "unit": "pt"},
        )
        self.assertEqual(
            updated_shadow["direction_degrees_clockwise"],
            315.12345,
        )
        updated_output = updated.document.to_bytes("docx")
        with ZipFile(io.BytesIO(updated_output)) as package:
            root = parse_xml(package.read("word/document.xml"))
            self.assertEqual(package.read("word/media/image1.png"), PNG)
        outer_shadow = root.find(f".//{_q(A, 'outerShdw')}")
        assert outer_shadow is not None
        self.assertEqual(
            outer_shadow.attrib,
            {
                "blurRad": "65068",
                "dist": "31750",
                "dir": "18907407",
                "algn": "br",
                "rotWithShape": "1",
            },
        )
        color = outer_shadow.find(f"./{_q(A, 'srgbClr')}")
        assert color is not None
        self.assertEqual(color.attrib, {"val": "ABCDEF"})
        alpha = color.find(f"./{_q(A, 'alpha')}")
        assert alpha is not None
        self.assertEqual(alpha.attrib, {"val": "62346"})
        reopened = Document.from_docx(updated_output)
        self.assertEqual(
            reopened.to_spec()["content"][0]["shadow"],
            updated_shadow,
        )

        replaced = reopened.replace_image(
            image["id"],
            JPEG,
            media_type="image/jpeg",
        )
        self.assertTrue(replaced.success, replaced.model_dump())
        assert replaced.document is not None
        self.assertEqual(
            replaced.document.to_spec()["content"][0]["shadow"],
            updated_shadow,
        )

        cleared = replaced.document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "clear": ["shadow"],
                }
            ]
        )
        self.assertTrue(cleared.success, cleared.model_dump())
        assert cleared.document is not None
        cleared_output = cleared.document.to_bytes("docx")
        with ZipFile(io.BytesIO(cleared_output)) as package:
            root = parse_xml(package.read("word/document.xml"))
        self.assertIsNone(root.find(f".//{_q(A, 'effectLst')}"))
        self.assertNotIn(
            "shadow",
            Document.from_docx(cleared_output).to_spec()["content"][0],
        )

    def test_empty_picture_effect_list_is_neutral_and_preserved(self) -> None:
        source = _image_document()
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
        shape = root.find(f".//{_q(PIC, 'spPr')}")
        assert shape is not None
        ET.SubElement(shape, _q(A, "effectLst"))
        source = _rewrite_package(
            source,
            replacements={
                "word/document.xml": serialize_xml(root),
            },
            additions={},
        )
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        self.assertNotIn("shadow", image)
        self.assertEqual(document.to_bytes("docx"), source)
        result = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {"alt_text": "Neutral effects preserved"},
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        with ZipFile(
            io.BytesIO(result.document.to_bytes("docx"))
        ) as package:
            root = parse_xml(package.read("word/document.xml"))
        effect_list = root.find(f".//{_q(A, 'effectLst')}")
        assert effect_list is not None
        self.assertEqual(effect_list.attrib, {})
        self.assertEqual(len(effect_list), 0)

    def test_picture_outer_shadow_projects_all_alignments_and_full_opacity(
        self,
    ) -> None:
        mappings = {
            "tl": "top_left",
            "t": "top",
            "tr": "top_right",
            "l": "left",
            "ctr": "center",
            "r": "right",
            "bl": "bottom_left",
            "b": "bottom",
            "br": "bottom_right",
        }
        for native_alignment, semantic_alignment in mappings.items():
            source = _image_document(shadowed=True)
            with ZipFile(io.BytesIO(source)) as package:
                root = parse_xml(package.read("word/document.xml"))
            outer_shadow = root.find(f".//{_q(A, 'outerShdw')}")
            assert outer_shadow is not None
            outer_shadow.set("algn", native_alignment)
            mutated = _rewrite_package(
                source,
                replacements={
                    "word/document.xml": serialize_xml(root),
                },
                additions={},
            )
            with self.subTest(alignment=native_alignment):
                image = Document.from_docx(mutated).to_spec()["content"][0]
                self.assertEqual(
                    image["shadow"]["alignment"],
                    semantic_alignment,
                )

        source = _image_document(shadowed=True)
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
        color = root.find(f".//{_q(A, 'outerShdw')}/{_q(A, 'srgbClr')}")
        assert color is not None
        alpha = color.find(f"./{_q(A, 'alpha')}")
        assert alpha is not None
        color.remove(alpha)
        without_alpha = _rewrite_package(
            source,
            replacements={
                "word/document.xml": serialize_xml(root),
            },
            additions={},
        )
        image = Document.from_docx(without_alpha).to_spec()["content"][0]
        self.assertEqual(image["shadow"]["opacity"], 100.0)

    def test_picture_outline_update_keeps_shadow_after_line(self) -> None:
        document = Document.from_docx(_image_document(shadowed=True))
        image = document.to_spec()["content"][0]
        result = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {
                        "outline": {
                            "width": {"value": 1, "unit": "pt"},
                            "color": "#AABBCC",
                        }
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        with ZipFile(
            io.BytesIO(result.document.to_bytes("docx"))
        ) as package:
            root = parse_xml(package.read("word/document.xml"))
        shape = root.find(f".//{_q(PIC, 'spPr')}")
        assert shape is not None
        self.assertEqual(
            [child.tag for child in shape],
            [
                _q(A, "xfrm"),
                _q(A, "prstGeom"),
                _q(A, "ln"),
                _q(A, "effectLst"),
            ],
        )
        reopened = Document.from_docx(result.document.to_bytes("docx"))
        reopened_image = reopened.to_spec()["content"][0]
        self.assertEqual(reopened_image["shadow"], image["shadow"])
        self.assertEqual(reopened_image["outline"]["color"], "#AABBCC")

    def test_inline_shadow_effect_extent_projects_updates_and_clears(
        self,
    ) -> None:
        source = _image_document(shadowed=True)
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
        effect_extent = root.find(f".//{_q(WP, 'effectExtent')}")
        assert effect_extent is not None
        effect_extent.attrib.update(
            {
                "l": "114935",
                "t": "0",
                "r": "114935",
                "b": "12700",
            }
        )
        source = _rewrite_package(
            source,
            replacements={
                "word/document.xml": serialize_xml(root),
            },
            additions={},
        )
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        self.assertEqual(
            image["shadow"]["effect_extent"],
            {
                "left": {"value": 9.05, "unit": "pt"},
                "top": {"value": 0.0, "unit": "pt"},
                "right": {"value": 9.05, "unit": "pt"},
                "bottom": {"value": 1.0, "unit": "pt"},
            },
        )
        self.assertIn(
            'data-aioffice-shadow-effect-extent-left-pt="9.05"',
            document.to_bytes("html").decode(),
        )

        updated_shadow = copy.deepcopy(image["shadow"])
        updated_shadow["effect_extent"] = {
            "left": {"value": 8, "unit": "pt"},
            "top": {"value": 2, "unit": "pt"},
            "right": {"value": 10, "unit": "pt"},
            "bottom": {"value": 4, "unit": "pt"},
        }
        result = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {"shadow": updated_shadow},
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        with ZipFile(
            io.BytesIO(result.document.to_bytes("docx"))
        ) as package:
            root = parse_xml(package.read("word/document.xml"))
        effect_extent = root.find(f".//{_q(WP, 'effectExtent')}")
        assert effect_extent is not None
        self.assertEqual(
            effect_extent.attrib,
            {
                "l": "101600",
                "t": "25400",
                "r": "127000",
                "b": "50800",
            },
        )

        cleared = result.document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "clear": ["shadow"],
                }
            ]
        )
        self.assertTrue(cleared.success, cleared.model_dump())
        assert cleared.document is not None
        with ZipFile(
            io.BytesIO(cleared.document.to_bytes("docx"))
        ) as package:
            root = parse_xml(package.read("word/document.xml"))
        effect_extent = root.find(f".//{_q(WP, 'effectExtent')}")
        assert effect_extent is not None
        self.assertEqual(
            effect_extent.attrib,
            {"l": "0", "t": "0", "r": "0", "b": "0"},
        )
        self.assertIsNone(root.find(f".//{_q(A, 'outerShdw')}"))

        unsupported = _image_document()
        with ZipFile(io.BytesIO(unsupported)) as package:
            root = parse_xml(package.read("word/document.xml"))
        effect_extent = root.find(f".//{_q(WP, 'effectExtent')}")
        assert effect_extent is not None
        effect_extent.set("r", "12700")
        unsupported = _rewrite_package(
            unsupported,
            replacements={
                "word/document.xml": serialize_xml(root),
            },
            additions={},
        )
        self.assertEqual(
            Document.from_docx(unsupported).to_spec()["content"][0]["type"],
            "opaque",
        )

    def test_invalid_native_picture_outer_shadow_remains_opaque(
        self,
    ) -> None:
        for case in (
            "effect_attribute",
            "duplicate_effect_list",
            "multiple_effects",
            "missing_attribute",
            "unknown_attribute",
            "non_integer",
            "negative_blur",
            "over_blur",
            "zero_geometry",
            "bad_direction",
            "bad_alignment",
            "bad_boolean",
            "wrong_color_model",
            "bad_hex",
            "duplicate_alpha",
            "zero_alpha",
            "over_alpha",
        ):
            source = _image_document(shadowed=True)
            with ZipFile(io.BytesIO(source)) as package:
                root = parse_xml(package.read("word/document.xml"))
            shape = root.find(f".//{_q(PIC, 'spPr')}")
            effect_list = root.find(f".//{_q(A, 'effectLst')}")
            outer_shadow = root.find(f".//{_q(A, 'outerShdw')}")
            color = outer_shadow.find(f"./{_q(A, 'srgbClr')}")
            alpha = root.find(f".//{_q(A, 'outerShdw')}//{_q(A, 'alpha')}")
            assert shape is not None
            assert effect_list is not None
            assert outer_shadow is not None
            assert color is not None
            assert alpha is not None
            if case == "effect_attribute":
                effect_list.set("future", "1")
            elif case == "duplicate_effect_list":
                shape.append(copy.deepcopy(effect_list))
            elif case == "multiple_effects":
                ET.SubElement(effect_list, _q(A, "glow"), {"rad": "1"})
            elif case == "missing_attribute":
                del outer_shadow.attrib["dist"]
            elif case == "unknown_attribute":
                outer_shadow.set("future", "1")
            elif case == "non_integer":
                outer_shadow.set("blurRad", "1pt")
            elif case == "negative_blur":
                outer_shadow.set("blurRad", "-1")
            elif case == "over_blur":
                outer_shadow.set("blurRad", "2147483648")
            elif case == "zero_geometry":
                outer_shadow.set("blurRad", "0")
                outer_shadow.set("dist", "0")
            elif case == "bad_direction":
                outer_shadow.set("dir", "21600000")
            elif case == "bad_alignment":
                outer_shadow.set("algn", "middle")
            elif case == "bad_boolean":
                outer_shadow.set("rotWithShape", "yes")
            elif case == "wrong_color_model":
                color.tag = _q(A, "schemeClr")
            elif case == "bad_hex":
                color.set("val", "XYZ123")
            elif case == "duplicate_alpha":
                color.append(copy.deepcopy(alpha))
            elif case == "zero_alpha":
                alpha.set("val", "0")
            elif case == "over_alpha":
                alpha.set("val", "100001")
            mutated = _rewrite_package(
                source,
                replacements={
                    "word/document.xml": serialize_xml(root),
                },
                additions={},
            )
            with self.subTest(case=case):
                opaque = Document.from_docx(mutated)
                self.assertEqual(
                    opaque.to_spec()["content"][0]["type"],
                    "opaque",
                )
                self.assertEqual(opaque.to_spec()["assets"], [])
                self.assertEqual(opaque.to_bytes("docx"), mutated)

    def test_opacity_update_preserves_blip_extension_metadata(
        self,
    ) -> None:
        source = _image_document()
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
        blip = root.find(f".//{_q(A, 'blip')}")
        assert blip is not None
        blip.set("cstate", "print")
        extension_list = ET.SubElement(blip, _q(A, "extLst"))
        extension = ET.SubElement(
            extension_list,
            _q(A, "ext"),
            {"uri": "{28A0092B-C50C-407E-A947-70E740481C1C}"},
        )
        ET.SubElement(
            extension,
            (
                "{http://schemas.microsoft.com/office/drawing/"
                "2010/main}useLocalDpi"
            ),
            {"val": "0"},
        )
        source = _rewrite_package(
            source,
            replacements={
                "word/document.xml": serialize_xml(root),
            },
            additions={},
        )
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        self.assertEqual(image["type"], "image")
        self.assertEqual(document.to_bytes("docx"), source)
        before_extension = serialize_xml(extension_list)

        result = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {"opacity": 50},
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        with ZipFile(
            io.BytesIO(result.document.to_bytes("docx"))
        ) as package:
            updated_root = parse_xml(package.read("word/document.xml"))
        updated_blip = updated_root.find(f".//{_q(A, 'blip')}")
        assert updated_blip is not None
        self.assertEqual(updated_blip.get("cstate"), "print")
        self.assertEqual(
            [child.tag for child in updated_blip],
            [_q(A, "alphaModFix"), _q(A, "extLst")],
        )
        updated_extension = updated_blip.find(f"./{_q(A, 'extLst')}")
        assert updated_extension is not None
        self.assertEqual(
            serialize_xml(updated_extension),
            before_extension,
        )

    def test_invalid_native_picture_opacity_remains_opaque(self) -> None:
        for case in (
            "missing_amount",
            "unknown_attribute",
            "child",
            "non_integer",
            "negative",
            "over_100_percent",
            "duplicate",
            "other_effect",
            "nested_opacity",
            "wrong_child_order",
            "unknown_blip_attribute",
        ):
            source = _image_document(opacity_amount="50000")
            with ZipFile(io.BytesIO(source)) as package:
                root = parse_xml(package.read("word/document.xml"))
            blip = root.find(f".//{_q(A, 'blip')}")
            opacity = root.find(f".//{_q(A, 'alphaModFix')}")
            assert blip is not None
            assert opacity is not None
            if case == "missing_amount":
                opacity.attrib.clear()
            elif case == "unknown_attribute":
                opacity.set("future", "1")
            elif case == "child":
                ET.SubElement(opacity, _q(A, "extLst"))
            elif case == "non_integer":
                opacity.set("amt", "50%")
            elif case == "negative":
                opacity.set("amt", "-1")
            elif case == "over_100_percent":
                opacity.set("amt", "100001")
            elif case == "duplicate":
                blip.append(copy.deepcopy(opacity))
            elif case == "other_effect":
                ET.SubElement(blip, _q(A, "grayscl"))
            elif case == "nested_opacity":
                blip.remove(opacity)
                extension = ET.SubElement(blip, _q(A, "extLst"))
                extension.append(opacity)
            elif case == "wrong_child_order":
                blip.insert(0, ET.Element(_q(A, "extLst")))
            elif case == "unknown_blip_attribute":
                blip.set("future", "1")
            mutated = _rewrite_package(
                source,
                replacements={
                    "word/document.xml": serialize_xml(root),
                },
                additions={},
            )
            with self.subTest(case=case):
                document = Document.from_docx(mutated)
                self.assertEqual(
                    document.to_spec()["content"][0]["type"],
                    "opaque",
                )
                self.assertEqual(document.to_spec()["assets"], [])
                self.assertEqual(document.to_bytes("docx"), mutated)

    def test_rectangular_source_crop_projects_updates_and_clears(
        self,
    ) -> None:
        source = _image_document(cropped=True)
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        self.assertEqual(
            image["crop"],
            {
                "left": 1.0,
                "top": 2.0,
                "right": 3.0,
                "bottom": 4.0,
            },
        )
        self.assertEqual(document.to_bytes("docx"), source)
        capabilities = document.capabilities()["assets"]
        self.assertIn("crop", capabilities["native_update_fields"])
        self.assertIn("crop", capabilities["clearable_update_fields"])
        self.assertEqual(capabilities["crop_unit"], "percentage_points")
        self.assertEqual(capabilities["crop_precision"], 0.001)
        html = document.to_bytes("html").decode()
        self.assertIn('data-aioffice-crop-left="1"', html)
        self.assertIn('data-aioffice-crop-bottom="4"', html)

        updated = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {
                        "crop": {
                            "left": 12.3454,
                            "top": 6.25,
                        }
                    },
                }
            ]
        )
        self.assertTrue(updated.success, updated.model_dump())
        self.assertEqual(
            updated.changes[0]["property_changes"],
            [
                {
                    "path": "crop",
                    "before": {
                        "left": 1.0,
                        "top": 2.0,
                        "right": 3.0,
                        "bottom": 4.0,
                    },
                    "after": {
                        "left": 12.345,
                        "top": 6.25,
                        "right": 0.0,
                        "bottom": 0.0,
                    },
                }
            ],
        )
        assert updated.document is not None
        updated_output = updated.document.to_bytes("docx")
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(updated_output)) as after,
        ):
            self.assertEqual(
                before.read("word/media/image1.png"),
                after.read("word/media/image1.png"),
            )
            self.assertEqual(
                before.read("word/_rels/document.xml.rels"),
                after.read("word/_rels/document.xml.rels"),
            )
            updated_root = parse_xml(after.read("word/document.xml"))
        source_rectangle = updated_root.find(f".//{_q(A, 'srcRect')}")
        assert source_rectangle is not None
        self.assertEqual(
            source_rectangle.attrib,
            {"l": "12345", "t": "6250"},
        )
        reopened_updated = Document.from_docx(updated_output)
        self.assertEqual(
            reopened_updated.to_spec()["content"][0]["crop"],
            {
                "left": 12.345,
                "top": 6.25,
                "right": 0.0,
                "bottom": 0.0,
            },
        )
        self.assertEqual(
            reopened_updated.to_bytes("docx"),
            updated_output,
        )

        cleared = reopened_updated.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "clear": ["crop"],
                }
            ]
        )
        self.assertTrue(cleared.success, cleared.model_dump())
        assert cleared.document is not None
        cleared_output = cleared.document.to_bytes("docx")
        with ZipFile(io.BytesIO(cleared_output)) as package:
            cleared_root = parse_xml(package.read("word/document.xml"))
            self.assertEqual(
                package.read("word/media/image1.png"),
                PNG,
            )
        self.assertIsNone(cleared_root.find(f".//{_q(A, 'srcRect')}"))
        reopened_cleared = Document.from_docx(cleared_output)
        self.assertNotIn(
            "crop",
            reopened_cleared.to_spec()["content"][0],
        )
        self.assertEqual(
            reopened_cleared.to_bytes("docx"),
            cleared_output,
        )

    def test_invalid_native_source_crops_remain_opaque(self) -> None:
        for attributes, duplicate in (
            ({"l": "-1"}, False),
            ({"l": "90000", "r": "10000"}, False),
            ({"l": "1000", "unexpected": "1"}, False),
            ({"l": "1000"}, True),
        ):
            source = _image_document(cropped=True)
            with ZipFile(io.BytesIO(source)) as archive:
                document_root = parse_xml(
                    archive.read("word/document.xml")
                )
            source_rectangle = document_root.find(
                f".//{_q(A, 'srcRect')}"
            )
            assert source_rectangle is not None
            source_rectangle.attrib.clear()
            source_rectangle.attrib.update(attributes)
            if duplicate:
                parent = document_root.find(
                    f".//{_q(PIC, 'blipFill')}"
                )
                assert parent is not None
                parent.insert(
                    list(parent).index(source_rectangle) + 1,
                    copy.deepcopy(source_rectangle),
                )
            source = _rewrite_package(
                source,
                replacements={
                    "word/document.xml": serialize_xml(document_root),
                },
                additions={},
            )
            with self.subTest(
                attributes=attributes,
                duplicate=duplicate,
            ):
                document = Document.from_docx(source)
                self.assertEqual(
                    document.to_spec()["content"][0]["type"],
                    "opaque",
                )
                self.assertEqual(document.to_spec()["assets"], [])
                self.assertEqual(document.to_bytes("docx"), source)

    def test_invalid_native_picture_transforms_remain_opaque(self) -> None:
        for case in (
            "unknown_attribute",
            "non_integer_rotation",
            "rotation_out_of_int32",
            "invalid_flip_boolean",
            "nonzero_offset",
            "extra_extent_attribute",
            "wrong_child_order",
        ):
            source = _image_document(
                transform_attributes={"rot": "5400000", "flipH": "1"}
            )
            with ZipFile(io.BytesIO(source)) as package:
                root = parse_xml(package.read("word/document.xml"))
            transform = root.find(
                f".//{_q(PIC, 'spPr')}/{_q(A, 'xfrm')}"
            )
            assert transform is not None
            offset = transform.find(f"./{_q(A, 'off')}")
            extent = transform.find(f"./{_q(A, 'ext')}")
            assert offset is not None
            assert extent is not None
            if case == "unknown_attribute":
                transform.set("future", "1")
            elif case == "non_integer_rotation":
                transform.set("rot", "90deg")
            elif case == "rotation_out_of_int32":
                transform.set("rot", str(2**31))
            elif case == "invalid_flip_boolean":
                transform.set("flipH", "yes")
            elif case == "nonzero_offset":
                offset.set("x", "1")
            elif case == "extra_extent_attribute":
                extent.set("future", "1")
            elif case == "wrong_child_order":
                transform.remove(offset)
                transform.append(offset)
            mutated = _rewrite_package(
                source,
                replacements={
                    "word/document.xml": serialize_xml(root),
                },
                additions={},
            )
            with self.subTest(case=case):
                document = Document.from_docx(mutated)
                self.assertEqual(
                    document.to_spec()["content"][0]["type"],
                    "opaque",
                )
                self.assertEqual(document.to_spec()["assets"], [])
                self.assertEqual(document.to_bytes("docx"), mutated)

    def test_neutral_zero_width_no_fill_outline_is_preserved(self) -> None:
        source = _image_document()
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
        shape = root.find(f".//{_q(PIC, 'spPr')}")
        assert shape is not None
        outline = ET.SubElement(shape, _q(A, "ln"), {"w": "0"})
        ET.SubElement(outline, _q(A, "noFill"))
        source = _rewrite_package(
            source,
            replacements={
                "word/document.xml": serialize_xml(root),
            },
            additions={},
        )
        document = Document.from_docx(source)
        image = document.to_spec()["content"][0]
        self.assertEqual(image["type"], "image")
        self.assertNotIn("outline", image)
        self.assertEqual(document.to_bytes("docx"), source)
        result = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {"alt_text": "Neutral outline preserved"},
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        with ZipFile(io.BytesIO(result.document.to_bytes("docx"))) as package:
            updated_root = parse_xml(package.read("word/document.xml"))
        updated_outline = updated_root.find(
            f".//{_q(PIC, 'spPr')}/{_q(A, 'ln')}"
        )
        assert updated_outline is not None
        self.assertEqual(updated_outline.attrib, {"w": "0"})
        self.assertEqual(
            [child.tag for child in updated_outline],
            [_q(A, "noFill")],
        )

    def test_invalid_native_picture_outlines_remain_opaque(self) -> None:
        for case in (
            "unknown_attribute",
            "zero_visible_width",
            "nondefault_cap",
            "bad_color",
            "color_transform",
            "theme_color",
            "bad_dash",
            "wrong_child_order",
            "shape_wrong_child_order",
            "bevel_join",
            "duplicate_outline",
            "nonzero_no_fill",
        ):
            source = _image_document(outlined=True)
            with ZipFile(io.BytesIO(source)) as package:
                root = parse_xml(package.read("word/document.xml"))
            shape = root.find(f".//{_q(PIC, 'spPr')}")
            outline = root.find(
                f".//{_q(PIC, 'spPr')}/{_q(A, 'ln')}"
            )
            assert shape is not None
            assert outline is not None
            color = outline.find(
                f"./{_q(A, 'solidFill')}/{_q(A, 'srgbClr')}"
            )
            dash = outline.find(f"./{_q(A, 'prstDash')}")
            join = outline.find(f"./{_q(A, 'round')}")
            assert color is not None
            assert dash is not None
            assert join is not None
            if case == "unknown_attribute":
                outline.set("future", "1")
            elif case == "zero_visible_width":
                outline.set("w", "0")
            elif case == "nondefault_cap":
                outline.set("cap", "rnd")
            elif case == "bad_color":
                color.set("val", "XYZ123")
            elif case == "color_transform":
                ET.SubElement(color, _q(A, "alpha"), {"val": "50000"})
            elif case == "theme_color":
                color.tag = _q(A, "schemeClr")
                color.set("val", "accent1")
            elif case == "bad_dash":
                dash.set("val", "futureDash")
            elif case == "wrong_child_order":
                solid_fill = outline.find(f"./{_q(A, 'solidFill')}")
                assert solid_fill is not None
                outline.remove(solid_fill)
                outline.append(solid_fill)
            elif case == "shape_wrong_child_order":
                shape.remove(outline)
                shape.insert(0, outline)
            elif case == "bevel_join":
                join.tag = _q(A, "bevel")
            elif case == "duplicate_outline":
                shape.append(copy.deepcopy(outline))
            elif case == "nonzero_no_fill":
                for child in list(outline):
                    outline.remove(child)
                outline.attrib.clear()
                outline.set("w", "1")
                ET.SubElement(outline, _q(A, "noFill"))
            mutated = _rewrite_package(
                source,
                replacements={
                    "word/document.xml": serialize_xml(root),
                },
                additions={},
            )
            with self.subTest(case=case):
                document = Document.from_docx(mutated)
                self.assertEqual(
                    document.to_spec()["content"][0]["type"],
                    "opaque",
                )
                self.assertEqual(document.to_spec()["assets"], [])
                self.assertEqual(document.to_bytes("docx"), mutated)

    def test_mismatched_picture_extents_remain_opaque(self) -> None:
        source = _image_document()
        with ZipFile(io.BytesIO(source)) as archive:
            document_root = parse_xml(
                archive.read("word/document.xml")
            )
        inner_extent = document_root.find(
            f".//{_q(PIC, 'spPr')}/{_q(A, 'xfrm')}/{_q(A, 'ext')}"
        )
        assert inner_extent is not None
        inner_extent.set("cx", "914400")
        source = _rewrite_package(
            source,
            replacements={
                "word/document.xml": serialize_xml(document_root),
            },
            additions={},
        )

        document = Document.from_docx(source)
        self.assertEqual(document.to_spec()["content"][0]["type"], "opaque")
        self.assertEqual(document.to_spec()["assets"], [])
        self.assertNotIn("image.update", document.capabilities()["operations"])
        self.assertEqual(document.to_bytes("docx"), source)

    def test_native_patch_reindexes_image_reference_and_preserves_identity(self) -> None:
        source = _image_document(preceding_text="Before")
        document = Document.from_docx(source)
        image_id = next(
            node["id"]
            for node in document.to_spec()["content"]
            if node["type"] == "image"
        )
        result = document.apply(
            [{"op": "node.remove", "target": "#before"}]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        self.assertEqual(result.document.image_bytes(image_id), PNG)

        reopened = Document.from_docx(
            result.document.to_bytes("docx")
        )
        reopened_image = next(
            node
            for node in reopened.to_spec()["content"]
            if node["type"] == "image"
        )
        self.assertEqual(reopened_image["id"], image_id)
        self.assertEqual(reopened.image_bytes(image_id), PNG)

        image_removed = document.apply(
            [{"op": "node.remove", "target": image_id}]
        )
        self.assertTrue(image_removed.success, image_removed.model_dump())
        assert image_removed.document is not None
        self.assertEqual(
            [node["id"] for node in image_removed.document.to_spec()["content"]],
            ["before"],
        )
        removed_output = image_removed.document.to_bytes("docx")
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(removed_output)) as after,
        ):
            self.assertEqual(
                after.read("word/media/image1.png"),
                before.read("word/media/image1.png"),
            )
            self.assertEqual(
                after.read("word/_rels/document.xml.rels"),
                before.read("word/_rels/document.xml.rels"),
            )
        removed_reopened = Document.from_docx(removed_output)
        self.assertEqual(
            [
                node["id"]
                for node in removed_reopened.to_spec()["content"]
            ],
            ["before"],
        )

    def test_image_update_patches_metadata_and_both_native_extents(self) -> None:
        source = _image_document(preceding_text="Before")
        document = Document.from_docx(source)
        image = next(
            node
            for node in document.to_spec()["content"]
            if node["type"] == "image"
        )
        result = document.apply(
            [
                {
                    "op": "image.update",
                    "target": f"#{image['id']}",
                    "set": {
                        "alt_text": "Updated accessible diagram",
                        "title": "Updated workflow",
                        "width": {"value": 1.5, "unit": "in"},
                        "height": {"value": 0.75, "unit": "in"},
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(
            result.changes[0]["resize_mode"],
            "exact",
        )
        self.assertEqual(
            result.fidelity.affected_parts if result.fidelity else None,
            ["/customXml/aioffice-manifest.xml", "/word/document.xml"],
        )
        assert result.document is not None
        output = result.document.to_bytes("docx")
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(output)) as after,
        ):
            for name in before.namelist():
                if name not in {
                    "word/document.xml",
                    "customXml/aioffice-manifest.xml",
                }:
                    self.assertEqual(
                        after.read(name),
                        before.read(name),
                        name,
                    )
            root = parse_xml(after.read("word/document.xml"))
            media_payload = after.read("word/media/image1.png")
        outer_extent = root.find(f".//{_q(WP, 'extent')}")
        inner_extent = root.find(
            f".//{_q(PIC, 'spPr')}/{_q(A, 'xfrm')}/{_q(A, 'ext')}"
        )
        document_properties = root.find(f".//{_q(WP, 'docPr')}")
        assert outer_extent is not None
        assert inner_extent is not None
        assert document_properties is not None
        self.assertEqual(
            (outer_extent.get("cx"), outer_extent.get("cy")),
            ("1371600", "685800"),
        )
        self.assertEqual(
            (inner_extent.get("cx"), inner_extent.get("cy")),
            ("1371600", "685800"),
        )
        self.assertEqual(
            document_properties.get("descr"),
            "Updated accessible diagram",
        )
        self.assertEqual(
            document_properties.get("title"),
            "Updated workflow",
        )
        self.assertEqual(media_payload, PNG)

        reopened = Document.from_docx(output)
        reopened_image = next(
            node
            for node in reopened.to_spec()["content"]
            if node["type"] == "image"
        )
        self.assertEqual(reopened_image["id"], image["id"])
        self.assertEqual(
            reopened_image["width"],
            {"value": 108.0, "unit": "pt"},
        )
        self.assertEqual(
            reopened_image["height"],
            {"value": 54.0, "unit": "pt"},
        )
        self.assertEqual(
            reopened_image["alt_text"],
            "Updated accessible diagram",
        )
        self.assertEqual(reopened.image_bytes(image["id"]), PNG)

    def test_paragraph_format_controls_projected_image_layout(self) -> None:
        source = _image_document(preceding_text="Before")
        document = Document.from_docx(source)
        image = next(
            node
            for node in document.to_spec()["content"]
            if node["type"] == "image"
        )
        operation = {
            "op": "paragraph.format",
            "target": f"#{image['id']}",
            "set": {
                "alignment": "center",
                "background_color": "#EFF4FB",
                "borders": {
                    "top": {
                        "style": "single",
                        "width": {"value": 1, "unit": "pt"},
                        "color": "#1F4E79",
                        "space": {"value": 2, "unit": "pt"},
                    },
                    "bottom": {
                        "style": "single",
                        "width": {"value": 1, "unit": "pt"},
                        "color": "#1F4E79",
                        "space": {"value": 2, "unit": "pt"},
                    },
                },
                "spacing_before": {"value": 12, "unit": "pt"},
                "spacing_after": {"value": 10, "unit": "pt"},
                "indent_left": {"value": 18, "unit": "pt"},
                "indent_right": {"value": 18, "unit": "pt"},
                "keep_together": True,
            },
        }
        result = document.apply([operation])
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(document.to_bytes("docx"), source)
        self.assertEqual(
            result.fidelity.affected_parts if result.fidelity else None,
            ["/customXml/aioffice-manifest.xml", "/word/document.xml"],
        )
        self.assertEqual(
            result.changes[0]["operation"],
            "paragraph.format",
        )
        self.assertEqual(
            {
                change["path"]
                for change in result.changes[0]["property_changes"]
            },
            {
                "paragraph_style.alignment",
                "paragraph_style.background_color",
                "paragraph_style.borders",
                "paragraph_style.indent_left",
                "paragraph_style.indent_right",
                "paragraph_style.keep_together",
                "paragraph_style.spacing_after",
                "paragraph_style.spacing_before",
            },
        )
        assert result.document is not None
        formatted_image = next(
            node
            for node in result.document.to_spec()["content"]
            if node["id"] == image["id"]
        )
        self.assertEqual(
            formatted_image["paragraph_style"]["alignment"],
            "center",
        )
        self.assertEqual(
            formatted_image["paragraph_style"]["background_color"],
            "#EFF4FB",
        )
        self.assertEqual(result.document.image_bytes(image["id"]), PNG)
        self.assertEqual(
            formatted_image["asset_id"],
            image["asset_id"],
        )

        output = result.document.to_bytes("docx")
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(output)) as after,
        ):
            before_root = parse_xml(before.read("word/document.xml"))
            after_root = parse_xml(after.read("word/document.xml"))
            for name in before.namelist():
                if name not in {
                    "word/document.xml",
                    "customXml/aioffice-manifest.xml",
                }:
                    self.assertEqual(after.read(name), before.read(name), name)
        before_drawing = before_root.find(f".//{_q(W, 'drawing')}")
        after_drawing = after_root.find(f".//{_q(W, 'drawing')}")
        assert before_drawing is not None
        assert after_drawing is not None
        self.assertEqual(
            ET.tostring(after_drawing),
            ET.tostring(before_drawing),
        )
        image_paragraph = next(
            paragraph
            for paragraph in after_root.findall(
                f"./{_q(W, 'body')}/{_q(W, 'p')}"
            )
            if paragraph.find(f".//{_q(W, 'drawing')}") is not None
        )
        paragraph_properties = image_paragraph.find(_q(W, "pPr"))
        assert paragraph_properties is not None
        alignment = paragraph_properties.find(_q(W, "jc"))
        spacing = paragraph_properties.find(_q(W, "spacing"))
        indentation = paragraph_properties.find(_q(W, "ind"))
        shading = paragraph_properties.find(_q(W, "shd"))
        borders = paragraph_properties.find(_q(W, "pBdr"))
        assert alignment is not None
        assert spacing is not None
        assert indentation is not None
        assert shading is not None
        assert borders is not None
        self.assertEqual(alignment.get(_q(W, "val")), "center")
        self.assertEqual(
            (spacing.get(_q(W, "before")), spacing.get(_q(W, "after"))),
            ("240", "200"),
        )
        self.assertEqual(
            (indentation.get(_q(W, "left")), indentation.get(_q(W, "right"))),
            ("360", "360"),
        )
        self.assertEqual(shading.get(_q(W, "fill")), "EFF4FB")
        self.assertIsNotNone(paragraph_properties.find(_q(W, "keepLines")))
        self.assertEqual(
            [
                edge.tag
                for edge in borders
            ],
            [_q(W, "top"), _q(W, "bottom")],
        )

        reopened = Document.from_docx(output)
        reopened_image = next(
            node
            for node in reopened.to_spec()["content"]
            if node["id"] == image["id"]
        )
        self.assertEqual(
            reopened_image["paragraph_style"],
            formatted_image["paragraph_style"],
        )
        self.assertEqual(reopened.image_bytes(image["id"]), PNG)

        cleared = reopened.apply(
            [
                {
                    "op": "paragraph.format",
                    "target": image["id"],
                    "clear": list(operation["set"]),
                }
            ]
        )
        self.assertTrue(cleared.success, cleared.model_dump())
        assert cleared.document is not None
        cleared_image = next(
            node
            for node in cleared.document.to_spec()["content"]
            if node["id"] == image["id"]
        )
        self.assertNotIn("paragraph_style", cleared_image)
        self.assertEqual(cleared.document.image_bytes(image["id"]), PNG)

        for invalid_operation in (
            {
                "op": "paragraph.format",
                "target": image["id"],
                "set": {"alignment": "diagonal"},
            },
            {
                "op": "text.format",
                "target": image["id"],
                "set": {"bold": True},
            },
        ):
            with self.subTest(operation=invalid_operation):
                invalid = document.apply([invalid_operation])
                self.assertFalse(invalid.success)
                self.assertEqual(
                    invalid.result_revision,
                    document.revision,
                )
        self.assertEqual(document.to_bytes("docx"), source)
        detached = Document.from_spec(document.to_spec())
        detached_result = detached.apply([operation])
        self.assertFalse(detached_result.success)
        self.assertEqual(
            detached_result.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )

    def test_image_paragraph_format_uses_cli_and_workspace_patch_paths(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            input_path = root / "source.docx"
            patch_path = root / "patch.json"
            output_path = root / "formatted.docx"
            input_path.write_bytes(_image_document())
            document = Document.from_docx(input_path)
            image_id = document.to_spec()["content"][0]["id"]
            operation = {
                "op": "paragraph.format",
                "target": image_id,
                "set": {
                    "alignment": "right",
                    "spacing_before": {"value": 7, "unit": "pt"},
                    "spacing_after": {"value": 9, "unit": "pt"},
                },
            }
            patch_path.write_text(
                json.dumps([operation]),
                encoding="utf-8",
            )
            stdout = io.StringIO()
            with redirect_stdout(stdout):
                exit_code = main(
                    [
                        "apply",
                        str(input_path),
                        str(patch_path),
                        "--output",
                        str(output_path),
                    ]
                )
            self.assertEqual(exit_code, 0)
            report = json.loads(stdout.getvalue())
            self.assertTrue(report["success"])
            cli_image = Document.from_docx(output_path).to_spec()["content"][0]
            self.assertEqual(
                cli_image["paragraph_style"]["alignment"],
                "right",
            )

            workspace = Workspace.init(root / "project")
            tracked = workspace.import_document(input_path)
            workspace_result = workspace.apply(
                tracked.id,
                [operation],
                base_revision=tracked.revision,
            )
            self.assertTrue(
                workspace_result.success,
                workspace_result.model_dump(),
            )
            workspace_image = workspace.open_document(
                tracked.id
            ).to_spec()["content"][0]
            self.assertEqual(
                workspace_image["paragraph_style"]["spacing_before"],
                {"value": 7.0, "unit": "pt"},
            )
            workspace_patch = (
                root
                / "project"
                / ".aioffice"
                / "artifacts"
                / tracked.id
                / "patches"
                / f"{workspace_result.result_revision:08d}.json"
            )
            persisted = json.loads(
                workspace_patch.read_text(encoding="utf-8")
            )
            self.assertEqual(
                persisted["operations"],
                [operation],
            )

    def test_single_dimension_resize_preserves_aspect_ratio(self) -> None:
        document = Document.from_docx(_image_document())
        image = document.to_spec()["content"][0]
        result = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {
                        "width": {"value": 3, "unit": "in"},
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(
            result.changes[0]["resize_mode"],
            "preserve_aspect_ratio",
        )
        assert result.document is not None
        updated = result.document.to_spec()["content"][0]
        self.assertEqual(
            updated["width"],
            {"value": 3.0, "unit": "in"},
        )
        self.assertEqual(
            updated["height"],
            {"value": 108.0, "unit": "pt"},
        )
        reopened = Document.from_docx(
            result.document.to_bytes("docx")
        ).to_spec()["content"][0]
        self.assertEqual(
            reopened["width"],
            {"value": 216.0, "unit": "pt"},
        )
        self.assertEqual(
            reopened["height"],
            {"value": 108.0, "unit": "pt"},
        )

    def test_height_only_resize_preserves_aspect_ratio(self) -> None:
        document = Document.from_docx(_image_document())
        image = document.to_spec()["content"][0]
        result = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {
                        "height": {"value": 0.5, "unit": "in"},
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(
            result.changes[0]["resize_mode"],
            "preserve_aspect_ratio",
        )
        assert result.document is not None
        reopened = Document.from_docx(
            result.document.to_bytes("docx")
        ).to_spec()["content"][0]
        self.assertEqual(
            reopened["width"],
            {"value": 72.0, "unit": "pt"},
        )
        self.assertEqual(
            reopened["height"],
            {"value": 36.0, "unit": "pt"},
        )

    def test_image_update_can_clear_accessibility_metadata(self) -> None:
        document = Document.from_docx(_image_document())
        image = document.to_spec()["content"][0]
        result = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "clear": ["alt_text", "title"],
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        output = result.document.to_bytes("docx")
        root = parse_xml(
            ZipFile(io.BytesIO(output)).read("word/document.xml")
        )
        document_properties = root.find(f".//{_q(WP, 'docPr')}")
        assert document_properties is not None
        self.assertNotIn("descr", document_properties.attrib)
        self.assertNotIn("title", document_properties.attrib)
        reopened = Document.from_docx(output)
        reopened_image = reopened.to_spec()["content"][0]
        self.assertNotIn("alt_text", reopened_image)
        self.assertNotIn("title", reopened_image)
        self.assertTrue(
            any(
                diagnostic.code == "IMAGE_ALT_TEXT_MISSING"
                for diagnostic in reopened.validate().warnings
            )
        )

    def test_image_update_rejects_unsafe_or_detached_requests(self) -> None:
        document = Document.from_docx(
            _image_document(preceding_text="Text")
        )
        image = next(
            node
            for node in document.to_spec()["content"]
            if node["type"] == "image"
        )
        invalid_operations = [
            {
                "op": "image.update",
                "target": image["id"],
                "clear": ["width"],
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {"alt_text": "   "},
            },
            {
                "op": "image.update",
                "target": "#before",
                "set": {"title": "Wrong node type"},
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {"width": {"value": 0, "unit": "pt"}},
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {"crop": {}},
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {
                    "crop": {
                        "left": 60,
                        "right": 40,
                    }
                },
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {"crop": {"top": -0.001}},
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {"transform": {}},
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {
                    "transform": {
                        "rotation_degrees_clockwise": 360,
                    }
                },
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {
                    "transform": {
                        "rotation_degrees_clockwise": "90",
                    }
                },
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {
                    "transform": {
                        "flip_horizontal": "true",
                    }
                },
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {
                    "outline": {
                        "width": {"value": 0, "unit": "pt"},
                        "color": "#112233",
                    }
                },
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {
                    "outline": {
                        "width": {"value": 1, "unit": "pt"},
                        "color": "112233",
                    }
                },
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {
                    "outline": {
                        "width": {"value": 1, "unit": "pt"},
                        "color": "#112233",
                        "dash": "custom",
                    }
                },
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {
                    "outline": {
                        "width": {"value": 1584.001, "unit": "pt"},
                        "color": "#112233",
                    }
                },
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {"opacity": -0.001},
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {"opacity": 100},
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {"opacity": 99.9999},
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {"opacity": True},
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {"opacity": "50"},
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {"shadow": {}},
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {
                    "shadow": {
                        "color": "#112233",
                        "opacity": 0,
                        "blur_radius": {"value": 1, "unit": "pt"},
                        "distance": {"value": 1, "unit": "pt"},
                    }
                },
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {
                    "shadow": {
                        "color": "#112233",
                        "blur_radius": {"value": 0, "unit": "pt"},
                        "distance": {"value": 0, "unit": "pt"},
                    }
                },
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {
                    "shadow": {
                        "color": "#112233",
                        "blur_radius": {"value": 1, "unit": "pt"},
                        "distance": {"value": -1, "unit": "pt"},
                    }
                },
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {
                    "shadow": {
                        "color": "#112233",
                        "blur_radius": {"value": 1, "unit": "pt"},
                        "distance": {"value": 1, "unit": "pt"},
                        "direction_degrees_clockwise": 360,
                    }
                },
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {
                    "shadow": {
                        "color": "#112233",
                        "blur_radius": {"value": 1, "unit": "pt"},
                        "distance": {"value": 1, "unit": "pt"},
                        "alignment": "middle",
                    }
                },
            },
            {
                "op": "image.update",
                "target": image["id"],
                "set": {
                    "shadow": {
                        "color": "#112233",
                        "blur_radius": {"value": 1, "unit": "pt"},
                        "distance": {"value": 1, "unit": "pt"},
                        "rotate_with_shape": "false",
                    }
                },
            },
        ]
        for operation in invalid_operations:
            with self.subTest(operation=operation):
                result = document.apply([operation])
                self.assertFalse(result.success)
                self.assertEqual(result.result_revision, document.revision)

        detached = Document.from_spec(document.to_spec())
        detached_result = detached.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {"title": "Detached"},
                }
            ]
        )
        self.assertFalse(detached_result.success)
        self.assertEqual(
            detached_result.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )

    def test_image_replace_is_occurrence_scoped_copy_on_write(self) -> None:
        source = _image_document(cropped=True)
        with ZipFile(io.BytesIO(source)) as archive:
            document_root = parse_xml(
                archive.read("word/document.xml")
            )
        body = document_root.find(_q(W, "body"))
        assert body is not None
        image_paragraph = body.find(_q(W, "p"))
        assert image_paragraph is not None
        repeated = copy.deepcopy(image_paragraph)
        repeated.attrib[_q(W14, "paraId")] = "ABCDEF12"
        body.insert(len(body) - 1, repeated)
        source = _rewrite_package(
            source,
            replacements={
                "word/document.xml": serialize_xml(document_root),
            },
            additions={},
        )

        document = Document.from_docx(source)
        images = [
            node
            for node in document.to_spec()["content"]
            if node["type"] == "image"
        ]
        self.assertEqual(len(images), 2)
        first_id, second_id = (image["id"] for image in images)
        replacement_sha256 = hashlib.sha256(JPEG).hexdigest()
        replacement_filename = f"aioffice-{replacement_sha256}.jpg"
        replacement_part = f"/word/media/{replacement_filename}"

        result = document.replace_image(
            first_id,
            JPEG,
            media_type="image/jpeg",
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(
            result.changes[0]["replacement_strategy"],
            "occurrence_copy_on_write",
        )
        self.assertEqual(
            result.changes[0]["binary_transport"],
            "out_of_band",
        )
        self.assertEqual(
            result.fidelity.affected_parts if result.fidelity else None,
            [
                "/[Content_Types].xml",
                "/customXml/aioffice-manifest.xml",
                "/word/_rels/document.xml.rels",
                "/word/document.xml",
                replacement_part,
            ],
        )

        assert result.document is not None
        self.assertEqual(result.document.image_bytes(first_id), JPEG)
        self.assertEqual(result.document.image_bytes(second_id), PNG)
        self.assertEqual(document.image_bytes(first_id), PNG)
        self.assertEqual(document.image_bytes(second_id), PNG)
        output = result.document.to_bytes("docx")

        with ZipFile(io.BytesIO(output)) as archive:
            self.assertEqual(archive.read("word/media/image1.png"), PNG)
            self.assertEqual(
                archive.read(replacement_part.lstrip("/")),
                JPEG,
            )
            output_root = parse_xml(archive.read("word/document.xml"))
            relationships = parse_xml(
                archive.read("word/_rels/document.xml.rels")
            )
            content_types = parse_xml(
                archive.read("[Content_Types].xml")
            )
        relationship_ids = [
            blip.get(_q(R, "embed"))
            for blip in output_root.findall(f".//{_q(A, 'blip')}")
        ]
        self.assertEqual(len(relationship_ids), 2)
        self.assertNotEqual(relationship_ids[0], relationship_ids[1])
        replacement_relationship = next(
            relationship
            for relationship in relationships.findall(
                _q(REL, "Relationship")
            )
            if relationship.get("Id") == relationship_ids[0]
        )
        self.assertEqual(
            replacement_relationship.get("Target"),
            f"media/{replacement_filename}",
        )
        replacement_override = next(
            override
            for override in content_types.findall(_q(CT, "Override"))
            if override.get("PartName") == replacement_part
        )
        self.assertEqual(
            replacement_override.get("ContentType"),
            "image/jpeg",
        )

        reopened = Document.from_docx(output)
        reopened_images = [
            node
            for node in reopened.to_spec()["content"]
            if node["type"] == "image"
        ]
        self.assertEqual(
            [image["id"] for image in reopened_images],
            [first_id, second_id],
        )
        self.assertEqual(
            reopened_images[0]["asset_id"],
            f"asset_{replacement_sha256}",
        )
        self.assertEqual(
            reopened_images[0]["width"],
            {"value": 144.0, "unit": "pt"},
        )
        self.assertEqual(
            reopened_images[0]["height"],
            {"value": 72.0, "unit": "pt"},
        )
        self.assertEqual(
            reopened_images[0]["crop"],
            {
                "left": 1.0,
                "top": 2.0,
                "right": 3.0,
                "bottom": 4.0,
            },
        )
        self.assertEqual(
            reopened_images[1]["crop"],
            reopened_images[0]["crop"],
        )
        self.assertEqual(
            reopened_images[0]["alt_text"],
            "A compact expert workflow diagram",
        )
        self.assertEqual(reopened.image_bytes(first_id), JPEG)
        self.assertEqual(reopened.image_bytes(second_id), PNG)
        self.assertEqual(len(reopened.to_spec()["assets"]), 2)

    def test_header_image_projects_reads_updates_and_renders_semantically(
        self,
    ) -> None:
        source = _header_image_document(
            transform_attributes={
                "rot": "10800000",
                "flipH": "1",
            }
        )
        document = Document.from_docx(source)
        spec = document.to_spec()
        header = spec["header_footers"][0]
        image = header["content"][0]
        self.assertEqual(image["type"], "image")
        self.assertEqual(
            image["capabilities"],
            ["inspect", "extract", "render"],
        )
        self.assertEqual(
            image["source_ref"]["part_uri"],
            "/word/header1.xml",
        )
        self.assertEqual(
            image["transform"],
            {
                "rotation_degrees_clockwise": 180.0,
                "flip_horizontal": True,
                "flip_vertical": False,
            },
        )
        self.assertEqual(document.image_bytes(image["id"]), PNG)
        self.assertEqual(document.inspect()["image_count"], 1)
        inspected = document.inspect()["header_footers"][0][
            "blocks"
        ][0]
        self.assertEqual(inspected["type"], "image")
        self.assertIn("image.update", inspected["supported_operations"])
        capabilities = document.capabilities()
        self.assertIn("image.replace", capabilities["operations"])
        self.assertEqual(
            capabilities["assets"]["projected_story_scopes"],
            ["document_body", "header_footer"],
        )
        header_footer_contract = capabilities["formatting"][
            "header_footer_contract"
        ]
        self.assertIn(
            "simple_native_image",
            header_footer_contract["editable_blocks"],
        )
        self.assertEqual(
            header_footer_contract["image_operations"],
            [
                "image.update",
                "image.anchor.update",
                "image.replace",
                "paragraph.format",
            ],
        )
        self.assertEqual(
            header_footer_contract["image_schema_kind"],
            "header-footer-image-block",
        )
        self.assertIn(
            'data-aioffice-asset-id="',
            document.to_bytes("html").decode(),
        )
        detached = Document.from_spec(spec)
        self.assertTrue(detached.validate().valid)
        detached_result = detached.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {"alt_text": "Detached update"},
                }
            ]
        )
        self.assertFalse(detached_result.success)
        self.assertEqual(
            detached_result.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )

        result = document.apply(
            [
                {
                    "op": "image.update",
                    "target": f"#{image['id']}",
                    "set": {
                        "width": {"value": 72, "unit": "pt"},
                        "alt_text": "Updated header logo",
                    },
                },
                {
                    "op": "paragraph.format",
                    "target": f"#{image['id']}",
                    "set": {"alignment": "right"},
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        output = result.document.to_bytes("docx")
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(output)) as after,
        ):
            self.assertEqual(
                before.read("word/document.xml"),
                after.read("word/document.xml"),
            )
            self.assertEqual(
                before.read("word/_rels/header1.xml.rels"),
                after.read("word/_rels/header1.xml.rels"),
            )
            self.assertEqual(
                before.read("word/media/image1.png"),
                after.read("word/media/image1.png"),
            )
        reopened = Document.from_docx(output)
        reopened_image = reopened.to_spec()["header_footers"][0][
            "content"
        ][0]
        self.assertEqual(
            reopened_image["width"],
            {"value": 72.0, "unit": "pt"},
        )
        self.assertEqual(
            reopened_image["height"],
            {"value": 36.0, "unit": "pt"},
        )
        self.assertEqual(
            reopened_image["alt_text"],
            "Updated header logo",
        )
        self.assertEqual(
            reopened_image["paragraph_style"]["alignment"],
            "right",
        )
        self.assertEqual(
            reopened_image["transform"],
            image["transform"],
        )
        self.assertEqual(reopened.to_bytes("docx"), output)

    def test_footer_image_projects_and_updates_in_its_own_part(self) -> None:
        source = _header_image_document(kind="footer")
        document = Document.from_docx(source)
        image = document.to_spec()["header_footers"][0]["content"][0]
        self.assertEqual(image["type"], "image")
        self.assertEqual(
            image["source_ref"]["part_uri"],
            "/word/footer1.xml",
        )
        result = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {"alt_text": "Accessible footer logo"},
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        output = result.document.to_bytes("docx")
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(output)) as after,
        ):
            self.assertEqual(
                before.read("word/document.xml"),
                after.read("word/document.xml"),
            )
            self.assertEqual(
                before.read("word/_rels/footer1.xml.rels"),
                after.read("word/_rels/footer1.xml.rels"),
            )
            self.assertEqual(
                before.read("word/media/image1.png"),
                after.read("word/media/image1.png"),
            )
        reopened = Document.from_docx(output)
        reopened_image = reopened.to_spec()["header_footers"][0][
            "content"
        ][0]
        self.assertEqual(
            reopened_image["alt_text"],
            "Accessible footer logo",
        )
        self.assertEqual(reopened.image_bytes(image["id"]), PNG)

    def test_header_image_crop_is_story_local(self) -> None:
        source = _header_image_document(cropped=True)
        document = Document.from_docx(source)
        image = document.to_spec()["header_footers"][0]["content"][0]
        self.assertEqual(
            image["crop"],
            {
                "left": 1.0,
                "top": 2.0,
                "right": 3.0,
                "bottom": 4.0,
            },
        )
        result = document.apply(
            [
                {
                    "op": "image.update",
                    "target": image["id"],
                    "set": {
                        "crop": {
                            "left": 8,
                            "right": 12,
                            "top": 5,
                            "bottom": 5,
                        }
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        output = result.document.to_bytes("docx")
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(output)) as after,
        ):
            self.assertEqual(
                before.read("word/document.xml"),
                after.read("word/document.xml"),
            )
            self.assertEqual(
                before.read("word/_rels/header1.xml.rels"),
                after.read("word/_rels/header1.xml.rels"),
            )
            self.assertEqual(
                before.read("word/media/image1.png"),
                after.read("word/media/image1.png"),
            )
            header_root = parse_xml(after.read("word/header1.xml"))
        source_rectangle = header_root.find(f".//{_q(A, 'srcRect')}")
        assert source_rectangle is not None
        self.assertEqual(
            source_rectangle.attrib,
            {
                "l": "8000",
                "t": "5000",
                "r": "12000",
                "b": "5000",
            },
        )
        reopened = Document.from_docx(output)
        reopened_image = reopened.to_spec()["header_footers"][0][
            "content"
        ][0]
        self.assertEqual(
            reopened_image["crop"],
            {
                "left": 8.0,
                "top": 5.0,
                "right": 12.0,
                "bottom": 5.0,
            },
        )
        self.assertEqual(reopened.image_bytes(image["id"]), PNG)
        self.assertEqual(reopened.to_bytes("docx"), output)

    def test_floating_header_image_update_is_story_local(self) -> None:
        source = _header_image_document(anchored=True, cropped=True)
        document = Document.from_docx(source)
        image = document.to_spec()["header_footers"][0]["content"][0]
        self.assertEqual(image["placement"], "floating")
        floating = image["floating"]
        self.assertEqual(floating["horizontal"]["relative_to"], "column")
        self.assertEqual(floating["vertical"]["relative_to"], "paragraph")
        self.assertEqual(floating["wrap"]["mode"], "square")

        result = document.apply(
            [
                {
                    "op": "image.update",
                    "target": f"#{image['id']}",
                    "set": {
                        "crop": {
                            "left": 15,
                            "top": 7.5,
                            "right": 15,
                            "bottom": 7.5,
                        }
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.diagnostics)
        assert result.document is not None
        output = result.document.to_bytes("docx")
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(output)) as after,
        ):
            self.assertEqual(
                before.read("word/document.xml"),
                after.read("word/document.xml"),
            )
            self.assertEqual(
                before.read("word/_rels/header1.xml.rels"),
                after.read("word/_rels/header1.xml.rels"),
            )
            self.assertEqual(
                before.read("word/media/image1.png"),
                after.read("word/media/image1.png"),
            )
        reopened = Document.from_docx(output)
        reopened_image = reopened.to_spec()["header_footers"][0][
            "content"
        ][0]
        self.assertEqual(reopened_image["placement"], "floating")
        self.assertEqual(reopened_image["floating"], floating)
        self.assertEqual(
            reopened_image["crop"],
            {
                "left": 15.0,
                "top": 7.5,
                "right": 15.0,
                "bottom": 7.5,
            },
        )

    def test_floating_header_anchor_update_is_story_local(self) -> None:
        source = _header_image_document(
            anchored=True,
            wrap_mode="none",
            cropped=True,
            relative_size=True,
        )
        document = Document.from_docx(source)
        image = document.to_spec()["header_footers"][0]["content"][0]
        inspected = document.inspect()["header_footers"][0]["blocks"][0]
        self.assertIn(
            "image.anchor.update",
            inspected["supported_operations"],
        )

        result = document.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": f"#{image['id']}",
                    "set": {
                        "horizontal": {
                            "relative_to": "margin",
                            "offset": {"value": 24, "unit": "pt"},
                        },
                        "anchor_distances": {
                            "top": {"value": 2, "unit": "pt"},
                            "right": {"value": 4, "unit": "pt"},
                            "bottom": {"value": 2, "unit": "pt"},
                            "left": {"value": 4, "unit": "pt"},
                        },
                        "anchor_effect_extent": {
                            "left": {"value": -0.5, "unit": "pt"},
                            "top": {"value": 1, "unit": "pt"},
                            "right": {"value": 1.5, "unit": "pt"},
                            "bottom": {"value": 2, "unit": "pt"},
                        },
                        "wrap": {
                            "mode": "top_and_bottom",
                            "distances": {
                                "top": {
                                    "value": 1,
                                    "unit": "pt",
                                },
                                "bottom": {
                                    "value": 2,
                                    "unit": "pt",
                                },
                            },
                            "effect_extent": {
                                "left": {
                                    "value": 3,
                                    "unit": "pt",
                                },
                                "top": {
                                    "value": 4,
                                    "unit": "pt",
                                },
                                "right": {
                                    "value": 5,
                                    "unit": "pt",
                                },
                                "bottom": {
                                    "value": 6,
                                    "unit": "pt",
                                },
                            },
                        },
                        "relative_size": {
                            "height": {
                                "relative_to": "top_margin",
                                "percentage": 35.5,
                            }
                        },
                        "behind_text": True,
                        "allow_overlap": False,
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        output = result.document.to_bytes("docx")
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(output)) as after,
        ):
            self.assertEqual(
                before.read("word/document.xml"),
                after.read("word/document.xml"),
            )
            self.assertEqual(
                before.read("word/_rels/header1.xml.rels"),
                after.read("word/_rels/header1.xml.rels"),
            )
            self.assertEqual(
                before.read("word/media/image1.png"),
                after.read("word/media/image1.png"),
            )
            header_root = parse_xml(after.read("word/header1.xml"))
        anchor = header_root.find(f".//{_q(WP, 'anchor')}")
        assert anchor is not None
        horizontal = anchor.find(f"./{_q(WP, 'positionH')}")
        assert horizontal is not None
        offset = horizontal.find(f"./{_q(WP, 'posOffset')}")
        assert offset is not None
        self.assertEqual(horizontal.get("relativeFrom"), "margin")
        self.assertEqual(offset.text, "304800")
        self.assertEqual(anchor.get("behindDoc"), "1")
        self.assertEqual(anchor.get("allowOverlap"), "0")
        self.assertIsNone(anchor.find(f"./{_q(WP, 'wrapNone')}"))
        wrap = anchor.find(f"./{_q(WP, 'wrapTopAndBottom')}")
        assert wrap is not None
        self.assertEqual(
            wrap.attrib,
            {"distT": "12700", "distB": "25400"},
        )
        wrap_effect = wrap.find(f"./{_q(WP, 'effectExtent')}")
        assert wrap_effect is not None
        self.assertEqual(
            wrap_effect.attrib,
            {
                "l": "38100",
                "t": "50800",
                "r": "63500",
                "b": "76200",
            },
        )
        anchor_effect = anchor.find(f"./{_q(WP, 'effectExtent')}")
        assert anchor_effect is not None
        self.assertEqual(
            anchor_effect.attrib,
            {
                "l": "-6350",
                "t": "12700",
                "r": "19050",
                "b": "25400",
            },
        )
        self.assertEqual(
            anchor.get(_q(WP14, "anchorId")),
            "A1B2C3D4",
        )
        self.assertIsNone(anchor.find(f"./{_q(WP14, 'sizeRelH')}"))
        relative_height = anchor.find(f"./{_q(WP14, 'sizeRelV')}")
        assert relative_height is not None
        self.assertEqual(
            relative_height.attrib,
            {"relativeFrom": "topMargin"},
        )
        self.assertEqual(relative_height[0].text, "35500")

        reopened = Document.from_docx(output)
        reopened_image = reopened.to_spec()["header_footers"][0][
            "content"
        ][0]
        self.assertEqual(
            reopened_image["floating"]["horizontal"],
            {
                "relative_to": "margin",
                "offset": {"value": 24.0, "unit": "pt"},
            },
        )
        self.assertTrue(reopened_image["floating"]["behind_text"])
        self.assertFalse(reopened_image["floating"]["allow_overlap"])
        self.assertEqual(
            reopened_image["floating"]["relative_size"],
            {
                "height": {
                    "relative_to": "top_margin",
                    "percentage": 35.5,
                }
            },
        )
        self.assertEqual(
            reopened_image["floating"]["wrap"],
            {
                "mode": "top_and_bottom",
                "distances": {
                    "top": {"value": 1.0, "unit": "pt"},
                    "bottom": {"value": 2.0, "unit": "pt"},
                },
                "effect_extent": {
                    "left": {"value": 3.0, "unit": "pt"},
                    "top": {"value": 4.0, "unit": "pt"},
                    "right": {"value": 5.0, "unit": "pt"},
                    "bottom": {"value": 6.0, "unit": "pt"},
                },
            },
        )
        self.assertNotIn(
            "side",
            reopened_image["floating"]["wrap"],
        )
        self.assertEqual(reopened_image["crop"], image["crop"])
        self.assertEqual(reopened.image_bytes(image["id"]), PNG)
        self.assertEqual(reopened.to_bytes("docx"), output)

    def test_aligned_floating_header_anchor_update_is_story_local(
        self,
    ) -> None:
        source = _header_image_document(
            anchored=True,
            aligned=True,
            cropped=True,
        )
        document = Document.from_docx(source)
        image = document.to_spec()["header_footers"][0]["content"][0]
        self.assertEqual(
            image["floating"]["horizontal"],
            {
                "relative_to": "margin",
                "alignment": "center",
            },
        )
        self.assertEqual(
            image["floating"]["vertical"],
            {
                "relative_to": "page",
                "alignment": "bottom",
            },
        )
        result = document.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": image["id"],
                    "set": {
                        "horizontal": {
                            "relative_to": "inside_margin",
                            "alignment": "inside",
                        },
                        "vertical": {
                            "relative_to": "paragraph",
                            "offset": {"value": 6, "unit": "pt"},
                        },
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        output = result.document.to_bytes("docx")
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(output)) as after,
        ):
            self.assertEqual(
                before.read("word/document.xml"),
                after.read("word/document.xml"),
            )
            self.assertEqual(
                before.read("word/_rels/header1.xml.rels"),
                after.read("word/_rels/header1.xml.rels"),
            )
            self.assertEqual(
                before.read("word/media/image1.png"),
                after.read("word/media/image1.png"),
            )
            header_root = parse_xml(after.read("word/header1.xml"))
        horizontal = header_root.find(f".//{_q(WP, 'positionH')}")
        vertical = header_root.find(f".//{_q(WP, 'positionV')}")
        assert horizontal is not None
        assert vertical is not None
        self.assertEqual(horizontal.get("relativeFrom"), "insideMargin")
        self.assertEqual(horizontal[0].tag, _q(WP, "align"))
        self.assertEqual(horizontal[0].text, "inside")
        self.assertEqual(vertical.get("relativeFrom"), "paragraph")
        self.assertEqual(vertical[0].tag, _q(WP, "posOffset"))
        self.assertEqual(vertical[0].text, "76200")

        reopened = Document.from_docx(output)
        reopened_image = reopened.to_spec()["header_footers"][0][
            "content"
        ][0]
        self.assertEqual(
            reopened_image["floating"]["horizontal"],
            {
                "relative_to": "inside_margin",
                "alignment": "inside",
            },
        )
        self.assertEqual(
            reopened_image["floating"]["vertical"],
            {
                "relative_to": "paragraph",
                "offset": {"value": 6.0, "unit": "pt"},
            },
        )
        self.assertEqual(reopened.image_bytes(image["id"]), PNG)

    def test_percentage_floating_header_clone_updates_story_locally(
        self,
    ) -> None:
        source = _header_image_document(
            anchored=True,
            percentage_position=True,
            relative_size=True,
            wrap_mode="none",
        )
        document = Document.from_docx(source)
        spec = document.to_spec()
        source_part = spec["header_footers"][0]
        source_image = source_part["content"][0]
        self.assertEqual(
            source_image["floating"]["horizontal"],
            {
                "relative_to": "column",
                "percentage_offset": 37.5,
            },
        )
        section_id = spec["sections"][0]["id"]
        cloned = document.apply(
            [
                {
                    "op": "header_footer.clone",
                    "target": f"#{source_part['id']}",
                    "part": {"id": "percentage_clone_header"},
                },
                {
                    "op": "section.header_footer.bind",
                    "target": f"#{section_id}",
                    "set": {
                        "header_default": "percentage_clone_header",
                    },
                },
            ]
        )
        self.assertTrue(cloned.success, cloned.model_dump())
        assert cloned.document is not None
        clone_part = next(
            part
            for part in cloned.document.to_spec()["header_footers"]
            if part["id"] == "percentage_clone_header"
        )
        clone_image = clone_part["content"][0]
        self.assertEqual(
            clone_image["floating"],
            source_image["floating"],
        )

        updated = cloned.document.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": f"#{clone_image['id']}",
                    "set": {
                        "horizontal": {
                            "relative_to": "page",
                            "percentage_offset": 75.25,
                        },
                        "relative_size": {
                            "width": {
                                "relative_to": "page",
                                "percentage": 80,
                            }
                        },
                    },
                }
            ]
        )
        self.assertTrue(updated.success, updated.model_dump())
        assert updated.document is not None
        updated_parts = {
            part["id"]: part
            for part in updated.document.to_spec()["header_footers"]
        }
        self.assertEqual(
            updated_parts[source_part["id"]]["content"][0][
                "floating"
            ]["horizontal"]["percentage_offset"],
            37.5,
        )
        self.assertEqual(
            updated_parts["percentage_clone_header"]["content"][0][
                "floating"
            ]["horizontal"],
            {
                "relative_to": "page",
                "percentage_offset": 75.25,
            },
        )
        self.assertEqual(
            updated_parts[source_part["id"]]["content"][0]["floating"][
                "relative_size"
            ],
            source_image["floating"]["relative_size"],
        )
        self.assertEqual(
            updated_parts["percentage_clone_header"]["content"][0][
                "floating"
            ]["relative_size"],
            {
                "width": {
                    "relative_to": "page",
                    "percentage": 80.0,
                }
            },
        )

        output = updated.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as package:
            source_root = parse_xml(package.read("word/header1.xml"))
            clone_root = parse_xml(package.read("word/header2.xml"))
        source_percentage = source_root.find(
            f".//{_q(WP14, 'pctPosHOffset')}"
        )
        clone_percentage = clone_root.find(
            f".//{_q(WP14, 'pctPosHOffset')}"
        )
        assert source_percentage is not None
        assert clone_percentage is not None
        self.assertEqual(source_percentage.text, "37500")
        self.assertEqual(clone_percentage.text, "75250")
        source_relative_width = source_root.find(
            f".//{_q(WP14, 'pctWidth')}"
        )
        clone_relative_width = clone_root.find(
            f".//{_q(WP14, 'pctWidth')}"
        )
        assert source_relative_width is not None
        assert clone_relative_width is not None
        self.assertEqual(source_relative_width.text, "50000")
        self.assertEqual(clone_relative_width.text, "80000")
        self.assertIsNotNone(
            source_root.find(f".//{_q(WP14, 'pctHeight')}")
        )
        self.assertIsNone(
            clone_root.find(f".//{_q(WP14, 'pctHeight')}")
        )
        self.assertIn(
            "wp14",
            (clone_root.get(_q(MC, "Ignorable")) or "").split(),
        )

    def test_cloned_header_image_can_be_replaced_copy_on_write(
        self,
    ) -> None:
        source = _header_image_document(
            cropped=True,
            anchored=True,
            relative_size=True,
            wrap_mode="tight",
            outlined=True,
            opacity_amount="62500",
            shadowed=True,
            transform_attributes={
                "rot": "1350000",
                "flipH": "1",
                "flipV": "1",
            },
        )
        document = Document.from_docx(source)
        spec = document.to_spec()
        source_part = spec["header_footers"][0]
        source_image = source_part["content"][0]
        section_id = spec["sections"][0]["id"]
        cloned = document.apply(
            [
                {
                    "op": "header_footer.clone",
                    "target": f"#{source_part['id']}",
                    "part": {"id": "alternate_logo_header"},
                },
                {
                    "op": "section.header_footer.bind",
                    "target": f"#{section_id}",
                    "set": {
                        "header_default": "alternate_logo_header",
                    },
                },
            ]
        )
        self.assertTrue(cloned.success, cloned.model_dump())
        assert cloned.document is not None
        cloned_part = next(
            part
            for part in cloned.document.to_spec()["header_footers"]
            if part["id"] == "alternate_logo_header"
        )
        cloned_image = cloned_part["content"][0]
        self.assertEqual(cloned_image["type"], "image")
        self.assertEqual(cloned_image["placement"], "floating")
        self.assertEqual(
            cloned_image["floating"],
            source_image["floating"],
        )
        self.assertNotEqual(cloned_image["id"], source_image["id"])
        self.assertEqual(
            cloned_image["asset_id"],
            source_image["asset_id"],
        )
        self.assertEqual(cloned_image["crop"], source_image["crop"])
        self.assertEqual(
            cloned_image["transform"],
            source_image["transform"],
        )
        self.assertEqual(
            cloned_image["outline"],
            source_image["outline"],
        )
        self.assertEqual(cloned_image["opacity"], source_image["opacity"])
        self.assertEqual(cloned_image["shadow"], source_image["shadow"])

        replaced = cloned.document.replace_image(
            cloned_image["id"],
            JPEG,
            media_type="image/jpeg",
        )
        self.assertTrue(replaced.success, replaced.model_dump())
        assert replaced.document is not None
        output = replaced.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as package:
            source_relationships = parse_xml(
                package.read("word/_rels/header1.xml.rels")
            )
            clone_relationships = parse_xml(
                package.read("word/_rels/header2.xml.rels")
            )
            clone_root = parse_xml(
                package.read("word/header2.xml")
            )
            source_root = parse_xml(
                package.read("word/header1.xml")
            )
            source_target = source_relationships.find(
                _q(REL, "Relationship")
            )
            clone_blip = clone_root.find(
                f".//{_q(A, 'blip')}"
            )
            assert source_target is not None
            assert clone_blip is not None
            source_anchor = source_root.find(f".//{_q(WP, 'anchor')}")
            clone_anchor = clone_root.find(f".//{_q(WP, 'anchor')}")
            assert source_anchor is not None
            assert clone_anchor is not None
            for identity_name in ("anchorId", "editId"):
                attribute = _q(WP14, identity_name)
                self.assertNotEqual(
                    source_anchor.get(attribute),
                    clone_anchor.get(attribute),
                )
                self.assertEqual(len(clone_anchor.get(attribute, "")), 8)
            clone_relationship_id = clone_blip.get(_q(R, "embed"))
            clone_target = next(
                relationship
                for relationship in clone_relationships.findall(
                    _q(REL, "Relationship")
                )
                if relationship.get("Id")
                == clone_relationship_id
            )
            self.assertEqual(
                source_target.get("Target"),
                "media/image1.png",
            )
            self.assertNotEqual(
                clone_target.get("Target"),
                "media/image1.png",
            )
            self.assertEqual(
                package.read("word/media/image1.png"),
                PNG,
            )
        reopened = Document.from_docx(output)
        self.assertEqual(
            reopened.image_bytes(cloned_image["id"]),
            JPEG,
        )
        self.assertEqual(
            reopened.image_bytes(source_image["id"]),
            PNG,
        )
        reopened_parts = {
            part["id"]: part
            for part in reopened.to_spec()["header_footers"]
        }
        reopened_source_image = next(
            block
            for block in reopened_parts[source_part["id"]]["content"]
            if block["id"] == source_image["id"]
        )
        reopened_clone_image = next(
            block
            for block in reopened_parts[cloned_part["id"]]["content"]
            if block["id"] == cloned_image["id"]
        )
        self.assertEqual(
            reopened_clone_image["crop"],
            reopened_source_image["crop"],
        )
        self.assertEqual(
            reopened_clone_image["floating"]["relative_size"],
            reopened_source_image["floating"]["relative_size"],
        )
        self.assertEqual(
            reopened_clone_image["transform"],
            reopened_source_image["transform"],
        )
        self.assertEqual(
            reopened_clone_image["outline"],
            reopened_source_image["outline"],
        )
        self.assertEqual(
            reopened_clone_image["opacity"],
            reopened_source_image["opacity"],
        )
        self.assertEqual(
            reopened_clone_image["shadow"],
            reopened_source_image["shadow"],
        )
        self.assertEqual(reopened.to_bytes("docx"), output)

    def test_workspace_replaces_header_image_without_binary_patch_data(
        self,
    ) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            input_path = root / "source.docx"
            input_path.write_bytes(_header_image_document())
            workspace = Workspace.init(root / "project")
            document = workspace.import_document(input_path)
            image = document.to_spec()["header_footers"][0]["content"][0]

            result = workspace.replace_image(
                document.id,
                image["id"],
                JPEG,
                media_type="image/jpeg",
                base_revision=document.revision,
            )
            self.assertTrue(result.success, result.model_dump())
            reopened_workspace = Workspace.open(root / "project")
            reopened = reopened_workspace.open_document(document.id)
            reopened_image = reopened.to_spec()["header_footers"][0][
                "content"
            ][0]
            self.assertEqual(reopened_image["id"], image["id"])
            self.assertEqual(reopened.image_bytes(image["id"]), JPEG)

            patch_path = (
                root
                / "project"
                / ".aioffice"
                / "artifacts"
                / document.id
                / "patches"
                / f"{result.result_revision:08d}.json"
            )
            patch_text = patch_path.read_text(encoding="utf-8")
            patch = json.loads(patch_text)
            self.assertEqual(
                patch["operations"][0]["op"],
                "image.replace",
            )
            self.assertEqual(
                patch["operations"][0]["target"],
                image["id"],
            )
            self.assertNotIn("base64", patch_text)
            self.assertNotIn("data", patch["operations"][0]["asset"])

    def test_image_replace_rejects_untrusted_binary_inputs(self) -> None:
        source = _image_document()
        document = Document.from_docx(source)
        image_id = document.to_spec()["content"][0]["id"]
        invalid_results = [
            document.replace_image(image_id, b"not an image"),
            document.replace_image(
                image_id,
                PNG,
                media_type="image/jpeg",
            ),
        ]
        bounded = Document.from_docx(
            source,
            security_policy=SecurityPolicy(
                max_file_size_mb=1,
                max_uncompressed_size_mb=1,
            ),
        )
        invalid_results.append(
            bounded.replace_image(
                image_id,
                b"\x89PNG\r\n\x1a\n" + b"x" * (1024 * 1024),
            )
        )
        for result in invalid_results:
            with self.subTest(result=result):
                self.assertFalse(result.success)
                self.assertEqual(result.result_revision, document.revision)
                self.assertEqual(
                    result.diagnostics[0].code,
                    "INVALID_ASSET_INPUT",
                )
        self.assertEqual(document.to_bytes("docx"), source)

        package_limit_result = bounded.replace_image(
            image_id,
            b"\x89PNG\r\n\x1a\n" + b"x" * (1024 * 1024 - 8),
        )
        self.assertFalse(package_limit_result.success)
        self.assertEqual(
            package_limit_result.diagnostics[0].code,
            "SECURITY_POLICY_VIOLATION",
        )

        raw_patch = document.apply(
            [
                {
                    "op": "image.replace",
                    "target": image_id,
                    "asset": document.to_spec()["assets"][0],
                }
            ]
        )
        self.assertFalse(raw_patch.success)
        self.assertEqual(
            raw_patch.diagnostics[0].code,
            "BINARY_ASSET_REQUIRED",
        )
        detached = Document.from_spec(document.to_spec())
        detached_result = detached.replace_image(image_id, JPEG)
        self.assertFalse(detached_result.success)
        self.assertEqual(
            detached_result.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )

    def test_replace_image_cli_uses_a_local_out_of_band_asset(self) -> None:
        document = Document.from_docx(_image_document())
        image_id = document.to_spec()["content"][0]["id"]
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            input_path = root / "source.docx"
            replacement_path = root / "replacement.jpg"
            output_path = root / "replaced.docx"
            input_path.write_bytes(document.to_bytes("docx"))
            replacement_path.write_bytes(JPEG)
            stdout = io.StringIO()
            with redirect_stdout(stdout):
                exit_code = main(
                    [
                        "replace-image",
                        str(input_path),
                        image_id,
                        str(replacement_path),
                        "--media-type",
                        "image/jpeg",
                        "-o",
                        str(output_path),
                    ]
                )
            self.assertEqual(exit_code, 0)
            report = json.loads(stdout.getvalue())
            self.assertTrue(report["success"])
            self.assertEqual(report["output"], str(output_path))
            reopened = Document.from_docx(output_path)
            self.assertEqual(reopened.image_bytes(image_id), JPEG)
            output_sha256 = hashlib.sha256(
                output_path.read_bytes()
            ).hexdigest()
            with redirect_stderr(io.StringIO()):
                repeated_exit = main(
                    [
                        "replace-image",
                        str(input_path),
                        image_id,
                        str(replacement_path),
                        "-o",
                        str(output_path),
                    ]
                )
            self.assertEqual(repeated_exit, 2)
            self.assertEqual(
                hashlib.sha256(output_path.read_bytes()).hexdigest(),
                output_sha256,
            )

    def test_workspace_image_replace_persists_binary_free_patch_log(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            input_path = root / "source.docx"
            input_path.write_bytes(_image_document())
            workspace = Workspace.init(root / "project")
            document = workspace.import_document(input_path)
            image_id = document.to_spec()["content"][0]["id"]

            result = workspace.replace_image(
                document.id,
                image_id,
                JPEG,
                media_type="image/jpeg",
                base_revision=document.revision,
            )
            self.assertTrue(result.success, result.model_dump())
            self.assertEqual(result.result_revision, document.revision + 1)
            reopened = workspace.open_document(document.id)
            self.assertEqual(reopened.revision, result.result_revision)
            self.assertEqual(reopened.image_bytes(image_id), JPEG)
            capabilities = workspace.capabilities(document.id)
            self.assertIn("replace_image", capabilities["operations"])
            self.assertFalse(
                capabilities["binary_operations"]["image.replace"][
                    "recorded_binary"
                ]
            )

            patch_path = (
                root
                / "project"
                / ".aioffice"
                / "artifacts"
                / document.id
                / "patches"
                / f"{result.result_revision:08d}.json"
            )
            patch = json.loads(patch_path.read_text(encoding="utf-8"))
            operation = patch["operations"][0]
            self.assertEqual(operation["op"], "image.replace")
            self.assertEqual(
                patch["changes"][0]["binary_transport"],
                "out_of_band",
            )
            self.assertEqual(
                operation["asset"]["sha256"],
                hashlib.sha256(JPEG).hexdigest(),
            )
            self.assertNotIn("data", operation["asset"])
            self.assertNotIn("base64", patch_path.read_text(encoding="utf-8"))

    def test_insert_image_after_creates_stable_native_inline_picture(self) -> None:
        source = _image_document(preceding_text="Before")
        document = Document.from_docx(source)
        original_image = next(
            node
            for node in document.to_spec()["content"]
            if node["type"] == "image"
        )
        replacement_sha256 = hashlib.sha256(JPEG).hexdigest()
        replacement_part = (
            f"/word/media/aioffice-{replacement_sha256}.jpg"
        )
        result = document.insert_image_after(
            "#before",
            JPEG,
            width={"value": 1.5, "unit": "in"},
            height={"value": 0.75, "unit": "in"},
            alt_text="Inserted expert workflow chart",
            media_type="image/jpeg",
            image_id="inserted_chart",
            name="Expert workflow chart",
            title="Workflow",
            transform={
                "rotation_degrees_clockwise": 45.123456,
                "flip_horizontal": True,
            },
            outline={
                "width": {"value": 1.5, "unit": "pt"},
                "color": "#2457A7",
                "dash": "dash",
            },
            opacity=72.3456,
            shadow={
                "color": "#102030",
                "opacity": 41.2346,
                "blur_radius": {"value": 6, "unit": "pt"},
                "distance": {"value": 4, "unit": "pt"},
                "direction_degrees_clockwise": 135,
                "alignment": "top_left",
                "rotate_with_shape": True,
                "effect_extent": {
                    "left": {"value": 8, "unit": "pt"},
                    "top": {"value": 2, "unit": "pt"},
                    "right": {"value": 10, "unit": "pt"},
                    "bottom": {"value": 4, "unit": "pt"},
                },
            },
            paragraph_style={
                "alignment": "center",
                "spacing_before": {"value": 6, "unit": "pt"},
                "spacing_after": {"value": 8, "unit": "pt"},
            },
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(
            result.changes[0],
            {
                "operation": "image.insert_after",
                "after": "before",
                "created_nodes": ["inserted_chart"],
                "asset_ids": [f"asset_{replacement_sha256}"],
                "binary_transport": "out_of_band",
                "placement": "inline",
            },
        )
        self.assertEqual(
            result.fidelity.affected_parts if result.fidelity else None,
            [
                "/[Content_Types].xml",
                "/customXml/aioffice-manifest.xml",
                "/word/_rels/document.xml.rels",
                "/word/document.xml",
                replacement_part,
            ],
        )
        assert result.document is not None
        content = result.document.to_spec()["content"]
        self.assertEqual(
            [(node["id"], node["type"]) for node in content],
            [
                ("before", "paragraph"),
                ("inserted_chart", "image"),
                (original_image["id"], "image"),
            ],
        )
        inserted = content[1]
        self.assertEqual(
            inserted["width"],
            {"value": 1.5, "unit": "in"},
        )
        self.assertEqual(
            inserted["height"],
            {"value": 0.75, "unit": "in"},
        )
        self.assertEqual(
            inserted["transform"],
            {
                "rotation_degrees_clockwise": 45.12345,
                "flip_horizontal": True,
                "flip_vertical": False,
            },
        )
        self.assertEqual(
            inserted["outline"],
            {
                "width": {"value": 1.5, "unit": "pt"},
                "color": "#2457A7",
                "dash": "dash",
            },
        )
        self.assertEqual(inserted["opacity"], 72.346)
        self.assertEqual(
            inserted["shadow"],
            {
                "color": "#102030",
                "opacity": 41.235,
                "blur_radius": {"value": 6.0, "unit": "pt"},
                "distance": {"value": 4.0, "unit": "pt"},
                "direction_degrees_clockwise": 135.0,
                "alignment": "top_left",
                "rotate_with_shape": True,
                "effect_extent": {
                    "left": {"value": 8.0, "unit": "pt"},
                    "top": {"value": 2.0, "unit": "pt"},
                    "right": {"value": 10.0, "unit": "pt"},
                    "bottom": {"value": 4.0, "unit": "pt"},
                },
            },
        )
        self.assertEqual(
            inserted["paragraph_style"]["alignment"],
            "center",
        )
        self.assertEqual(
            result.document.image_bytes("inserted_chart"),
            JPEG,
        )
        self.assertEqual(
            result.document.image_bytes(original_image["id"]),
            PNG,
        )
        self.assertEqual(document.to_bytes("docx"), source)
        output = result.document.to_bytes("docx")

        with ZipFile(io.BytesIO(output)) as archive:
            root = parse_xml(archive.read("word/document.xml"))
            relationships = parse_xml(
                archive.read("word/_rels/document.xml.rels")
            )
            self.assertEqual(
                archive.read(replacement_part.lstrip("/")),
                JPEG,
            )
        body = root.find(_q(W, "body"))
        assert body is not None
        body_paragraphs = body.findall(_q(W, "p"))
        self.assertEqual(len(body_paragraphs), 3)
        inserted_paragraph = body_paragraphs[1]
        self.assertIsNotNone(inserted_paragraph.get(_q(W14, "paraId")))
        alignment = inserted_paragraph.find(
            f"./{_q(W, 'pPr')}/{_q(W, 'jc')}"
        )
        assert alignment is not None
        self.assertEqual(alignment.get(_q(W, "val")), "center")
        outer_extent = inserted_paragraph.find(
            f".//{_q(WP, 'extent')}"
        )
        inner_extent = inserted_paragraph.find(
            f".//{_q(PIC, 'spPr')}/{_q(A, 'xfrm')}/{_q(A, 'ext')}"
        )
        assert outer_extent is not None
        assert inner_extent is not None
        self.assertEqual(
            (outer_extent.get("cx"), outer_extent.get("cy")),
            ("1371600", "685800"),
        )
        self.assertEqual(
            (inner_extent.get("cx"), inner_extent.get("cy")),
            ("1371600", "685800"),
        )
        inserted_effect_extent = inserted_paragraph.find(
            f".//{_q(WP, 'effectExtent')}"
        )
        assert inserted_effect_extent is not None
        self.assertEqual(
            inserted_effect_extent.attrib,
            {
                "l": "101600",
                "t": "25400",
                "r": "127000",
                "b": "50800",
            },
        )
        inserted_transform = inserted_paragraph.find(
            f".//{_q(PIC, 'spPr')}/{_q(A, 'xfrm')}"
        )
        assert inserted_transform is not None
        self.assertEqual(
            inserted_transform.attrib,
            {"rot": "2707407", "flipH": "1"},
        )
        inserted_outline = inserted_paragraph.find(
            f".//{_q(PIC, 'spPr')}/{_q(A, 'ln')}"
        )
        assert inserted_outline is not None
        self.assertEqual(
            inserted_outline.attrib,
            {
                "w": "19050",
                "cap": "flat",
                "cmpd": "sng",
                "algn": "ctr",
            },
        )
        inserted_opacity = inserted_paragraph.find(
            f".//{_q(A, 'blip')}/{_q(A, 'alphaModFix')}"
        )
        assert inserted_opacity is not None
        self.assertEqual(inserted_opacity.attrib, {"amt": "72346"})
        inserted_shadow = inserted_paragraph.find(
            f".//{_q(PIC, 'spPr')}/{_q(A, 'effectLst')}/"
            f"{_q(A, 'outerShdw')}"
        )
        assert inserted_shadow is not None
        self.assertEqual(
            inserted_shadow.attrib,
            {
                "blurRad": "76200",
                "dist": "50800",
                "dir": "8100000",
                "algn": "tl",
                "rotWithShape": "1",
            },
        )
        inserted_relationship_id = inserted_paragraph.find(
            f".//{_q(A, 'blip')}"
        ).get(_q(R, "embed"))
        inserted_relationship = next(
            relationship
            for relationship in relationships.findall(
                _q(REL, "Relationship")
            )
            if relationship.get("Id") == inserted_relationship_id
        )
        self.assertEqual(
            inserted_relationship.get("Target"),
            replacement_part.removeprefix("/word/"),
        )

        reopened = Document.from_docx(output)
        reopened_content = reopened.to_spec()["content"]
        self.assertEqual(
            [(node["id"], node["type"]) for node in reopened_content],
            [
                ("before", "paragraph"),
                ("inserted_chart", "image"),
                (original_image["id"], "image"),
            ],
        )
        reopened_inserted = reopened_content[1]
        self.assertEqual(
            reopened_inserted["width"],
            {"value": 108.0, "unit": "pt"},
        )
        self.assertEqual(
            reopened_inserted["height"],
            {"value": 54.0, "unit": "pt"},
        )
        self.assertEqual(
            reopened_inserted["alt_text"],
            "Inserted expert workflow chart",
        )
        self.assertEqual(
            reopened_inserted["transform"],
            inserted["transform"],
        )
        self.assertEqual(
            reopened_inserted["outline"],
            inserted["outline"],
        )
        self.assertEqual(
            reopened_inserted["opacity"],
            inserted["opacity"],
        )
        self.assertEqual(
            reopened_inserted["shadow"],
            inserted["shadow"],
        )
        self.assertEqual(reopened.image_bytes("inserted_chart"), JPEG)

    def test_insert_image_after_creates_editable_native_floating_picture(
        self,
    ) -> None:
        source = _image_document(preceding_text="Before")
        document = Document.from_docx(source)
        layout = {
            "horizontal": {
                "relative_to": "column",
                "offset": {"value": 1, "unit": "in"},
            },
            "vertical": {
                "relative_to": "paragraph",
                "offset": {"value": 0.5, "unit": "in"},
            },
            "anchor_distances": {
                "top": {"value": 2, "unit": "pt"},
                "right": {"value": 6, "unit": "pt"},
                "bottom": {"value": 2, "unit": "pt"},
                "left": {"value": 6, "unit": "pt"},
            },
            "anchor_effect_extent": {
                "left": {"value": -0.5, "unit": "pt"},
                "top": {"value": 1, "unit": "pt"},
                "right": {"value": 1.5, "unit": "pt"},
                "bottom": {"value": 2, "unit": "pt"},
            },
            "wrap": {
                "mode": "square",
                "side": "both_sides",
            },
            "relative_height": 1536,
            "behind_text": False,
            "locked": True,
            "layout_in_cell": True,
            "allow_overlap": True,
        }
        result = document.insert_image_after(
            "#before",
            JPEG,
            width={"value": 2, "unit": "in"},
            height={"value": 1, "unit": "in"},
            alt_text="Inserted floating expert diagram",
            media_type="image/jpeg",
            image_id="floating_diagram",
            name="Floating expert diagram",
            title="Expert diagram",
            floating=layout,
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(
            result.changes[0]["placement"],
            "floating",
        )
        assert result.document is not None
        inserted = next(
            node
            for node in result.document.to_spec()["content"]
            if node["id"] == "floating_diagram"
        )
        self.assertEqual(inserted["placement"], "floating")
        self.assertEqual(inserted["floating"], layout)
        self.assertEqual(
            result.document.read_image("floating_diagram").data,
            JPEG,
        )
        self.assertIn(
            "image.anchor.update",
            next(
                node
                for node in result.document.inspect()["nodes"]
                if node["id"] == "floating_diagram"
            )["supported_operations"],
        )
        capabilities = result.document.capabilities()["assets"]
        self.assertEqual(
            capabilities["insert_placements"],
            [
                "inline",
                (
                    "floating_offset_alignment_or_percentage_"
                    "supported_wrap"
                ),
            ],
        )
        self.assertEqual(
            capabilities["insert_default_placement"],
            "inline",
        )
        self.assertEqual(
            capabilities["insert_floating_layout_schema"],
            "floating-image-layout",
        )

        output = result.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as archive:
            root = parse_xml(archive.read("word/document.xml"))
            relationships = parse_xml(
                archive.read("word/_rels/document.xml.rels")
            )
        body = root.find(_q(W, "body"))
        assert body is not None
        inserted_paragraph = body.findall(_q(W, "p"))[1]
        anchor = inserted_paragraph.find(f".//{_q(WP, 'anchor')}")
        assert anchor is not None
        self.assertEqual(
            [child.tag for child in anchor],
            [
                _q(WP, "simplePos"),
                _q(WP, "positionH"),
                _q(WP, "positionV"),
                _q(WP, "extent"),
                _q(WP, "effectExtent"),
                _q(WP, "wrapSquare"),
                _q(WP, "docPr"),
                _q(WP, "cNvGraphicFramePr"),
                _q(A, "graphic"),
            ],
        )
        self.assertEqual(
            anchor.attrib,
            {
                "distT": "25400",
                "distR": "76200",
                "distB": "25400",
                "distL": "76200",
                "simplePos": "0",
                "relativeHeight": "1536",
                "behindDoc": "0",
                "locked": "1",
                "layoutInCell": "1",
                "allowOverlap": "1",
            },
        )
        horizontal = anchor.find(f"./{_q(WP, 'positionH')}")
        vertical = anchor.find(f"./{_q(WP, 'positionV')}")
        assert horizontal is not None
        assert vertical is not None
        self.assertEqual(horizontal.get("relativeFrom"), "column")
        self.assertEqual(vertical.get("relativeFrom"), "paragraph")
        self.assertEqual(
            horizontal.find(f"./{_q(WP, 'posOffset')}").text,
            "914400",
        )
        self.assertEqual(
            vertical.find(f"./{_q(WP, 'posOffset')}").text,
            "457200",
        )
        self.assertEqual(
            anchor.find(f"./{_q(WP, 'wrapSquare')}").attrib,
            {"wrapText": "bothSides"},
        )
        effect_extent = anchor.find(f"./{_q(WP, 'effectExtent')}")
        assert effect_extent is not None
        self.assertEqual(
            effect_extent.attrib,
            {
                "l": "-6350",
                "t": "12700",
                "r": "19050",
                "b": "25400",
            },
        )
        document_properties = anchor.find(f"./{_q(WP, 'docPr')}")
        assert document_properties is not None
        self.assertGreater(int(document_properties.get("id", "0")), 0)
        blip = anchor.find(f".//{_q(A, 'blip')}")
        assert blip is not None
        relationship_id = blip.get(_q(R, "embed"))
        self.assertEqual(
            len(
                [
                    relationship
                    for relationship in relationships.findall(
                        _q(REL, "Relationship")
                    )
                    if relationship.get("Id") == relationship_id
                    and relationship.get("Type")
                    == IMAGE_RELATIONSHIP_TYPE
                ]
            ),
            1,
        )

        reopened = Document.from_docx(output)
        reopened_inserted = next(
            node
            for node in reopened.to_spec()["content"]
            if node["id"] == "floating_diagram"
        )
        self.assertEqual(
            reopened_inserted["floating"]["horizontal"]["offset"],
            {"value": 72.0, "unit": "pt"},
        )
        self.assertEqual(
            reopened_inserted["floating"]["vertical"]["offset"],
            {"value": 36.0, "unit": "pt"},
        )
        moved = reopened.apply(
            [
                {
                    "op": "image.anchor.update",
                    "target": "#floating_diagram",
                    "set": {
                        "horizontal": {
                            "relative_to": "page",
                            "offset": {"value": 96, "unit": "pt"},
                        }
                    },
                }
            ]
        )
        self.assertTrue(moved.success, moved.model_dump())
        assert moved.document is not None
        self.assertEqual(
            moved.document.to_spec()["content"][1]["floating"][
                "horizontal"
            ],
            {
                "relative_to": "page",
                "offset": {"value": 96.0, "unit": "pt"},
            },
        )

    def test_insert_image_after_creates_aligned_floating_picture(
        self,
    ) -> None:
        document = Document.from_docx(
            _image_document(preceding_text="Before")
        )
        layout = {
            "horizontal": {
                "relative_to": "margin",
                "alignment": "center",
            },
            "vertical": {
                "relative_to": "page",
                "alignment": "bottom",
            },
            "anchor_distances": {
                "top": {"value": 3, "unit": "pt"},
                "right": {"value": 5, "unit": "pt"},
                "bottom": {"value": 3, "unit": "pt"},
                "left": {"value": 5, "unit": "pt"},
            },
            "wrap": {
                "mode": "square",
                "side": "largest",
            },
            "relative_height": 4096,
            "behind_text": False,
            "locked": False,
            "layout_in_cell": True,
            "allow_overlap": False,
        }
        result = document.insert_image_after(
            "#before",
            JPEG,
            width={"value": 2, "unit": "in"},
            height={"value": 1, "unit": "in"},
            alt_text="Aligned floating expert diagram",
            media_type="image/jpeg",
            image_id="aligned_floating_diagram",
            floating=layout,
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        inserted = result.document.to_spec()["content"][1]
        self.assertEqual(inserted["id"], "aligned_floating_diagram")
        self.assertEqual(inserted["floating"], layout)
        output = result.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as package:
            root = parse_xml(package.read("word/document.xml"))
        anchor = root.find(f".//{_q(WP, 'anchor')}")
        assert anchor is not None
        horizontal = anchor.find(f"./{_q(WP, 'positionH')}")
        vertical = anchor.find(f"./{_q(WP, 'positionV')}")
        assert horizontal is not None
        assert vertical is not None
        self.assertEqual(horizontal.get("relativeFrom"), "margin")
        self.assertEqual(horizontal[0].tag, _q(WP, "align"))
        self.assertEqual(horizontal[0].text, "center")
        self.assertEqual(vertical.get("relativeFrom"), "page")
        self.assertEqual(vertical[0].tag, _q(WP, "align"))
        self.assertEqual(vertical[0].text, "bottom")
        self.assertIsNone(horizontal.find(f"./{_q(WP, 'posOffset')}"))
        self.assertIsNone(vertical.find(f"./{_q(WP, 'posOffset')}"))

        reopened = Document.from_docx(output)
        reopened_image = reopened.to_spec()["content"][1]
        self.assertEqual(reopened_image["floating"], layout)
        resized = reopened.apply(
            [
                {
                    "op": "image.update",
                    "target": "#aligned_floating_diagram",
                    "set": {
                        "width": {"value": 180, "unit": "pt"},
                        "alt_text": "Resized aligned diagram",
                    },
                }
            ]
        )
        self.assertTrue(resized.success, resized.model_dump())
        assert resized.document is not None
        self.assertEqual(
            resized.document.to_spec()["content"][1]["floating"],
            layout,
        )
        self.assertEqual(
            resized.document.image_bytes("aligned_floating_diagram"),
            JPEG,
        )

    def test_insert_image_after_creates_percentage_floating_picture(
        self,
    ) -> None:
        document = Document.from_docx(
            _image_document(preceding_text="Before")
        )
        layout = {
            "horizontal": {
                "relative_to": "page",
                "percentage_offset": 50.125,
            },
            "vertical": {
                "relative_to": "margin",
                "percentage_offset": -7.5,
            },
            "wrap": {"mode": "none"},
            "relative_size": {
                "width": {
                    "relative_to": "margin",
                    "percentage": 75.125,
                },
                "height": {
                    "relative_to": "page",
                    "percentage": 40,
                },
            },
            "relative_height": 8192,
            "behind_text": True,
            "locked": False,
            "layout_in_cell": True,
            "allow_overlap": True,
        }
        result = document.insert_image_after(
            "#before",
            JPEG,
            width={"value": 2, "unit": "in"},
            height={"value": 1, "unit": "in"},
            alt_text="Percentage-positioned floating diagram",
            media_type="image/jpeg",
            image_id="percentage_floating_diagram",
            floating=layout,
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        inserted = result.document.to_spec()["content"][1]
        self.assertEqual(inserted["floating"], layout)

        output = result.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as package:
            native_xml = package.read("word/document.xml")
            root = parse_xml(native_xml)
        anchor = root.find(f".//{_q(WP, 'anchor')}")
        assert anchor is not None
        horizontal = anchor.find(f"./{_q(WP, 'positionH')}")
        vertical = anchor.find(f"./{_q(WP, 'positionV')}")
        assert horizontal is not None
        assert vertical is not None
        self.assertEqual(
            horizontal[0].tag,
            _q(WP14, "pctPosHOffset"),
        )
        self.assertEqual(horizontal[0].text, "50125")
        self.assertEqual(
            vertical[0].tag,
            _q(WP14, "pctPosVOffset"),
        )
        self.assertEqual(vertical[0].text, "-7500")
        relative_width = anchor.find(f"./{_q(WP14, 'sizeRelH')}")
        relative_height = anchor.find(f"./{_q(WP14, 'sizeRelV')}")
        assert relative_width is not None
        assert relative_height is not None
        self.assertEqual(
            relative_width.attrib,
            {"relativeFrom": "margin"},
        )
        self.assertEqual(relative_width[0].tag, _q(WP14, "pctWidth"))
        self.assertEqual(relative_width[0].text, "75125")
        self.assertEqual(
            relative_height.attrib,
            {"relativeFrom": "page"},
        )
        self.assertEqual(relative_height[0].tag, _q(WP14, "pctHeight"))
        self.assertEqual(relative_height[0].text, "40000")
        self.assertEqual(
            [child.tag for child in anchor[-3:]],
            [
                _q(A, "graphic"),
                _q(WP14, "sizeRelH"),
                _q(WP14, "sizeRelV"),
            ],
        )
        self.assertIn(
            "wp14",
            (root.get(_q(MC, "Ignorable")) or "").split(),
        )
        self.assertIn(b"xmlns:wp14=", native_xml)

        reopened = Document.from_docx(output)
        reopened_image = reopened.to_spec()["content"][1]
        self.assertEqual(reopened_image["floating"], layout)
        self.assertEqual(
            reopened.image_bytes("percentage_floating_diagram"),
            JPEG,
        )
        self.assertEqual(reopened.to_bytes("docx"), output)

    def test_insert_image_after_rejects_unsafe_requests_atomically(self) -> None:
        source = _image_document(preceding_text="Before")
        document = Document.from_docx(source)
        invalid_results = [
            document.insert_image_after(
                "#missing",
                JPEG,
                width={"value": 1, "unit": "in"},
                height={"value": 1, "unit": "in"},
                alt_text="Missing target",
            ),
            document.insert_image_after(
                "#before",
                JPEG,
                width={"value": 0, "unit": "pt"},
                height={"value": 1, "unit": "in"},
                alt_text="Invalid width",
            ),
            document.insert_image_after(
                "#before",
                JPEG,
                width={"value": 1, "unit": "in"},
                height={"value": 1, "unit": "in"},
                alt_text="   ",
            ),
            document.insert_image_after(
                "#before",
                JPEG,
                width={"value": 1, "unit": "in"},
                height={"value": 1, "unit": "in"},
                alt_text="Duplicate ID",
                image_id="before",
            ),
            document.insert_image_after(
                "#before",
                JPEG,
                width={"value": 1, "unit": "in"},
                height={"value": 1, "unit": "in"},
                alt_text="Incomplete floating layout",
                floating={
                    "horizontal": {
                        "relative_to": "page",
                        "offset": {"value": 1, "unit": "in"},
                    }
                },
            ),
        ]
        for result in invalid_results:
            with self.subTest(result=result):
                self.assertFalse(result.success)
                self.assertEqual(result.result_revision, document.revision)
        self.assertEqual(document.to_bytes("docx"), source)

        raw = document.apply(
            [
                {
                    "op": "image.insert_after",
                    "target": "#before",
                    "image": {},
                    "asset": {},
                }
            ]
        )
        self.assertFalse(raw.success)
        self.assertEqual(
            raw.diagnostics[0].code,
            "BINARY_ASSET_REQUIRED",
        )
        detached = Document.from_spec(document.to_spec())
        detached_result = detached.insert_image_after(
            "#before",
            JPEG,
            width={"value": 1, "unit": "in"},
            height={"value": 1, "unit": "in"},
            alt_text="Detached",
        )
        self.assertFalse(detached_result.success)
        self.assertEqual(
            detached_result.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )

    def test_insert_image_after_multi_paragraph_list_uses_last_anchor(
        self,
    ) -> None:
        source = (
            DocumentBuilder()
            .bullet_list(["One", "Two"], id="steps")
            .paragraph("After", id="after")
            .build()
            .to_bytes("docx")
        )
        document = Document.from_docx(source)
        result = document.insert_image_after(
            "#steps",
            JPEG,
            width={"value": 1, "unit": "in"},
            height={"value": 0.5, "unit": "in"},
            alt_text="List result",
            image_id="list_chart",
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        self.assertEqual(
            [
                (node["id"], node["type"])
                for node in result.document.to_spec()["content"]
            ],
            [
                ("steps", "bullet_list"),
                ("list_chart", "image"),
                ("after", "paragraph"),
            ],
        )
        reopened = Document.from_docx(
            result.document.to_bytes("docx")
        )
        self.assertEqual(
            [
                (node["id"], node["type"])
                for node in reopened.to_spec()["content"]
            ],
            [
                ("steps", "bullet_list"),
                ("list_chart", "image"),
                ("after", "paragraph"),
            ],
        )

    def test_insert_image_embeds_identity_manifest_for_third_party_docx(
        self,
    ) -> None:
        source = _image_document(preceding_text="Before")
        with ZipFile(io.BytesIO(source)) as archive:
            root_relationships = parse_xml(
                archive.read("_rels/.rels")
            )
        for relationship in list(root_relationships):
            if (
                relationship.get("Type")
                == MANIFEST_RELATIONSHIP_TYPE
            ):
                root_relationships.remove(relationship)
        source = _rewrite_package(
            source,
            replacements={
                "_rels/.rels": serialize_xml(root_relationships),
            },
            additions={},
            deletions={"customXml/aioffice-manifest.xml"},
        )
        document = Document.from_docx(source)
        target_id = next(
            node["id"]
            for node in document.to_spec()["content"]
            if node["type"] == "paragraph"
        )
        result = document.insert_image_after(
            f"#{target_id}",
            JPEG,
            width={"value": 2, "unit": "in"},
            height={"value": 1, "unit": "in"},
            alt_text="Third-party inserted chart",
            image_id="third_party_chart",
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        output = result.document.to_bytes("docx")
        self.assertIn(
            "/_rels/.rels",
            result.fidelity.affected_parts if result.fidelity else [],
        )
        self.assertIn(
            "/customXml/aioffice-manifest.xml",
            result.fidelity.affected_parts if result.fidelity else [],
        )
        with ZipFile(io.BytesIO(output)) as archive:
            self.assertIn(
                "customXml/aioffice-manifest.xml",
                archive.namelist(),
            )
            output_relationships = parse_xml(
                archive.read("_rels/.rels")
            )
        manifest_relationships = [
            relationship
            for relationship in output_relationships.findall(
                _q(REL, "Relationship")
            )
            if relationship.get("Type")
            == MANIFEST_RELATIONSHIP_TYPE
        ]
        self.assertEqual(len(manifest_relationships), 1)

        reopened = Document.from_docx(output)
        inserted = next(
            node
            for node in reopened.to_spec()["content"]
            if node["type"] == "image"
            and node["id"] == "third_party_chart"
        )
        self.assertEqual(
            inserted["alt_text"],
            "Third-party inserted chart",
        )
        self.assertEqual(
            reopened.image_bytes("third_party_chart"),
            JPEG,
        )

    def test_move_image_relative_content_attaches_third_party_identity(
        self,
    ) -> None:
        source = _image_document(preceding_text="Before")
        with ZipFile(io.BytesIO(source)) as archive:
            root_relationships = parse_xml(
                archive.read("_rels/.rels")
            )
            content_types = parse_xml(
                archive.read("[Content_Types].xml")
            )
        for relationship in list(root_relationships):
            if (
                relationship.get("Type")
                == MANIFEST_RELATIONSHIP_TYPE
            ):
                root_relationships.remove(relationship)
        for override in list(content_types):
            if (
                override.tag == _q(CT, "Override")
                and override.get("PartName")
                == "/customXml/aioffice-manifest.xml"
            ):
                content_types.remove(override)
        source = _rewrite_package(
            source,
            replacements={
                "_rels/.rels": serialize_xml(root_relationships),
                "[Content_Types].xml": serialize_xml(content_types),
            },
            additions={},
            deletions={"customXml/aioffice-manifest.xml"},
        )
        document = Document.from_docx(source)
        before_content = document.to_spec()["content"]
        paragraph = next(
            node
            for node in before_content
            if node["type"] == "paragraph"
        )
        image = next(
            node
            for node in before_content
            if node["type"] == "image"
        )
        original_image = document.image_bytes(image["id"])
        before_root = parse_xml(
            ZipFile(io.BytesIO(source)).read("word/document.xml")
        )
        before_drawing = before_root.find(f".//{_q(W, 'drawing')}")
        assert before_drawing is not None

        result = document.apply(
            [
                {
                    "op": "node.move_before",
                    "target": image["id"],
                    "before": paragraph["id"],
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(document.to_bytes("docx"), source)
        self.assertEqual(
            result.fidelity.affected_parts if result.fidelity else None,
            [
                "/[Content_Types].xml",
                "/_rels/.rels",
                "/customXml/aioffice-manifest.xml",
                "/word/document.xml",
            ],
        )
        assert result.document is not None
        output = result.document.to_bytes("docx")
        self.assertEqual(
            [
                node["id"]
                for node in result.document.to_spec()["content"]
            ],
            [image["id"], paragraph["id"]],
        )
        self.assertEqual(
            result.document.image_bytes(image["id"]),
            original_image,
        )
        with ZipFile(io.BytesIO(output)) as archive:
            self.assertIn(
                "customXml/aioffice-manifest.xml",
                archive.namelist(),
            )
            after_root = parse_xml(
                archive.read("word/document.xml")
            )
            after_relationships = parse_xml(
                archive.read("_rels/.rels")
            )
        after_drawing = after_root.find(f".//{_q(W, 'drawing')}")
        assert after_drawing is not None
        self.assertEqual(
            ET.tostring(after_drawing),
            ET.tostring(before_drawing),
        )
        self.assertEqual(
            len(
                [
                    relationship
                    for relationship in after_relationships.findall(
                        _q(REL, "Relationship")
                    )
                    if relationship.get("Type")
                    == MANIFEST_RELATIONSHIP_TYPE
                ]
            ),
            1,
        )
        reopened = Document.from_docx(output)
        self.assertEqual(
            [
                node["id"]
                for node in reopened.to_spec()["content"]
            ],
            [image["id"], paragraph["id"]],
        )
        self.assertEqual(
            reopened.image_bytes(image["id"]),
            original_image,
        )

    def test_insert_image_after_supports_none_and_top_bottom_wrap(
        self,
    ) -> None:
        native_tags = {
            "none": "wrapNone",
            "top_and_bottom": "wrapTopAndBottom",
        }
        for mode, native_tag in native_tags.items():
            with self.subTest(mode=mode):
                source = _image_document(preceding_text="Before")
                document = Document.from_docx(source)
                layout = {
                    "horizontal": {
                        "relative_to": "margin",
                        "alignment": "center",
                    },
                    "vertical": {
                        "relative_to": "page",
                        "alignment": "center",
                    },
                    "anchor_distances": {
                        "top": {
                            "value": 3,
                            "unit": "pt",
                        },
                        "right": {
                            "value": 5,
                            "unit": "pt",
                        },
                        "bottom": {
                            "value": 3,
                            "unit": "pt",
                        },
                        "left": {
                            "value": 5,
                            "unit": "pt",
                        },
                    },
                    "wrap": {
                        "mode": mode,
                    },
                    "relative_height": 2048,
                    "behind_text": mode == "none",
                    "locked": False,
                    "layout_in_cell": True,
                    "allow_overlap": True,
                }
                image_id = f"inserted_{mode}"
                result = document.insert_image_after(
                    "#before",
                    JPEG,
                    width={"value": 2, "unit": "in"},
                    height={"value": 1, "unit": "in"},
                    alt_text=f"Inserted {mode} image",
                    media_type="image/jpeg",
                    image_id=image_id,
                    floating=layout,
                )
                self.assertTrue(result.success, result.model_dump())
                assert result.document is not None
                inserted = next(
                    node
                    for node in result.document.to_spec()["content"]
                    if node["id"] == image_id
                )
                self.assertEqual(inserted["floating"], layout)
                self.assertNotIn(
                    "side",
                    inserted["floating"]["wrap"],
                )
                output = result.document.to_bytes("docx")
                with ZipFile(io.BytesIO(output)) as package:
                    root = parse_xml(package.read("word/document.xml"))
                anchors = root.findall(f".//{_q(WP, 'anchor')}")
                self.assertEqual(len(anchors), 1)
                wrap = anchors[0].find(f"./{_q(WP, native_tag)}")
                assert wrap is not None
                self.assertFalse(wrap.attrib)
                self.assertFalse(len(wrap))
                reopened = Document.from_docx(output)
                reopened_image = next(
                    node
                    for node in reopened.to_spec()["content"]
                    if node["id"] == image_id
                )
                self.assertEqual(reopened_image["floating"], layout)
                self.assertEqual(reopened.image_bytes(image_id), JPEG)

    def test_insert_image_after_supports_tight_and_through_polygons(
        self,
    ) -> None:
        for mode, native_tag in (
            ("tight", "wrapTight"),
            ("through", "wrapThrough"),
        ):
            with self.subTest(mode=mode):
                source = _image_document(preceding_text="Before")
                document = Document.from_docx(source)
                layout = {
                    "horizontal": {
                        "relative_to": "margin",
                        "alignment": "center",
                    },
                    "vertical": {
                        "relative_to": "page",
                        "alignment": "center",
                    },
                    "wrap": {
                        "mode": mode,
                        "side": "both_sides",
                        "distances": {
                            "left": {"value": 2, "unit": "pt"},
                            "right": {"value": 3, "unit": "pt"},
                        },
                        "polygon": {
                            "start": {"x": 0, "y": 0},
                            "line_to": [
                                {"x": 0, "y": 21600},
                                {"x": 21600, "y": 21600},
                                {"x": 21600, "y": 0},
                            ],
                        },
                    },
                    "relative_height": 2048,
                    "behind_text": False,
                    "locked": False,
                    "layout_in_cell": True,
                    "allow_overlap": True,
                }
                image_id = f"inserted_{mode}"
                result = document.insert_image_after(
                    "#before",
                    JPEG,
                    width={"value": 2, "unit": "in"},
                    height={"value": 1, "unit": "in"},
                    alt_text=f"Inserted {mode} image",
                    media_type="image/jpeg",
                    image_id=image_id,
                    floating=layout,
                )
                self.assertTrue(result.success, result.model_dump())
                assert result.document is not None
                inserted = next(
                    node
                    for node in result.document.to_spec()["content"]
                    if node["id"] == image_id
                )
                self.assertEqual(inserted["floating"], layout)
                output = result.document.to_bytes("docx")
                with ZipFile(io.BytesIO(output)) as package:
                    root = parse_xml(package.read("word/document.xml"))
                wrap = root.find(f".//{_q(WP, native_tag)}")
                assert wrap is not None
                self.assertEqual(
                    wrap.attrib,
                    {
                        "wrapText": "bothSides",
                        "distR": "38100",
                        "distL": "25400",
                    },
                )
                polygon = wrap.find(f"./{_q(WP, 'wrapPolygon')}")
                assert polygon is not None
                self.assertFalse(polygon.attrib)
                self.assertEqual(len(polygon), 4)
                reopened = Document.from_docx(output)
                reopened_image = next(
                    node
                    for node in reopened.to_spec()["content"]
                    if node["id"] == image_id
                )
                self.assertEqual(reopened_image["floating"], layout)
                self.assertEqual(reopened.image_bytes(image_id), JPEG)

    def test_insert_image_after_cli_and_workspace_paths(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            input_path = root / "source.docx"
            replacement_path = root / "replacement.jpg"
            output_path = root / "inserted.docx"
            floating_layout_path = root / "floating-layout.json"
            transform_path = root / "transform.json"
            outline_path = root / "outline.json"
            shadow_path = root / "shadow.json"
            transform = {
                "rotation_degrees_clockwise": 15,
                "flip_vertical": True,
            }
            outline = {
                "width": {"value": 1, "unit": "pt"},
                "color": "#0F6B4F",
                "dash": "dot",
            }
            shadow = {
                "color": "#243040",
                "opacity": 48.125,
                "blur_radius": {"value": 5, "unit": "pt"},
                "distance": {"value": 2, "unit": "pt"},
                "direction_degrees_clockwise": 90,
                "alignment": "center",
                "rotate_with_shape": False,
            }
            floating_layout = {
                "horizontal": {
                    "relative_to": "column",
                    "offset": {"value": 72, "unit": "pt"},
                },
                "vertical": {
                    "relative_to": "paragraph",
                    "offset": {"value": 24, "unit": "pt"},
                },
                "anchor_distances": {
                    "top": {"value": 2, "unit": "pt"},
                    "right": {"value": 4, "unit": "pt"},
                    "bottom": {"value": 2, "unit": "pt"},
                    "left": {"value": 4, "unit": "pt"},
                },
                "wrap": {
                    "mode": "square",
                    "side": "both_sides",
                },
                "relative_height": 1024,
                "behind_text": False,
                "locked": False,
                "layout_in_cell": True,
                "allow_overlap": True,
            }
            input_path.write_bytes(
                _image_document(preceding_text="Before")
            )
            replacement_path.write_bytes(JPEG)
            floating_layout_path.write_text(
                json.dumps(floating_layout),
                encoding="utf-8",
            )
            transform_path.write_text(
                json.dumps(transform),
                encoding="utf-8",
            )
            outline_path.write_text(
                json.dumps(outline),
                encoding="utf-8",
            )
            shadow_path.write_text(
                json.dumps(shadow),
                encoding="utf-8",
            )
            stdout = io.StringIO()
            with redirect_stdout(stdout):
                exit_code = main(
                    [
                        "insert-image-after",
                        str(input_path),
                        "#before",
                        str(replacement_path),
                        "--width",
                        "1.5",
                        "--width-unit",
                        "in",
                        "--height",
                        "0.75",
                        "--height-unit",
                        "in",
                        "--alt-text",
                        "CLI inserted chart",
                        "--image-id",
                        "cli_chart",
                        "--transform",
                        str(transform_path),
                        "--outline",
                        str(outline_path),
                        "--opacity",
                        "64.3214",
                        "--shadow",
                        str(shadow_path),
                        "--floating-layout",
                        str(floating_layout_path),
                        "--align",
                        "center",
                        "-o",
                        str(output_path),
                    ]
                )
            self.assertEqual(exit_code, 0)
            report = json.loads(stdout.getvalue())
            self.assertTrue(report["success"])
            cli_document = Document.from_docx(output_path)
            self.assertEqual(cli_document.image_bytes("cli_chart"), JPEG)
            cli_image = next(
                node
                for node in cli_document.to_spec()["content"]
                if node["id"] == "cli_chart"
            )
            self.assertEqual(cli_image["placement"], "floating")
            self.assertEqual(
                cli_image["transform"],
                {
                    "rotation_degrees_clockwise": 15.0,
                    "flip_horizontal": False,
                    "flip_vertical": True,
                },
            )
            self.assertEqual(cli_image["outline"], outline)
            self.assertEqual(cli_image["opacity"], 64.321)
            self.assertEqual(cli_image["shadow"], shadow)
            self.assertEqual(
                cli_image["floating"]["horizontal"],
                floating_layout["horizontal"],
            )

            workspace = Workspace.init(root / "project")
            tracked = workspace.import_document(input_path)
            workspace_result = workspace.insert_image_after(
                tracked.id,
                "#before",
                replacement_path,
                width={"value": 2, "unit": "in"},
                height={"value": 1, "unit": "in"},
                alt_text="Workspace inserted chart",
                image_id="workspace_chart",
                transform=transform,
                outline=outline,
                opacity=64.3214,
                shadow=shadow,
                floating=floating_layout,
                paragraph_style={"alignment": "right"},
                base_revision=tracked.revision,
            )
            self.assertTrue(
                workspace_result.success,
                workspace_result.model_dump(),
            )
            committed = workspace.open_document(tracked.id)
            self.assertEqual(
                committed.image_bytes("workspace_chart"),
                JPEG,
            )
            workspace_image = next(
                node
                for node in committed.to_spec()["content"]
                if node["id"] == "workspace_chart"
            )
            self.assertEqual(
                workspace_image["placement"],
                "floating",
            )
            self.assertEqual(
                workspace_image["transform"],
                cli_image["transform"],
            )
            self.assertEqual(workspace_image["outline"], outline)
            self.assertEqual(workspace_image["opacity"], 64.321)
            self.assertEqual(workspace_image["shadow"], shadow)
            self.assertIn(
                "insert_image_after",
                workspace.capabilities(tracked.id)["operations"],
            )
            workspace_capabilities = workspace.capabilities(tracked.id)
            self.assertIn(
                "image.anchor.update",
                workspace_capabilities["patch_operations"],
            )
            self.assertEqual(
                workspace_capabilities["binary_operations"][
                    "image.insert_after"
                ]["placements"],
                [
                    "inline",
                    (
                        "floating_offset_alignment_or_percentage_"
                        "supported_wrap"
                    ),
                ],
            )
            patch_path = (
                root
                / "project"
                / ".aioffice"
                / "artifacts"
                / tracked.id
                / "patches"
                / f"{workspace_result.result_revision:08d}.json"
            )
            patch = json.loads(patch_path.read_text(encoding="utf-8"))
            self.assertEqual(
                patch["operations"][0]["op"],
                "image.insert_after",
            )
            self.assertEqual(
                patch["operations"][0]["image"]["placement"],
                "floating",
            )
            self.assertEqual(
                patch["operations"][0]["image"]["floating"],
                floating_layout,
            )
            self.assertEqual(
                patch["operations"][0]["image"]["transform"],
                cli_image["transform"],
            )
            self.assertEqual(
                patch["operations"][0]["image"]["outline"],
                outline,
            )
            self.assertEqual(
                patch["operations"][0]["image"]["opacity"],
                64.321,
            )
            self.assertEqual(
                patch["operations"][0]["image"]["shadow"],
                shadow,
            )
            self.assertEqual(
                patch["changes"][0]["binary_transport"],
                "out_of_band",
            )
            self.assertNotIn(
                "base64",
                patch_path.read_text(encoding="utf-8"),
            )

    def test_semantic_image_without_native_package_is_rejected(self) -> None:
        document = Document.from_spec(
            {
                "assets": [
                    {
                        "id": "asset_"
                        + "a" * 64,
                        "sha256": "a" * 64,
                        "media_type": "image/png",
                        "size_bytes": 1,
                    }
                ],
                "content": [
                    {
                        "id": "image_semantic",
                        "type": "image",
                        "asset_id": "asset_" + "a" * 64,
                        "width": {"value": 1, "unit": "in"},
                        "height": {"value": 1, "unit": "in"},
                    }
                ],
            }
        )
        validation = document.validate()
        self.assertFalse(validation.valid)
        self.assertTrue(
            any(
                diagnostic.code == "UNSUPPORTED_FEATURE"
                for diagnostic in validation.errors
            )
        )


if __name__ == "__main__":
    unittest.main()
