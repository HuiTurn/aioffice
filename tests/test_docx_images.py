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


def _image_document(
    *,
    preceding_text: str | None = None,
    mixed_text: str | None = None,
    anchored: bool = False,
    aligned: bool = False,
    wrap_mode: str = "square",
    cropped: bool = False,
    alt_text: str | None = "A compact expert workflow diagram",
) -> bytes:
    if aligned and not anchored:
        raise ValueError("Aligned positioning requires a floating anchor.")
    if wrap_mode not in {"square", "none", "top_and_bottom"}:
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
            _q(WP, "align" if aligned else "posOffset"),
        ).text = "center" if aligned else "457200"
        vertical = ET.SubElement(
            placement,
            _q(WP, "positionV"),
            {"relativeFrom": "page" if aligned else "paragraph"},
        )
        ET.SubElement(
            vertical,
            _q(WP, "align" if aligned else "posOffset"),
        ).text = "bottom" if aligned else "5080"
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
        }[wrap_mode]
        ET.SubElement(
            placement,
            _q(WP, wrap_tag),
            {"wrapText": "bothSides"} if wrap_mode == "square" else {},
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
    ET.SubElement(
        blip_fill,
        _q(A, "blip"),
        {_q(R, "embed"): "rIdImage1"},
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
    transform = ET.SubElement(shape, _q(A, "xfrm"))
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
    wrap_mode: str = "square",
) -> bytes:
    assert kind in {"header", "footer"}
    body_image = _image_document(
        cropped=cropped,
        anchored=anchored,
        aligned=aligned,
        wrap_mode=wrap_mode,
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
                "floating-image-horizontal-position",
                {"relative_to", "offset", "alignment"},
            ),
            (
                "floating-image-vertical-position",
                {"relative_to", "offset", "alignment"},
            ),
            (
                "floating-image-text-wrap",
                {
                    "mode",
                    "side",
                    "distance_top",
                    "distance_right",
                    "distance_bottom",
                    "distance_left",
                },
            ),
            (
                "floating-image-layout",
                {
                    "horizontal",
                    "vertical",
                    "wrap",
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
                    "wrap",
                    "relative_height",
                    "behind_text",
                    "locked",
                    "layout_in_cell",
                    "allow_overlap",
                },
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
                    "alt_text",
                    "paragraph_style",
                },
            ),
            (
                "image-update",
                {"width", "height", "crop", "alt_text", "title"},
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
                    {("offset",), ("alignment",)},
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
                    {"square", "none", "top_and_bottom"},
                )
                self.assertIn("side", branches["square"]["required"])
                self.assertEqual(
                    branches["square"]["properties"]["side"],
                    {"not": {"type": "null"}},
                )
                for mode in ("none", "top_and_bottom"):
                    self.assertIn("mode", branches[mode]["required"])
                    self.assertEqual(
                        branches[mode]["not"],
                        {"required": ["side"]},
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

        self.assertEqual(spec["spec_version"], "0.2-draft.38")
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
            ["width", "height", "crop", "alt_text", "title"],
        )
        self.assertEqual(
            capabilities["clearable_update_fields"],
            ["crop", "alt_text", "title"],
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
            "wrap": {
                "mode": "square",
                "side": "both_sides",
                "distance_top": {"value": 1.0, "unit": "pt"},
                "distance_right": {"value": 4.0, "unit": "pt"},
                "distance_bottom": {"value": 2.0, "unit": "pt"},
                "distance_left": {"value": 3.0, "unit": "pt"},
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
                    {
                        "mode": mode,
                        "distance_top": {
                            "value": 1.0,
                            "unit": "pt",
                        },
                        "distance_right": {
                            "value": 4.0,
                            "unit": "pt",
                        },
                        "distance_bottom": {
                            "value": 2.0,
                            "unit": "pt",
                        },
                        "distance_left": {
                            "value": 3.0,
                            "unit": "pt",
                        },
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
        self.assertEqual(
            reopened.to_spec()["content"][0]["floating"]["wrap"],
            {
                "mode": "square",
                "side": "largest",
                **{
                    name: {
                        "value": float(value["value"]),
                        "unit": "pt",
                    }
                    for name, value in distances.items()
                },
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
                "floating_offset_or_alignment_supported_wrap",
            ],
        )
        self.assertEqual(
            capabilities["floating_position_modes"],
            ["offset", "alignment"],
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
            ["square", "none", "top_and_bottom"],
        )
        self.assertEqual(
            capabilities["floating_square_wrap_sides"],
            ["both_sides", "largest", "left", "right"],
        )
        self.assertEqual(
            capabilities["floating_wrap_distance_authority"],
            "four_native_anchor_distances",
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
            "wrap": {
                "mode": "square",
                "side": "right",
                "distance_top": {"value": 5, "unit": "pt"},
                "distance_right": {"value": 6, "unit": "pt"},
                "distance_bottom": {"value": 7, "unit": "pt"},
                "distance_left": {"value": 8, "unit": "pt"},
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
                        "distance_top": {
                            "value": -1,
                            "unit": "pt",
                        },
                        "distance_right": {
                            "value": 0,
                            "unit": "pt",
                        },
                        "distance_bottom": {
                            "value": 0,
                            "unit": "pt",
                        },
                        "distance_left": {
                            "value": 0,
                            "unit": "pt",
                        },
                    }
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "wrap": {
                        "mode": "square",
                        "distance_top": {
                            "value": 0,
                            "unit": "pt",
                        },
                        "distance_right": {
                            "value": 0,
                            "unit": "pt",
                        },
                        "distance_bottom": {
                            "value": 0,
                            "unit": "pt",
                        },
                        "distance_left": {
                            "value": 0,
                            "unit": "pt",
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
                        "side": "both_sides",
                        "distance_top": {
                            "value": 0,
                            "unit": "pt",
                        },
                        "distance_right": {
                            "value": 0,
                            "unit": "pt",
                        },
                        "distance_bottom": {
                            "value": 0,
                            "unit": "pt",
                        },
                        "distance_left": {
                            "value": 0,
                            "unit": "pt",
                        },
                    }
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {
                    "wrap": {
                        "mode": "top_and_bottom",
                        "side": None,
                        "distance_top": {
                            "value": 0,
                            "unit": "pt",
                        },
                        "distance_right": {
                            "value": 0,
                            "unit": "pt",
                        },
                        "distance_bottom": {
                            "value": 0,
                            "unit": "pt",
                        },
                        "distance_left": {
                            "value": 0,
                            "unit": "pt",
                        },
                    }
                },
            },
            {
                "op": "image.anchor.update",
                "target": image["id"],
                "set": {"behind_text": True},
                "clear": [],
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
            "simple_position",
            "tight_wrap",
            "wrap_none_attribute",
            "wrap_text_content",
            "wrap_top_bottom_child",
            "duplicate_wrap",
            "nondefault_bw_mode",
            "duplicate_no_fill",
            "negative_distance",
            "malformed_extension_id",
            "missing_required_attribute",
        )
        for case in cases:
            source = _image_document(anchored=True)
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
            elif case == "simple_position":
                anchor.attrib["simplePos"] = "1"
            elif case == "tight_wrap":
                wrap = anchor.find(f"./{_q(WP, 'wrapSquare')}")
                assert wrap is not None
                wrap.tag = _q(WP, "wrapTight")
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
            elif case == "wrap_top_bottom_child":
                wrap = anchor.find(f"./{_q(WP, 'wrapSquare')}")
                assert wrap is not None
                wrap.tag = _q(WP, "wrapTopAndBottom")
                wrap.attrib.clear()
                ET.SubElement(
                    wrap,
                    _q(WP, "effectExtent"),
                    {"l": "0", "t": "0", "r": "0", "b": "0"},
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
        source = _header_image_document()
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
                        "wrap": {
                            "mode": "top_and_bottom",
                            "distance_top": {
                                "value": 2,
                                "unit": "pt",
                            },
                            "distance_right": {
                                "value": 4,
                                "unit": "pt",
                            },
                            "distance_bottom": {
                                "value": 2,
                                "unit": "pt",
                            },
                            "distance_left": {
                                "value": 4,
                                "unit": "pt",
                            },
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
        self.assertFalse(wrap.attrib)
        self.assertFalse(len(wrap))
        self.assertEqual(
            anchor.get(_q(WP14, "anchorId")),
            "A1B2C3D4",
        )

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
            reopened_image["floating"]["wrap"]["mode"],
            "top_and_bottom",
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

    def test_cloned_header_image_can_be_replaced_copy_on_write(
        self,
    ) -> None:
        source = _header_image_document(cropped=True, anchored=True)
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
            "wrap": {
                "mode": "square",
                "side": "both_sides",
                "distance_top": {"value": 2, "unit": "pt"},
                "distance_right": {"value": 6, "unit": "pt"},
                "distance_bottom": {"value": 2, "unit": "pt"},
                "distance_left": {"value": 6, "unit": "pt"},
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
                "floating_offset_or_alignment_supported_wrap",
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
            "wrap": {
                "mode": "square",
                "side": "largest",
                "distance_top": {"value": 3, "unit": "pt"},
                "distance_right": {"value": 5, "unit": "pt"},
                "distance_bottom": {"value": 3, "unit": "pt"},
                "distance_left": {"value": 5, "unit": "pt"},
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
                    "wrap": {
                        "mode": mode,
                        "distance_top": {
                            "value": 3,
                            "unit": "pt",
                        },
                        "distance_right": {
                            "value": 5,
                            "unit": "pt",
                        },
                        "distance_bottom": {
                            "value": 3,
                            "unit": "pt",
                        },
                        "distance_left": {
                            "value": 5,
                            "unit": "pt",
                        },
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

    def test_insert_image_after_cli_and_workspace_paths(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            input_path = root / "source.docx"
            replacement_path = root / "replacement.jpg"
            output_path = root / "inserted.docx"
            floating_layout_path = root / "floating-layout.json"
            floating_layout = {
                "horizontal": {
                    "relative_to": "column",
                    "offset": {"value": 72, "unit": "pt"},
                },
                "vertical": {
                    "relative_to": "paragraph",
                    "offset": {"value": 24, "unit": "pt"},
                },
                "wrap": {
                    "mode": "square",
                    "side": "both_sides",
                    "distance_top": {"value": 2, "unit": "pt"},
                    "distance_right": {"value": 4, "unit": "pt"},
                    "distance_bottom": {"value": 2, "unit": "pt"},
                    "distance_left": {"value": 4, "unit": "pt"},
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
                    "floating_offset_or_alignment_supported_wrap",
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
