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
    cropped: bool = False,
    alt_text: str | None = "A compact expert workflow diagram",
) -> bytes:
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
    placement = ET.SubElement(
        drawing,
        _q(WP, "anchor" if anchored else "inline"),
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
    if cropped:
        ET.SubElement(blip_fill, _q(A, "srcRect"), {"l": "1000"})
    ET.SubElement(
        blip_fill,
        _q(A, "blip"),
        {_q(R, "embed"): "rIdImage1"},
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


class DocxImageTests(unittest.TestCase):
    def test_cli_exposes_strict_image_and_asset_schemas(self) -> None:
        for kind, required_properties in (
            (
                "image-block",
                {"asset_id", "placement", "width", "height", "editable"},
            ),
            (
                "asset-ref",
                {"id", "sha256", "media_type", "size_bytes"},
            ),
            (
                "image-insert",
                {
                    "id",
                    "width",
                    "height",
                    "alt_text",
                    "paragraph_style",
                },
            ),
            (
                "image-update",
                {"width", "height", "alt_text", "title"},
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

    def test_projects_metadata_and_reads_verified_native_bytes(self) -> None:
        source = _image_document()
        document = Document.from_docx(source)
        spec = document.to_spec()

        self.assertEqual(spec["spec_version"], "0.2-draft.17")
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
            capabilities["native_update_fields"],
            ["width", "height", "alt_text", "title"],
        )
        self.assertEqual(
            capabilities["clearable_update_fields"],
            ["alt_text", "title"],
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

    def test_mixed_floating_and_cropped_images_remain_opaque(self) -> None:
        for source in (
            _image_document(mixed_text="Caption"),
            _image_document(anchored=True),
            _image_document(cropped=True),
        ):
            with self.subTest():
                document = Document.from_docx(source)
                spec = document.to_spec()
                self.assertEqual(spec["content"][0]["type"], "opaque")
                self.assertEqual(spec["assets"], [])
                self.assertEqual(document.to_bytes("docx"), source)
        mixed = Document.from_docx(
            _image_document(mixed_text="Caption")
        ).to_spec()["content"][0]
        self.assertIn("with text", mixed["summary"])

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
        document = Document.from_docx(
            _image_document(preceding_text="Before")
        )
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
            reopened_images[0]["alt_text"],
            "A compact expert workflow diagram",
        )
        self.assertEqual(reopened.image_bytes(first_id), JPEG)
        self.assertEqual(reopened.image_bytes(second_id), PNG)
        self.assertEqual(len(reopened.to_spec()["assets"]), 2)

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

    def test_insert_image_after_cli_and_workspace_paths(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            input_path = root / "source.docx"
            replacement_path = root / "replacement.jpg"
            output_path = root / "inserted.docx"
            input_path.write_bytes(
                _image_document(preceding_text="Before")
            )
            replacement_path.write_bytes(JPEG)
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
            self.assertIn(
                "insert_image_after",
                workspace.capabilities(tracked.id)["operations"],
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
