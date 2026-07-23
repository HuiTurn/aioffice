from __future__ import annotations

import base64
import copy
import io
import json
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from aioffice.cli.main import main
from aioffice.core.errors import NativePackageError
from aioffice.documents import Document, DocumentBuilder
from aioffice.native.xml import parse_xml, serialize_xml

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


def _q(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def _rewrite_package(
    source: bytes,
    *,
    replacements: dict[str, bytes],
    additions: dict[str, bytes],
) -> bytes:
    output = io.BytesIO()
    with (
        ZipFile(io.BytesIO(source)) as before,
        ZipFile(output, "w", compression=ZIP_DEFLATED) as after,
    ):
        for info in before.infolist():
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

        self.assertEqual(spec["spec_version"], "0.2-draft.14")
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
        capabilities = document.capabilities()["assets"]
        self.assertFalse(capabilities["binary_in_json"])
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
