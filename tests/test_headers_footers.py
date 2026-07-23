from __future__ import annotations

import copy
import io
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from aioffice import Document, DocumentBuilder, SpecValidationError
from aioffice.native.xml import parse_xml, serialize_xml

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"


def _q(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def _rewrite_part(source: bytes, name: str, payload: bytes) -> bytes:
    output = io.BytesIO()
    with (
        ZipFile(io.BytesIO(source)) as before,
        ZipFile(output, "w", compression=ZIP_DEFLATED) as after,
    ):
        for info in before.infolist():
            after.writestr(
                copy.copy(info),
                payload if info.filename == name else before.read(info.filename),
            )
    return output.getvalue()


def _document_with_regions() -> Document:
    return DocumentBuilder(
        settings={"even_and_odd_headers": True},
        header_footers=[
            {
                "id": "report_header",
                "kind": "header",
                "content": [
                    {
                        "id": "header_text",
                        "type": "paragraph",
                        "text": "Confidential report",
                        "paragraph_style": {
                            "alignment": "right",
                            "background_color": "#EAF2F8",
                            "borders": {
                                "bottom": {
                                    "style": "single",
                                    "width": {
                                        "value": 1,
                                        "unit": "pt",
                                    },
                                    "color": "#1F4E78",
                                    "space": {
                                        "value": 2,
                                        "unit": "pt",
                                    },
                                }
                            },
                        },
                    }
                ],
            },
            {
                "id": "even_header",
                "kind": "header",
                "content": [
                    {
                        "id": "even_header_text",
                        "type": "paragraph",
                        "text": "Even page",
                    }
                ],
            },
            {
                "id": "report_footer",
                "kind": "footer",
                "content": [
                    {
                        "id": "footer_text",
                        "type": "paragraph",
                        "text": "AiOffice",
                    }
                ],
            },
        ],
        sections=[
            {
                "id": "front",
                "layout": {"different_first_page": False},
                "header_footer": {
                    "header_default": "report_header",
                    "header_even": "even_header",
                    "footer_default": "report_footer",
                },
            },
            {
                "id": "body_section",
                "start_at": "body",
                "layout": {"start_type": "next_page"},
            },
        ],
    ).paragraph("Cover", id="cover").paragraph("Body", id="body").build()


class HeaderFooterTests(unittest.TestCase):
    def test_generation_projection_inheritance_and_settings(self) -> None:
        document = _document_with_regions()
        self.assertTrue(document.validate().valid, document.validate().diagnostics)
        source = document.to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            names = set(package.namelist())
            self.assertIn("word/header1.xml", names)
            self.assertIn("word/header2.xml", names)
            self.assertIn("word/footer1.xml", names)
            self.assertIn("word/settings.xml", names)

            relationships = parse_xml(
                package.read("word/_rels/document.xml.rels")
            )
            relationship_types = {
                element.attrib["Type"] for element in relationships
            }
            self.assertIn(
                "http://schemas.openxmlformats.org/"
                "officeDocument/2006/relationships/header",
                relationship_types,
            )
            self.assertIn(
                "http://schemas.openxmlformats.org/"
                "officeDocument/2006/relationships/footer",
                relationship_types,
            )
            self.assertIn(
                "http://schemas.openxmlformats.org/"
                "officeDocument/2006/relationships/settings",
                relationship_types,
            )

            content_types = parse_xml(package.read("[Content_Types].xml"))
            overrides = {
                element.attrib["PartName"]: element.attrib["ContentType"]
                for element in content_types.findall(_q(CT, "Override"))
            }
            self.assertIn("/word/header1.xml", overrides)
            self.assertIn("/word/footer1.xml", overrides)
            self.assertIn("/word/settings.xml", overrides)

            settings = parse_xml(package.read("word/settings.xml"))
            even_odd = settings.find(_q(W, "evenAndOddHeaders"))
            assert even_odd is not None
            self.assertEqual(even_odd.attrib[_q(W, "val")], "1")

            document_root = parse_xml(package.read("word/document.xml"))
            references = list(document_root.iter(_q(W, "headerReference")))
            self.assertEqual(
                {reference.attrib[_q(W, "type")] for reference in references},
                {"default", "even"},
            )

        reopened = Document.from_docx(source)
        spec = reopened.to_spec()
        self.assertEqual(
            [part["id"] for part in spec["header_footers"]],
            ["report_header", "even_header", "report_footer"],
        )
        self.assertEqual(
            spec["sections"][0]["header_footer"],
            {
                "header_default": "report_header",
                "header_even": "even_header",
                "footer_default": "report_footer",
            },
        )
        self.assertNotIn("header_footer", spec["sections"][1])
        self.assertTrue(spec["settings"]["even_and_odd_headers"])
        self.assertEqual(reopened.to_bytes("docx"), source)

        inspection = reopened.inspect()
        self.assertEqual(inspection["header_footer_count"], 3)
        self.assertEqual(
            inspection["header_footers"][0]["blocks"][0]["text"],
            "Confidential report",
        )
        html = reopened.to_bytes("html").decode()
        self.assertEqual(html.count("Confidential report"), 2)
        self.assertEqual(html.count("AiOffice"), 3)  # title + two inherited footers
        self.assertIn("background-color:#EAF2F8", html)
        self.assertIn(
            "border-bottom:1pt solid #1F4E78",
            html,
        )

    def test_native_header_text_and_format_patch_touch_only_header_part(self) -> None:
        source = _document_with_regions().to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            header_root = parse_xml(package.read("word/header1.xml"))
            original_document_xml = package.read("word/document.xml")
        paragraph = header_root.find(_q(W, "p"))
        assert paragraph is not None
        paragraph.set("{urn:aioffice:test}futureAttribute", "preserve")
        future = ET.SubElement(header_root, "{urn:aioffice:test}futureHeader")
        future.text = "keep"
        source = _rewrite_part(
            source,
            "word/header1.xml",
            serialize_xml(header_root),
        )

        imported = Document.from_docx(source)
        result = imported.apply(
            [
                {
                    "op": "text.replace",
                    "target": "#header_text",
                    "search": "Confidential",
                    "replacement": "Public",
                },
                {
                    "op": "paragraph.format",
                    "target": "#header_text",
                    "set": {"alignment": "center"},
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        assert result.fidelity is not None
        self.assertEqual(
            result.fidelity.affected_parts,
            ["/customXml/aioffice-manifest.xml", "/word/header1.xml"],
        )
        output = result.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as package:
            self.assertEqual(package.read("word/document.xml"), original_document_xml)
            patched = parse_xml(package.read("word/header1.xml"))
        patched_paragraph = patched.find(_q(W, "p"))
        assert patched_paragraph is not None
        self.assertEqual(
            patched_paragraph.attrib["{urn:aioffice:test}futureAttribute"],
            "preserve",
        )
        self.assertIsNotNone(patched.find("{urn:aioffice:test}futureHeader"))
        self.assertEqual(
            "".join(node.text or "" for node in patched.iter(_q(W, "t"))),
            "Public report",
        )
        alignment = patched_paragraph.find(
            f"./{_q(W, 'pPr')}/{_q(W, 'jc')}"
        )
        assert alignment is not None
        self.assertEqual(alignment.attrib[_q(W, "val")], "center")

        reopened = Document.from_docx(output)
        block = reopened.inspect()["header_footers"][0]["blocks"][0]
        self.assertEqual(block["id"], "header_text")
        self.assertEqual(block["text"], "Public report")

    def test_complex_native_header_is_opaque_not_reconstructed(self) -> None:
        source = _document_with_regions().to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            header_root = parse_xml(package.read("word/header1.xml"))
        paragraph = header_root.find(_q(W, "p"))
        assert paragraph is not None
        run = paragraph.find(_q(W, "r"))
        assert run is not None
        field_start = ET.Element(_q(W, "fldChar"), {_q(W, "fldCharType"): "begin"})
        run.insert(0, field_start)
        source = _rewrite_part(
            source,
            "word/header1.xml",
            serialize_xml(header_root),
        )

        imported = Document.from_docx(source)
        header = next(
            part
            for part in imported.to_spec()["header_footers"]
            if part["kind"] == "header"
            and part["metadata"]["native_part_uri"] == "/word/header1.xml"
        )
        self.assertFalse(header["metadata"]["projection_complete"])
        self.assertEqual(header["content"][0]["type"], "opaque")
        self.assertIn("field", header["content"][0]["summary"])
        self.assertEqual(imported.to_bytes("docx"), source)

    def test_identical_empty_parts_keep_distinct_persistent_ids(self) -> None:
        document = DocumentBuilder(
            header_footers=[
                {"id": "default_empty", "kind": "header"},
                {"id": "first_empty", "kind": "header"},
            ],
            sections=[
                {
                    "id": "only",
                    "layout": {"different_first_page": True},
                    "header_footer": {
                        "header_default": "default_empty",
                        "header_first": "first_empty",
                    },
                }
            ],
        ).paragraph("Body").build()
        reopened = Document.from_docx(document.to_bytes("docx"))
        self.assertEqual(
            [part.id for part in reopened.spec.header_footers],
            ["default_empty", "first_empty"],
        )
        self.assertFalse(
            any(
                diagnostic.code == "IDENTITY_AMBIGUOUS"
                for diagnostic in reopened.import_diagnostics
            )
        )

    def test_header_hyperlink_has_part_scoped_relationship(self) -> None:
        document = DocumentBuilder(
            header_footers=[
                {
                    "id": "linked_header",
                    "kind": "header",
                    "content": [
                        {
                            "id": "linked_header_text",
                            "type": "paragraph",
                            "content": [
                                {"text": "Docs: "},
                                {
                                    "text": "AiOffice",
                                    "marks": ["link"],
                                    "href": "https://example.com/aioffice",
                                },
                            ],
                        }
                    ],
                }
            ],
            sections=[
                {
                    "id": "only",
                    "header_footer": {
                        "header_default": "linked_header"
                    },
                }
            ],
        ).paragraph("Body").build()
        source = document.to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            self.assertIn("word/_rels/header1.xml.rels", package.namelist())
            header_relationships = parse_xml(
                package.read("word/_rels/header1.xml.rels")
            )
            hyperlink = next(iter(header_relationships))
            self.assertEqual(
                hyperlink.attrib["Target"],
                "https://example.com/aioffice",
            )
            self.assertEqual(hyperlink.attrib["TargetMode"], "External")

        reopened = Document.from_docx(source)
        paragraph = reopened.to_spec()["header_footers"][0]["content"][0]
        linked_span = next(
            span for span in paragraph["content"] if span.get("href")
        )
        self.assertEqual(linked_span["text"], "AiOffice")
        self.assertEqual(linked_span["marks"], ["link"])

    def test_binding_kind_mismatch_is_rejected(self) -> None:
        with self.assertRaises(SpecValidationError):
            Document.from_spec(
                {
                    "header_footers": [
                        {"id": "footer_part", "kind": "footer"}
                    ],
                    "sections": [
                        {
                            "id": "only",
                            "header_footer": {
                                "header_default": "footer_part"
                            },
                        }
                    ],
                }
            )


if __name__ == "__main__":
    unittest.main()
