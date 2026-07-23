from __future__ import annotations

import copy
import io
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from aioffice import Document, DocumentBuilder, SpecValidationError
from aioffice.native.xml import parse_xml, serialize_xml

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(local: str) -> str:
    return f"{{{W}}}{local}"


def _rewrite_document_xml(source: bytes, root: ET.Element) -> bytes:
    output = io.BytesIO()
    with (
        ZipFile(io.BytesIO(source)) as before,
        ZipFile(output, "w", compression=ZIP_DEFLATED) as after,
    ):
        for info in before.infolist():
            payload = (
                serialize_xml(root)
                if info.filename == "word/document.xml"
                else before.read(info.filename)
            )
            after.writestr(copy.copy(info), payload)
    return output.getvalue()


def _two_section_document() -> Document:
    return DocumentBuilder(
        sections=[
            {
                "id": "front_section",
                "start_at": None,
                "layout": {
                    "page_size": {
                        "preset": "letter",
                        "orientation": "portrait",
                    },
                    "margin_top": {"value": 72, "unit": "pt"},
                    "margin_right": {"value": 72, "unit": "pt"},
                    "margin_bottom": {"value": 72, "unit": "pt"},
                    "margin_left": {"value": 72, "unit": "pt"},
                    "gutter": {"value": 0, "unit": "pt"},
                    "header_distance": {"value": 36, "unit": "pt"},
                    "footer_distance": {"value": 36, "unit": "pt"},
                    "columns": {"count": 1},
                    "vertical_alignment": "top",
                    "different_first_page": True,
                },
            },
            {
                "id": "body_section",
                "start_at": "body",
                "layout": {
                    "start_type": "next_page",
                    "page_size": {
                        "preset": "a4",
                        "orientation": "landscape",
                    },
                    "margin_top": {"value": 54, "unit": "pt"},
                    "margin_right": {"value": 48, "unit": "pt"},
                    "margin_bottom": {"value": 54, "unit": "pt"},
                    "margin_left": {"value": 48, "unit": "pt"},
                    "gutter": {"value": 9, "unit": "pt"},
                    "header_distance": {"value": 30, "unit": "pt"},
                    "footer_distance": {"value": 30, "unit": "pt"},
                    "columns": {
                        "count": 2,
                        "spacing": {"value": 24, "unit": "pt"},
                        "separator": True,
                    },
                    "vertical_alignment": "center",
                    "different_first_page": False,
                },
            },
        ]
    ).paragraph("Cover", id="cover").paragraph("Body", id="body").build()


class SectionTests(unittest.TestCase):
    def test_strict_page_size_and_section_anchor_validation(self) -> None:
        with self.assertRaises(SpecValidationError):
            Document.from_spec(
                {
                    "sections": [
                        {
                            "id": "bad_page",
                            "layout": {
                                "page_size": {
                                    "preset": "a4",
                                    "width": {"value": 8, "unit": "in"},
                                }
                            },
                        }
                    ]
                }
            )

        invalid_anchor = Document.from_spec(
            {
                "sections": [
                    {"id": "first"},
                    {
                        "id": "second",
                        "start_at": "missing",
                        "layout": {"start_type": "next_page"},
                    },
                ],
                "content": [{"id": "actual", "type": "paragraph", "text": "Body"}],
            }
        )
        self.assertIn(
            "INVALID_SECTION_ANCHOR",
            {diagnostic.code for diagnostic in invalid_anchor.validate().errors},
        )

        overflow = Document.from_spec(
            {
                "sections": [
                    {
                        "id": "columns",
                        "layout": {
                            "page_size": {"preset": "letter"},
                            "margin_left": {"value": 72, "unit": "pt"},
                            "margin_right": {"value": 72, "unit": "pt"},
                            "columns": {
                                "count": 2,
                                "equal_width": False,
                                "columns": [
                                    {
                                        "width": {"value": 300, "unit": "pt"},
                                        "space_after": {"value": 24, "unit": "pt"},
                                    },
                                    {
                                        "width": {"value": 200, "unit": "pt"},
                                        "space_after": {"value": 0, "unit": "pt"},
                                    },
                                ],
                            },
                        },
                    }
                ],
                "content": [{"type": "paragraph", "text": "Body"}],
            }
        )
        self.assertIn(
            "SECTION_COLUMNS_OVERFLOW",
            {diagnostic.code for diagnostic in overflow.validate().errors},
        )

    def test_generated_multi_section_docx_round_trips_semantics_and_identity(self) -> None:
        document = _two_section_document()
        self.assertTrue(document.validate().valid, document.validate().diagnostics)
        source = document.to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as archive:
            root = parse_xml(archive.read("word/document.xml"))
        body = root.find(_q("body"))
        assert body is not None
        children = list(body)
        self.assertEqual(
            [child.tag for child in children],
            [_q("p"), _q("p"), _q("p"), _q("sectPr")],
        )
        carrier = children[1].find(f"./{_q('pPr')}/{_q('sectPr')}")
        assert carrier is not None
        carrier_size = carrier.find(_q("pgSz"))
        assert carrier_size is not None
        self.assertEqual(carrier_size.attrib[_q("w")], "12240")
        self.assertEqual(carrier_size.attrib[_q("h")], "15840")

        final_size = children[-1].find(_q("pgSz"))
        assert final_size is not None
        self.assertEqual(final_size.attrib[_q("w")], "16838")
        self.assertEqual(final_size.attrib[_q("h")], "11906")
        self.assertEqual(final_size.attrib[_q("orient")], "landscape")

        reopened = Document.from_docx(source)
        spec = reopened.to_spec()
        self.assertEqual([node["id"] for node in spec["content"]], ["cover", "body"])
        self.assertEqual(
            [section["id"] for section in spec["sections"]],
            ["front_section", "body_section"],
        )
        self.assertIsNone(spec["sections"][0].get("start_at"))
        self.assertEqual(spec["sections"][1]["start_at"], "body")
        self.assertEqual(
            spec["sections"][1]["layout"]["page_size"],
            {"preset": "a4", "orientation": "landscape"},
        )
        self.assertEqual(spec["sections"][1]["layout"]["columns"]["count"], 2)
        self.assertEqual(reopened.to_bytes("docx"), source)

        inspection = reopened.inspect()
        self.assertEqual(inspection["section_count"], 2)
        self.assertEqual(inspection["sections"][1]["start_at"], "body")
        html = reopened.to_bytes("html").decode()
        self.assertIn('data-aioffice-section="front_section"', html)
        self.assertIn('data-aioffice-section="body_section"', html)
        self.assertIn("width:841.9pt", html)
        self.assertIn("column-count:2", html)

    def test_unequal_columns_project_exact_native_widths(self) -> None:
        document = DocumentBuilder(
            sections=[
                {
                    "id": "unequal",
                    "layout": {
                        "page_size": {"preset": "letter"},
                        "margin_left": {"value": 54, "unit": "pt"},
                        "margin_right": {"value": 54, "unit": "pt"},
                        "columns": {
                            "count": 2,
                            "equal_width": False,
                            "spacing": {"value": 18, "unit": "pt"},
                            "columns": [
                                {
                                    "width": {"value": 180, "unit": "pt"},
                                    "space_after": {"value": 18, "unit": "pt"},
                                },
                                {
                                    "width": {"value": 270, "unit": "pt"},
                                    "space_after": {"value": 0, "unit": "pt"},
                                },
                            ],
                        },
                    },
                }
            ]
        ).paragraph("Columns", id="columns_body").build()
        source = document.to_bytes("docx")
        reopened = Document.from_docx(source)
        columns = reopened.to_spec()["sections"][0]["layout"]["columns"]
        self.assertFalse(columns["equal_width"])
        self.assertEqual(
            [column["width"]["value"] for column in columns["columns"]],
            [180.0, 270.0],
        )
        self.assertEqual(
            [column["space_after"]["value"] for column in columns["columns"]],
            [18.0, 0.0],
        )

    def test_native_section_patch_is_minimal_and_preserves_unknown_xml(self) -> None:
        source = _two_section_document().to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as archive:
            root = parse_xml(archive.read("word/document.xml"))
        body = root.find(_q("body"))
        assert body is not None
        first_section = body.find(f"./{_q('p')}/{_q('pPr')}/{_q('sectPr')}")
        final_section = body.find(_q("sectPr"))
        assert first_section is not None
        assert final_section is not None
        first_before = ET.tostring(first_section, encoding="utf-8")
        margins = final_section.find(_q("pgMar"))
        assert margins is not None
        margins.set("{urn:aioffice:test}futureMargin", "preserve")
        future = ET.SubElement(final_section, "{urn:aioffice:test}futureSection")
        future.set("mode", "keep")
        source = _rewrite_document_xml(source, root)

        imported = Document.from_docx(source)
        target = imported.to_spec()["sections"][1]["id"]
        result = imported.apply(
            [
                {
                    "op": "section.format",
                    "target": f"#{target}",
                    "set": {
                        "margin_left": {"value": 63, "unit": "pt"},
                        "different_first_page": True,
                    },
                    "clear": ["footer_distance"],
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        assert result.fidelity is not None
        self.assertEqual(
            result.fidelity.affected_parts,
            ["/customXml/aioffice-manifest.xml", "/word/document.xml"],
        )
        output = result.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as archive:
            patched_root = parse_xml(archive.read("word/document.xml"))
        patched_body = patched_root.find(_q("body"))
        assert patched_body is not None
        patched_first = patched_body.find(
            f"./{_q('p')}/{_q('pPr')}/{_q('sectPr')}"
        )
        patched_final = patched_body.find(_q("sectPr"))
        assert patched_first is not None
        assert patched_final is not None
        self.assertEqual(
            ET.tostring(patched_first, encoding="utf-8"),
            first_before,
        )
        patched_margins = patched_final.find(_q("pgMar"))
        assert patched_margins is not None
        self.assertEqual(patched_margins.attrib[_q("left")], "1260")
        self.assertNotIn(_q("footer"), patched_margins.attrib)
        self.assertEqual(
            patched_margins.attrib["{urn:aioffice:test}futureMargin"],
            "preserve",
        )
        self.assertIsNotNone(
            patched_final.find("{urn:aioffice:test}futureSection")
        )
        title_page = patched_final.find(_q("titlePg"))
        assert title_page is not None
        self.assertEqual(title_page.attrib[_q("val")], "1")

        reopened = Document.from_docx(output)
        reopened_section = next(
            section
            for section in reopened.to_spec()["sections"]
            if section["id"] == target
        )
        self.assertEqual(
            reopened_section["layout"]["margin_left"],
            {"value": 63.0, "unit": "pt"},
        )
        self.assertTrue(reopened_section["layout"]["different_first_page"])
        self.assertNotIn("footer_distance", reopened_section["layout"])

    def test_unknown_section_format_field_is_atomic(self) -> None:
        document = Document.from_docx(_two_section_document().to_bytes("docx"))
        source = document.to_bytes("docx")
        result = document.apply(
            [
                {
                    "op": "section.format",
                    "target": "#body_section",
                    "set": {"imaginary_margin": {"value": 1, "unit": "pt"}},
                }
            ]
        )
        self.assertFalse(result.success)
        self.assertEqual(result.diagnostics[0].code, "INVALID_SPEC")
        self.assertEqual(document.to_bytes("docx"), source)


if __name__ == "__main__":
    unittest.main()
