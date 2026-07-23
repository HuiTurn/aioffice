from __future__ import annotations

import copy
import io
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from aioffice import Document, DocumentBuilder
from aioffice.formats.docx_fields import parse_paragraph_fields
from aioffice.native.xml import parse_xml, serialize_xml

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(local: str) -> str:
    return f"{{{W}}}{local}"


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


def _document_with_fields() -> Document:
    return DocumentBuilder(
        header_footers=[
            {
                "id": "report_header",
                "kind": "header",
                "content": [
                    {
                        "id": "header_line",
                        "type": "paragraph",
                        "content": [
                            {"text": "Section "},
                            {
                                "id": "section_number",
                                "type": "field",
                                "kind": "section_number",
                                "cached_result": "1",
                            },
                        ],
                    }
                ],
            },
            {
                "id": "report_footer",
                "kind": "footer",
                "content": [
                    {
                        "id": "footer_line",
                        "type": "paragraph",
                        "paragraph_style": {"alignment": "center"},
                        "content": [
                            {"text": "Page "},
                            {
                                "id": "current_page",
                                "type": "field",
                                "kind": "page_number",
                                "number_format": "upper_roman",
                                "cached_result": "I",
                            },
                            {"text": " of "},
                            {
                                "id": "total_pages",
                                "type": "field",
                                "kind": "page_count",
                                "cached_result": "12",
                            },
                        ],
                    }
                ],
            },
        ],
        sections=[
            {
                "id": "report_section",
                "layout": {
                    "page_number_start": 5,
                    "page_number_format": "lower_roman",
                },
                "header_footer": {
                    "header_default": "report_header",
                    "footer_default": "report_footer",
                },
            }
        ],
    ).paragraph("Body", id="body").build()


class DynamicFieldTests(unittest.TestCase):
    def test_body_field_semantic_and_native_update(self) -> None:
        semantic = DocumentBuilder().rich_paragraph(
            [
                {"text": "See page "},
                {
                    "id": "body_page",
                    "type": "field",
                    "kind": "page_number",
                    "cached_result": "1",
                },
            ],
            id="field_paragraph",
        ).build()
        semantic_result = semantic.apply(
            [
                {
                    "op": "field.update",
                    "target": "#body_page",
                    "set": {
                        "kind": "section_page_count",
                        "number_format": "upper_letter",
                    },
                }
            ]
        )
        self.assertTrue(semantic_result.success, semantic_result.model_dump())
        assert semantic_result.document is not None
        semantic_field = semantic_result.document.to_spec()["content"][0][
            "content"
        ][1]
        self.assertEqual(semantic_field["kind"], "section_page_count")
        self.assertEqual(semantic_field["number_format"], "upper_letter")

        source = semantic.to_bytes("docx")
        imported = Document.from_docx(source)
        native_result = imported.apply(
            [
                {
                    "op": "field.update",
                    "target": "#body_page",
                    "set": {"kind": "page_count"},
                }
            ]
        )
        self.assertTrue(native_result.success, native_result.model_dump())
        assert native_result.document is not None
        assert native_result.fidelity is not None
        self.assertEqual(
            native_result.fidelity.affected_parts,
            ["/customXml/aioffice-manifest.xml", "/word/document.xml"],
        )
        with ZipFile(
            io.BytesIO(native_result.document.to_bytes("docx"))
        ) as package:
            root = parse_xml(package.read("word/document.xml"))
        instruction = root.find(f".//{_q('instrText')}")
        assert instruction is not None
        self.assertIn("NUMPAGES", instruction.text or "")

    def test_generation_projection_identity_preview_and_section_numbering(self) -> None:
        document = _document_with_fields()
        self.assertTrue(document.validate().valid, document.validate().diagnostics)
        source = document.to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            settings = parse_xml(package.read("word/settings.xml"))
            update_fields = settings.find(_q("updateFields"))
            assert update_fields is not None
            self.assertEqual(update_fields.attrib[_q("val")], "1")

            footer = parse_xml(package.read("word/footer1.xml"))
            instructions = [
                node.text or "" for node in footer.iter(_q("instrText"))
            ]
            self.assertEqual(len(instructions), 2)
            self.assertIn("PAGE \\* ROMAN", instructions[0])
            self.assertIn("NUMPAGES", instructions[1])
            self.assertEqual(
                [
                    node.attrib[_q("fldCharType")]
                    for node in footer.iter(_q("fldChar"))
                ],
                ["begin", "separate", "end", "begin", "separate", "end"],
            )

            document_root = parse_xml(package.read("word/document.xml"))
            section = document_root.find(f".//{_q('sectPr')}")
            assert section is not None
            page_numbering = section.find(_q("pgNumType"))
            assert page_numbering is not None
            self.assertEqual(page_numbering.attrib[_q("start")], "5")
            self.assertEqual(page_numbering.attrib[_q("fmt")], "lowerRoman")

            manifest = package.read("customXml/aioffice-manifest.xml").decode()
            self.assertIn('id="current_page"', manifest)
            self.assertIn('subIndex="0"', manifest)
            self.assertIn('id="total_pages"', manifest)
            self.assertIn('subIndex="1"', manifest)

        reopened = Document.from_docx(source)
        spec = reopened.to_spec()
        footer = next(
            part
            for part in spec["header_footers"]
            if part["id"] == "report_footer"
        )
        fields = [
            inline
            for inline in footer["content"][0]["content"]
            if inline["type"] == "field"
        ]
        self.assertEqual(
            [(field["id"], field["kind"]) for field in fields],
            [
                ("current_page", "page_number"),
                ("total_pages", "page_count"),
            ],
        )
        self.assertEqual(fields[0]["number_format"], "upper_roman")
        self.assertEqual(fields[0]["source_ref"]["sub_index"], 0)
        self.assertEqual(fields[1]["source_ref"]["sub_index"], 1)
        self.assertTrue(spec["settings"]["update_fields_on_open"])
        self.assertEqual(reopened.to_bytes("docx"), source)

        inspection = reopened.inspect()
        self.assertEqual(inspection["field_count"], 3)
        footer_block = next(
            part
            for part in inspection["header_footers"]
            if part["id"] == "report_footer"
        )["blocks"][0]
        self.assertEqual(
            [field["id"] for field in footer_block["fields"]],
            ["current_page", "total_pages"],
        )
        html = reopened.to_bytes("html").decode()
        self.assertIn('data-aioffice-field-kind="page_number"', html)
        self.assertIn('data-aioffice-field-kind="page_count"', html)
        self.assertIn(">I</span>", html)

    def test_native_field_update_touches_only_target_part_and_manifest(self) -> None:
        source = _document_with_fields().to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            footer = parse_xml(package.read("word/footer1.xml"))
            original_document_xml = package.read("word/document.xml")
            original_header = package.read("word/header1.xml")
        instruction = footer.find(f".//{_q('instrText')}")
        assert instruction is not None
        instruction.set("{urn:aioffice:test}futureInstruction", "preserve")
        future = ET.SubElement(footer, "{urn:aioffice:test}futureFooter")
        future.text = "keep"
        source = _rewrite_part(
            source,
            "word/footer1.xml",
            serialize_xml(footer),
        )

        imported = Document.from_docx(source)
        result = imported.apply(
            [
                {
                    "op": "field.update",
                    "target": "#current_page",
                    "set": {
                        "kind": "page_count",
                        "number_format": "lower_roman",
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        assert result.fidelity is not None
        self.assertEqual(
            result.fidelity.affected_parts,
            ["/customXml/aioffice-manifest.xml", "/word/footer1.xml"],
        )
        output = result.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as package:
            self.assertEqual(package.read("word/document.xml"), original_document_xml)
            self.assertEqual(package.read("word/header1.xml"), original_header)
            patched = parse_xml(package.read("word/footer1.xml"))
        patched_instruction = patched.find(f".//{_q('instrText')}")
        assert patched_instruction is not None
        self.assertIn("NUMPAGES \\* roman", patched_instruction.text or "")
        self.assertEqual(
            patched_instruction.attrib[
                "{urn:aioffice:test}futureInstruction"
            ],
            "preserve",
        )
        begin = next(
            node
            for node in patched.iter(_q("fldChar"))
            if node.attrib.get(_q("fldCharType")) == "begin"
        )
        self.assertEqual(begin.attrib[_q("dirty")], "1")
        self.assertIsNotNone(
            patched.find("{urn:aioffice:test}futureFooter")
        )
        self.assertIn(
            "Page I of 12",
            "".join(node.text or "" for node in patched.iter(_q("t"))),
        )

        reopened = Document.from_docx(output)
        current = next(
            field
            for part in reopened.to_spec()["header_footers"]
            for block in part["content"]
            if block["type"] == "paragraph"
            for field in block.get("content", [])
            if field.get("id") == "current_page"
        )
        self.assertEqual(current["kind"], "page_count")
        self.assertEqual(current["number_format"], "lower_roman")
        immediate = next(
            field
            for part in result.document.to_spec()["header_footers"]
            for block in part["content"]
            if block["type"] == "paragraph"
            for field in block.get("content", [])
            if field.get("id") == "current_page"
        )
        self.assertIn(
            "NUMPAGES \\* roman",
            immediate["metadata"]["native_instruction"],
        )
        self.assertTrue(immediate["metadata"]["dirty"])

    def test_simple_field_projects_and_patches_without_reconstruction(self) -> None:
        source = _document_with_fields().to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            footer = parse_xml(package.read("word/footer1.xml"))
        paragraph = footer.find(_q("p"))
        assert paragraph is not None
        first = parse_paragraph_fields(paragraph)[0]
        for element in first.elements:
            paragraph.remove(element)
        simple = ET.Element(
            _q("fldSimple"),
            {
                _q("instr"): " PAGE \\* Arabic ",
                "{urn:aioffice:test}futureField": "preserve",
            },
        )
        run = ET.SubElement(simple, _q("r"))
        ET.SubElement(run, _q("t")).text = "7"
        paragraph.insert(first.start_index, simple)
        source = _rewrite_part(
            source,
            "word/footer1.xml",
            serialize_xml(footer),
        )

        imported = Document.from_docx(source)
        field = next(
            inline
            for part in imported.to_spec()["header_footers"]
            if part["kind"] == "footer"
            for block in part["content"]
            if block["type"] == "paragraph"
            for inline in block["content"]
            if inline["type"] == "field"
            and inline["source_ref"]["native_kind"] == "w:fldSimple"
        )
        self.assertEqual(field["kind"], "page_number")
        self.assertEqual(field["number_format"], "decimal")
        self.assertEqual(field["cached_result"], "7")
        self.assertEqual(imported.to_bytes("docx"), source)

        result = imported.apply(
            [
                {
                    "op": "field.update",
                    "target": f"#{field['id']}",
                    "set": {"kind": "section_page_count"},
                    "clear": ["number_format"],
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        with ZipFile(io.BytesIO(result.document.to_bytes("docx"))) as package:
            patched = parse_xml(package.read("word/footer1.xml"))
        patched_simple = patched.find(f".//{_q('fldSimple')}")
        assert patched_simple is not None
        self.assertIn("SECTIONPAGES", patched_simple.attrib[_q("instr")])
        self.assertEqual(patched_simple.attrib[_q("dirty")], "1")
        self.assertEqual(
            patched_simple.attrib["{urn:aioffice:test}futureField"],
            "preserve",
        )
        self.assertEqual(
            "".join(node.text or "" for node in patched_simple.iter(_q("t"))),
            "7",
        )

    def test_field_without_cached_result_remains_structured(self) -> None:
        source = DocumentBuilder().rich_paragraph(
            [
                {
                    "id": "uncalculated_page",
                    "type": "field",
                    "kind": "page_number",
                }
            ],
            id="field_only_paragraph",
        ).build().to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            document_root = parse_xml(package.read("word/document.xml"))
        paragraph = document_root.find(f".//{_q('p')}")
        assert paragraph is not None
        field = parse_paragraph_fields(paragraph)[0]
        result_text = next(
            node
            for element in field.result_elements
            for node in element.iter(_q("t"))
        )
        result_text.text = None
        source = _rewrite_part(
            source,
            "word/document.xml",
            serialize_xml(document_root),
        )

        imported = Document.from_docx(source)
        paragraph_spec = imported.to_spec()["content"][0]
        self.assertEqual(paragraph_spec["type"], "paragraph")
        self.assertEqual(paragraph_spec["id"], "field_only_paragraph")
        self.assertEqual(
            [
                (inline["id"], inline["kind"], inline.get("cached_result"))
                for inline in paragraph_spec["content"]
            ],
            [("uncalculated_page", "page_number", None)],
        )
        self.assertEqual(imported.to_bytes("docx"), source)

    def test_unknown_field_is_structured_read_only_and_text_ops_refuse_fields(self) -> None:
        source = _document_with_fields().to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            footer = parse_xml(package.read("word/footer1.xml"))
        paragraph = footer.find(_q("p"))
        assert paragraph is not None
        first = parse_paragraph_fields(paragraph)[0]
        first.instruction_nodes[0].text = (
            " DOCPROPERTY Secret \\* MERGEFORMAT "
        )
        source = _rewrite_part(
            source,
            "word/footer1.xml",
            serialize_xml(footer),
        )

        imported = Document.from_docx(source)
        footer_part = next(
            part
            for part in imported.to_spec()["header_footers"]
            if part["kind"] == "footer"
        )
        paragraph_spec = footer_part["content"][0]
        native_field = next(
            inline
            for inline in paragraph_spec["content"]
            if inline["type"] == "field"
            and inline["kind"] == "native"
        )
        self.assertFalse(native_field["editable"])
        self.assertIn("DOCPROPERTY", native_field["instruction"])
        self.assertEqual(imported.to_bytes("docx"), source)

        field_result = imported.apply(
            [
                {
                    "op": "field.update",
                    "target": f"#{native_field['id']}",
                    "set": {"kind": "page_number"},
                }
            ]
        )
        self.assertFalse(field_result.success)
        self.assertEqual(
            field_result.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )
        text_result = imported.apply(
            [
                {
                    "op": "text.replace",
                    "target": f"#{paragraph_spec['id']}",
                    "search": "Page",
                    "replacement": "Sheet",
                }
            ]
        )
        self.assertFalse(text_result.success)
        self.assertEqual(
            text_result.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )
        range_result = imported.apply(
            [
                {
                    "op": "text.format",
                    "target": f"#{paragraph_spec['id']}",
                    "range": {"start": 0, "end": 4},
                    "set": {"bold": True},
                }
            ]
        )
        self.assertFalse(range_result.success)
        self.assertEqual(
            range_result.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )
        self.assertEqual(imported.to_bytes("docx"), source)

    def test_section_page_numbering_native_patch_preserves_unknown_attributes(
        self,
    ) -> None:
        source = _document_with_fields().to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
        section = root.find(f".//{_q('sectPr')}")
        assert section is not None
        page_numbering = section.find(_q("pgNumType"))
        assert page_numbering is not None
        page_numbering.set("{urn:aioffice:test}futureNumbering", "keep")
        source = _rewrite_part(
            source,
            "word/document.xml",
            serialize_xml(root),
        )
        imported = Document.from_docx(source)
        section_id = imported.to_spec()["sections"][0]["id"]
        result = imported.apply(
            [
                {
                    "op": "section.format",
                    "target": f"#{section_id}",
                    "set": {
                        "page_number_start": 20,
                        "page_number_format": "upper_letter",
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        with ZipFile(io.BytesIO(result.document.to_bytes("docx"))) as package:
            patched_root = parse_xml(package.read("word/document.xml"))
        patched_numbering = patched_root.find(f".//{_q('pgNumType')}")
        assert patched_numbering is not None
        self.assertEqual(patched_numbering.attrib[_q("start")], "20")
        self.assertEqual(patched_numbering.attrib[_q("fmt")], "upperLetter")
        self.assertEqual(
            patched_numbering.attrib[
                "{urn:aioffice:test}futureNumbering"
            ],
            "keep",
        )


if __name__ == "__main__":
    unittest.main()
