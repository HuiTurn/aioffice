from __future__ import annotations

import copy
import io
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from aioffice import Document, DocumentBuilder, SpecValidationError
from aioffice.formats.docx_import import _unique_id
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


def _section_elements(root: ET.Element) -> list[ET.Element]:
    body = root.find(_q(W, "body"))
    assert body is not None
    result: list[ET.Element] = []
    for element in list(body):
        if element.tag == _q(W, "p"):
            section = element.find(
                f"./{_q(W, 'pPr')}/{_q(W, 'sectPr')}"
            )
            if section is not None:
                result.append(section)
        elif element.tag == _q(W, "sectPr"):
            result.append(element)
    return result


class HeaderFooterTests(unittest.TestCase):
    def test_unanchored_native_ids_resolve_repeated_collisions(
        self,
    ) -> None:
        seen: set[str] = set()
        self.assertEqual(
            [
                _unique_id("para", "000000", 0, seen)
                for _ in range(5)
            ],
            [
                "para_000000",
                "para_000000_000000",
                "para_000000_000000_02",
                "para_000000_000000_03",
                "para_000000_000000_04",
            ],
        )

    def test_semantic_header_footer_create_compiles_normally(
        self,
    ) -> None:
        document = (
            DocumentBuilder()
            .paragraph("Semantic body", id="semantic_body")
            .build()
        )
        result = document.apply(
            [
                {
                    "op": "header_footer.create",
                    "part": {
                        "id": "semantic_footer",
                        "kind": "footer",
                        "content": [
                            {
                                "id": "semantic_footer_text",
                                "type": "paragraph",
                                "text": "Semantic footer",
                            }
                        ],
                    },
                },
                {
                    "op": "section.header_footer.bind",
                    "target": "#section_default",
                    "set": {
                        "footer_default": "semantic_footer",
                    },
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        self.assertIsNone(result.fidelity)
        output = result.document.to_bytes("docx")
        reopened = Document.from_docx(output)
        self.assertEqual(
            reopened.to_spec()["sections"][0]["header_footer"],
            {"footer_default": "semantic_footer"},
        )
        self.assertEqual(
            next(
                part
                for part in reopened.to_spec()["header_footers"]
                if part["id"] == "semantic_footer"
            )["content"][0]["text"],
            "Semantic footer",
        )

    def test_created_empty_header_remains_semantically_empty(
        self,
    ) -> None:
        source = (
            DocumentBuilder()
            .paragraph("Body", id="empty_header_body")
            .build()
            .to_bytes("docx")
        )
        imported = Document.from_docx(source)
        result = imported.apply(
            [
                {
                    "op": "header_footer.create",
                    "part": {
                        "id": "explicit_blank_header",
                        "kind": "header",
                        "content": [],
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        output = result.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as package:
            root = parse_xml(package.read("word/header1.xml"))
        self.assertEqual(
            [child.tag for child in list(root)],
            [_q(W, "p")],
        )
        reopened = Document.from_docx(output)
        created = next(
            part
            for part in reopened.to_spec()["header_footers"]
            if part["id"] == "explicit_blank_header"
        )
        self.assertEqual(created["content"], [])
        self.assertEqual(reopened.to_bytes("docx"), output)

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

    def test_semantic_header_footer_binding_is_explicit_and_stable(
        self,
    ) -> None:
        document = _document_with_regions()
        source = document.to_bytes("docx")
        result = document.apply(
            [
                {
                    "op": "section.header_footer.bind",
                    "target": "#front",
                    "set": {
                        "header_default": "#even_header",
                    },
                    "clear": [
                        "header_even",
                        "footer_default",
                    ],
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(document.to_bytes("docx"), source)
        self.assertEqual(
            result.changes,
            [
                {
                    "operation": "section.header_footer.bind",
                    "section_ids": ["front"],
                    "binding_changes": [
                        {
                            "slot": "footer_default",
                            "before": "report_footer",
                            "after": None,
                        },
                        {
                            "slot": "header_default",
                            "before": "report_header",
                            "after": "even_header",
                        },
                        {
                            "slot": "header_even",
                            "before": "even_header",
                            "after": None,
                        },
                    ],
                }
            ],
        )
        assert result.document is not None
        self.assertEqual(
            result.document.to_spec()["sections"][0]["header_footer"],
            {"header_default": "even_header"},
        )

        output = result.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as package:
            root = parse_xml(package.read("word/document.xml"))
        first_section = _section_elements(root)[0]
        header_references = first_section.findall(
            _q(W, "headerReference")
        )
        self.assertEqual(len(header_references), 1)
        self.assertEqual(
            header_references[0].get(_q(W, "type")),
            "default",
        )
        self.assertEqual(
            first_section.findall(_q(W, "footerReference")),
            [],
        )

        reopened = Document.from_docx(output)
        self.assertEqual(
            reopened.to_spec()["sections"][0]["header_footer"],
            {"header_default": "even_header"},
        )
        self.assertEqual(reopened.to_bytes("docx"), output)

    def test_native_binding_preserves_parts_and_supports_new_section(
        self,
    ) -> None:
        raw_spec = _document_with_regions().to_spec()
        raw_spec["content"].append(
            {
                "id": "body_detail",
                "type": "paragraph",
                "text": "Body detail",
            }
        )
        source = Document.from_spec(raw_spec).to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
        sections = _section_elements(root)
        sections[0].set(
            "{urn:aioffice:test}futureBinding",
            "preserve",
        )
        future = ET.SubElement(
            sections[-1],
            "{urn:aioffice:test}futureSectionProperty",
        )
        future.set("mode", "keep")
        source = _rewrite_part(
            source,
            "word/document.xml",
            serialize_xml(root),
        )
        with ZipFile(io.BytesIO(source)) as package:
            before_parts = {
                name: package.read(name)
                for name in package.namelist()
            }

        imported = Document.from_docx(source)
        source_spec = imported.to_spec()
        front_section_id = source_spec["sections"][0]["id"]
        body_section_id = source_spec["sections"][1]["id"]
        detail_id = next(
            node["id"]
            for node in source_spec["content"]
            if node["id"] == "body_detail"
        )
        parts_by_text = {
            "".join(
                block.get("text", "")
                for block in part["content"]
            ): part["id"]
            for part in source_spec["header_footers"]
        }
        report_header_id = parts_by_text["Confidential report"]
        even_header_id = parts_by_text["Even page"]
        report_footer_id = parts_by_text["AiOffice"]
        original_nodes = {
            node["id"]: [
                ET.tostring(list(root.find(_q(W, "body")))[index])
                for index in node["source_ref"]["element_indices"]
            ]
            for node in source_spec["content"]
        }

        result = imported.apply(
            [
                {
                    "op": "section.header_footer.bind",
                    "target": f"#{front_section_id}",
                    "set": {
                        "header_default": even_header_id,
                    },
                    "clear": [
                        "header_even",
                        "footer_default",
                    ],
                },
                {
                    "op": "section.header_footer.bind",
                    "target": f"#{body_section_id}",
                    "set": {
                        "header_default": report_header_id,
                    },
                },
                {
                    "op": "section.insert_before",
                    "target": f"#{detail_id}",
                    "section": {
                        "id": "detail_section",
                    },
                },
                {
                    "op": "section.header_footer.bind",
                    "target": "#detail_section",
                    "set": {
                        "footer_default": report_footer_id,
                    },
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(imported.to_bytes("docx"), source)
        assert result.document is not None
        assert result.fidelity is not None
        output = result.document.to_bytes("docx")
        result_spec = result.document.to_spec()
        self.assertEqual(
            result_spec["sections"][0]["header_footer"],
            {"header_default": even_header_id},
        )
        self.assertEqual(
            result_spec["sections"][-1]["header_footer"],
            {
                "header_default": report_header_id,
                "footer_default": report_footer_id,
            },
        )
        self.assertEqual(
            result_spec["sections"][-1]["start_at"],
            detail_id,
        )

        with ZipFile(io.BytesIO(output)) as package:
            after_parts = {
                name: package.read(name)
                for name in package.namelist()
            }
            patched_root = parse_xml(
                package.read("word/document.xml")
            )
        self.assertEqual(set(before_parts), set(after_parts))
        changed_parts = {
            name
            for name in before_parts
            if before_parts[name] != after_parts[name]
        }
        self.assertEqual(
            changed_parts,
            {
                "customXml/aioffice-manifest.xml",
                "word/document.xml",
            },
        )
        for name in {
            "word/header1.xml",
            "word/header2.xml",
            "word/footer1.xml",
            "word/_rels/document.xml.rels",
            "[Content_Types].xml",
        }:
            self.assertEqual(
                before_parts[name],
                after_parts[name],
            )

        patched_body = patched_root.find(_q(W, "body"))
        assert patched_body is not None
        for node in result_spec["content"]:
            self.assertEqual(
                [
                    ET.tostring(list(patched_body)[index])
                    for index in (
                        node["source_ref"]["element_indices"]
                    )
                ],
                original_nodes[node["id"]],
            )
        patched_sections = _section_elements(patched_root)
        self.assertEqual(
            patched_sections[0].get(
                "{urn:aioffice:test}futureBinding"
            ),
            "preserve",
        )
        self.assertIsNotNone(
            patched_sections[-1].find(
                "{urn:aioffice:test}futureSectionProperty"
            )
        )

        reopened = Document.from_docx(output)
        reopened_spec = reopened.to_spec()
        self.assertEqual(
            [
                (
                    section["id"],
                    section.get("start_at"),
                    section.get("header_footer"),
                )
                for section in reopened_spec["sections"]
            ],
            [
                (
                    result_spec["sections"][0]["id"],
                    None,
                    {"header_default": even_header_id},
                ),
                (
                    result_spec["sections"][1]["id"],
                    result_spec["sections"][1]["start_at"],
                    {
                        "header_default": report_header_id,
                    },
                ),
                (
                    "detail_section",
                    detail_id,
                    {
                        "header_default": report_header_id,
                        "footer_default": report_footer_id,
                    },
                ),
            ],
        )
        self.assertEqual(reopened.to_bytes("docx"), output)

    def test_native_binding_failures_are_atomic(self) -> None:
        document = _document_with_regions()
        source = document.to_bytes("docx")
        kind_mismatch = document.apply(
            [
                {
                    "op": "section.header_footer.bind",
                    "target": "#front",
                    "set": {
                        "header_default": "report_footer",
                    },
                }
            ]
        )
        self.assertFalse(kind_mismatch.success)
        self.assertEqual(
            kind_mismatch.diagnostics[0].code,
            "TARGET_TYPE_MISMATCH",
        )
        self.assertEqual(document.to_bytes("docx"), source)

        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
        first_section = _section_elements(root)[0]
        default_header = next(
            reference
            for reference in first_section.findall(
                _q(W, "headerReference")
            )
            if reference.get(_q(W, "type")) == "default"
        )
        first_section.insert(0, copy.deepcopy(default_header))
        duplicate_source = _rewrite_part(
            source,
            "word/document.xml",
            serialize_xml(root),
        )
        imported = Document.from_docx(duplicate_source)
        first_section_id = imported.to_spec()["sections"][0]["id"]
        duplicate_result = imported.apply(
            [
                {
                    "op": "section.header_footer.bind",
                    "target": f"#{first_section_id}",
                    "clear": ["header_default"],
                }
            ]
        )
        self.assertFalse(duplicate_result.success)
        self.assertEqual(
            duplicate_result.diagnostics[0].code,
            "NATIVE_PATCH_FAILED",
        )
        self.assertIn(
            "existing native header_default reference",
            duplicate_result.diagnostics[0].message,
        )
        self.assertEqual(
            imported.to_bytes("docx"),
            duplicate_source,
        )

    def test_native_header_footer_create_then_bind_is_transactional(
        self,
    ) -> None:
        source = (
            DocumentBuilder()
            .paragraph("Independent body", id="body")
            .build()
            .to_bytes("docx")
        )
        imported = Document.from_docx(source)
        section_id = imported.to_spec()["sections"][0]["id"]
        result = imported.apply(
            [
                {
                    "op": "header_footer.create",
                    "part": {
                        "id": "created_header",
                        "kind": "header",
                        "metadata": {"role": "running_header"},
                        "content": [
                            {
                                "id": "created_header_text",
                                "type": "paragraph",
                                "paragraph_style": {
                                    "alignment": "right",
                                },
                                "content": [
                                    {
                                        "type": "text",
                                        "text": "Expert report · ",
                                        "marks": ["strong"],
                                    },
                                    {
                                        "id": "created_page_field",
                                        "type": "field",
                                        "kind": "page_number",
                                        "number_format": "decimal",
                                    },
                                    {
                                        "type": "text",
                                        "text": " · AiOffice",
                                        "marks": ["link"],
                                        "href": "https://aioffice.dev",
                                    },
                                ],
                            }
                        ],
                    },
                },
                {
                    "op": "section.header_footer.bind",
                    "target": f"#{section_id}",
                    "set": {
                        "header_default": "#created_header",
                    },
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(imported.to_bytes("docx"), source)
        self.assertEqual(
            result.changes,
            [
                {
                    "operation": "header_footer.create",
                    "part_ids": ["created_header"],
                    "kind": "header",
                    "created_nodes": ["created_header_text"],
                },
                {
                    "operation": "section.header_footer.bind",
                    "section_ids": [section_id],
                    "binding_changes": [
                        {
                            "slot": "header_default",
                            "before": None,
                            "after": "created_header",
                        }
                    ],
                },
            ],
        )
        assert result.document is not None
        assert result.fidelity is not None
        output = result.document.to_bytes("docx")
        output_spec = result.document.to_spec()
        created = next(
            part
            for part in output_spec["header_footers"]
            if part["id"] == "created_header"
        )
        self.assertEqual(
            created["source_ref"]["part_uri"],
            "/word/header1.xml",
        )
        self.assertEqual(
            created["content"][0]["source_ref"]["part_uri"],
            "/word/header1.xml",
        )
        field = next(
            inline
            for inline in created["content"][0]["content"]
            if inline["type"] == "field"
        )
        self.assertEqual(field["revision_added"], 2)
        self.assertEqual(
            field["source_ref"]["native_kind"],
            "w:complex-field",
        )
        self.assertEqual(
            output_spec["sections"][0]["header_footer"],
            {"header_default": "created_header"},
        )

        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(output)) as after,
        ):
            before_parts = {
                name: before.read(name)
                for name in before.namelist()
            }
            after_parts = {
                name: after.read(name)
                for name in after.namelist()
            }
            self.assertEqual(
                set(after_parts) - set(before_parts),
                {
                    "word/header1.xml",
                    "word/_rels/header1.xml.rels",
                },
            )
            changed_existing = {
                name
                for name, payload in before_parts.items()
                if after_parts[name] != payload
            }
            self.assertEqual(
                changed_existing,
                {
                    "[Content_Types].xml",
                    "customXml/aioffice-manifest.xml",
                    "word/_rels/document.xml.rels",
                    "word/document.xml",
                },
            )
            relationships = parse_xml(
                after.read("word/_rels/document.xml.rels")
            )
            self.assertNotIn(
                b"ns0:",
                after.read("word/_rels/document.xml.rels"),
            )
            self.assertIn(
                (
                    b'<Relationships xmlns="http://schemas.'
                    b"openxmlformats.org/package/2006/relationships"
                    b'">'
                ),
                after.read("word/_rels/document.xml.rels"),
            )
            created_relationships = [
                relationship
                for relationship in relationships.findall(
                    _q(REL, "Relationship")
                )
                if relationship.get("Type", "").endswith(
                    "/header"
                )
                and relationship.get("Target") == "header1.xml"
            ]
            self.assertEqual(len(created_relationships), 1)
            local_relationships = parse_xml(
                after.read("word/_rels/header1.xml.rels")
            )
            hyperlinks = [
                relationship
                for relationship in local_relationships.findall(
                    _q(REL, "Relationship")
                )
                if relationship.get("Type", "").endswith(
                    "/hyperlink"
                )
            ]
            self.assertEqual(len(hyperlinks), 1)
            self.assertEqual(
                hyperlinks[0].get("Target"),
                "https://aioffice.dev",
            )
            self.assertEqual(
                hyperlinks[0].get("TargetMode"),
                "External",
            )
            content_types = parse_xml(
                after.read("[Content_Types].xml")
            )
            overrides = [
                override
                for override in content_types.findall(
                    _q(CT, "Override")
                )
                if override.get("PartName")
                == "/word/header1.xml"
            ]
            self.assertEqual(len(overrides), 1)

        reopened = Document.from_docx(output)
        reopened_spec = reopened.to_spec()
        self.assertEqual(
            reopened_spec["sections"][0]["header_footer"],
            {"header_default": "created_header"},
        )
        reopened_created = next(
            part
            for part in reopened_spec["header_footers"]
            if part["id"] == "created_header"
        )
        self.assertEqual(
            "".join(
                inline.get("text", "")
                for inline in reopened_created["content"][0][
                    "content"
                ]
            ),
            "Expert report ·  · AiOffice",
        )
        self.assertEqual(reopened.to_bytes("docx"), output)

    def test_native_header_footer_create_allocates_after_existing_parts(
        self,
    ) -> None:
        source = _document_with_regions().to_bytes("docx")
        imported = Document.from_docx(source)
        result = imported.apply(
            [
                {
                    "op": "header_footer.create",
                    "part": {
                        "id": "third_header",
                        "kind": "header",
                        "content": [
                            {
                                "id": "third_header_text",
                                "type": "paragraph",
                                "text": "Unbound but reusable",
                            }
                        ],
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        created = next(
            part
            for part in result.document.to_spec()["header_footers"]
            if part["id"] == "third_header"
        )
        self.assertEqual(
            created["source_ref"]["part_uri"],
            "/word/header3.xml",
        )
        output = result.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as package:
            self.assertIn("word/header3.xml", package.namelist())
            self.assertNotIn(
                "word/_rels/header3.xml.rels",
                package.namelist(),
            )
        reopened = Document.from_docx(output)
        reopened_created = next(
            part
            for part in reopened.to_spec()["header_footers"]
            if part["id"] == "third_header"
        )
        self.assertEqual(
            reopened_created["content"][0]["text"],
            "Unbound but reusable",
        )
        self.assertEqual(reopened.to_bytes("docx"), output)

    def test_native_header_footer_create_failures_are_atomic(self) -> None:
        source = (
            DocumentBuilder()
            .paragraph("Body", id="body")
            .build()
            .to_bytes("docx")
        )
        imported = Document.from_docx(source)
        claimed = imported.apply(
            [
                {
                    "op": "header_footer.create",
                    "part": {
                        "id": "claimed_header",
                        "kind": "header",
                        "source_ref": {
                            "format": "docx",
                            "part_uri": "/word/header99.xml",
                        },
                    },
                }
            ]
        )
        self.assertFalse(claimed.success)
        self.assertEqual(
            claimed.diagnostics[0].code,
            "INVALID_SPEC",
        )
        self.assertEqual(imported.to_bytes("docx"), source)

        with ZipFile(io.BytesIO(source)) as package:
            relationships = parse_xml(
                package.read("word/_rels/document.xml.rels")
            )
        first = relationships.find(_q(REL, "Relationship"))
        assert first is not None
        relationships.append(copy.deepcopy(first))
        ambiguous_source = _rewrite_part(
            source,
            "word/_rels/document.xml.rels",
            serialize_xml(relationships),
        )
        ambiguous = Document.from_docx(ambiguous_source)
        failed = ambiguous.apply(
            [
                {
                    "op": "header_footer.create",
                    "part": {
                        "id": "safe_header",
                        "kind": "header",
                        "content": [
                            {
                                "id": "safe_header_text",
                                "type": "paragraph",
                                "text": "Safe",
                            }
                        ],
                    },
                }
            ]
        )
        self.assertFalse(failed.success)
        self.assertEqual(
            failed.diagnostics[0].code,
            "NATIVE_PATCH_FAILED",
        )
        self.assertIn(
            "relationship IDs",
            failed.diagnostics[0].message,
        )
        self.assertEqual(
            ambiguous.to_bytes("docx"),
            ambiguous_source,
        )


if __name__ == "__main__":
    unittest.main()
