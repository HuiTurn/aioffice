from __future__ import annotations

import copy
import io
import tempfile
import unittest
from pathlib import Path
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from aioffice.documents import Document, DocumentBuilder
from aioffice.formats.docx import compile_docx
from aioffice.native import (
    FidelityLevel,
    MANIFEST_PART_URI,
    MANIFEST_RELATIONSHIP_TYPE,
)
from aioffice.native.xml import parse_xml, serialize_xml

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
HYPERLINK_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)
NUMBERING_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/"
    "relationships/numbering"
)


def _q(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def _semantic_text(node: dict[str, object]) -> str:
    text = node.get("text")
    if isinstance(text, str):
        return text
    content = node.get("content", [])
    assert isinstance(content, list)
    return "".join(str(span.get("text", "")) for span in content if isinstance(span, dict))


def _rewrite_package(
    source: bytes,
    replacements: dict[str, bytes],
    additions: dict[str, bytes] | None = None,
    deletions: set[str] | None = None,
) -> bytes:
    removed = deletions or set()
    output = io.BytesIO()
    with (
        ZipFile(io.BytesIO(source)) as input_archive,
        ZipFile(output, "w", compression=ZIP_DEFLATED) as output_archive,
    ):
        for info in input_archive.infolist():
            if info.filename in removed:
                continue
            payload = replacements.get(info.filename, input_archive.read(info.filename))
            output_archive.writestr(copy.copy(info), payload)
        for name, payload in (additions or {}).items():
            output_archive.writestr(name, payload)
    return output.getvalue()


class NativeDocxTests(unittest.TestCase):
    def _source_document(self) -> bytes:
        document = (
            DocumentBuilder(title="Native round trip")
            .heading("Native round trip", id="title")
            .rich_paragraph(
                [
                    {"text": "Alpha ", "marks": ["strong"]},
                    {"text": "Beta", "marks": ["emphasis"]},
                    {"text": " Gamma"},
                ],
                id="body",
            )
            .build()
        )
        return compile_docx(document.spec)

    def test_noop_roundtrip_is_exact_package(self) -> None:
        source = self._source_document()
        document = Document.from_docx(source, roundtrip="strict")
        self.assertEqual(document.origin, "native")
        self.assertIsNotNone(document.fidelity)
        assert document.fidelity is not None
        self.assertEqual(document.fidelity.level, FidelityLevel.EXACT_PACKAGE)

        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "roundtrip.docx"
            document.export(target)
            self.assertEqual(target.read_bytes(), source)

    def test_cross_run_patch_preserves_unaffected_and_unknown_parts(self) -> None:
        source = self._source_document()
        document_xml = parse_xml(ZipFile(io.BytesIO(source)).read("word/document.xml"))
        body = document_xml.find(_q(W, "body"))
        assert body is not None
        first_paragraph = next(child for child in body if child.tag == _q(W, "p"))
        unknown = ET.SubElement(first_paragraph, "{urn:aioffice:test}futureFeature")
        unknown.text = "preserve-me"
        source = _rewrite_package(
            source,
            {"word/document.xml": serialize_xml(document_xml)},
            {"customXml/future.xml": b"<future xmlns='urn:aioffice:test'>keep</future>"},
        )

        document = Document.from_docx(source)
        paragraph = next(
            node
            for node in document.to_spec()["content"]
            if _semantic_text(node) == "Alpha Beta Gamma"
        )
        result = document.apply(
            [
                {
                    "op": "text.replace",
                    "target": f"#{paragraph['id']}",
                    "search": "ha Be",
                    "replacement": "HA-BE",
                }
            ],
            dry_run=True,
        )
        self.assertTrue(result.success)
        self.assertIsNotNone(result.document)
        self.assertIsNotNone(result.fidelity)
        assert result.document is not None
        assert result.fidelity is not None
        self.assertEqual(
            result.fidelity.affected_parts,
            ["/customXml/aioffice-manifest.xml", "/word/document.xml"],
        )
        self.assertTrue(result.fidelity.visual_verification_required)

        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "patched.docx"
            result.document.export(target)
            with ZipFile(io.BytesIO(source)) as before, ZipFile(target) as after:
                for name in before.namelist():
                    if name not in {
                        "word/document.xml",
                        "customXml/aioffice-manifest.xml",
                    }:
                        self.assertEqual(before.read(name), after.read(name), name)
                patched_xml = after.read("word/document.xml")
                self.assertIn(b"futureFeature", patched_xml)
                self.assertIn(b"preserve-me", patched_xml)
                self.assertEqual(
                    after.read("customXml/future.xml"),
                    b"<future xmlns='urn:aioffice:test'>keep</future>",
                )
            reopened = Document.from_docx(target)
            texts = [
                _semantic_text(node)
                for node in reopened.to_spec()["content"]
                if node["type"] in {"heading", "paragraph"}
            ]
            self.assertIn("AlpHA-BEta Gamma", texts)

    def test_native_append_preserves_terminal_section_and_existing_xml(
        self,
    ) -> None:
        source = (
            DocumentBuilder(
                sections=[
                    {"id": "front", "start_at": None},
                    {
                        "id": "body_section",
                        "start_at": "body",
                        "layout": {"start_type": "next_page"},
                    },
                ]
            )
            .paragraph("Cover", id="cover")
            .paragraph("Body", id="body")
            .build()
            .to_bytes("docx")
        )
        document = Document.from_docx(source)
        before_spec = document.to_spec()
        before_root = parse_xml(
            ZipFile(io.BytesIO(source)).read("word/document.xml")
        )
        before_body = before_root.find(_q(W, "body"))
        assert before_body is not None
        terminal_section = ET.tostring(list(before_body)[-1])
        original_payloads = {
            node["id"]: [
                ET.tostring(list(before_body)[index])
                for index in node["source_ref"]["element_indices"]
            ]
            for node in before_spec["content"]
        }
        result = document.apply(
            [
                {
                    "op": "node.append",
                    "target": "$",
                    "content": {
                        "id": "appendix",
                        "type": "heading",
                        "level": 2,
                        "text": "Appendix",
                    },
                },
                {
                    "op": "node.append",
                    "target": f"#{document.id}",
                    "content": {
                        "id": "tail",
                        "type": "paragraph",
                        "text": "Draft tail",
                    },
                },
                {
                    "op": "text.replace",
                    "target": "#tail",
                    "search": "Draft",
                    "replacement": "Final",
                },
                {
                    "op": "paragraph.format",
                    "target": "#tail",
                    "set": {"keep_with_next": True},
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(document.to_bytes("docx"), source)
        assert result.document is not None
        result_spec = result.document.to_spec()
        self.assertEqual(
            [node["id"] for node in result_spec["content"]],
            ["cover", "body", "appendix", "tail"],
        )
        self.assertEqual(
            [
                section.get("start_at")
                for section in result_spec["sections"]
            ],
            [
                section.get("start_at")
                for section in before_spec["sections"]
            ],
        )
        output = result.document.to_bytes("docx")
        output_root = parse_xml(
            ZipFile(io.BytesIO(output)).read("word/document.xml")
        )
        output_body = output_root.find(_q(W, "body"))
        assert output_body is not None
        self.assertEqual(
            list(output_body)[-1].tag,
            _q(W, "sectPr"),
        )
        self.assertEqual(
            ET.tostring(list(output_body)[-1]),
            terminal_section,
        )
        for node in result_spec["content"]:
            if node["id"] not in original_payloads:
                continue
            self.assertEqual(
                [
                    ET.tostring(list(output_body)[index])
                    for index in node["source_ref"]["element_indices"]
                ],
                original_payloads[node["id"]],
                node["id"],
            )
        reopened = Document.from_docx(output)
        reopened_nodes = reopened.to_spec()["content"]
        self.assertEqual(
            [node["id"] for node in reopened_nodes],
            ["cover", "body", "appendix", "tail"],
        )
        self.assertEqual(_semantic_text(reopened_nodes[-1]), "Final tail")
        self.assertTrue(
            reopened_nodes[-1]["paragraph_style"]["keep_with_next"]
        )

    def test_native_append_populates_an_empty_document_body(
        self,
    ) -> None:
        source = DocumentBuilder().build().to_bytes("docx")
        document = Document.from_docx(source)
        self.assertEqual(document.to_spec()["content"], [])
        result = document.apply(
            [
                {
                    "op": "node.append",
                    "target": document.id,
                    "content": {
                        "id": "first",
                        "type": "paragraph",
                        "text": "First content",
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        output = result.document.to_bytes("docx")
        output_root = parse_xml(
            ZipFile(io.BytesIO(output)).read("word/document.xml")
        )
        output_body = output_root.find(_q(W, "body"))
        assert output_body is not None
        self.assertEqual(
            [
                element.tag
                for element in list(output_body)
            ],
            [_q(W, "p"), _q(W, "sectPr")],
        )
        reopened = Document.from_docx(output)
        self.assertEqual(
            [
                node["id"]
                for node in reopened.to_spec()["content"]
            ],
            ["first"],
        )

    def test_native_append_table_populates_an_empty_document_body(
        self,
    ) -> None:
        source = DocumentBuilder().build().to_bytes("docx")
        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "node.append",
                    "target": "$",
                    "content": {
                        "id": "first_table",
                        "type": "table",
                        "columns": [
                            {
                                "id": "first_column",
                                "key": "value",
                                "title": "Value",
                            }
                        ],
                        "rows": [
                            {
                                "id": "first_row",
                                "cells": [
                                    {
                                        "id": "first_cell",
                                        "column_key": "value",
                                        "value": "First content",
                                    }
                                ],
                            }
                        ],
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(document.to_bytes("docx"), source)
        assert result.document is not None
        output = result.document.to_bytes("docx")
        output_root = parse_xml(
            ZipFile(io.BytesIO(output)).read("word/document.xml")
        )
        output_body = output_root.find(_q(W, "body"))
        assert output_body is not None
        self.assertEqual(
            [element.tag for element in list(output_body)],
            [_q(W, "tbl"), _q(W, "sectPr")],
        )
        reopened_table = Document.from_docx(output).to_spec()[
            "content"
        ][0]
        self.assertEqual(reopened_table["id"], "first_table")
        self.assertEqual(
            reopened_table["rows"][0]["cells"][0]["source_ref"][
                "native_kind"
            ],
            "w:tc",
        )

    def test_native_append_refuses_nonterminal_body_section_properties(
        self,
    ) -> None:
        source = (
            DocumentBuilder()
            .paragraph("Existing", id="existing")
            .build()
            .to_bytes("docx")
        )
        source_root = parse_xml(
            ZipFile(io.BytesIO(source)).read("word/document.xml")
        )
        source_body = source_root.find(_q(W, "body"))
        assert source_body is not None
        terminal_section = list(source_body)[-1]
        self.assertEqual(terminal_section.tag, _q(W, "sectPr"))
        source_body.remove(terminal_section)
        source_body.insert(0, terminal_section)
        malformed = _rewrite_package(
            source,
            {
                "word/document.xml": serialize_xml(source_root),
            },
        )
        document = Document.from_docx(malformed)
        result = document.apply(
            [
                {
                    "op": "node.append",
                    "target": "$",
                    "content": {
                        "id": "unsafe_append",
                        "type": "paragraph",
                        "text": "Unsafe",
                    },
                }
            ]
        )
        self.assertFalse(result.success)
        self.assertEqual(
            result.diagnostics[0].code,
            "NATIVE_PATCH_FAILED",
        )
        self.assertIn(
            "final and only direct section properties",
            result.diagnostics[0].message,
        )
        self.assertEqual(document.to_bytes("docx"), malformed)

    def test_native_append_supports_body_without_section_properties(
        self,
    ) -> None:
        source = (
            DocumentBuilder()
            .paragraph("Existing", id="existing")
            .build()
            .to_bytes("docx")
        )
        source_root = parse_xml(
            ZipFile(io.BytesIO(source)).read("word/document.xml")
        )
        source_body = source_root.find(_q(W, "body"))
        assert source_body is not None
        terminal_section = list(source_body)[-1]
        self.assertEqual(terminal_section.tag, _q(W, "sectPr"))
        source_body.remove(terminal_section)
        without_section_properties = _rewrite_package(
            source,
            {
                "word/document.xml": serialize_xml(source_root),
            },
        )
        document = Document.from_docx(without_section_properties)
        result = document.apply(
            [
                {
                    "op": "node.append",
                    "target": "$",
                    "content": {
                        "id": "appended",
                        "type": "paragraph",
                        "text": "Appended",
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(
            document.to_bytes("docx"),
            without_section_properties,
        )
        assert result.document is not None
        output_root = parse_xml(
            ZipFile(
                io.BytesIO(result.document.to_bytes("docx"))
            ).read("word/document.xml")
        )
        output_body = output_root.find(_q(W, "body"))
        assert output_body is not None
        self.assertEqual(
            [
                _semantic_text(node)
                for node in Document.from_docx(
                    result.document.to_bytes("docx")
                ).to_spec()["content"]
            ],
            ["Existing", "Appended"],
        )
        self.assertEqual(
            list(output_body)[-1].tag,
            _q(W, "p"),
        )

    def test_native_page_breaks_support_all_insertion_positions(
        self,
    ) -> None:
        source = (
            DocumentBuilder(
                sections=[
                    {"id": "front", "start_at": None},
                    {
                        "id": "body_section",
                        "start_at": "body",
                        "layout": {"start_type": "continuous"},
                    },
                ]
            )
            .paragraph("Cover", id="cover")
            .paragraph("Body", id="body")
            .paragraph("Conclusion", id="conclusion")
            .build()
            .to_bytes("docx")
        )
        document = Document.from_docx(source)
        before_spec = document.to_spec()
        before_root = parse_xml(
            ZipFile(io.BytesIO(source)).read("word/document.xml")
        )
        before_body = before_root.find(_q(W, "body"))
        assert before_body is not None
        terminal_section = ET.tostring(list(before_body)[-1])
        original_payloads = {
            node["id"]: [
                ET.tostring(list(before_body)[index])
                for index in node["source_ref"]["element_indices"]
            ]
            for node in before_spec["content"]
        }
        result = document.apply(
            [
                {
                    "op": "node.insert_after",
                    "target": "#cover",
                    "content": {
                        "id": "after_cover",
                        "type": "page_break",
                    },
                },
                {
                    "op": "node.insert_before",
                    "target": "#body",
                    "content": {
                        "id": "body_break",
                        "type": "page_break",
                    },
                },
                {
                    "op": "node.append",
                    "target": "$",
                    "content": {
                        "id": "final_break",
                        "type": "page_break",
                    },
                },
                {
                    "op": "node.insert_before",
                    "target": "#final_break",
                    "content": {
                        "id": "closing_label",
                        "type": "paragraph",
                        "text": "Closing label",
                    },
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(document.to_bytes("docx"), source)
        assert result.document is not None
        result_spec = result.document.to_spec()
        self.assertEqual(
            [node["id"] for node in result_spec["content"]],
            [
                "cover",
                "after_cover",
                "body_break",
                "body",
                "conclusion",
                "closing_label",
                "final_break",
            ],
        )
        self.assertEqual(
            result_spec["sections"][1]["start_at"],
            "body_break",
        )
        self.assertEqual(
            result.changes[1]["section_start_updated"],
            {
                "section_id": "body_section",
                "from": "body",
                "to": "body_break",
            },
        )
        output = result.document.to_bytes("docx")
        output_root = parse_xml(
            ZipFile(io.BytesIO(output)).read("word/document.xml")
        )
        output_body = output_root.find(_q(W, "body"))
        assert output_body is not None
        self.assertEqual(
            ET.tostring(list(output_body)[-1]),
            terminal_section,
        )
        for node in result_spec["content"]:
            if node["id"] not in original_payloads:
                continue
            self.assertEqual(
                [
                    ET.tostring(list(output_body)[index])
                    for index in node["source_ref"]["element_indices"]
                ],
                original_payloads[node["id"]],
                node["id"],
            )
        for break_id in {
            "after_cover",
            "body_break",
            "final_break",
        }:
            break_node = next(
                node
                for node in result_spec["content"]
                if node["id"] == break_id
            )
            self.assertEqual(break_node["type"], "page_break")
            self.assertEqual(
                break_node["source_ref"]["native_kind"],
                "w:page-break",
            )
            break_element = list(output_body)[
                break_node["source_ref"]["element_index"]
            ]
            native_break = break_element.find(
                f".//{_q(W, 'br')}"
            )
            assert native_break is not None
            self.assertEqual(
                native_break.get(_q(W, "type")),
                "page",
            )
            self.assertIsNone(
                break_element.find(f".//{_q(W, 't')}")
            )
        reopened = Document.from_docx(output)
        self.assertEqual(
            [
                (node["id"], node["type"])
                for node in reopened.to_spec()["content"]
            ],
            [
                ("cover", "paragraph"),
                ("after_cover", "page_break"),
                ("body_break", "page_break"),
                ("body", "paragraph"),
                ("conclusion", "paragraph"),
                ("closing_label", "paragraph"),
                ("final_break", "page_break"),
            ],
        )
        self.assertEqual(
            reopened.to_spec()["sections"][1]["start_at"],
            "body_break",
        )

    def test_native_lists_support_all_positions_numbering_and_batch_ops(
        self,
    ) -> None:
        source = (
            DocumentBuilder(
                sections=[
                    {"id": "front", "start_at": None},
                    {
                        "id": "body_section",
                        "start_at": "body",
                        "layout": {"start_type": "continuous"},
                    },
                ]
            )
            .paragraph("Cover", id="cover")
            .bullet_list(["Existing item"], id="existing_list")
            .paragraph("Body", id="body")
            .paragraph("Conclusion", id="conclusion")
            .build()
            .to_bytes("docx")
        )
        with ZipFile(io.BytesIO(source)) as before:
            before_root = parse_xml(before.read("word/document.xml"))
            before_numbering = parse_xml(
                before.read("word/numbering.xml")
            )
            before_relationships = before.read(
                "word/_rels/document.xml.rels"
            )
            before_content_types = before.read(
                "[Content_Types].xml"
            )
            before_styles = before.read("word/styles.xml")
        before_body = before_root.find(_q(W, "body"))
        assert before_body is not None
        before_spec = Document.from_docx(source).to_spec()
        original_payloads = {
            node["id"]: [
                ET.tostring(list(before_body)[index])
                for index in node["source_ref"]["element_indices"]
            ]
            for node in before_spec["content"]
        }
        original_numbering_children = [
            ET.tostring(child)
            for child in list(before_numbering)
        ]
        original_number_ids = {
            int(number.get(_q(W, "numId")))
            for number in before_numbering.findall(_q(W, "num"))
        }

        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "node.insert_after",
                    "target": "#cover",
                    "content": {
                        "id": "after_cover_list",
                        "type": "bullet_list",
                        "items": [
                            "Review scope",
                            "Confirm owner",
                        ],
                    },
                },
                {
                    "op": "node.move_after",
                    "target": "#after_cover_list",
                    "after": "#existing_list",
                },
                {
                    "op": "node.insert_before",
                    "target": "#body",
                    "content": {
                        "id": "body_steps",
                        "type": "ordered_list",
                        "items": [
                            "Inspect evidence",
                            "Approve decision",
                            "Record outcome",
                        ],
                    },
                },
                {
                    "op": "node.insert_after",
                    "target": "#body_steps",
                    "content": {
                        "id": "steps_note",
                        "type": "paragraph",
                        "text": "Steps complete",
                    },
                },
                {
                    "op": "node.insert_before",
                    "target": "#conclusion",
                    "content": {
                        "id": "temporary_list",
                        "type": "ordered_list",
                        "items": ["Temporary"],
                    },
                },
                {
                    "op": "node.remove",
                    "target": "#temporary_list",
                },
                {
                    "op": "node.append",
                    "target": "$",
                    "content": {
                        "id": "final_list",
                        "type": "bullet_list",
                        "items": [
                            "Archive evidence",
                            "Notify stakeholders",
                        ],
                    },
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(document.to_bytes("docx"), source)
        assert result.document is not None
        result_spec = result.document.to_spec()
        self.assertEqual(
            [
                (node["id"], node["type"])
                for node in result_spec["content"]
            ],
            [
                ("cover", "paragraph"),
                ("existing_list", "bullet_list"),
                ("after_cover_list", "bullet_list"),
                ("body_steps", "ordered_list"),
                ("steps_note", "paragraph"),
                ("body", "paragraph"),
                ("conclusion", "paragraph"),
                ("final_list", "bullet_list"),
            ],
        )
        self.assertEqual(
            result_spec["sections"][1]["start_at"],
            "body_steps",
        )
        self.assertEqual(
            result.changes[2]["section_start_updated"],
            {
                "section_id": "body_section",
                "from": "body",
                "to": "body_steps",
            },
        )
        self.assertEqual(
            result.changes[5]["removed_nodes"],
            ["temporary_list"],
        )
        for list_id, item_count in {
            "after_cover_list": 2,
            "body_steps": 3,
            "final_list": 2,
        }.items():
            inserted_list = next(
                node
                for node in result_spec["content"]
                if node["id"] == list_id
            )
            self.assertEqual(
                inserted_list["source_ref"]["native_kind"],
                "w:p-group",
            )
            self.assertEqual(
                len(
                    inserted_list["source_ref"]["element_indices"]
                ),
                item_count,
            )

        output = result.document.to_bytes("docx")
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(output)) as after,
        ):
            self.assertEqual(
                after.read("word/_rels/document.xml.rels"),
                before_relationships,
            )
            self.assertEqual(
                after.read("[Content_Types].xml"),
                before_content_types,
            )
            self.assertEqual(
                after.read("word/styles.xml"),
                before_styles,
            )
            for name in before.namelist():
                if name not in {
                    "word/document.xml",
                    "word/numbering.xml",
                    "customXml/aioffice-manifest.xml",
                }:
                    self.assertEqual(before.read(name), after.read(name), name)
            after_root = parse_xml(after.read("word/document.xml"))
            after_numbering = parse_xml(
                after.read("word/numbering.xml")
            )
        after_body = after_root.find(_q(W, "body"))
        assert after_body is not None
        for node in result_spec["content"]:
            if node["id"] not in original_payloads:
                continue
            self.assertEqual(
                [
                    ET.tostring(list(after_body)[index])
                    for index in node["source_ref"][
                        "element_indices"
                    ]
                ],
                original_payloads[node["id"]],
                node["id"],
            )
        after_numbering_payloads = [
            ET.tostring(child)
            for child in list(after_numbering)
        ]
        original_cursor = iter(after_numbering_payloads)
        for original_child in original_numbering_children:
            self.assertIn(original_child, original_cursor)
        numbering_child_tags = [
            child.tag for child in list(after_numbering)
        ]
        first_number_index = numbering_child_tags.index(_q(W, "num"))
        self.assertTrue(
            all(
                tag != _q(W, "abstractNum")
                for tag in numbering_child_tags[first_number_index:]
            )
        )
        self.assertEqual(
            len(after_numbering.findall(_q(W, "abstractNum"))),
            len(before_numbering.findall(_q(W, "abstractNum"))) + 4,
        )
        self.assertEqual(
            len(after_numbering.findall(_q(W, "num"))),
            len(before_numbering.findall(_q(W, "num"))) + 4,
        )
        number_to_abstract = {
            int(number.get(_q(W, "numId"))): int(
                number.find(_q(W, "abstractNumId")).get(
                    _q(W, "val")
                )
            )
            for number in after_numbering.findall(_q(W, "num"))
        }
        abstract_formats = {
            int(abstract.get(_q(W, "abstractNumId"))): abstract.find(
                f"./{_q(W, 'lvl')}/{_q(W, 'numFmt')}"
            ).get(_q(W, "val"))
            for abstract in after_numbering.findall(
                _q(W, "abstractNum")
            )
        }
        active_number_ids: dict[str, int] = {}
        for list_id in {
            "after_cover_list",
            "body_steps",
            "final_list",
        }:
            list_node = next(
                node
                for node in result_spec["content"]
                if node["id"] == list_id
            )
            paragraphs = [
                list(after_body)[index]
                for index in list_node["source_ref"][
                    "element_indices"
                ]
            ]
            item_number_ids = {
                int(
                    paragraph.find(
                        f"./{_q(W, 'pPr')}/{_q(W, 'numPr')}/"
                        f"{_q(W, 'numId')}"
                    ).get(_q(W, "val"))
                )
                for paragraph in paragraphs
            }
            self.assertEqual(len(item_number_ids), 1)
            active_number_ids[list_id] = item_number_ids.pop()
        self.assertEqual(
            len(set(active_number_ids.values())),
            len(active_number_ids),
        )
        self.assertTrue(
            set(active_number_ids.values()).isdisjoint(
                original_number_ids
            )
        )
        self.assertEqual(
            abstract_formats[
                number_to_abstract[
                    active_number_ids["after_cover_list"]
                ]
            ],
            "bullet",
        )
        self.assertEqual(
            abstract_formats[
                number_to_abstract[
                    active_number_ids["body_steps"]
                ]
            ],
            "decimal",
        )
        para_ids = [
            paragraph.get(
                "{http://schemas.microsoft.com/office/"
                "word/2010/wordml}paraId"
            )
            for paragraph in after_root.iter(_q(W, "p"))
            if paragraph.get(
                "{http://schemas.microsoft.com/office/"
                "word/2010/wordml}paraId"
            )
            is not None
        ]
        self.assertEqual(len(para_ids), len(set(para_ids)))

        reopened = Document.from_docx(output)
        self.assertEqual(
            [
                (
                    node["id"],
                    node["type"],
                    node.get("items"),
                )
                for node in reopened.to_spec()["content"]
            ],
            [
                ("cover", "paragraph", None),
                (
                    "existing_list",
                    "bullet_list",
                    ["Existing item"],
                ),
                (
                    "after_cover_list",
                    "bullet_list",
                    ["Review scope", "Confirm owner"],
                ),
                (
                    "body_steps",
                    "ordered_list",
                    [
                        "Inspect evidence",
                        "Approve decision",
                        "Record outcome",
                    ],
                ),
                ("steps_note", "paragraph", None),
                ("body", "paragraph", None),
                ("conclusion", "paragraph", None),
                (
                    "final_list",
                    "bullet_list",
                    [
                        "Archive evidence",
                        "Notify stakeholders",
                    ],
                ),
            ],
        )
        self.assertEqual(
            reopened.to_spec()["sections"][1]["start_at"],
            "body_steps",
        )

    def test_native_list_insert_creates_missing_numbering_parts(
        self,
    ) -> None:
        source = (
            DocumentBuilder()
            .paragraph("Anchor", id="anchor")
            .build()
            .to_bytes("docx")
        )
        with ZipFile(io.BytesIO(source)) as package:
            relationships = parse_xml(
                package.read("word/_rels/document.xml.rels")
            )
            content_types = parse_xml(
                package.read("[Content_Types].xml")
            )
        for relationship in list(relationships):
            if (
                relationship.get("Type")
                == NUMBERING_RELATIONSHIP_TYPE
            ):
                relationships.remove(relationship)
        for override in list(content_types):
            if (
                override.tag == _q(CT, "Override")
                and override.get("PartName")
                == "/word/numbering.xml"
            ):
                content_types.remove(override)
        source = _rewrite_package(
            source,
            {
                "word/_rels/document.xml.rels": serialize_xml(
                    relationships
                ),
                "[Content_Types].xml": serialize_xml(content_types),
            },
            deletions={"word/numbering.xml"},
        )
        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "node.append",
                    "target": "$",
                    "content": {
                        "id": "created_steps",
                        "type": "ordered_list",
                        "items": ["First", "Second"],
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(document.to_bytes("docx"), source)
        assert result.document is not None
        output = result.document.to_bytes("docx")
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(output)) as after,
        ):
            self.assertIn("word/numbering.xml", after.namelist())
            numbering = parse_xml(after.read("word/numbering.xml"))
            self.assertEqual(
                numbering.find(
                    f"./{_q(W, 'abstractNum')}/"
                    f"{_q(W, 'lvl')}/{_q(W, 'numFmt')}"
                ).get(_q(W, "val")),
                "decimal",
            )
            output_relationships = parse_xml(
                after.read("word/_rels/document.xml.rels")
            )
            numbering_relationships = [
                relationship
                for relationship in output_relationships.findall(
                    _q(REL, "Relationship")
                )
                if relationship.get("Type")
                == NUMBERING_RELATIONSHIP_TYPE
            ]
            self.assertEqual(len(numbering_relationships), 1)
            self.assertEqual(
                numbering_relationships[0].get("Target"),
                "numbering.xml",
            )
            output_content_types = parse_xml(
                after.read("[Content_Types].xml")
            )
            numbering_overrides = [
                override
                for override in output_content_types.findall(
                    _q(CT, "Override")
                )
                if override.get("PartName")
                == "/word/numbering.xml"
            ]
            self.assertEqual(len(numbering_overrides), 1)
            self.assertEqual(
                numbering_overrides[0].get("ContentType"),
                (
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.numbering+xml"
                ),
            )
            for name in before.namelist():
                if name not in {
                    "[Content_Types].xml",
                    "word/_rels/document.xml.rels",
                    "word/document.xml",
                    "customXml/aioffice-manifest.xml",
                }:
                    self.assertEqual(before.read(name), after.read(name), name)
        reopened = Document.from_docx(output)
        self.assertEqual(
            reopened.to_spec()["content"][-1]["items"],
            ["First", "Second"],
        )
        self.assertEqual(
            reopened.to_spec()["content"][-1]["source_ref"][
                "native_kind"
            ],
            "w:p-group",
        )

    def test_native_list_insert_failures_are_atomic(self) -> None:
        source = (
            DocumentBuilder()
            .paragraph("Anchor", id="anchor")
            .build()
            .to_bytes("docx")
        )
        document = Document.from_docx(source)
        anchor_ref = document.to_spec()["content"][0]["source_ref"]
        forged = document.apply(
            [
                {
                    "op": "node.append",
                    "target": "$",
                    "content": {
                        "id": "forged_list",
                        "type": "bullet_list",
                        "items": ["Unsafe"],
                        "source_ref": anchor_ref,
                    },
                }
            ]
        )
        self.assertFalse(forged.success)
        self.assertIn(
            "cannot claim an existing native source reference",
            forged.diagnostics[0].message,
        )
        unsafe_xml = document.apply(
            [
                {
                    "op": "node.append",
                    "target": "$",
                    "content": {
                        "id": "unsafe_list",
                        "type": "ordered_list",
                        "items": ["Unsafe\u0001item"],
                    },
                }
            ]
        )
        self.assertFalse(unsafe_xml.success)
        self.assertIn(
            "valid, safe XML",
            unsafe_xml.diagnostics[0].message,
        )
        self.assertEqual(document.to_bytes("docx"), source)

        with ZipFile(io.BytesIO(source)) as package:
            numbering = parse_xml(
                package.read("word/numbering.xml")
            )
        duplicate = copy.deepcopy(
            numbering.find(_q(W, "abstractNum"))
        )
        assert duplicate is not None
        first_number = numbering.find(_q(W, "num"))
        assert first_number is not None
        numbering.insert(
            list(numbering).index(first_number),
            duplicate,
        )
        malformed_source = _rewrite_package(
            source,
            {"word/numbering.xml": serialize_xml(numbering)},
        )
        malformed_document = Document.from_docx(
            malformed_source
        )
        malformed = malformed_document.apply(
            [
                {
                    "op": "node.append",
                    "target": "$",
                    "content": {
                        "id": "blocked_list",
                        "type": "bullet_list",
                        "items": ["Blocked"],
                    },
                }
            ]
        )
        self.assertFalse(malformed.success)
        self.assertIn(
            "duplicate abstractNumId",
            malformed.diagnostics[0].message,
        )
        self.assertEqual(
            malformed_document.to_bytes("docx"),
            malformed_source,
        )

        with ZipFile(io.BytesIO(source)) as package:
            numbering = parse_xml(
                package.read("word/numbering.xml")
            )
        displaced_abstract = numbering.find(
            _q(W, "abstractNum")
        )
        assert displaced_abstract is not None
        numbering.remove(displaced_abstract)
        numbering.append(displaced_abstract)
        malformed_order_source = _rewrite_package(
            source,
            {"word/numbering.xml": serialize_xml(numbering)},
        )
        malformed_order_document = Document.from_docx(
            malformed_order_source
        )
        malformed_order = malformed_order_document.apply(
            [
                {
                    "op": "node.append",
                    "target": "$",
                    "content": {
                        "id": "blocked_order_list",
                        "type": "ordered_list",
                        "items": ["Blocked"],
                    },
                }
            ]
        )
        self.assertFalse(malformed_order.success)
        self.assertIn(
            "not in OOXML schema order",
            malformed_order.diagnostics[0].message,
        )
        self.assertEqual(
            malformed_order_document.to_bytes("docx"),
            malformed_order_source,
        )

    def test_unsupported_native_operation_is_atomic(self) -> None:
        source = self._source_document()
        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "node.append",
                    "target": "$",
                    "content": {
                        "id": "unsupported_opaque",
                        "type": "opaque",
                        "summary": "Unsupported native block",
                    },
                }
            ]
        )
        self.assertFalse(result.success)
        self.assertEqual(result.diagnostics[0].code, "NATIVE_PATCH_FAILED")
        self.assertIn(
            "only paragraph, heading, page_break, bullet_list, "
            "ordered_list, and table",
            result.diagnostics[0].message,
        )
        self.assertEqual(document.revision, 1)
        assert document.fidelity is not None
        self.assertEqual(document.fidelity.level, FidelityLevel.EXACT_PACKAGE)

        unsupported_insert = document.apply(
            [
                {
                    "op": "node.insert_after",
                    "target": "#body",
                    "content": {
                        "id": "unsupported_opaque",
                        "type": "opaque",
                        "summary": "Unsupported native block",
                    },
                }
            ]
        )
        self.assertFalse(unsupported_insert.success)
        self.assertEqual(
            unsupported_insert.diagnostics[0].code,
            "NATIVE_PATCH_FAILED",
        )
        self.assertIn(
            "only paragraph, heading, page_break, bullet_list, "
            "ordered_list, and table",
            unsupported_insert.diagnostics[0].message,
        )
        invalid_xml_insert = document.apply(
            [
                {
                    "op": "node.insert_after",
                    "target": "#body",
                    "content": {
                        "id": "invalid_xml",
                        "type": "paragraph",
                        "text": "Unsafe\u0001text",
                    },
                }
            ]
        )
        self.assertFalse(invalid_xml_insert.success)
        self.assertEqual(
            invalid_xml_insert.diagnostics[0].code,
            "NATIVE_PATCH_FAILED",
        )
        self.assertIn(
            "valid, safe XML",
            invalid_xml_insert.diagnostics[0].message,
        )
        self.assertEqual(document.to_bytes("docx"), source)

    def test_native_insert_rich_paragraph_preserves_existing_xml(self) -> None:
        source = (
            DocumentBuilder()
            .paragraph("Anchor", id="anchor")
            .paragraph("Tail", id="tail")
            .build()
            .to_bytes("docx")
        )
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
                and override.get("PartName") == MANIFEST_PART_URI
            ):
                content_types.remove(override)
        source = _rewrite_package(
            source,
            {
                "_rels/.rels": serialize_xml(root_relationships),
                "[Content_Types].xml": serialize_xml(content_types),
            },
            deletions={MANIFEST_PART_URI.lstrip("/")},
        )
        before_root = parse_xml(
            ZipFile(io.BytesIO(source)).read("word/document.xml")
        )
        before_body = before_root.find(_q(W, "body"))
        assert before_body is not None
        existing_payloads = {
            "anchor": ET.tostring(list(before_body)[0]),
            "tail": ET.tostring(list(before_body)[1]),
        }
        document = Document.from_docx(source)
        anchor_id = next(
            str(node["id"])
            for node in document.to_spec()["content"]
            if _semantic_text(node) == "Anchor"
        )
        result = document.apply(
            [
                {
                    "op": "node.insert_after",
                    "target": f"#{anchor_id}",
                    "content": {
                        "type": "paragraph",
                        "content": [
                            {
                                "type": "text",
                                "text": "Executive ",
                                "marks": ["strong"],
                            },
                            {
                                "id": "inserted_page",
                                "type": "field",
                                "kind": "page_number",
                                "number_format": "upper_roman",
                                "cached_result": "IV",
                            },
                            {
                                "type": "text",
                                "text": " / site",
                                "marks": ["link"],
                                "href": "https://example.com/report",
                            },
                            {
                                "type": "text",
                                "text": " / local",
                                "marks": ["link"],
                                "href": "#appendix",
                            },
                        ],
                        "paragraph_style": {
                            "alignment": "center",
                            "spacing_after": {
                                "value": 8,
                                "unit": "pt",
                            },
                        },
                        "text_style": {
                            "font_size": {
                                "value": 11,
                                "unit": "pt",
                            }
                        },
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(document.to_bytes("docx"), source)
        created_id = result.changes[0]["created_nodes"][0]
        self.assertTrue(str(created_id).startswith("para_"))
        assert result.document is not None
        inserted = next(
            node
            for node in result.document.to_spec()["content"]
            if node["id"] == created_id
        )
        self.assertEqual(
            inserted["paragraph_style"]["alignment"],
            "center",
        )
        inserted_field = next(
            inline
            for inline in inserted["content"]
            if inline["type"] == "field"
        )
        self.assertEqual(inserted_field["id"], "inserted_page")
        self.assertEqual(
            inserted_field["source_ref"]["native_kind"],
            "w:complex-field",
        )

        output = result.document.to_bytes("docx")
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(output)) as after,
        ):
            for name in before.namelist():
                if name not in {
                    "[Content_Types].xml",
                    "_rels/.rels",
                    "word/document.xml",
                    "word/_rels/document.xml.rels",
                    "customXml/aioffice-manifest.xml",
                }:
                    self.assertEqual(before.read(name), after.read(name), name)
            relationships = parse_xml(
                after.read("word/_rels/document.xml.rels")
            )
            hyperlinks = [
                relationship
                for relationship in relationships.findall(
                    _q(REL, "Relationship")
                )
                if relationship.get("Type")
                == HYPERLINK_RELATIONSHIP_TYPE
            ]
            self.assertEqual(len(hyperlinks), 1)
            self.assertEqual(
                hyperlinks[0].get("Target"),
                "https://example.com/report",
            )
            self.assertEqual(
                hyperlinks[0].get("TargetMode"),
                "External",
            )
            after_root = parse_xml(after.read("word/document.xml"))
            self.assertIn(
                MANIFEST_PART_URI.lstrip("/"),
                after.namelist(),
            )
        after_body = after_root.find(_q(W, "body"))
        assert after_body is not None
        self.assertEqual(
            ET.tostring(list(after_body)[0]),
            existing_payloads["anchor"],
        )
        self.assertEqual(
            ET.tostring(list(after_body)[2]),
            existing_payloads["tail"],
        )
        native_inserted = list(after_body)[1]
        internal_link = next(
            hyperlink
            for hyperlink in native_inserted.findall(_q(W, "hyperlink"))
            if hyperlink.get(_q(W, "anchor")) is not None
        )
        self.assertEqual(
            internal_link.get(_q(W, "anchor")),
            "appendix",
        )
        external_link = next(
            hyperlink
            for hyperlink in native_inserted.findall(_q(W, "hyperlink"))
            if hyperlink.get(_q(R, "id")) is not None
        )
        self.assertEqual(
            external_link.get(_q(R, "id")),
            hyperlinks[0].get("Id"),
        )
        instruction = native_inserted.find(f".//{_q(W, 'instrText')}")
        assert instruction is not None
        self.assertIn("PAGE", instruction.text or "")
        self.assertIn("ROMAN", instruction.text or "")

        reopened = Document.from_docx(output)
        reopened_inserted = next(
            node
            for node in reopened.to_spec()["content"]
            if node["id"] == created_id
        )
        self.assertEqual(
            reopened_inserted["paragraph_style"]["alignment"],
            "center",
        )
        self.assertEqual(
            next(
                inline
                for inline in reopened_inserted["content"]
                if inline["type"] == "field"
            )["id"],
            "inserted_page",
        )
        self.assertEqual(
            [
                inline.get("href")
                for inline in reopened_inserted["content"]
                if inline.get("href") is not None
            ],
            [
                "https://example.com/report",
                "#appendix",
            ],
        )

    def test_native_insert_table_maps_components_and_is_batch_addressable(
        self,
    ) -> None:
        source = (
            DocumentBuilder()
            .paragraph("Anchor", id="anchor")
            .paragraph("Tail", id="tail")
            .build()
            .to_bytes("docx")
        )
        with ZipFile(io.BytesIO(source)) as before:
            before_root = parse_xml(before.read("word/document.xml"))
            before_relationships = parse_xml(
                before.read("word/_rels/document.xml.rels")
            )
        before_body = before_root.find(_q(W, "body"))
        assert before_body is not None
        existing_payloads = {
            "anchor": ET.tostring(list(before_body)[0]),
            "tail": ET.tostring(list(before_body)[1]),
            "section": ET.tostring(list(before_body)[2]),
        }
        existing_relationships = {
            tuple(sorted(relationship.attrib.items()))
            for relationship in before_relationships.findall(
                _q(REL, "Relationship")
            )
        }

        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "node.insert_after",
                    "target": "#anchor",
                    "content": {
                        "id": "metrics_table",
                        "type": "table",
                        "columns": [
                            {
                                "key": "metric",
                                "title": "Metric",
                                "width": {
                                    "value": 120,
                                    "unit": "pt",
                                },
                            },
                            {
                                "id": "value_column",
                                "key": "value",
                                "title": "Value",
                                "width": {
                                    "value": 96,
                                    "unit": "pt",
                                },
                            },
                        ],
                        "rows": [
                            {
                                "id": "growth_row",
                                "cells": [
                                    {
                                        "id": "growth_label",
                                        "column_key": "metric",
                                        "value": "Growth",
                                    },
                                    {
                                        "id": "growth_value",
                                        "column_key": "value",
                                        "value": "18%",
                                    },
                                ],
                            },
                            {
                                "id": "evidence_row",
                                "cells": [
                                    {
                                        "id": "evidence_cell",
                                        "column_key": "metric",
                                        "content": [
                                            {
                                                "id": "evidence_paragraph",
                                                "type": "paragraph",
                                                "content": [
                                                    {
                                                        "type": "text",
                                                        "text": "Evidence",
                                                        "marks": [
                                                            "strong"
                                                        ],
                                                    },
                                                    {
                                                        "type": "text",
                                                        "text": " source",
                                                        "marks": [
                                                            "link"
                                                        ],
                                                        "href": (
                                                            "https://"
                                                            "example.com/"
                                                            "evidence"
                                                        ),
                                                    },
                                                    {
                                                        "type": "text",
                                                        "text": " / appendix",
                                                        "marks": [
                                                            "link"
                                                        ],
                                                        "href": "#appendix",
                                                    },
                                                ],
                                                "paragraph_style": {
                                                    "alignment": "center"
                                                },
                                            }
                                        ],
                                    },
                                    {
                                        "id": "evidence_status",
                                        "column_key": "value",
                                        "value": "Reviewed",
                                    },
                                ],
                            },
                            {
                                "cells": [
                                    {
                                        "column_key": "metric",
                                        "value": "Generated IDs",
                                    },
                                    {
                                        "column_key": "value",
                                        "value": "mapped",
                                    },
                                ]
                            },
                        ],
                        "layout": {
                            "style_ref": "TableGrid",
                            "alignment": "center",
                            "algorithm": "fixed",
                            "repeat_header": True,
                        },
                    },
                },
                {
                    "op": "table.format",
                    "target": "#metrics_table",
                    "set": {
                        "alignment": "right",
                        "cell_margin_left": {
                            "value": 6,
                            "unit": "pt",
                        },
                    },
                },
                {
                    "op": "table.column.format",
                    "target": "#metrics_table",
                    "column": "#value_column",
                    "set": {
                        "width": {
                            "value": 144,
                            "unit": "pt",
                        }
                    },
                },
                {
                    "op": "table.cell.format",
                    "target": "#metrics_table",
                    "cell": "#growth_value",
                    "set": {
                        "vertical_alignment": "center",
                        "background_color": "#E2F0D9",
                    },
                },
                {
                    "op": "text.replace",
                    "target": "#evidence_paragraph",
                    "search": "Evidence",
                    "replacement": "Audited evidence",
                },
                {
                    "op": "node.insert_after",
                    "target": "#metrics_table",
                    "content": {
                        "id": "after_table",
                        "type": "paragraph",
                        "text": "After table",
                    },
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(document.to_bytes("docx"), source)
        assert result.document is not None
        self.assertEqual(
            [
                node["id"]
                for node in result.document.to_spec()["content"]
            ],
            [
                "anchor",
                "metrics_table",
                "after_table",
                "tail",
            ],
        )
        inserted_table = next(
            node
            for node in result.document.to_spec()["content"]
            if node["id"] == "metrics_table"
        )
        self.assertEqual(
            inserted_table["source_ref"]["native_kind"],
            "w:tbl",
        )
        self.assertTrue(
            all(
                component["source_ref"]["native_kind"]
                == expected_kind
                for components, expected_kind in (
                    (inserted_table["columns"], "w:gridCol"),
                    (inserted_table["rows"], "w:tr"),
                    (
                        [
                            cell
                            for row in inserted_table["rows"]
                            for cell in row["cells"]
                        ],
                        "w:tc",
                    ),
                )
                for component in components
            )
        )
        evidence_paragraph = inserted_table["rows"][1]["cells"][0][
            "content"
        ][0]
        self.assertEqual(
            evidence_paragraph["source_ref"]["native_kind"],
            "w:tc/w:p",
        )
        self.assertEqual(
            _semantic_text(evidence_paragraph),
            "Audited evidence source / appendix",
        )

        output = result.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as after:
            after_root = parse_xml(after.read("word/document.xml"))
            after_relationships = parse_xml(
                after.read("word/_rels/document.xml.rels")
            )
        after_body = after_root.find(_q(W, "body"))
        assert after_body is not None
        self.assertEqual(
            ET.tostring(list(after_body)[0]),
            existing_payloads["anchor"],
        )
        self.assertEqual(
            ET.tostring(list(after_body)[3]),
            existing_payloads["tail"],
        )
        self.assertEqual(
            ET.tostring(list(after_body)[4]),
            existing_payloads["section"],
        )
        native_table = list(after_body)[1]
        self.assertEqual(native_table.tag, _q(W, "tbl"))
        properties = native_table.find(_q(W, "tblPr"))
        assert properties is not None
        self.assertEqual(
            properties.find(_q(W, "tblStyle")).get(_q(W, "val")),
            "TableGrid",
        )
        self.assertEqual(
            properties.find(_q(W, "jc")).get(_q(W, "val")),
            "right",
        )
        margins = properties.find(_q(W, "tblCellMar"))
        assert margins is not None
        self.assertEqual(
            margins.find(_q(W, "left")).get(_q(W, "w")),
            "120",
        )
        self.assertEqual(
            native_table.findall(
                f"./{_q(W, 'tblGrid')}/{_q(W, 'gridCol')}"
            )[1].get(_q(W, "w")),
            "2880",
        )
        growth_cell = native_table.findall(_q(W, "tr"))[1].findall(
            _q(W, "tc")
        )[1]
        growth_properties = growth_cell.find(_q(W, "tcPr"))
        assert growth_properties is not None
        self.assertEqual(
            growth_properties.find(_q(W, "vAlign")).get(
                _q(W, "val")
            ),
            "center",
        )
        self.assertEqual(
            growth_properties.find(_q(W, "shd")).get(
                _q(W, "fill")
            ),
            "E2F0D9",
        )
        self.assertIn(
            "Audited evidence source / appendix",
            "".join(
                text.text or ""
                for text in native_table.iter(_q(W, "t"))
            ),
        )
        internal_link = next(
            hyperlink
            for hyperlink in native_table.iter(_q(W, "hyperlink"))
            if hyperlink.get(_q(W, "anchor")) is not None
        )
        self.assertEqual(
            internal_link.get(_q(W, "anchor")),
            "appendix",
        )
        external_link = next(
            hyperlink
            for hyperlink in native_table.iter(_q(W, "hyperlink"))
            if hyperlink.get(_q(R, "id")) is not None
        )
        hyperlink_relationships = [
            relationship
            for relationship in after_relationships.findall(
                _q(REL, "Relationship")
            )
            if relationship.get("Type")
            == HYPERLINK_RELATIONSHIP_TYPE
        ]
        self.assertEqual(len(hyperlink_relationships), 1)
        self.assertEqual(
            hyperlink_relationships[0].get("Target"),
            "https://example.com/evidence",
        )
        self.assertEqual(
            hyperlink_relationships[0].get("TargetMode"),
            "External",
        )
        self.assertEqual(
            external_link.get(_q(R, "id")),
            hyperlink_relationships[0].get("Id"),
        )
        self.assertTrue(
            existing_relationships.issubset(
                {
                    tuple(sorted(relationship.attrib.items()))
                    for relationship in after_relationships.findall(
                        _q(REL, "Relationship")
                    )
                }
            )
        )
        para_ids = [
            paragraph.get(
                "{http://schemas.microsoft.com/office/"
                "word/2010/wordml}paraId"
            )
            for paragraph in after_root.iter(_q(W, "p"))
            if paragraph.get(
                "{http://schemas.microsoft.com/office/"
                "word/2010/wordml}paraId"
            )
            is not None
        ]
        self.assertEqual(len(para_ids), len(set(para_ids)))

        reopened = Document.from_docx(output)
        reopened_table = next(
            node
            for node in reopened.to_spec()["content"]
            if node["id"] == "metrics_table"
        )
        self.assertEqual(
            [
                column["id"]
                for column in reopened_table["columns"]
            ],
            [
                column["id"]
                for column in inserted_table["columns"]
            ],
        )
        self.assertEqual(
            [
                row["id"]
                for row in reopened_table["rows"]
            ],
            [
                row["id"]
                for row in inserted_table["rows"]
            ],
        )
        self.assertEqual(
            [
                cell["id"]
                for row in reopened_table["rows"]
                for cell in row["cells"]
            ],
            [
                cell["id"]
                for row in inserted_table["rows"]
                for cell in row["cells"]
            ],
        )
        self.assertEqual(
            reopened_table["rows"][1]["cells"][0]["content"][0]["id"],
            "evidence_paragraph",
        )

    def test_native_merged_tables_support_all_insertion_positions(
        self,
    ) -> None:
        source = (
            DocumentBuilder(
                sections=[
                    {"id": "front", "start_at": None},
                    {
                        "id": "body_section",
                        "start_at": "body",
                        "layout": {"start_type": "continuous"},
                    },
                ]
            )
            .paragraph("Cover", id="cover")
            .paragraph("Body", id="body")
            .paragraph("Conclusion", id="conclusion")
            .build()
            .to_bytes("docx")
        )
        before_root = parse_xml(
            ZipFile(io.BytesIO(source)).read("word/document.xml")
        )
        before_body = before_root.find(_q(W, "body"))
        assert before_body is not None
        terminal_section = ET.tostring(list(before_body)[-1])

        def merged_table(table_id: str, label: str) -> dict[str, object]:
            return {
                "id": table_id,
                "type": "table",
                "columns": [
                    {
                        "id": f"{table_id}_left",
                        "key": "left",
                        "title": "Left",
                    },
                    {
                        "id": f"{table_id}_right",
                        "key": "right",
                        "title": "Right",
                    },
                ],
                "rows": [
                    {
                        "id": f"{table_id}_row",
                        "cells": [
                            {
                                "id": f"{table_id}_cell",
                                "column_key": "left",
                                "column_span": 2,
                                "content": [
                                    {
                                        "id": f"{table_id}_paragraph",
                                        "type": "paragraph",
                                        "text": label,
                                    }
                                ],
                            }
                        ],
                    }
                ],
                "layout": {
                    "style_ref": "TableGrid",
                    "algorithm": "fixed",
                },
            }

        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "node.insert_after",
                    "target": "#cover",
                    "content": merged_table(
                        "after_cover_table",
                        "After cover",
                    ),
                },
                {
                    "op": "node.insert_before",
                    "target": "#body",
                    "content": merged_table(
                        "body_table",
                        "New section start",
                    ),
                },
                {
                    "op": "node.append",
                    "target": "$",
                    "content": merged_table(
                        "final_table",
                        "Final table",
                    ),
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(document.to_bytes("docx"), source)
        assert result.document is not None
        result_spec = result.document.to_spec()
        self.assertEqual(
            [node["id"] for node in result_spec["content"]],
            [
                "cover",
                "after_cover_table",
                "body_table",
                "body",
                "conclusion",
                "final_table",
            ],
        )
        self.assertEqual(
            result_spec["sections"][1]["start_at"],
            "body_table",
        )
        self.assertEqual(
            result.changes[1]["section_start_updated"],
            {
                "section_id": "body_section",
                "from": "body",
                "to": "body_table",
            },
        )

        output = result.document.to_bytes("docx")
        output_root = parse_xml(
            ZipFile(io.BytesIO(output)).read("word/document.xml")
        )
        output_body = output_root.find(_q(W, "body"))
        assert output_body is not None
        self.assertEqual(
            ET.tostring(list(output_body)[-1]),
            terminal_section,
        )
        inserted_tables = [
            node
            for node in result_spec["content"]
            if node["type"] == "table"
        ]
        self.assertEqual(len(inserted_tables), 3)
        for table in inserted_tables:
            self.assertEqual(
                table["source_ref"]["native_kind"],
                "w:tbl",
            )
            native_table = list(output_body)[
                table["source_ref"]["element_index"]
            ]
            self.assertEqual(native_table.tag, _q(W, "tbl"))
            merged_cell = native_table.findall(_q(W, "tr"))[1].find(
                _q(W, "tc")
            )
            assert merged_cell is not None
            self.assertEqual(
                merged_cell.find(
                    f"./{_q(W, 'tcPr')}/{_q(W, 'gridSpan')}"
                ).get(_q(W, "val")),
                "2",
            )
        reopened = Document.from_docx(output)
        self.assertEqual(
            [
                (node["id"], node["type"])
                for node in reopened.to_spec()["content"]
            ],
            [
                ("cover", "paragraph"),
                ("after_cover_table", "table"),
                ("body_table", "table"),
                ("body", "paragraph"),
                ("conclusion", "paragraph"),
                ("final_table", "table"),
            ],
        )
        self.assertEqual(
            reopened.to_spec()["sections"][1]["start_at"],
            "body_table",
        )

    def test_native_table_insert_rejects_forged_refs_and_missing_style(
        self,
    ) -> None:
        source = (
            DocumentBuilder()
            .paragraph("Anchor", id="anchor")
            .build()
            .to_bytes("docx")
        )
        document = Document.from_docx(source)
        anchor_ref = document.to_spec()["content"][0]["source_ref"]
        forged = document.apply(
            [
                {
                    "op": "node.insert_after",
                    "target": "#anchor",
                    "content": {
                        "id": "forged_table",
                        "type": "table",
                        "columns": [
                            {
                                "id": "forged_column",
                                "key": "value",
                                "title": "Value",
                                "source_ref": anchor_ref,
                            }
                        ],
                        "rows": [
                            {
                                "id": "forged_row",
                                "cells": [
                                    {
                                        "id": "forged_cell",
                                        "column_key": "value",
                                        "value": "Unsafe",
                                    }
                                ],
                            }
                        ],
                    },
                }
            ]
        )
        self.assertFalse(forged.success)
        self.assertEqual(
            forged.diagnostics[0].code,
            "NATIVE_PATCH_FAILED",
        )
        self.assertIn(
            "cannot claim existing native source references: "
            "forged_column",
            forged.diagnostics[0].message,
        )
        self.assertEqual(document.to_bytes("docx"), source)

        with ZipFile(io.BytesIO(source)) as package:
            styles_root = parse_xml(package.read("word/styles.xml"))
        for style in list(styles_root):
            if (
                style.tag == _q(W, "style")
                and style.get(_q(W, "type")) == "table"
                and style.get(_q(W, "styleId")) == "TableGrid"
            ):
                styles_root.remove(style)
        without_table_grid = _rewrite_package(
            source,
            {"word/styles.xml": serialize_xml(styles_root)},
        )
        missing_style_document = Document.from_docx(
            without_table_grid
        )
        missing_style = missing_style_document.apply(
            [
                {
                    "op": "node.append",
                    "target": "$",
                    "content": {
                        "id": "unstyled_table",
                        "type": "table",
                        "columns": [
                            {
                                "id": "unstyled_column",
                                "key": "value",
                                "title": "Value",
                            }
                        ],
                        "rows": [
                            {
                                "id": "unstyled_row",
                                "cells": [
                                    {
                                        "id": "unstyled_cell",
                                        "column_key": "value",
                                        "value": "Safe failure",
                                    }
                                ],
                            }
                        ],
                    },
                }
            ]
        )
        self.assertFalse(missing_style.success)
        self.assertEqual(
            missing_style.diagnostics[0].code,
            "NATIVE_PATCH_FAILED",
        )
        self.assertIn(
            "requires exactly one table style 'TableGrid'",
            missing_style.diagnostics[0].message,
        )
        self.assertEqual(
            missing_style_document.to_bytes("docx"),
            without_table_grid,
        )

    def test_native_inserted_tables_can_move_and_remove_in_same_patch(
        self,
    ) -> None:
        source = (
            DocumentBuilder()
            .paragraph("Anchor", id="anchor")
            .paragraph("Tail", id="tail")
            .build()
            .to_bytes("docx")
        )

        def table_payload(table_id: str) -> dict[str, object]:
            return {
                "id": table_id,
                "type": "table",
                "columns": [
                    {
                        "id": f"{table_id}_column",
                        "key": "value",
                        "title": "Value",
                    }
                ],
                "rows": [
                    {
                        "id": f"{table_id}_row",
                        "cells": [
                            {
                                "id": f"{table_id}_cell",
                                "column_key": "value",
                                "value": table_id,
                            }
                        ],
                    }
                ],
            }

        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "node.insert_after",
                    "target": "#anchor",
                    "content": table_payload("movable_table"),
                },
                {
                    "op": "node.move_after",
                    "target": "#movable_table",
                    "after": "#tail",
                },
                {
                    "op": "node.insert_before",
                    "target": "#tail",
                    "content": table_payload("temporary_table"),
                },
                {
                    "op": "node.remove",
                    "target": "#temporary_table",
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(document.to_bytes("docx"), source)
        assert result.document is not None
        self.assertEqual(
            [
                (node["id"], node["type"])
                for node in result.document.to_spec()["content"]
            ],
            [
                ("anchor", "paragraph"),
                ("tail", "paragraph"),
                ("movable_table", "table"),
            ],
        )
        output = result.document.to_bytes("docx")
        output_root = parse_xml(
            ZipFile(io.BytesIO(output)).read("word/document.xml")
        )
        output_body = output_root.find(_q(W, "body"))
        assert output_body is not None
        self.assertEqual(
            [element.tag for element in list(output_body)],
            [
                _q(W, "p"),
                _q(W, "p"),
                _q(W, "tbl"),
                _q(W, "sectPr"),
            ],
        )
        reopened = Document.from_docx(output)
        self.assertEqual(
            [
                node["id"]
                for node in reopened.to_spec()["content"]
            ],
            ["anchor", "tail", "movable_table"],
        )

    def test_native_inserted_nodes_are_batch_addressable(self) -> None:
        source = (
            DocumentBuilder()
            .paragraph("Anchor", id="anchor")
            .paragraph("Tail", id="tail")
            .build()
            .to_bytes("docx")
        )
        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "node.insert_after",
                    "target": "#anchor",
                    "content": {
                        "id": "new_body",
                        "type": "paragraph",
                        "text": "Alpha",
                    },
                },
                {
                    "op": "node.insert_after",
                    "target": "#new_body",
                    "content": {
                        "id": "new_heading",
                        "type": "heading",
                        "level": 2,
                        "text": "New section",
                    },
                },
                {
                    "op": "text.replace",
                    "target": "#new_body",
                    "search": "Alpha",
                    "replacement": "Revised",
                },
                {
                    "op": "paragraph.format",
                    "target": "#new_body",
                    "set": {"alignment": "right"},
                },
                {
                    "op": "node.move_before",
                    "target": "#new_heading",
                    "before": "#new_body",
                },
                {
                    "op": "node.insert_after",
                    "target": "#new_body",
                    "content": {
                        "id": "temporary",
                        "type": "paragraph",
                        "text": "Temporary",
                    },
                },
                {
                    "op": "node.remove",
                    "target": "#temporary",
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(document.to_bytes("docx"), source)
        assert result.document is not None
        self.assertEqual(
            [
                node["id"]
                for node in result.document.to_spec()["content"]
            ],
            ["anchor", "new_heading", "new_body", "tail"],
        )
        self.assertNotIn(
            "temporary",
            {
                node["id"]
                for node in result.document.to_spec()["content"]
            },
        )
        output = result.document.to_bytes("docx")
        reopened = Document.from_docx(output)
        reopened_nodes = reopened.to_spec()["content"]
        self.assertEqual(
            [node["id"] for node in reopened_nodes],
            ["anchor", "new_heading", "new_body", "tail"],
        )
        self.assertEqual(reopened_nodes[1]["type"], "heading")
        self.assertEqual(reopened_nodes[1]["level"], 2)
        self.assertEqual(_semantic_text(reopened_nodes[1]), "New section")
        self.assertEqual(_semantic_text(reopened_nodes[2]), "Revised")
        self.assertEqual(
            reopened_nodes[2]["paragraph_style"]["alignment"],
            "right",
        )
        self.assertEqual(
            [
                node["source_ref"]["element_index"]
                for node in reopened_nodes
            ],
            [0, 1, 2, 3],
        )

    def test_native_insert_before_handles_head_and_multi_element_anchor(
        self,
    ) -> None:
        source = (
            DocumentBuilder()
            .paragraph("Before", id="before")
            .bullet_list(["One", "Two"], id="steps")
            .paragraph("After", id="after")
            .build()
            .to_bytes("docx")
        )
        document = Document.from_docx(source)
        before_root = parse_xml(
            ZipFile(io.BytesIO(source)).read("word/document.xml")
        )
        before_body = before_root.find(_q(W, "body"))
        assert before_body is not None
        original_payloads = {
            node["id"]: [
                ET.tostring(list(before_body)[index])
                for index in node["source_ref"]["element_indices"]
            ]
            for node in document.to_spec()["content"]
        }
        result = document.apply(
            [
                {
                    "op": "node.insert_before",
                    "target": "#before",
                    "content": {
                        "id": "document_prelude",
                        "type": "heading",
                        "level": 2,
                        "text": "Document prelude",
                    },
                },
                {
                    "op": "node.insert_before",
                    "target": "#steps",
                    "content": {
                        "id": "list_intro",
                        "type": "paragraph",
                        "text": "List intro",
                    },
                },
                {
                    "op": "node.insert_before",
                    "target": "#list_intro",
                    "content": {
                        "id": "list_label",
                        "type": "paragraph",
                        "text": "List label",
                    },
                },
                {
                    "op": "paragraph.format",
                    "target": "#list_label",
                    "set": {"keep_with_next": True},
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(document.to_bytes("docx"), source)
        assert result.document is not None
        self.assertEqual(
            [
                node["id"]
                for node in result.document.to_spec()["content"]
            ],
            [
                "document_prelude",
                "before",
                "list_label",
                "list_intro",
                "steps",
                "after",
            ],
        )
        output = result.document.to_bytes("docx")
        output_root = parse_xml(
            ZipFile(io.BytesIO(output)).read("word/document.xml")
        )
        output_body = output_root.find(_q(W, "body"))
        assert output_body is not None
        for node in result.document.to_spec()["content"]:
            if node["id"] not in original_payloads:
                continue
            self.assertEqual(
                [
                    ET.tostring(list(output_body)[index])
                    for index in node["source_ref"]["element_indices"]
                ],
                original_payloads[node["id"]],
                node["id"],
            )
        reopened = Document.from_docx(output)
        reopened_nodes = reopened.to_spec()["content"]
        self.assertEqual(
            [node["id"] for node in reopened_nodes],
            [
                "document_prelude",
                "before",
                "list_label",
                "list_intro",
                "steps",
                "after",
            ],
        )
        self.assertEqual(
            reopened_nodes[2]["paragraph_style"]["keep_with_next"],
            True,
        )
        self.assertEqual(
            reopened_nodes[4]["source_ref"]["element_indices"],
            [4, 5],
        )

    def test_identity_map_is_refreshed_after_removal(self) -> None:
        document = Document.from_docx(self._source_document())
        nodes = document.to_spec()["content"]
        removed = document.apply([{"op": "node.remove", "target": f"#{nodes[0]['id']}"}])
        self.assertTrue(removed.success)
        assert removed.document is not None

        remaining = removed.document.to_spec()["content"][0]
        self.assertEqual(remaining["source_ref"]["element_index"], 0)
        edited = removed.document.apply(
            [
                {
                    "op": "text.replace",
                    "target": f"#{remaining['id']}",
                    "search": "Beta",
                    "replacement": "Delta",
                }
            ]
        )
        self.assertTrue(edited.success)
        assert edited.document is not None

        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "twice-patched.docx"
            edited.document.export(target)
            reopened = Document.from_docx(target)
            texts = [
                _semantic_text(node)
                for node in reopened.to_spec()["content"]
                if node["type"] in {"heading", "paragraph"}
            ]
            self.assertIn("Alpha Delta Gamma", texts)

    def test_third_party_remove_attaches_identity_and_preserves_payloads(
        self,
    ) -> None:
        source = (
            DocumentBuilder()
            .paragraph("A", id="a")
            .paragraph("B", id="b")
            .paragraph("C", id="c")
            .build()
            .to_bytes("docx")
        )
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
                and override.get("PartName") == MANIFEST_PART_URI
            ):
                content_types.remove(override)
        source = _rewrite_package(
            source,
            {
                "_rels/.rels": serialize_xml(root_relationships),
                "[Content_Types].xml": serialize_xml(content_types),
            },
            deletions={MANIFEST_PART_URI.lstrip("/")},
        )
        document = Document.from_docx(source)
        before_spec = document.to_spec()
        ids = {
            _semantic_text(node): str(node["id"])
            for node in before_spec["content"]
        }
        detached = Document.from_spec(before_spec)
        self.assertNotIn(
            "node.append",
            detached.capabilities()["operations"],
        )
        self.assertNotIn(
            "node.insert_after",
            detached.capabilities()["operations"],
        )
        self.assertNotIn(
            "node.insert_before",
            detached.capabilities()["operations"],
        )
        self.assertNotIn(
            "node.remove",
            detached.capabilities()["operations"],
        )
        detached_insert = detached.apply(
            [
                {
                    "op": "node.insert_after",
                    "target": ids["A"],
                    "content": {
                        "id": "detached_new",
                        "type": "paragraph",
                        "text": "Detached",
                    },
                }
            ]
        )
        self.assertFalse(detached_insert.success)
        self.assertEqual(
            detached_insert.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )
        detached_insert_before = detached.apply(
            [
                {
                    "op": "node.insert_before",
                    "target": ids["A"],
                    "content": {
                        "id": "detached_before",
                        "type": "paragraph",
                        "text": "Detached before",
                    },
                }
            ]
        )
        self.assertFalse(detached_insert_before.success)
        self.assertEqual(
            detached_insert_before.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )
        detached_append = detached.apply(
            [
                {
                    "op": "node.append",
                    "target": "$",
                    "content": {
                        "id": "detached_append",
                        "type": "paragraph",
                        "text": "Detached append",
                    },
                }
            ]
        )
        self.assertFalse(detached_append.success)
        self.assertEqual(
            detached_append.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )
        detached_result = detached.apply(
            [{"op": "node.remove", "target": ids["B"]}]
        )
        self.assertFalse(detached_result.success)
        self.assertEqual(
            detached_result.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )

        before_root = parse_xml(
            ZipFile(io.BytesIO(source)).read("word/document.xml")
        )
        before_body = before_root.find(_q(W, "body"))
        assert before_body is not None
        before_payloads = {
            str(node["id"]): ET.tostring(
                list(before_body)[
                    int(node["source_ref"]["element_index"])
                ]
            )
            for node in before_spec["content"]
        }
        result = document.apply(
            [{"op": "node.remove", "target": ids["B"]}]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(document.to_bytes("docx"), source)
        assert result.diff is not None
        self.assertEqual(result.diff.summary["removed"], 1)
        self.assertEqual(result.diff.summary["moved"], 0)
        self.assertNotIn(
            "content.order",
            [entry.path for entry in result.diff.entries],
        )
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
        self.assertEqual(
            [
                node["id"]
                for node in result.document.to_spec()["content"]
            ],
            [ids["A"], ids["C"]],
        )
        output = result.document.to_bytes("docx")
        with (
            ZipFile(io.BytesIO(source)) as before,
            ZipFile(io.BytesIO(output)) as after,
        ):
            for name in before.namelist():
                if name not in {
                    "[Content_Types].xml",
                    "_rels/.rels",
                    "word/document.xml",
                }:
                    self.assertEqual(after.read(name), before.read(name), name)
            self.assertIn(
                MANIFEST_PART_URI.lstrip("/"),
                after.namelist(),
            )
            after_root = parse_xml(after.read("word/document.xml"))
        after_body = after_root.find(_q(W, "body"))
        assert after_body is not None
        for node in result.document.to_spec()["content"]:
            self.assertEqual(
                ET.tostring(
                    list(after_body)[
                        int(node["source_ref"]["element_index"])
                    ]
                ),
                before_payloads[str(node["id"])],
            )
        reopened = Document.from_docx(output)
        self.assertEqual(
            [node["id"] for node in reopened.to_spec()["content"]],
            [ids["A"], ids["C"]],
        )

    def test_native_move_after_tracks_objects_across_sequential_moves(
        self,
    ) -> None:
        source = (
            DocumentBuilder()
            .paragraph("A", id="a")
            .paragraph("B", id="b")
            .paragraph("C", id="c")
            .paragraph("D", id="d")
            .build()
            .to_bytes("docx")
        )
        document = Document.from_docx(source)
        before_spec = document.to_spec()
        detached = Document.from_spec(before_spec)
        self.assertNotIn(
            "node.move_after",
            detached.capabilities()["operations"],
        )
        self.assertNotIn(
            "node.move_before",
            detached.capabilities()["operations"],
        )
        self.assertFalse(
            detached.capabilities()["structural_editing"]["available"]
        )
        detached_move = detached.apply(
            [
                {
                    "op": "node.move_before",
                    "target": "#c",
                    "before": "#a",
                }
            ]
        )
        self.assertFalse(detached_move.success)
        self.assertEqual(
            detached_move.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )
        self.assertIn(
            "attached native DOCX package",
            detached_move.diagnostics[0].message,
        )
        before_root = parse_xml(
            ZipFile(io.BytesIO(source)).read("word/document.xml")
        )
        before_body = before_root.find(_q(W, "body"))
        assert before_body is not None
        before_payloads = {
            node["id"]: ET.tostring(
                list(before_body)[
                    node["source_ref"]["element_index"]
                ]
            )
            for node in before_spec["content"]
        }

        result = document.apply(
            [
                {
                    "op": "node.move_after",
                    "target": "#a",
                    "after": "#c",
                },
                {
                    "op": "node.move_before",
                    "target": "#d",
                    "before": "#b",
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(document.to_bytes("docx"), source)
        self.assertEqual(
            result.fidelity.affected_parts if result.fidelity else None,
            ["/customXml/aioffice-manifest.xml", "/word/document.xml"],
        )
        assert result.document is not None
        self.assertEqual(
            [
                node["id"]
                for node in result.document.to_spec()["content"]
            ],
            ["d", "b", "c", "a"],
        )
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
                    self.assertEqual(after.read(name), before.read(name), name)
            after_root = parse_xml(after.read("word/document.xml"))
        after_body = after_root.find(_q(W, "body"))
        assert after_body is not None
        after_spec = result.document.to_spec()
        for node in after_spec["content"]:
            self.assertEqual(
                ET.tostring(
                    list(after_body)[
                        node["source_ref"]["element_index"]
                    ]
                ),
                before_payloads[node["id"]],
                node["id"],
            )

        reopened = Document.from_docx(output)
        reopened_spec = reopened.to_spec()
        self.assertEqual(
            [node["id"] for node in reopened_spec["content"]],
            ["d", "b", "c", "a"],
        )
        self.assertEqual(
            [
                node["source_ref"]["element_index"]
                for node in reopened_spec["content"]
            ],
            [0, 1, 2, 3],
        )

    def test_native_move_keeps_multi_paragraph_list_contiguous(
        self,
    ) -> None:
        source = (
            DocumentBuilder()
            .paragraph("Before", id="before")
            .bullet_list(["One", "Two"], id="steps")
            .paragraph("Middle", id="middle")
            .paragraph("After", id="after")
            .build()
            .to_bytes("docx")
        )
        document = Document.from_docx(source)
        steps = next(
            node
            for node in document.to_spec()["content"]
            if node["id"] == "steps"
        )
        before_root = parse_xml(
            ZipFile(io.BytesIO(source)).read("word/document.xml")
        )
        before_body = before_root.find(_q(W, "body"))
        assert before_body is not None
        list_payloads = [
            ET.tostring(list(before_body)[index])
            for index in steps["source_ref"]["element_indices"]
        ]

        result = document.apply(
            [
                {
                    "op": "node.move_before",
                    "target": "#steps",
                    "before": "#before",
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        self.assertEqual(
            [
                node["id"]
                for node in result.document.to_spec()["content"]
            ],
            ["steps", "before", "middle", "after"],
        )
        moved_steps = next(
            node
            for node in result.document.to_spec()["content"]
            if node["id"] == "steps"
        )
        self.assertEqual(
            moved_steps["source_ref"]["element_indices"],
            [0, 1],
        )
        output = result.document.to_bytes("docx")
        output_root = parse_xml(
            ZipFile(io.BytesIO(output)).read("word/document.xml")
        )
        output_body = output_root.find(_q(W, "body"))
        assert output_body is not None
        self.assertEqual(
            [
                ET.tostring(list(output_body)[index])
                for index in moved_steps["source_ref"]["element_indices"]
            ],
            list_payloads,
        )
        reopened = Document.from_docx(output)
        self.assertEqual(
            [
                node["id"]
                for node in reopened.to_spec()["content"]
            ],
            ["steps", "before", "middle", "after"],
        )
        reopened_steps = reopened.to_spec()["content"][0]
        self.assertEqual(
            reopened_steps["source_ref"]["element_indices"],
            [0, 1],
        )

    def test_native_move_after_protects_section_carriers(self) -> None:
        source = (
            DocumentBuilder(
                sections=[
                    {"id": "front", "start_at": None},
                    {
                        "id": "body_section",
                        "start_at": "body",
                        "layout": {"start_type": "next_page"},
                    },
                ]
            )
            .paragraph("Cover", id="cover")
            .paragraph("Front end", id="front_end")
            .paragraph("Body", id="body")
            .paragraph("Analysis", id="analysis")
            .paragraph("Conclusion", id="conclusion")
            .build()
            .to_bytes("docx")
        )
        document_root = parse_xml(
            ZipFile(io.BytesIO(source)).read("word/document.xml")
        )
        body = document_root.find(_q(W, "body"))
        assert body is not None
        body_paragraphs = body.findall(_q(W, "p"))
        carrier = next(
            paragraph
            for paragraph in body_paragraphs
            if paragraph.find(f".//{_q(W, 'sectPr')}") is not None
        )
        front_end_paragraph = next(
            paragraph
            for paragraph in body_paragraphs
            if "".join(
                node.text or ""
                for node in paragraph.iter(_q(W, "t"))
            )
            == "Front end"
        )
        carrier_properties = carrier.find(_q(W, "pPr"))
        assert carrier_properties is not None
        section_properties = carrier_properties.find(_q(W, "sectPr"))
        assert section_properties is not None
        carrier_properties.remove(section_properties)
        front_properties = front_end_paragraph.find(_q(W, "pPr"))
        if front_properties is None:
            front_properties = ET.Element(_q(W, "pPr"))
            front_end_paragraph.insert(0, front_properties)
        front_properties.append(section_properties)
        body.remove(carrier)

        root_relationships = parse_xml(
            ZipFile(io.BytesIO(source)).read("_rels/.rels")
        )
        for relationship in list(root_relationships):
            if (
                relationship.get("Type")
                == MANIFEST_RELATIONSHIP_TYPE
            ):
                root_relationships.remove(relationship)
        source = _rewrite_package(
            source,
            {
                "word/document.xml": serialize_xml(document_root),
                "_rels/.rels": serialize_xml(root_relationships),
            },
            deletions={"customXml/aioffice-manifest.xml"},
        )
        document = Document.from_docx(source)
        body_section_id = document.to_spec()["sections"][1]["id"]
        ids = {
            _semantic_text(node): str(node["id"])
            for node in document.to_spec()["content"]
        }
        carrier_move = document.apply(
            [
                {
                    "op": "node.move_after",
                    "target": ids["Cover"],
                    "after": ids["Front end"],
                }
            ]
        )
        self.assertFalse(carrier_move.success)
        self.assertEqual(
            carrier_move.diagnostics[0].code,
            "NATIVE_PATCH_FAILED",
        )
        self.assertIn(
            "section boundary",
            carrier_move.diagnostics[0].message,
        )
        cross_section = document.apply(
            [
                {
                    "op": "node.move_after",
                    "target": ids["Cover"],
                    "after": ids["Analysis"],
                }
            ]
        )
        self.assertFalse(cross_section.success)
        self.assertEqual(
            cross_section.diagnostics[0].code,
            "CROSS_SECTION_MOVE_UNSUPPORTED",
        )
        anchor_move = document.apply(
            [
                {
                    "op": "node.move_after",
                    "target": ids["Body"],
                    "after": ids["Analysis"],
                }
            ]
        )
        self.assertFalse(anchor_move.success)
        self.assertEqual(
            anchor_move.diagnostics[0].code,
            "UNSUPPORTED_FEATURE",
        )
        carrier_remove = document.apply(
            [
                {
                    "op": "node.remove",
                    "target": ids["Front end"],
                }
            ]
        )
        self.assertFalse(carrier_remove.success)
        self.assertEqual(
            carrier_remove.diagnostics[0].code,
            "NATIVE_PATCH_FAILED",
        )
        self.assertIn(
            "section boundary",
            carrier_remove.diagnostics[0].message,
        )
        safe_insert = document.apply(
            [
                {
                    "op": "node.insert_before",
                    "target": ids["Front end"],
                    "content": {
                        "id": "front_note",
                        "type": "paragraph",
                        "text": "Front note",
                    },
                },
                {
                    "op": "node.insert_before",
                    "target": ids["Body"],
                    "content": {
                        "id": "body_preface",
                        "type": "heading",
                        "level": 2,
                        "text": "Body preface",
                    },
                },
                {
                    "op": "node.insert_before",
                    "target": "#body_preface",
                    "content": {
                        "id": "body_label",
                        "type": "paragraph",
                        "text": "Body label",
                    },
                },
            ]
        )
        self.assertTrue(safe_insert.success, safe_insert.model_dump())
        self.assertEqual(document.to_bytes("docx"), source)
        assert safe_insert.document is not None
        self.assertEqual(
            safe_insert.document.to_spec()["sections"][1][
                "start_at"
            ],
            "body_label",
        )
        self.assertEqual(
            safe_insert.changes[1]["section_start_updated"],
            {
                "section_id": body_section_id,
                "from": ids["Body"],
                "to": "body_preface",
            },
        )
        self.assertEqual(
            safe_insert.changes[2]["section_start_updated"],
            {
                "section_id": body_section_id,
                "from": "body_preface",
                "to": "body_label",
            },
        )
        safe_output = safe_insert.document.to_bytes("docx")
        safe_root = parse_xml(
            ZipFile(io.BytesIO(safe_output)).read(
                "word/document.xml"
            )
        )
        safe_body = safe_root.find(_q(W, "body"))
        assert safe_body is not None
        safe_paragraphs = safe_body.findall(_q(W, "p"))
        safe_texts = [
            "".join(
                node.text or ""
                for node in paragraph.iter(_q(W, "t"))
            )
            for paragraph in safe_paragraphs
        ]
        self.assertLess(
            safe_texts.index("Front note"),
            safe_texts.index("Front end"),
        )
        self.assertLess(
            safe_texts.index("Front end"),
            safe_texts.index("Body label"),
        )
        self.assertLess(
            safe_texts.index("Body label"),
            safe_texts.index("Body preface"),
        )
        self.assertLess(
            safe_texts.index("Body preface"),
            safe_texts.index("Body"),
        )
        boundary_paragraph = safe_paragraphs[
            safe_texts.index("Front end")
        ]
        self.assertIsNotNone(
            boundary_paragraph.find(f".//{_q(W, 'sectPr')}")
        )
        safe_reopened = Document.from_docx(safe_output)
        self.assertEqual(
            safe_reopened.to_spec()["sections"][1]["start_at"],
            "body_label",
        )

        carrier_insert = document.apply(
            [
                {
                    "op": "node.insert_after",
                    "target": ids["Front end"],
                    "content": {
                        "id": "unsafe_insert",
                        "type": "paragraph",
                        "text": "Wrong section",
                    },
                }
            ]
        )
        self.assertFalse(carrier_insert.success)
        self.assertEqual(
            carrier_insert.diagnostics[0].code,
            "NATIVE_PATCH_FAILED",
        )
        self.assertIn(
            "section boundary",
            carrier_insert.diagnostics[0].message,
        )
        section_start_remove = document.apply(
            [{"op": "node.remove", "target": ids["Body"]}]
        )
        self.assertFalse(section_start_remove.success)
        self.assertEqual(document.to_bytes("docx"), source)

        successful = document.apply(
            [
                {
                    "op": "node.move_before",
                    "target": ids["Conclusion"],
                    "before": ids["Body"],
                }
            ]
        )
        self.assertTrue(successful.success, successful.model_dump())
        assert successful.document is not None
        self.assertEqual(
            [
                node["id"]
                for node in successful.document.to_spec()["content"]
            ],
            [
                ids["Cover"],
                ids["Front end"],
                ids["Conclusion"],
                ids["Body"],
                ids["Analysis"],
            ],
        )
        self.assertEqual(
            [
                section.get("start_at")
                for section in successful.document.to_spec()["sections"]
            ],
            [None, ids["Conclusion"]],
        )

    def test_native_format_patch_changes_only_known_properties(self) -> None:
        source = self._source_document()
        document_xml = parse_xml(ZipFile(io.BytesIO(source)).read("word/document.xml"))
        body = document_xml.find(_q(W, "body"))
        assert body is not None
        paragraph = next(
            element
            for element in body.findall(_q(W, "p"))
            if "Alpha Beta Gamma" == "".join(node.text or "" for node in element.iter(_q(W, "t")))
        )
        properties = paragraph.find(_q(W, "pPr"))
        if properties is None:
            properties = ET.Element(_q(W, "pPr"))
            paragraph.insert(0, properties)
        future = ET.SubElement(properties, "{urn:aioffice:test}futureLayout")
        future.set("mode", "preserve")
        source = _rewrite_package(
            source,
            {"word/document.xml": serialize_xml(document_xml)},
        )

        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "paragraph.format",
                    "target": "#body",
                    "set": {
                        "alignment": "center",
                        "spacing_before": {"value": 12, "unit": "pt"},
                    },
                },
                {
                    "op": "text.format",
                    "target": "#body",
                    "set": {
                        "font_size": {"value": 13, "unit": "pt"},
                        "color": "#C00000",
                    },
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        output = result.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as archive:
            patched_root = parse_xml(archive.read("word/document.xml"))
            patched_body = patched_root.find(_q(W, "body"))
            assert patched_body is not None
            patched = next(
                element
                for element in patched_body.findall(_q(W, "p"))
                if "Alpha Beta Gamma"
                == "".join(node.text or "" for node in element.iter(_q(W, "t")))
            )
            patched_properties = patched.find(_q(W, "pPr"))
            assert patched_properties is not None
            self.assertIsNotNone(patched_properties.find("{urn:aioffice:test}futureLayout"))
            alignment = patched_properties.find(_q(W, "jc"))
            spacing = patched_properties.find(_q(W, "spacing"))
            assert alignment is not None
            assert spacing is not None
            self.assertEqual(alignment.attrib[_q(W, "val")], "center")
            self.assertEqual(spacing.attrib[_q(W, "before")], "240")
            for run in patched.iter(_q(W, "r")):
                run_properties = run.find(_q(W, "rPr"))
                assert run_properties is not None
                size = run_properties.find(_q(W, "sz"))
                color = run_properties.find(_q(W, "color"))
                assert size is not None
                assert color is not None
                self.assertEqual(size.attrib[_q(W, "val")], "26")
                self.assertEqual(color.attrib[_q(W, "val")], "C00000")

        reopened = Document.from_docx(output)
        reopened_body = next(node for node in reopened.to_spec()["content"] if node["id"] == "body")
        self.assertEqual(reopened_body["paragraph_style"]["alignment"], "center")
        self.assertEqual(reopened_body["text_style"]["font_size"]["value"], 13.0)
        self.assertEqual(reopened_body["text_style"]["color"], "#C00000")

    def test_format_then_remove_in_one_native_patch_is_atomic(self) -> None:
        document = Document.from_docx(self._source_document())
        result = document.apply(
            [
                {
                    "op": "paragraph.format",
                    "target": "#body",
                    "set": {"alignment": "right"},
                },
                {"op": "node.remove", "target": "#body"},
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        reopened = Document.from_docx(result.document.to_bytes("docx"))
        self.assertNotIn(
            "Alpha Beta Gamma",
            [
                _semantic_text(node)
                for node in reopened.to_spec()["content"]
                if node["type"] in {"heading", "paragraph"}
            ],
        )

    def test_empty_native_paragraph_keeps_text_style_on_paragraph_mark(self) -> None:
        source = compile_docx(DocumentBuilder().paragraph("", id="empty").build().spec)
        document_xml = parse_xml(ZipFile(io.BytesIO(source)).read("word/document.xml"))
        body = document_xml.find(_q(W, "body"))
        assert body is not None
        paragraph = body.find(_q(W, "p"))
        assert paragraph is not None
        for run in list(paragraph.findall(_q(W, "r"))):
            paragraph.remove(run)
        source = _rewrite_package(
            source,
            {"word/document.xml": serialize_xml(document_xml)},
        )

        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "text.format",
                    "target": "#empty",
                    "set": {
                        "font_size": {"value": 11, "unit": "pt"},
                        "color": "#112233",
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        reopened = Document.from_docx(result.document.to_bytes("docx"))
        node = next(node for node in reopened.to_spec()["content"] if node["id"] == "empty")
        self.assertEqual(node["text_style"]["font_size"]["value"], 11.0)
        self.assertEqual(node["text_style"]["color"], "#112233")

    def test_native_range_format_splits_cross_run_selection_exactly(self) -> None:
        document = Document.from_docx(self._source_document())
        result = document.apply(
            [
                {
                    "op": "text.format",
                    "target": "#body",
                    "match": {"text": "ha Be"},
                    "set": {"color": "#FF0000"},
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        root = parse_xml(
            ZipFile(io.BytesIO(result.document.to_bytes("docx"))).read("word/document.xml")
        )
        body = root.find(_q(W, "body"))
        assert body is not None
        paragraph = next(
            element
            for element in body.findall(_q(W, "p"))
            if _semantic_text(
                {"text": "".join(text.text or "" for text in element.iter(_q(W, "t")))}
            )
            == "Alpha Beta Gamma"
        )
        runs = []
        for run in paragraph.iter(_q(W, "r")):
            text = "".join(node.text or "" for node in run.iter(_q(W, "t")))
            color = run.find(f"./{_q(W, 'rPr')}/{_q(W, 'color')}")
            runs.append(
                (
                    text,
                    color.attrib.get(_q(W, "val")) if color is not None else None,
                )
            )
        self.assertEqual(
            runs,
            [
                ("Alp", None),
                ("ha ", "FF0000"),
                ("Be", "FF0000"),
                ("ta", None),
                (" Gamma", None),
            ],
        )
        reopened = Document.from_docx(result.document.to_bytes("docx"))
        node = next(node for node in reopened.to_spec()["content"] if node["id"] == "body")
        self.assertEqual(_semantic_text(node), "Alpha Beta Gamma")
        self.assertEqual(
            [
                span["text"]
                for span in node["content"]
                if span.get("style", {}).get("color") == "#FF0000"
            ],
            ["ha ", "Be"],
        )

    def test_partial_range_refuses_complex_run_without_data_loss(self) -> None:
        source = self._source_document()
        document_xml = parse_xml(ZipFile(io.BytesIO(source)).read("word/document.xml"))
        body = document_xml.find(_q(W, "body"))
        assert body is not None
        paragraph = next(
            element
            for element in body.findall(_q(W, "p"))
            if "Alpha Beta Gamma" == "".join(node.text or "" for node in element.iter(_q(W, "t")))
        )
        first_run = next(paragraph.iter(_q(W, "r")))
        future = ET.SubElement(first_run, "{urn:aioffice:test}futureInline")
        future.text = "preserve"
        source = _rewrite_package(
            source,
            {"word/document.xml": serialize_xml(document_xml)},
        )

        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "text.format",
                    "target": "#body",
                    "range": {"start": 1, "end": 3},
                    "set": {"bold": True},
                }
            ]
        )
        self.assertFalse(result.success)
        self.assertEqual(result.diagnostics[0].code, "NATIVE_PATCH_FAILED")
        self.assertIn("complex native run", result.diagnostics[0].message)
        self.assertEqual(document.to_bytes("docx"), source)

    def test_hyperlink_projection_and_range_split_preserve_target(self) -> None:
        source = (
            DocumentBuilder()
            .rich_paragraph(
                [
                    {"text": "See "},
                    {
                        "text": "docs",
                        "marks": ["link"],
                        "href": "https://example.com/docs",
                    },
                    {"text": " now"},
                ],
                id="link_para",
            )
            .build()
            .to_bytes("docx")
        )
        document = Document.from_docx(source)
        projected = document.to_spec()["content"][0]
        link_span = next(
            span for span in projected["content"] if span.get("href") == "https://example.com/docs"
        )
        self.assertEqual(link_span["text"], "docs")
        self.assertEqual(link_span["marks"], ["link"])

        result = document.apply(
            [
                {
                    "op": "text.format",
                    "target": "#link_para",
                    "match": {"text": "oc"},
                    "set": {"bold": True},
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        reopened = Document.from_docx(result.document.to_bytes("docx"))
        spans = reopened.to_spec()["content"][0]["content"]
        linked = [span for span in spans if span.get("href") == "https://example.com/docs"]
        self.assertEqual([span["text"] for span in linked], ["d", "oc", "s"])
        self.assertEqual(linked[1]["style"]["bold"], True)
        self.assertTrue(all(span["marks"] == ["link"] for span in linked))

    def test_range_selection_uses_text_after_earlier_operation(self) -> None:
        document = Document.from_docx(self._source_document())
        result = document.apply(
            [
                {
                    "op": "text.replace",
                    "target": "#body",
                    "search": "Alpha",
                    "replacement": "A",
                },
                {
                    "op": "text.format",
                    "target": "#body",
                    "match": {"text": "Beta"},
                    "set": {"underline": True},
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        reopened = Document.from_docx(result.document.to_bytes("docx"))
        node = next(node for node in reopened.to_spec()["content"] if node["id"] == "body")
        self.assertEqual(_semantic_text(node), "A Beta Gamma")
        beta = next(span for span in node["content"] if span["text"] == "Beta")
        self.assertEqual(beta["style"]["underline"], True)

    def test_range_split_preserves_multiple_text_nodes_in_one_run(self) -> None:
        source = DocumentBuilder().paragraph("Alpha", id="multi_text").build().to_bytes("docx")
        document_xml = parse_xml(ZipFile(io.BytesIO(source)).read("word/document.xml"))
        body = document_xml.find(_q(W, "body"))
        assert body is not None
        paragraph = body.find(_q(W, "p"))
        assert paragraph is not None
        run = paragraph.find(_q(W, "r"))
        assert run is not None
        text = run.find(_q(W, "t"))
        assert text is not None
        text.text = "Al"
        second = copy.deepcopy(text)
        second.text = "pha"
        run.append(second)
        source = _rewrite_package(
            source,
            {"word/document.xml": serialize_xml(document_xml)},
        )

        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "text.format",
                    "target": "#multi_text",
                    "range": {"start": 1, "end": 4},
                    "set": {"italic": True},
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        root = parse_xml(
            ZipFile(io.BytesIO(result.document.to_bytes("docx"))).read("word/document.xml")
        )
        paragraph = next(root.iter(_q(W, "p")))
        runs = [
            (
                "".join(node.text or "" for node in run.iter(_q(W, "t"))),
                run.find(f"./{_q(W, 'rPr')}/{_q(W, 'i')}") is not None,
            )
            for run in paragraph.iter(_q(W, "r"))
        ]
        self.assertEqual(runs, [("A", False), ("lph", True), ("a", False)])

    def test_native_range_clear_removes_only_selected_direct_property(self) -> None:
        source = (
            DocumentBuilder()
            .paragraph(
                "ABCD",
                id="clear_range",
                text_style={"bold": True, "color": "#1F4E78"},
            )
            .build()
            .to_bytes("docx")
        )
        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "text.format",
                    "target": "#clear_range",
                    "range": {"start": 1, "end": 3},
                    "clear": ["bold"],
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        reopened = Document.from_docx(result.document.to_bytes("docx"))
        node = reopened.to_spec()["content"][0]
        self.assertEqual(
            [(span["text"], span.get("style", {}).get("bold")) for span in node["content"]],
            [("A", True), ("BC", None), ("D", True)],
        )
        self.assertEqual(node["text_style"]["color"], "#1F4E78")


if __name__ == "__main__":
    unittest.main()
