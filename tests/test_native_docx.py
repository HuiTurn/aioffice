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

    def test_unsupported_native_operation_is_atomic(self) -> None:
        source = self._source_document()
        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "node.append",
                    "target": "$",
                    "content": {"type": "paragraph", "text": "New"},
                }
            ]
        )
        self.assertFalse(result.success)
        self.assertEqual(result.diagnostics[0].code, "NATIVE_PATCH_FAILED")
        self.assertEqual(document.revision, 1)
        assert document.fidelity is not None
        self.assertEqual(document.fidelity.level, FidelityLevel.EXACT_PACKAGE)

        unsupported_insert = document.apply(
            [
                {
                    "op": "node.insert_after",
                    "target": "#body",
                    "content": {
                        "id": "unsupported_break",
                        "type": "page_break",
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
            "only paragraph and heading",
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
