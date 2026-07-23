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
from aioffice.native import FidelityLevel
from aioffice.native.xml import parse_xml, serialize_xml

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _q(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def _rewrite_package(
    source: bytes,
    replacements: dict[str, bytes],
    additions: dict[str, bytes] | None = None,
) -> bytes:
    output = io.BytesIO()
    with ZipFile(io.BytesIO(source)) as input_archive, ZipFile(
        output, "w", compression=ZIP_DEFLATED
    ) as output_archive:
        for info in input_archive.infolist():
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
        document_xml = parse_xml(
            ZipFile(io.BytesIO(source)).read("word/document.xml")
        )
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
            if node.get("text") == "Alpha Beta Gamma"
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
        self.assertEqual(result.fidelity.affected_parts, ["/word/document.xml"])
        self.assertTrue(result.fidelity.visual_verification_required)

        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / "patched.docx"
            result.document.export(target)
            with ZipFile(io.BytesIO(source)) as before, ZipFile(target) as after:
                for name in before.namelist():
                    if name != "word/document.xml":
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
                node.get("text")
                for node in reopened.to_spec()["content"]
                if node["type"] in {"heading", "paragraph"}
            ]
            self.assertIn("AlpHA-BEta Gamma", texts)

    def test_unsupported_native_operation_is_atomic(self) -> None:
        document = Document.from_docx(self._source_document())
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

    def test_identity_map_is_refreshed_after_removal(self) -> None:
        document = Document.from_docx(self._source_document())
        nodes = document.to_spec()["content"]
        removed = document.apply(
            [{"op": "node.remove", "target": f"#{nodes[0]['id']}"}]
        )
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
                node.get("text")
                for node in reopened.to_spec()["content"]
                if node["type"] in {"heading", "paragraph"}
            ]
            self.assertIn("Alpha Delta Gamma", texts)


if __name__ == "__main__":
    unittest.main()
