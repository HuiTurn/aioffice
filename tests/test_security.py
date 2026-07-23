from __future__ import annotations

import copy
import io
import unittest
import warnings
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from aioffice.core.errors import NativePackageError, SecurityError
from aioffice.documents import Document, DocumentBuilder
from aioffice.formats.docx import compile_docx
from aioffice.native import (
    MANIFEST_PART_URI,
    MANIFEST_RELATIONSHIP_TYPE,
    NativePackage,
)


def _append(source: bytes, name: str, payload: bytes) -> bytes:
    output = io.BytesIO()
    with ZipFile(io.BytesIO(source)) as input_archive, ZipFile(
        output, "w", compression=ZIP_DEFLATED
    ) as output_archive:
        for info in input_archive.infolist():
            output_archive.writestr(copy.copy(info), input_archive.read(info.filename))
        output_archive.writestr(name, payload)
    return output.getvalue()


def _replace(source: bytes, name: str, payload: bytes) -> bytes:
    output = io.BytesIO()
    with ZipFile(io.BytesIO(source)) as input_archive, ZipFile(
        output, "w", compression=ZIP_DEFLATED
    ) as output_archive:
        for info in input_archive.infolist():
            value = payload if info.filename == name else input_archive.read(info.filename)
            output_archive.writestr(copy.copy(info), value)
    return output.getvalue()


class NativeSecurityTests(unittest.TestCase):
    def setUp(self) -> None:
        self.source = compile_docx(
            DocumentBuilder(title="Security").paragraph("Safe").build().spec
        )

    def test_path_traversal_is_rejected(self) -> None:
        malicious = _append(self.source, "../escape.xml", b"<escape/>")
        with self.assertRaises(SecurityError):
            NativePackage.open(malicious, format_name="docx")

    def test_duplicate_members_are_rejected(self) -> None:
        with warnings.catch_warnings():
            warnings.simplefilter("ignore")
            malicious = _append(self.source, "word/document.xml", b"<duplicate/>")
        with self.assertRaises(SecurityError):
            NativePackage.open(malicious, format_name="docx")

    def test_macro_payload_is_rejected(self) -> None:
        malicious = _append(self.source, "word/vbaProject.bin", b"macro")
        with self.assertRaises(SecurityError):
            NativePackage.open(malicious, format_name="docx")

    def test_suspicious_compression_ratio_is_rejected(self) -> None:
        malicious = _append(
            self.source,
            "customXml/compression-bomb.xml",
            b"0" * (1024 * 1024),
        )
        with self.assertRaises(SecurityError):
            NativePackage.open(malicious, format_name="docx")

    def test_dtd_and_entity_expansion_are_rejected(self) -> None:
        malicious_xml = b"""<?xml version="1.0"?>
<!DOCTYPE doc [<!ENTITY x "unsafe">]>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body><w:p><w:r><w:t>&x;</w:t></w:r></w:p></w:body>
</w:document>
"""
        malicious = _replace(self.source, "word/document.xml", malicious_xml)
        with self.assertRaises(NativePackageError):
            Document.from_docx(malicious)

    def test_duplicate_embedded_identity_ids_are_rejected(self) -> None:
        with ZipFile(io.BytesIO(self.source)) as archive:
            manifest = ET.fromstring(
                archive.read(MANIFEST_PART_URI.lstrip("/"))
            )
        first_node = next(iter(manifest))
        manifest.append(copy.deepcopy(first_node))
        malicious = _replace(
            self.source,
            MANIFEST_PART_URI.lstrip("/"),
            ET.tostring(manifest, encoding="utf-8", xml_declaration=True),
        )
        with self.assertRaises(NativePackageError):
            Document.from_docx(malicious)

    def test_identity_relationship_must_target_manifest(self) -> None:
        with ZipFile(io.BytesIO(self.source)) as archive:
            relationships = ET.fromstring(archive.read("_rels/.rels"))
        identity_relationship = next(
            relationship
            for relationship in relationships
            if relationship.attrib.get("Type") == MANIFEST_RELATIONSHIP_TYPE
        )
        identity_relationship.set("Target", "customXml/other.xml")
        malicious = _replace(
            self.source,
            "_rels/.rels",
            ET.tostring(
                relationships,
                encoding="utf-8",
                xml_declaration=True,
            ),
        )
        with self.assertRaises(NativePackageError):
            Document.from_docx(malicious)


if __name__ == "__main__":
    unittest.main()
