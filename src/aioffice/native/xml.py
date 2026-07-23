"""Hardened XML parsing helpers that retain comments and namespace declarations."""

from __future__ import annotations

from io import BytesIO
from xml.etree import ElementTree as ET

from defusedxml import ElementTree as DET

from aioffice.core.errors import NativePackageError, SecurityError


def register_document_namespaces(data: bytes) -> None:
    try:
        for _, namespace in DET.iterparse(
            BytesIO(data),
            events=("start-ns",),
            forbid_dtd=True,
            forbid_entities=True,
            forbid_external=True,
        ):
            prefix, uri = namespace
            try:
                ET.register_namespace(prefix or "", uri)
            except ValueError:
                # ElementTree reserves prefixes matching ns\d+.
                continue
    except (ET.ParseError, ValueError, SecurityError) as error:
        raise NativePackageError(f"Invalid or unsafe XML namespace declarations: {error}") from error


def parse_xml(data: bytes) -> ET.Element:
    register_document_namespaces(data)
    parser = DET.DefusedXMLParser(
        target=ET.TreeBuilder(insert_comments=True, insert_pis=True),
        forbid_dtd=True,
        forbid_entities=True,
        forbid_external=True,
    )
    try:
        return ET.fromstring(data, parser=parser)
    except (ET.ParseError, ValueError) as error:
        raise NativePackageError(f"Invalid or unsafe XML: {error}") from error


def serialize_xml(root: ET.Element) -> bytes:
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)
