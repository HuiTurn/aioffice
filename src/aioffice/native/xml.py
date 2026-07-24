"""Hardened XML parsing helpers that retain comments and namespace declarations."""

from __future__ import annotations

from io import BytesIO
from xml.etree import ElementTree as ET

from defusedxml import ElementTree as DET

from aioffice.core.errors import NativePackageError, SecurityError


def _namespace_declaration_items(data: bytes) -> list[tuple[str, str]]:
    declarations: list[tuple[str, str]] = []
    try:
        for _, namespace in DET.iterparse(
            BytesIO(data),
            events=("start-ns",),
            forbid_dtd=True,
            forbid_entities=True,
            forbid_external=True,
        ):
            prefix, uri = namespace
            declarations.append((prefix or "", uri))
    except (ET.ParseError, ValueError, SecurityError) as error:
        raise NativePackageError(
            f"Invalid or unsafe XML namespace declarations: {error}"
        ) from error
    return declarations


def namespace_declarations(data: bytes) -> dict[str, str]:
    """Return declared XML prefixes without resolving document content."""

    return dict(_namespace_declaration_items(data))


def register_document_namespaces(data: bytes) -> None:
    for prefix, uri in _namespace_declaration_items(data):
        try:
            ET.register_namespace(prefix, uri)
        except ValueError:
            # ElementTree reserves prefixes matching ns\d+.
            continue


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
