"""Hardened XML parsing helpers that retain comments and namespace declarations."""

from __future__ import annotations

from io import BytesIO
from xml.etree import ElementTree as ET
from xml.sax.saxutils import quoteattr

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


def namespace_declaration_values(
    data: bytes,
) -> dict[str, frozenset[str]]:
    """Return every URI bound to each prefix anywhere in one XML part."""

    values: dict[str, set[str]] = {}
    for prefix, uri in _namespace_declaration_items(data):
        values.setdefault(prefix, set()).add(uri)
    return {
        prefix: frozenset(uris)
        for prefix, uris in values.items()
    }


def preserve_namespace_prefixes(
    data: bytes,
    *,
    source: bytes,
    prefixes: frozenset[str],
) -> bytes:
    """Restore selected unambiguous prefixes omitted by ElementTree.

    ElementTree drops namespace declarations that are referenced only by
    lexical values such as ``mc:Choice/@Requires``. This helper restores only
    explicitly requested, uniquely bound non-default prefixes on the root.
    """

    source_values = namespace_declaration_values(source)
    current_values = namespace_declaration_values(data)
    declarations: list[str] = []
    for prefix in sorted(prefixes):
        if not prefix or prefix == "xml":
            raise NativePackageError(
                "Only explicit non-reserved namespace prefixes can be "
                "preserved."
            )
        values = source_values.get(prefix)
        if values is None:
            continue
        if len(values) != 1:
            raise NativePackageError(
                f"Namespace prefix {prefix!r} is rebound in the source XML."
            )
        current = current_values.get(prefix)
        if current is not None:
            if current != values:
                raise NativePackageError(
                    f"Namespace prefix {prefix!r} changed during XML editing."
                )
            continue
        uri = next(iter(values))
        declarations.append(
            f" xmlns:{prefix}={quoteattr(uri)}"
        )
    if not declarations:
        return data
    declaration_end = data.find(b"?>")
    root_end = data.find(b">", declaration_end + 2)
    if root_end < 0:
        raise NativePackageError(
            "Serialized XML has no root start tag for namespace restoration."
        )
    insertion = "".join(declarations).encode("utf-8")
    return data[:root_end] + insertion + data[root_end:]


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
