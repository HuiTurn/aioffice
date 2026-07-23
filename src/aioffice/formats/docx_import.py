"""Semantic projection of an existing DOCX over a lossless native package."""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from aioffice.core.diagnostics import Diagnostic
from aioffice.core.errors import NativePackageError
from aioffice.native import (
    FidelityPolicy,
    IdentityManifest,
    MANIFEST_PART_URI,
    MANIFEST_RELATIONSHIP_TYPE,
    NativePackage,
    apply_identity_manifest,
    native_ref_for_elements,
    parse_identity_manifest,
)
from aioffice.native.xml import parse_xml
from aioffice.security import SecurityPolicy
from aioffice.spec.models import AiOfficeDocumentSpec

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
DC = "http://purl.org/dc/elements/1.1/"


def _q(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def _unique_id(prefix: str, hint: str, index: int, seen: set[str]) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_.:-]", "_", hint)[:80]
    candidate = f"{prefix}_{cleaned}" if cleaned else f"{prefix}_{index:06d}"
    if candidate in seen:
        candidate = f"{candidate}_{index:06d}"
    seen.add(candidate)
    return candidate


def _paragraph_text(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.iter(_q(W, "t")))


def _source_ref(
    elements: list[ET.Element],
    indices: list[int],
    native_kind: str,
    *,
    native_id: str | None = None,
) -> dict[str, Any]:
    return native_ref_for_elements(
        elements,
        indices,
        native_kind=native_kind,
        native_id=native_id,
    ).model_dump(mode="json", exclude_none=True)


def _paragraph_projection(
    element: ET.Element,
    index: int,
    seen_ids: set[str],
) -> dict[str, Any]:
    text = _paragraph_text(element)
    para_id = element.attrib.get(_q(W14, "paraId"), f"{index:06d}")
    node_id = _unique_id("para", para_id, index, seen_ids)
    style = element.find(f"./{_q(W, 'pPr')}/{_q(W, 'pStyle')}")
    style_id = style.attrib.get(_q(W, "val"), "") if style is not None else ""
    heading_match = re.fullmatch(r"Heading([1-6])", style_id, flags=re.IGNORECASE)
    native_features: list[str] = []
    if element.find(f".//{_q(W, 'drawing')}") is not None:
        native_features.append("drawing")
    if element.find(f".//{_q(W, 'object')}") is not None:
        native_features.append("object")
    common: dict[str, Any] = {
        "id": node_id,
        "source_ref": _source_ref(
            [element],
            [index],
            "w:p",
            native_id=element.attrib.get(_q(W14, "paraId")),
        ),
        "metadata": {
            "native_style_id": style_id or None,
            "native_features": native_features,
        },
    }
    if heading_match:
        return {
            **common,
            "type": "heading",
            "level": int(heading_match.group(1)),
            "text": text,
        }
    if not text and native_features:
        return {
            **common,
            "id": _unique_id("opaque", para_id, index, seen_ids),
            "type": "opaque",
            "summary": "DOCX paragraph containing unsupported drawing or embedded content.",
            "capabilities": ["inspect", "move", "delete", "render"],
            "editable": False,
        }
    return {**common, "type": "paragraph", "text": text}


def _table_projection(
    element: ET.Element,
    index: int,
    seen_ids: set[str],
) -> dict[str, Any]:
    raw_rows: list[list[str]] = []
    for row in element.findall(f"./{_q(W, 'tr')}"):
        raw_rows.append(
            [_paragraph_text(cell) for cell in row.findall(f"./{_q(W, 'tc')}")]
        )
    column_count = max((len(row) for row in raw_rows), default=1)
    header = raw_rows[0] if raw_rows else []
    columns = [
        {
            "key": f"column_{column_index + 1}",
            "title": (
                header[column_index]
                if column_index < len(header) and header[column_index]
                else f"Column {column_index + 1}"
            ),
        }
        for column_index in range(column_count)
    ]
    table_id = _unique_id(
        "table",
        hashlib.sha256(ET.tostring(element, encoding="utf-8")).hexdigest()[:12],
        index,
        seen_ids,
    )
    rows = []
    for row_index, values in enumerate(raw_rows[1:], start=1):
        rows.append(
            {
                "id": _unique_id("row", f"{index}_{row_index}", row_index, seen_ids),
                "values": {
                    column["key"]: values[column_index]
                    if column_index < len(values)
                    else ""
                    for column_index, column in enumerate(columns)
                },
            }
        )
    return {
        "id": table_id,
        "type": "table",
        "columns": columns,
        "rows": rows,
        "source_ref": _source_ref([element], [index], "w:tbl"),
        "metadata": {"projection": "heuristic", "header_row_assumed": bool(raw_rows)},
    }


def _page_break_projection(
    element: ET.Element,
    index: int,
    seen_ids: set[str],
) -> dict[str, Any]:
    para_id = element.attrib.get(_q(W14, "paraId"), f"{index:06d}")
    return {
        "id": _unique_id("break", para_id, index, seen_ids),
        "type": "page_break",
        "source_ref": _source_ref(
            [element],
            [index],
            "w:page-break",
            native_id=element.attrib.get(_q(W14, "paraId")),
        ),
    }


def _is_page_break(element: ET.Element) -> bool:
    if _paragraph_text(element):
        return False
    return any(
        node.attrib.get(_q(W, "type")) == "page"
        for node in element.iter(_q(W, "br"))
    )


def _paragraph_numbering(element: ET.Element) -> tuple[str, int] | None:
    properties = element.find(f"./{_q(W, 'pPr')}/{_q(W, 'numPr')}")
    if properties is None:
        return None
    number = properties.find(_q(W, "numId"))
    level = properties.find(_q(W, "ilvl"))
    if number is None:
        return None
    number_id = number.attrib.get(_q(W, "val"))
    if not number_id:
        return None
    try:
        level_value = int(level.attrib.get(_q(W, "val"), "0")) if level is not None else 0
    except ValueError:
        level_value = 0
    return number_id, level_value


def _numbering_formats(package: NativePackage) -> dict[tuple[str, int], str]:
    try:
        root = parse_xml(package.get_part("/word/numbering.xml"))
    except NativePackageError:
        return {}
    abstract_formats: dict[tuple[str, int], str] = {}
    for abstract in root.findall(_q(W, "abstractNum")):
        abstract_id = abstract.attrib.get(_q(W, "abstractNumId"))
        if abstract_id is None:
            continue
        for level in abstract.findall(_q(W, "lvl")):
            try:
                level_index = int(level.attrib.get(_q(W, "ilvl"), "0"))
            except ValueError:
                level_index = 0
            number_format = level.find(_q(W, "numFmt"))
            if number_format is not None:
                abstract_formats[(abstract_id, level_index)] = number_format.attrib.get(
                    _q(W, "val"), ""
                )
    result: dict[tuple[str, int], str] = {}
    for number in root.findall(_q(W, "num")):
        number_id = number.attrib.get(_q(W, "numId"))
        abstract = number.find(_q(W, "abstractNumId"))
        if number_id is None or abstract is None:
            continue
        abstract_id = abstract.attrib.get(_q(W, "val"))
        for (candidate_id, level), format_name in abstract_formats.items():
            if candidate_id == abstract_id:
                result[(number_id, level)] = format_name
    return result


def _list_projection(
    elements: list[ET.Element],
    indices: list[int],
    seen_ids: set[str],
    *,
    number_id: str,
    level: int,
    format_name: str,
) -> dict[str, Any]:
    first = elements[0]
    first_index = indices[0]
    para_id = first.attrib.get(_q(W14, "paraId"))
    hint = para_id or hashlib.sha256(
        b"".join(ET.tostring(element, encoding="utf-8") for element in elements)
    ).hexdigest()[:12]
    list_type = "bullet_list" if format_name == "bullet" else "ordered_list"
    return {
        "id": _unique_id("list", hint, first_index, seen_ids),
        "type": list_type,
        "items": [_paragraph_text(element) for element in elements],
        "source_ref": _source_ref(
            elements,
            indices,
            "w:p-group",
            native_id=para_id,
        ),
        "metadata": {
            "native_num_id": number_id,
            "native_level": level,
            "native_num_format": format_name,
        },
    }


def _core_metadata(package: NativePackage) -> dict[str, Any]:
    try:
        payload = package.get_part("/docProps/core.xml")
    except NativePackageError:
        return {}
    root = parse_xml(payload)
    title = root.find(_q(DC, "title"))
    author = root.find(_q(DC, "creator"))
    return {
        key: value
        for key, value in {
            "title": title.text if title is not None else None,
            "author": author.text if author is not None else None,
        }.items()
        if value
    }


def _embedded_identity_manifest(
    package: NativePackage,
) -> IdentityManifest | None:
    relationships = [
        relationship
        for relationship in package.relationships
        if relationship.source_part == "/"
        and relationship.relationship_type == MANIFEST_RELATIONSHIP_TYPE
    ]
    if not relationships:
        return None
    if len(relationships) != 1:
        raise NativePackageError(
            "DOCX package contains multiple AiOffice identity relationships."
        )
    relationship = relationships[0]
    if relationship.external:
        raise NativePackageError("AiOffice identity relationship cannot be external.")
    target_uri = "/" + relationship.target.lstrip("/")
    if target_uri != MANIFEST_PART_URI:
        raise NativePackageError(
            "AiOffice identity relationship targets an unexpected package part."
        )
    if not package.has_part(MANIFEST_PART_URI):
        raise NativePackageError(
            "DOCX package identity relationship points to a missing manifest."
        )
    return parse_identity_manifest(package.get_part(MANIFEST_PART_URI))


@dataclass(slots=True)
class ImportedDocx:
    spec: AiOfficeDocumentSpec
    native: NativePackage
    diagnostics: list[Diagnostic]


def import_docx(
    source: str | Path | bytes,
    *,
    roundtrip: FidelityPolicy | str = FidelityPolicy.PRESERVE_UNKNOWN,
    security_policy: SecurityPolicy | None = None,
    identity_manifest: IdentityManifest | None = None,
) -> ImportedDocx:
    package = NativePackage.open(
        source,
        format_name="docx",
        policy=roundtrip,
        security_policy=security_policy,
    )
    document_root = parse_xml(package.get_part("/word/document.xml"))
    body = document_root.find(_q(W, "body"))
    if body is None:
        raise NativePackageError("DOCX main document part has no w:body.")

    content: list[dict[str, Any]] = []
    seen_ids: set[str] = set()
    body_elements = list(body)
    numbering_formats = _numbering_formats(package)
    index = 0
    while index < len(body_elements):
        element = body_elements[index]
        if element.tag == _q(W, "p"):
            numbering = _paragraph_numbering(element)
            if numbering is not None and numbering in numbering_formats:
                group_elements = [element]
                group_indices = [index]
                next_index = index + 1
                while next_index < len(body_elements):
                    candidate = body_elements[next_index]
                    if (
                        candidate.tag != _q(W, "p")
                        or _paragraph_numbering(candidate) != numbering
                    ):
                        break
                    group_elements.append(candidate)
                    group_indices.append(next_index)
                    next_index += 1
                content.append(
                    _list_projection(
                        group_elements,
                        group_indices,
                        seen_ids,
                        number_id=numbering[0],
                        level=numbering[1],
                        format_name=numbering_formats[numbering],
                    )
                )
                index = next_index
                continue
            if _is_page_break(element):
                content.append(_page_break_projection(element, index, seen_ids))
            else:
                content.append(_paragraph_projection(element, index, seen_ids))
        elif element.tag == _q(W, "tbl"):
            content.append(_table_projection(element, index, seen_ids))
        elif element.tag == _q(W, "sectPr"):
            index += 1
            continue
        else:
            native_kind = element.tag.rsplit("}", 1)[-1]
            content.append(
                {
                    "id": _unique_id("opaque", native_kind, index, seen_ids),
                    "type": "opaque",
                    "summary": f"Unsupported DOCX body element {native_kind}.",
                    "source_ref": _source_ref([element], [index], native_kind),
                    "capabilities": ["inspect", "move", "delete", "render"],
                    "editable": False,
                }
            )
        index += 1

    active_identity_manifest = identity_manifest
    identity_source = "workspace" if identity_manifest is not None else None
    if active_identity_manifest is None:
        active_identity_manifest = _embedded_identity_manifest(package)
        if active_identity_manifest is not None:
            identity_source = "embedded"
    diagnostics: list[Diagnostic] = []
    if active_identity_manifest is not None:
        if active_identity_manifest.format != "docx":
            raise NativePackageError("Identity manifest format does not match DOCX input.")
        diagnostics.extend(
            apply_identity_manifest(
                content,
                active_identity_manifest,
                package_sha256=package.source_sha256,
            )
        )

    spec = AiOfficeDocumentSpec.model_validate(
        {
            "artifact": {
                "id": (
                    active_identity_manifest.artifact_id
                    if active_identity_manifest is not None
                    else f"doc_{package.source_sha256[:24]}"
                ),
                "kind": "document",
                "revision": (
                    active_identity_manifest.revision
                    if active_identity_manifest is not None
                    else 1
                ),
            },
            "metadata": _core_metadata(package),
            "theme": {"ref": "business-clean"},
            "content": content,
            "extensions": {
                "dev.aioffice.native": {
                    "format": "docx",
                    "authority": "native",
                    "roundtrip_policy": FidelityPolicy(roundtrip).value,
                    "source_sha256": package.source_sha256,
                    "identity_source": identity_source,
                }
            },
        }
    )
    return ImportedDocx(spec=spec, native=package, diagnostics=diagnostics)


__all__ = ["ImportedDocx", "import_docx"]
