"""Semantic projection of an existing DOCX over a lossless native package."""

from __future__ import annotations

import hashlib
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from aioffice.core.errors import NativePackageError
from aioffice.native import FidelityPolicy, NativePackage
from aioffice.native.xml import parse_xml, serialize_xml
from aioffice.security import SecurityPolicy
from aioffice.spec.models import AiOfficeDocumentSpec

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
DC = "http://purl.org/dc/elements/1.1/"


def _q(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def _fingerprint(element: ET.Element) -> str:
    return "sha256:" + hashlib.sha256(serialize_xml(element)).hexdigest()


def _unique_id(prefix: str, hint: str, index: int, seen: set[str]) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9_.:-]", "_", hint)[:80]
    candidate = f"{prefix}_{cleaned}" if cleaned else f"{prefix}_{index:06d}"
    if candidate in seen:
        candidate = f"{candidate}_{index:06d}"
    seen.add(candidate)
    return candidate


def _paragraph_text(element: ET.Element) -> str:
    return "".join(node.text or "" for node in element.iter(_q(W, "t")))


def _source_ref(element: ET.Element, index: int, native_kind: str) -> dict[str, Any]:
    return {
        "format": "docx",
        "part_uri": "/word/document.xml",
        "native_kind": native_kind,
        "element_index": index,
        "path_hint": f"/w:document/w:body/*[{index + 1}]",
        "fingerprint": _fingerprint(element),
    }


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
        "source_ref": _source_ref(element, index, "w:p"),
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
        hashlib.sha256(serialize_xml(element)).hexdigest()[:12],
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
        "source_ref": _source_ref(element, index, "w:tbl"),
        "metadata": {"projection": "heuristic", "header_row_assumed": bool(raw_rows)},
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


@dataclass(slots=True)
class ImportedDocx:
    spec: AiOfficeDocumentSpec
    native: NativePackage


def import_docx(
    source: str | Path | bytes,
    *,
    roundtrip: FidelityPolicy | str = FidelityPolicy.PRESERVE_UNKNOWN,
    security_policy: SecurityPolicy | None = None,
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
    for index, element in enumerate(list(body)):
        if element.tag == _q(W, "p"):
            content.append(_paragraph_projection(element, index, seen_ids))
        elif element.tag == _q(W, "tbl"):
            content.append(_table_projection(element, index, seen_ids))
        elif element.tag == _q(W, "sectPr"):
            continue
        else:
            native_kind = element.tag.rsplit("}", 1)[-1]
            content.append(
                {
                    "id": _unique_id("opaque", native_kind, index, seen_ids),
                    "type": "opaque",
                    "summary": f"Unsupported DOCX body element {native_kind}.",
                    "source_ref": _source_ref(element, index, native_kind),
                    "capabilities": ["inspect", "move", "delete", "render"],
                    "editable": False,
                }
            )

    spec = AiOfficeDocumentSpec.model_validate(
        {
            "artifact": {
                "id": f"doc_{package.source_sha256[:24]}",
                "kind": "document",
                "revision": 1,
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
                }
            },
        }
    )
    return ImportedDocx(spec=spec, native=package)


__all__ = ["ImportedDocx", "import_docx"]
