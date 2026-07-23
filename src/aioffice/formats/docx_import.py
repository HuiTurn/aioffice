"""Semantic projection of an existing DOCX over a lossless native package."""

from __future__ import annotations

import hashlib
import re
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any
from xml.etree import ElementTree as ET

from aioffice.core.diagnostics import Diagnostic, Severity
from aioffice.core.errors import NativePackageError
from aioffice.formats.docx_style import (
    common_text_style,
    read_paragraph_style,
    read_text_style,
)
from aioffice.formats.docx_named_styles import (
    read_document_defaults,
    read_named_styles,
)
from aioffice.formats.docx_header_footer import (
    FOOTER_RELATIONSHIP_TYPE,
    HEADER_RELATIONSHIP_TYPE,
    binding_field,
    native_ref_for_header_footer_part,
    read_even_and_odd_headers,
    read_update_fields_on_open,
    resolve_relationship_target,
)
from aioffice.formats.docx_fields import (
    FieldMatch,
    FieldStructureError,
    field_payload,
    parse_paragraph_fields,
)
from aioffice.formats.docx_section import (
    native_ref_for_section,
    read_section_layout,
)
from aioffice.native import (
    FidelityPolicy,
    IdentityManifest,
    MANIFEST_PART_URI,
    MANIFEST_RELATIONSHIP_TYPE,
    NativePackage,
    apply_identity_manifest,
    native_ref_for_elements,
    native_ref_for_part_elements,
    parse_identity_manifest,
)
from aioffice.native.xml import parse_xml
from aioffice.security import SecurityPolicy
from aioffice.spec.models import AiOfficeDocumentSpec, NamedStyle, TextStyle

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
DC = "http://purl.org/dc/elements/1.1/"
HYPERLINK_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
)


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


def _paragraph_section(element: ET.Element) -> ET.Element | None:
    return element.find(f"./{_q(W, 'pPr')}/{_q(W, 'sectPr')}")


def _is_section_carrier(element: ET.Element) -> bool:
    return (
        _paragraph_section(element) is not None
        and not _paragraph_text(element)
        and element.find(f".//{_q(W, 'drawing')}") is None
        and element.find(f".//{_q(W, 'object')}") is None
    )


def _hyperlink_targets(
    package: NativePackage,
    *,
    source_part: str = "/word/document.xml",
) -> dict[str, str]:
    return {
        relationship.relationship_id: relationship.target
        for relationship in package.relationships
        if relationship.source_part == source_part
        and relationship.relationship_type == HYPERLINK_RELATIONSHIP_TYPE
    }


def _iter_text_runs(
    element: ET.Element,
    hyperlink_targets: dict[str, str],
    *,
    inherited_href: str | None = None,
) -> list[tuple[ET.Element, str | None]]:
    result: list[tuple[ET.Element, str | None]] = []
    for child in list(element):
        if child.tag == _q(W, "pPr"):
            continue
        if child.tag == _q(W, "hyperlink"):
            relationship_id = child.attrib.get(_q(R, "id"))
            anchor = child.attrib.get(_q(W, "anchor"))
            href = (
                hyperlink_targets.get(relationship_id, "")
                if relationship_id
                else f"#{anchor}"
                if anchor
                else None
            )
            result.extend(
                _iter_text_runs(
                    child,
                    hyperlink_targets,
                    inherited_href=href or inherited_href,
                )
            )
        elif child.tag == _q(W, "r"):
            result.append((child, inherited_href))
        else:
            result.extend(
                _iter_text_runs(
                    child,
                    hyperlink_targets,
                    inherited_href=inherited_href,
                )
            )
    return result


def _merge_projected_spans(spans: list[dict[str, Any]]) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for span in spans:
        if not span["text"]:
            continue
        if result and {key: value for key, value in result[-1].items() if key != "text"} == {
            key: value for key, value in span.items() if key != "text"
        }:
            result[-1]["text"] += span["text"]
        else:
            result.append(span)
    return result


def _merge_inline_content(
    content: list[dict[str, Any]],
) -> list[dict[str, Any]]:
    result: list[dict[str, Any]] = []
    for inline in content:
        if (
            result
            and result[-1].get("type") == "text"
            and inline.get("type") == "text"
            and {
                key: value
                for key, value in result[-1].items()
                if key != "text"
            }
            == {
                key: value
                for key, value in inline.items()
                if key != "text"
            }
        ):
            result[-1]["text"] += inline.get("text", "")
        else:
            result.append(inline)
    return result


def _text_projection(
    element: ET.Element,
    hyperlink_targets: dict[str, str],
) -> tuple[dict[str, Any], dict[str, Any] | None]:
    common_style = common_text_style(element)
    common_payload = (
        common_style.model_dump(mode="json", exclude_none=True) if common_style is not None else {}
    )
    spans: list[dict[str, Any]] = []
    for run, href in _iter_text_runs(element, hyperlink_targets):
        text = _paragraph_text(run)
        if not text:
            continue
        run_style = read_text_style(run)
        style_payload = (
            run_style.model_dump(mode="json", exclude_none=True) if run_style is not None else {}
        )
        for field_name, common_value in common_payload.items():
            if style_payload.get(field_name) == common_value:
                style_payload.pop(field_name)
        span: dict[str, Any] = {"type": "text", "text": text}
        if href:
            span["marks"] = ["link"]
            span["href"] = href
        if style_payload:
            span["style"] = style_payload
        spans.append(span)
    spans = _merge_projected_spans(spans)
    if len(spans) == 1 and not spans[0].get("marks") and not spans[0].get("style"):
        return {"text": spans[0]["text"]}, common_payload or None
    if spans:
        return {"content": spans}, common_payload or None
    return {"text": _paragraph_text(element)}, common_payload or None


def _materialized_text_spans(
    elements: list[ET.Element],
    hyperlink_targets: dict[str, str],
) -> list[dict[str, Any]]:
    if not elements:
        return []
    wrapper = ET.Element(_q(W, "p"))
    for element in elements:
        wrapper.append(deepcopy(element))
    projection, common_style = _text_projection(wrapper, hyperlink_targets)
    spans = (
        [{"type": "text", "text": projection["text"]}]
        if "text" in projection
        else [deepcopy(span) for span in projection.get("content", [])]
    )
    if common_style:
        for span in spans:
            style = {
                **common_style,
                **(
                    span.get("style", {})
                    if isinstance(span.get("style"), dict)
                    else {}
                ),
            }
            if style:
                span["style"] = style
    return spans


def _field_result_style(match: FieldMatch) -> TextStyle | None:
    if not match.result_elements:
        return None
    wrapper = ET.Element(_q(W, "p"))
    for element in match.result_elements:
        wrapper.append(deepcopy(element))
    return common_text_style(wrapper)


def _field_aware_text_projection(
    element: ET.Element,
    index: int,
    seen_ids: set[str],
    hyperlink_targets: dict[str, str],
    *,
    part_uri: str,
    root_path: str,
) -> tuple[dict[str, Any], dict[str, Any] | None]:
    matches = parse_paragraph_fields(element)
    if not matches:
        return _text_projection(element, hyperlink_targets)
    children = list(element)
    content: list[dict[str, Any]] = []
    cursor = 0
    para_id = element.attrib.get(_q(W14, "paraId"), f"{index:06d}")
    for match in matches:
        content.extend(
            _materialized_text_spans(
                children[cursor : match.start_index],
                hyperlink_targets,
            )
        )
        fingerprint = hashlib.sha256(
            b"".join(
                ET.tostring(field_element, encoding="utf-8")
                for field_element in match.elements
            )
        ).hexdigest()[:12]
        field_id = _unique_id(
            "field",
            f"{para_id}_{match.ordinal}_{fingerprint}",
            match.ordinal,
            seen_ids,
        )
        content.append(
            field_payload(
                element,
                index,
                match,
                field_id=field_id,
                part_uri=part_uri,
                root_path=root_path,
                style=_field_result_style(match),
            )
        )
        cursor = match.end_index + 1
    content.extend(
        _materialized_text_spans(
            children[cursor:],
            hyperlink_targets,
        )
    )
    return {"content": _merge_inline_content(content)}, None


def _source_ref(
    elements: list[ET.Element],
    indices: list[int],
    native_kind: str,
    *,
    native_id: str | None = None,
    part_uri: str = "/word/document.xml",
    root_path: str = "/w:document/w:body",
) -> dict[str, Any]:
    factory = (
        native_ref_for_elements
        if part_uri == "/word/document.xml"
        and root_path == "/w:document/w:body"
        else None
    )
    source_ref = (
        factory(
            elements,
            indices,
            native_kind=native_kind,
            native_id=native_id,
        )
        if factory is not None
        else native_ref_for_part_elements(
            elements,
            indices,
            part_uri=part_uri,
            native_kind=native_kind,
            root_path=root_path,
            native_id=native_id,
        )
    )
    return source_ref.model_dump(mode="json", exclude_none=True)


def _paragraph_projection(
    element: ET.Element,
    index: int,
    seen_ids: set[str],
    hyperlink_targets: dict[str, str],
    named_styles: dict[str, NamedStyle],
    *,
    part_uri: str = "/word/document.xml",
    root_path: str = "/w:document/w:body",
    allow_heading: bool = True,
) -> dict[str, Any]:
    try:
        text_projection, text_style = _field_aware_text_projection(
            element,
            index,
            seen_ids,
            hyperlink_targets,
            part_uri=part_uri,
            root_path=root_path,
        )
    except FieldStructureError as error:
        para_id = element.attrib.get(_q(W14, "paraId"), f"{index:06d}")
        return {
            "id": _unique_id("opaque", para_id, index, seen_ids),
            "type": "opaque",
            "summary": f"Unsupported or malformed DOCX field structure: {error}",
            "source_ref": _source_ref(
                [element],
                [index],
                "w:p",
                native_id=element.attrib.get(_q(W14, "paraId")),
                part_uri=part_uri,
                root_path=root_path,
            ),
            "capabilities": ["inspect", "render"],
            "editable": False,
            "metadata": {"native_features": ["field_structure"]},
        }
    text = _paragraph_text(element)
    para_id = element.attrib.get(_q(W14, "paraId"), f"{index:06d}")
    node_id = _unique_id("para", para_id, index, seen_ids)
    style = element.find(f"./{_q(W, 'pPr')}/{_q(W, 'pStyle')}")
    style_id = style.attrib.get(_q(W, "val"), "") if style is not None else ""
    named_style = named_styles.get(style_id)
    heading_level = (
        named_style.heading_level
        if named_style is not None
        and named_style.semantic_role == "heading"
        and named_style.heading_level is not None
        and named_style.heading_level <= 6
        else None
    )
    native_features: list[str] = []
    if any(
        inline.get("type") == "field"
        for inline in text_projection.get("content", [])
        if isinstance(inline, dict)
    ):
        native_features.append("field")
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
            part_uri=part_uri,
            root_path=root_path,
        ),
        "metadata": {
            "native_style_id": style_id or None,
            "native_features": native_features,
        },
    }
    implicit_heading_style = (
        heading_level is not None
        and style_id.casefold() == f"Heading{heading_level}".casefold()
    )
    if style_id and not implicit_heading_style:
        common["style_ref"] = style_id
    paragraph_style = read_paragraph_style(element)
    if paragraph_style is not None:
        common["paragraph_style"] = paragraph_style.model_dump(
            mode="json",
            exclude_none=True,
        )
    if text_style is not None:
        common["text_style"] = text_style
    if heading_level is not None and allow_heading:
        return {
            **common,
            "type": "heading",
            "level": heading_level,
            **text_projection,
        }
    if not text and any(
        feature != "field" for feature in native_features
    ):
        return {
            **common,
            "id": _unique_id("opaque", para_id, index, seen_ids),
            "type": "opaque",
            "summary": "DOCX paragraph containing unsupported drawing or embedded content.",
            "capabilities": ["inspect", "move", "delete", "render"],
            "editable": False,
        }
    return {**common, "type": "paragraph", **text_projection}


def _header_footer_part_projection(
    package: NativePackage,
    *,
    part_uri: str,
    kind: str,
    seen_ids: set[str],
    named_styles: dict[str, NamedStyle],
) -> dict[str, Any]:
    root = parse_xml(package.get_part(part_uri))
    expected_root = _q(W, "hdr" if kind == "header" else "ftr")
    if root.tag != expected_root:
        raise NativePackageError(
            f"Header/footer relationship targets unexpected root {root.tag!r}."
        )
    fingerprint = hashlib.sha256(
        ET.tostring(root, encoding="utf-8")
    ).hexdigest()[:12]
    part_id = _unique_id(kind, fingerprint, 0, seen_ids)
    root_path = "/w:hdr" if kind == "header" else "/w:ftr"
    hyperlinks = _hyperlink_targets(package, source_part=part_uri)
    content: list[dict[str, Any]] = []
    for index, element in enumerate(list(root)):
        complex_features = [
            feature
            for feature, query in (
                ("drawing", f".//{_q(W, 'drawing')}"),
                ("object", f".//{_q(W, 'object')}"),
            )
            if element.find(query) is not None
        ]
        if element.tag == _q(W, "p") and not complex_features:
            content.append(
                _paragraph_projection(
                    element,
                    index,
                    seen_ids,
                    hyperlinks,
                    named_styles,
                    part_uri=part_uri,
                    root_path=root_path,
                    allow_heading=False,
                )
            )
            continue
        native_kind = element.tag.rsplit("}", 1)[-1]
        summary = (
            f"Native {kind} paragraph contains "
            f"{', '.join(complex_features)}."
            if element.tag == _q(W, "p")
            else f"Unsupported native {kind} element {native_kind}."
        )
        content.append(
            {
                "id": _unique_id(
                    "opaque",
                    f"{kind}_{native_kind}_{fingerprint}",
                    index,
                    seen_ids,
                ),
                "type": "opaque",
                "summary": summary,
                "source_ref": _source_ref(
                    [element],
                    [index],
                    native_kind,
                    part_uri=part_uri,
                    root_path=root_path,
                ),
                "capabilities": ["inspect", "render"],
                "editable": False,
                "metadata": {"native_features": complex_features},
            }
        )
    return {
        "id": part_id,
        "type": "header_footer",
        "kind": kind,
        "content": content,
        "source_ref": native_ref_for_header_footer_part(
            root,
            part_uri,
            kind=kind,
        ).model_dump(mode="json", exclude_none=True),
        "metadata": {
            "native_part_uri": part_uri,
            "projection_complete": not any(
                block["type"] == "opaque" for block in content
            ),
        },
    }


def _table_projection(
    element: ET.Element,
    index: int,
    seen_ids: set[str],
) -> dict[str, Any]:
    raw_rows: list[list[str]] = []
    for row in element.findall(f"./{_q(W, 'tr')}"):
        raw_rows.append([_paragraph_text(cell) for cell in row.findall(f"./{_q(W, 'tc')}")])
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
                    column["key"]: values[column_index] if column_index < len(values) else ""
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
    return any(node.attrib.get(_q(W, "type")) == "page" for node in element.iter(_q(W, "br")))


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
    hint = (
        para_id
        or hashlib.sha256(
            b"".join(ET.tostring(element, encoding="utf-8") for element in elements)
        ).hexdigest()[:12]
    )
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
        raise NativePackageError("DOCX package contains multiple AiOffice identity relationships.")
    relationship = relationships[0]
    if relationship.external:
        raise NativePackageError("AiOffice identity relationship cannot be external.")
    target_uri = "/" + relationship.target.lstrip("/")
    if target_uri != MANIFEST_PART_URI:
        raise NativePackageError(
            "AiOffice identity relationship targets an unexpected package part."
        )
    if not package.has_part(MANIFEST_PART_URI):
        raise NativePackageError("DOCX package identity relationship points to a missing manifest.")
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
    header_footer_diagnostics: list[Diagnostic] = []
    document_relationships = {
        relationship.relationship_id: relationship
        for relationship in package.relationships
        if relationship.source_part == "/word/document.xml"
    }
    section_endpoints: list[tuple[int, ET.Element, str]] = []
    for body_index, body_element in enumerate(body_elements):
        if body_element.tag == _q(W, "p"):
            paragraph_section = _paragraph_section(body_element)
            if paragraph_section is not None:
                section_endpoints.append(
                    (body_index, paragraph_section, "paragraph")
                )
        elif body_element.tag == _q(W, "sectPr"):
            section_endpoints.append((body_index, body_element, "body"))
    sections: list[dict[str, Any]] = []
    section_binding_targets: list[dict[str, tuple[str, str]]] = []
    for section_index, (body_index, section, container) in enumerate(
        section_endpoints
    ):
        fingerprint = hashlib.sha256(
            ET.tostring(section, encoding="utf-8")
        ).hexdigest()[:12]
        section_id = _unique_id(
            "section",
            fingerprint,
            section_index,
            seen_ids,
        )
        sections.append(
            {
                "id": section_id,
                "type": "section",
                "start_at": None,
                "layout": read_section_layout(
                    section,
                    first=section_index == 0,
                ).model_dump(mode="json", exclude_none=True),
                "source_ref": native_ref_for_section(
                    section,
                    body_index,
                    container=container,
                ).model_dump(mode="json", exclude_none=True),
                "metadata": {
                    "native_container": container,
                },
            }
        )
        binding_targets: dict[str, tuple[str, str]] = {}
        for reference in list(section):
            if reference.tag == _q(W, "headerReference"):
                kind = "header"
                expected_relationship_type = HEADER_RELATIONSHIP_TYPE
            elif reference.tag == _q(W, "footerReference"):
                kind = "footer"
                expected_relationship_type = FOOTER_RELATIONSHIP_TYPE
            else:
                continue
            variant = reference.get(_q(W, "type"), "default")
            field_name = binding_field(kind, variant)
            relationship_id = reference.get(_q(R, "id"))
            relationship = (
                document_relationships.get(relationship_id)
                if relationship_id is not None
                else None
            )
            if (
                field_name is None
                or relationship is None
                or relationship.external
                or relationship.relationship_type != expected_relationship_type
            ):
                header_footer_diagnostics.append(
                    Diagnostic(
                        severity=Severity.WARNING,
                        code="HEADER_FOOTER_REFERENCE_INVALID",
                        message=(
                            f"Section {section_id!r} contains an invalid "
                            f"{kind} reference."
                        ),
                        node_ids=[section_id],
                        recoverable=True,
                        suggested_actions=[
                            {"action": "inspect_native_header_footer_reference"}
                        ],
                    )
                )
                continue
            part_uri = resolve_relationship_target(
                relationship.source_part,
                relationship.target,
            )
            if not package.has_part(part_uri) or field_name in binding_targets:
                header_footer_diagnostics.append(
                    Diagnostic(
                        severity=Severity.WARNING,
                        code="HEADER_FOOTER_REFERENCE_INVALID",
                        message=(
                            f"Section {section_id!r} has a missing or duplicate "
                            f"{field_name} target."
                        ),
                        node_ids=[section_id],
                        recoverable=True,
                    )
                )
                continue
            binding_targets[field_name] = (kind, part_uri)
        section_binding_targets.append(binding_targets)
    body_index_to_node: dict[int, dict[str, Any]] = {}
    numbering_formats = _numbering_formats(package)
    hyperlink_targets = _hyperlink_targets(package)
    try:
        styles_root = parse_xml(package.get_part("/word/styles.xml"))
    except NativePackageError:
        styles_root = None
    raw_projected_styles = (
        read_named_styles(styles_root) if styles_root is not None else []
    )
    projected_styles: list[NamedStyle] = []
    projected_style_ids: set[str] = set()
    duplicate_style_ids: set[str] = set()
    for style in raw_projected_styles:
        if style.id in projected_style_ids:
            duplicate_style_ids.add(style.id)
            continue
        projected_style_ids.add(style.id)
        projected_styles.append(style)
    named_styles = {style.id: style for style in projected_styles}
    document_defaults = (
        read_document_defaults(styles_root) if styles_root is not None else None
    )
    header_footers: list[dict[str, Any]] = []
    header_footer_by_uri: dict[str, dict[str, Any]] = {}
    section_binding_parts: list[dict[str, dict[str, Any]]] = []
    for section_index, binding_targets in enumerate(section_binding_targets):
        binding_parts: dict[str, dict[str, Any]] = {}
        for field_name, (kind, part_uri) in binding_targets.items():
            projected_part = header_footer_by_uri.get(part_uri)
            if projected_part is None:
                try:
                    projected_part = _header_footer_part_projection(
                        package,
                        part_uri=part_uri,
                        kind=kind,
                        seen_ids=seen_ids,
                        named_styles=named_styles,
                    )
                except NativePackageError as error:
                    header_footer_diagnostics.append(
                        Diagnostic(
                            severity=Severity.WARNING,
                            code="HEADER_FOOTER_PROJECTION_FAILED",
                            message=str(error),
                            node_ids=[sections[section_index]["id"]],
                            recoverable=True,
                        )
                    )
                    continue
                header_footer_by_uri[part_uri] = projected_part
                header_footers.append(projected_part)
            elif projected_part["kind"] != kind:
                header_footer_diagnostics.append(
                    Diagnostic(
                        severity=Severity.WARNING,
                        code="HEADER_FOOTER_REFERENCE_INVALID",
                        message=(
                            f"Part {part_uri!r} is referenced as both a header "
                            "and footer."
                        ),
                        node_ids=[sections[section_index]["id"]],
                        recoverable=True,
                    )
                )
                continue
            binding_parts[field_name] = projected_part
        section_binding_parts.append(binding_parts)
        if binding_parts:
            sections[section_index]["header_footer"] = {
                field_name: part["id"]
                for field_name, part in binding_parts.items()
            }
    index = 0
    while index < len(body_elements):
        element = body_elements[index]
        if element.tag == _q(W, "p"):
            if _is_section_carrier(element):
                index += 1
                continue
            numbering = _paragraph_numbering(element)
            if numbering is not None and numbering in numbering_formats:
                group_elements = [element]
                group_indices = [index]
                next_index = index + 1
                while (
                    next_index < len(body_elements)
                    and _paragraph_section(group_elements[-1]) is None
                ):
                    candidate = body_elements[next_index]
                    if candidate.tag != _q(W, "p") or _paragraph_numbering(candidate) != numbering:
                        break
                    group_elements.append(candidate)
                    group_indices.append(next_index)
                    next_index += 1
                projected = _list_projection(
                    group_elements,
                    group_indices,
                    seen_ids,
                    number_id=numbering[0],
                    level=numbering[1],
                    format_name=numbering_formats[numbering],
                )
                content.append(projected)
                for group_index in group_indices:
                    body_index_to_node[group_index] = projected
                index = next_index
                continue
            if _is_page_break(element):
                projected = _page_break_projection(element, index, seen_ids)
            else:
                projected = _paragraph_projection(
                    element,
                    index,
                    seen_ids,
                    hyperlink_targets,
                    named_styles,
                )
            content.append(projected)
            body_index_to_node[index] = projected
        elif element.tag == _q(W, "tbl"):
            projected = _table_projection(element, index, seen_ids)
            content.append(projected)
            body_index_to_node[index] = projected
        elif element.tag == _q(W, "sectPr"):
            index += 1
            continue
        else:
            native_kind = element.tag.rsplit("}", 1)[-1]
            projected = {
                "id": _unique_id("opaque", native_kind, index, seen_ids),
                "type": "opaque",
                "summary": f"Unsupported DOCX body element {native_kind}.",
                "source_ref": _source_ref([element], [index], native_kind),
                "capabilities": ["inspect", "move", "delete", "render"],
                "editable": False,
            }
            content.append(projected)
            body_index_to_node[index] = projected
        index += 1

    active_identity_manifest = identity_manifest
    identity_source = "workspace" if identity_manifest is not None else None
    if active_identity_manifest is None:
        active_identity_manifest = _embedded_identity_manifest(package)
        if active_identity_manifest is not None:
            identity_source = "embedded"
    diagnostics: list[Diagnostic] = [
        *header_footer_diagnostics,
        *[
            Diagnostic(
                severity=Severity.WARNING,
                code="STYLE_PROJECTION_AMBIGUOUS",
                message=(
                    f"Native DOCX contains duplicate paragraph style ID {style_id!r}; "
                    "the first definition is projected and all native definitions "
                    "are preserved."
                ),
                recoverable=True,
                suggested_actions=[
                    {
                        "action": "repair_duplicate_native_style",
                        "style_id": style_id,
                    }
                ],
            )
            for style_id in sorted(duplicate_style_ids)
        ],
    ]
    if active_identity_manifest is not None:
        if active_identity_manifest.format != "docx":
            raise NativePackageError("Identity manifest format does not match DOCX input.")
        diagnostics.extend(
            apply_identity_manifest(
                content,
                active_identity_manifest,
                package_sha256=package.source_sha256,
                sections=sections,
                header_footers=header_footers,
            )
        )

    for section_index, section in enumerate(sections):
        binding_parts = section_binding_parts[section_index]
        if binding_parts:
            section["header_footer"] = {
                field_name: part["id"]
                for field_name, part in binding_parts.items()
            }
        if section_index == 0:
            section["start_at"] = None
            continue
        previous_endpoint = section_endpoints[section_index - 1][0]
        current_endpoint = section_endpoints[section_index][0]
        start_node = next(
            (
                body_index_to_node[body_index]
                for body_index in range(previous_endpoint + 1, current_endpoint + 1)
                if body_index in body_index_to_node
            ),
            None,
        )
        section["start_at"] = start_node["id"] if start_node is not None else None

    try:
        settings_root = parse_xml(package.get_part("/word/settings.xml"))
    except NativePackageError:
        settings_root = None
    even_and_odd_headers = read_even_and_odd_headers(settings_root)
    update_fields_on_open = read_update_fields_on_open(settings_root)

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
                    active_identity_manifest.revision if active_identity_manifest is not None else 1
                ),
            },
            "metadata": _core_metadata(package),
            "theme": {"ref": "native-docx"},
            "defaults": (
                document_defaults.model_dump(mode="json", exclude_none=True)
                if document_defaults is not None
                else {}
            ),
            "settings": (
                {
                    key: value
                    for key, value in {
                        "even_and_odd_headers": even_and_odd_headers,
                        "update_fields_on_open": update_fields_on_open,
                    }.items()
                    if value is not None
                }
                if even_and_odd_headers is not None
                or update_fields_on_open is not None
                else None
            ),
            "styles": [
                style.model_dump(mode="json", exclude_none=True)
                for style in projected_styles
            ],
            "sections": sections or [{"id": "section_default"}],
            "header_footers": header_footers,
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
