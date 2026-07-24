"""DOCX header/footer bindings, part references, and document settings."""

from __future__ import annotations

from collections.abc import Iterable, Mapping
import posixpath
from xml.etree import ElementTree as ET

from aioffice.native.identity import fingerprint_elements
from aioffice.spec.models import HeaderFooterBindings, NativeRef

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"

HEADER_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
)
FOOTER_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
)
SETTINGS_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
)
HEADER_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument."
    "wordprocessingml.header+xml"
)
FOOTER_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument."
    "wordprocessingml.footer+xml"
)
SETTINGS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument."
    "wordprocessingml.settings+xml"
)

HEADER_FOOTER_BINDING_SLOTS = (
    ("header_default", "header", "default"),
    ("header_first", "header", "first"),
    ("header_even", "header", "even"),
    ("footer_default", "footer", "default"),
    ("footer_first", "footer", "first"),
    ("footer_even", "footer", "even"),
)


def _q(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def resolve_relationship_target(source_part: str, target: str) -> str:
    """Resolve an internal OPC relationship target to an absolute part URI."""

    if target.startswith("/"):
        normalized = posixpath.normpath(target)
    else:
        base = posixpath.dirname(source_part)
        normalized = posixpath.normpath(posixpath.join(base, target))
    return "/" + normalized.lstrip("/")


def binding_field(kind: str, variant: str) -> str | None:
    return next(
        (
            field_name
            for field_name, candidate_kind, candidate_variant
            in HEADER_FOOTER_BINDING_SLOTS
            if candidate_kind == kind and candidate_variant == variant
        ),
        None,
    )


def binding_slot(field_name: str) -> tuple[str, str] | None:
    return next(
        (
            (kind, variant)
            for candidate, kind, variant in HEADER_FOOTER_BINDING_SLOTS
            if candidate == field_name
        ),
        None,
    )


def binding_references(
    section: ET.Element,
    field_name: str,
) -> list[ET.Element]:
    slot = binding_slot(field_name)
    if slot is None:
        raise ValueError(f"Unknown header/footer binding field {field_name!r}.")
    kind, variant = slot
    return [
        child
        for child in list(section)
        if child.tag == _q(W, f"{kind}Reference")
        and child.get(_q(W, "type"), "default") == variant
    ]


def patch_header_footer_bindings(
    section: ET.Element,
    *,
    set_relationship_ids: Mapping[str, str],
    clear: Iterable[str],
) -> bool:
    """Patch selected section references while preserving all other XML."""

    clear_fields = tuple(clear)
    known_fields = {
        field_name
        for field_name, _, _ in HEADER_FOOTER_BINDING_SLOTS
    }
    unknown = (
        set(set_relationship_ids)
        | set(clear_fields)
    ) - known_fields
    overlap = set(set_relationship_ids) & set(clear_fields)
    if unknown:
        raise ValueError(
            "Unknown header/footer binding fields: "
            f"{', '.join(sorted(unknown))}."
        )
    if overlap:
        raise ValueError(
            "Header/footer binding fields cannot be both set and cleared: "
            f"{', '.join(sorted(overlap))}."
        )

    changed = False
    for field_name, _, _ in HEADER_FOOTER_BINDING_SLOTS:
        references = binding_references(section, field_name)
        if len(references) > 1:
            raise ValueError(
                f"Section contains duplicate {field_name} references."
            )
        if field_name in clear_fields:
            if references:
                section.remove(references[0])
                changed = True
            continue
        relationship_id = set_relationship_ids.get(field_name)
        if relationship_id is None:
            continue
        if not relationship_id:
            raise ValueError(
                f"{field_name} relationship ID cannot be empty."
            )
        if references:
            reference = references[0]
            if reference.get(_q(R, "id")) != relationship_id:
                reference.set(_q(R, "id"), relationship_id)
                changed = True
            continue

        slot = binding_slot(field_name)
        assert slot is not None
        kind, variant = slot
        reference = ET.Element(
            _q(W, f"{kind}Reference"),
            {
                _q(R, "id"): relationship_id,
                _q(W, "type"): variant,
            },
        )
        children = list(section)
        if kind == "header":
            insert_index = next(
                (
                    index
                    for index, child in enumerate(children)
                    if child.tag != _q(W, "headerReference")
                ),
                len(children),
            )
        else:
            insert_index = next(
                (
                    index
                    for index, child in enumerate(children)
                    if child.tag
                    not in {
                        _q(W, "headerReference"),
                        _q(W, "footerReference"),
                    }
                ),
                len(children),
            )
        section.insert(insert_index, reference)
        changed = True
    return changed


def native_ref_for_header_footer_part(
    root: ET.Element,
    part_uri: str,
    *,
    kind: str,
) -> NativeRef:
    root_name = "hdr" if kind == "header" else "ftr"
    return NativeRef(
        format="docx",
        part_uri=part_uri,
        native_kind=f"w:{root_name}-part",
        path_hint=f"/w:{root_name}",
        fingerprint=fingerprint_elements([root]),
    )


def apply_header_footer_bindings(
    section: ET.Element,
    bindings: HeaderFooterBindings | None,
    relationship_ids: dict[str, str],
) -> None:
    """Write explicit references; omitted slots retain Word inheritance."""

    if bindings is None:
        return
    insert_index = 0
    for field_name, kind, variant in HEADER_FOOTER_BINDING_SLOTS:
        part_id = getattr(bindings, field_name)
        if part_id is None:
            continue
        try:
            relationship_id = relationship_ids[part_id]
        except KeyError as error:
            raise ValueError(
                f"Header/footer part {part_id!r} has no document relationship."
            ) from error
        reference = ET.Element(
            _q(W, f"{kind}Reference"),
            {
                _q(R, "id"): relationship_id,
                _q(W, "type"): variant,
            },
        )
        section.insert(insert_index, reference)
        insert_index += 1


def read_even_and_odd_headers(settings_root: ET.Element | None) -> bool | None:
    if settings_root is None:
        return None
    element = settings_root.find(_q(W, "evenAndOddHeaders"))
    if element is None:
        return None
    value = element.get(_q(W, "val"))
    return value is None or value.casefold() not in {"0", "false", "off", "no"}


def read_update_fields_on_open(settings_root: ET.Element | None) -> bool | None:
    if settings_root is None:
        return None
    element = settings_root.find(_q(W, "updateFields"))
    if element is None:
        return None
    value = element.get(_q(W, "val"))
    return value is None or value.casefold() not in {"0", "false", "off", "no"}


def settings_xml(
    *,
    even_and_odd_headers: bool | None = None,
    update_fields_on_open: bool | None = None,
) -> bytes:
    root = ET.Element(_q(W, "settings"))
    if even_and_odd_headers is not None:
        ET.SubElement(
            root,
            _q(W, "evenAndOddHeaders"),
            {_q(W, "val"): "1" if even_and_odd_headers else "0"},
        )
    if update_fields_on_open is not None:
        ET.SubElement(
            root,
            _q(W, "updateFields"),
            {_q(W, "val"): "1" if update_fields_on_open else "0"},
        )
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


__all__ = [
    "FOOTER_CONTENT_TYPE",
    "FOOTER_RELATIONSHIP_TYPE",
    "HEADER_FOOTER_BINDING_SLOTS",
    "HEADER_CONTENT_TYPE",
    "HEADER_RELATIONSHIP_TYPE",
    "SETTINGS_CONTENT_TYPE",
    "SETTINGS_RELATIONSHIP_TYPE",
    "apply_header_footer_bindings",
    "binding_field",
    "binding_references",
    "binding_slot",
    "native_ref_for_header_footer_part",
    "patch_header_footer_bindings",
    "read_even_and_odd_headers",
    "read_update_fields_on_open",
    "resolve_relationship_target",
    "settings_xml",
]
