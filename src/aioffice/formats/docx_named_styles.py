"""Projection and minimal mutation of WordprocessingML paragraph styles."""

from __future__ import annotations

import re
from collections.abc import Iterable
from xml.etree import ElementTree as ET

from aioffice.core.errors import NativePackageError
from aioffice.formats.docx_style import (
    apply_paragraph_style,
    apply_text_style,
    patch_paragraph_style,
    patch_text_style,
    read_paragraph_style,
    read_text_style,
)
from aioffice.spec.models import (
    DocumentDefaults,
    NamedStyle,
    ParagraphStyle,
    SemanticStyleRole,
    TextStyle,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
_FALSE_VALUES = {"0", "false", "off", "no", "none"}
_STYLE_CHILD_ORDER = {
    name: index
    for index, name in enumerate(
        (
            "name",
            "aliases",
            "basedOn",
            "next",
            "link",
            "autoRedefine",
            "hidden",
            "uiPriority",
            "semiHidden",
            "unhideWhenUsed",
            "qFormat",
            "locked",
            "personal",
            "personalCompose",
            "personalReply",
            "rsid",
            "pPr",
            "rPr",
            "tblPr",
            "trPr",
            "tcPr",
            "tblStylePr",
        )
    )
}


def _q(local: str) -> str:
    return f"{{{W}}}{local}"


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _ensure_style_child(parent: ET.Element, name: str) -> ET.Element:
    existing = parent.find(_q(name))
    if existing is not None:
        return existing
    child = ET.Element(_q(name))
    rank = _STYLE_CHILD_ORDER.get(name, 10_000)
    for index, candidate in enumerate(list(parent)):
        if _STYLE_CHILD_ORDER.get(_local(candidate.tag), 10_000) > rank:
            parent.insert(index, child)
            break
    else:
        parent.append(child)
    return child


def _native_bool(parent: ET.Element, name: str) -> bool:
    child = parent.find(_q(name))
    if child is None:
        return False
    value = child.attrib.get(_q("val"))
    return value is None or value.lower() not in _FALSE_VALUES


def _set_native_bool(parent: ET.Element, name: str, value: bool | None) -> None:
    child = parent.find(_q(name))
    if value is None:
        if child is not None:
            parent.remove(child)
        return
    child = _ensure_style_child(parent, name)
    if value:
        child.attrib.pop(_q("val"), None)
    else:
        child.set(_q("val"), "0")


def _style_ref(parent: ET.Element, name: str) -> str | None:
    child = parent.find(_q(name))
    if child is None:
        return None
    return child.attrib.get(_q("val")) or None


def _semantic_role(
    style_id: str,
    name: str,
    paragraph_style: ParagraphStyle | None,
) -> tuple[SemanticStyleRole, int | None]:
    if paragraph_style is not None and paragraph_style.outline_level is not None:
        return "heading", paragraph_style.outline_level
    heading_match = re.fullmatch(
        r"heading[\s_-]*([1-9])",
        style_id,
        flags=re.IGNORECASE,
    ) or re.fullmatch(r"heading[\s_-]*([1-9])", name, flags=re.IGNORECASE)
    if heading_match:
        return "heading", int(heading_match.group(1))
    normalized = re.sub(r"[\s_-]+", "", f"{style_id} {name}").lower()
    candidates_by_role: tuple[
        tuple[SemanticStyleRole, tuple[str, ...]],
        ...,
    ] = (
        ("body", ("normal", "bodytext")),
        ("subtitle", ("subtitle",)),
        ("title", ("title",)),
        ("quote", ("quote", "intensequote")),
        ("caption", ("caption",)),
        ("code", ("code", "codeblock")),
        ("list", ("listparagraph",)),
    )
    for role, candidates in candidates_by_role:
        if any(candidate in normalized for candidate in candidates):
            return role, None
    return "custom", None


def read_document_defaults(styles_root: ET.Element) -> DocumentDefaults:
    defaults = styles_root.find(_q("docDefaults"))
    if defaults is None:
        return DocumentDefaults()
    paragraph_default = defaults.find(_q("pPrDefault"))
    run_default = defaults.find(_q("rPrDefault"))
    return DocumentDefaults(
        paragraph_style=(
            read_paragraph_style(paragraph_default)
            if paragraph_default is not None
            else None
        ),
        text_style=read_text_style(run_default) if run_default is not None else None,
    )


def read_named_styles(styles_root: ET.Element) -> list[NamedStyle]:
    projected: list[NamedStyle] = []
    for element in styles_root.findall(_q("style")):
        if element.attrib.get(_q("type"), "paragraph") != "paragraph":
            continue
        style_id = element.attrib.get(_q("styleId"))
        if not style_id:
            continue
        name = _style_ref(element, "name") or style_id
        paragraph_style = read_paragraph_style(element)
        text_style = read_text_style(element)
        semantic_role, heading_level = _semantic_role(style_id, name, paragraph_style)
        projected.append(
            NamedStyle(
                id=style_id,
                name=name,
                semantic_role=semantic_role,
                heading_level=heading_level,
                based_on=_style_ref(element, "basedOn"),
                next_style=_style_ref(element, "next"),
                paragraph_style=paragraph_style,
                text_style=text_style,
                quick_style=_native_bool(element, "qFormat"),
                hidden=_native_bool(element, "hidden"),
                metadata={
                    "native_custom_style": (
                        element.attrib.get(_q("customStyle"), "").lower()
                        not in {"", *_FALSE_VALUES}
                    ),
                },
            )
        )

    by_id = {style.id: style for style in projected}

    def inherited_heading(style: NamedStyle, visiting: set[str]) -> int | None:
        if style.heading_level is not None:
            return style.heading_level
        if style.id in visiting or style.based_on is None:
            return None
        parent = by_id.get(style.based_on)
        if parent is None:
            return None
        return inherited_heading(parent, visiting | {style.id})

    result: list[NamedStyle] = []
    for style in projected:
        level = inherited_heading(style, set())
        if level is not None and style.semantic_role != "heading":
            style = style.model_copy(
                update={"semantic_role": "heading", "heading_level": level}
            )
        result.append(style)
    return result


def find_named_style(styles_root: ET.Element, style_id: str) -> ET.Element | None:
    matches = [
        element
        for element in styles_root.findall(_q("style"))
        if element.attrib.get(_q("type"), "paragraph") == "paragraph"
        and element.attrib.get(_q("styleId")) == style_id
    ]
    if len(matches) > 1:
        raise NativePackageError(
            f"Native DOCX paragraph style {style_id!r} is ambiguous; "
            f"{len(matches)} definitions use that ID."
        )
    return matches[0] if matches else None


def _set_style_ref(parent: ET.Element, name: str, value: str | None) -> None:
    child = parent.find(_q(name))
    if value is None:
        if child is not None:
            parent.remove(child)
        return
    child = _ensure_style_child(parent, name)
    child.set(_q("val"), value)


def upsert_named_style(
    styles_root: ET.Element,
    style: NamedStyle,
    *,
    custom_style: bool,
) -> ET.Element:
    """Create or fully lower supported fields of one paragraph style."""

    element = find_named_style(styles_root, style.id)
    if element is None:
        attributes = {
            _q("type"): "paragraph",
            _q("styleId"): style.id,
        }
        if custom_style:
            attributes[_q("customStyle")] = "1"
        element = ET.SubElement(styles_root, _q("style"), attributes)
    name = _ensure_style_child(element, "name")
    name.set(_q("val"), style.name)
    _set_style_ref(element, "basedOn", style.based_on)
    _set_style_ref(element, "next", style.next_style)
    _set_native_bool(element, "qFormat", style.quick_style)
    _set_native_bool(element, "hidden", style.hidden)
    apply_paragraph_style(element, style.paragraph_style)
    apply_text_style(element, style.text_style)
    return element


def format_named_style(
    styles_root: ET.Element,
    style_id: str,
    *,
    paragraph_style: ParagraphStyle,
    paragraph_fields: Iterable[str],
    text_style: TextStyle,
    text_fields: Iterable[str],
) -> None:
    element = find_named_style(styles_root, style_id)
    if element is None:
        raise NativePackageError(f"Native DOCX has no paragraph style {style_id!r}.")
    patch_paragraph_style(element, paragraph_style, paragraph_fields)
    patch_text_style(element, text_style, text_fields)


__all__ = [
    "find_named_style",
    "format_named_style",
    "read_document_defaults",
    "read_named_styles",
    "upsert_named_style",
]
