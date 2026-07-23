"""WordprocessingML section projection and minimal native patching."""

from __future__ import annotations

from typing import Literal, cast
from xml.etree import ElementTree as ET

from aioffice.native.identity import native_ref_for_elements
from aioffice.spec.models import (
    ColumnLayout,
    Length,
    NativeRef,
    PageSize,
    SectionColumn,
    SectionLayout,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

SectionStart = Literal[
    "continuous",
    "next_page",
    "even_page",
    "odd_page",
    "next_column",
]
VerticalAlignment = Literal["top", "center", "both", "bottom"]

_START_FROM_NATIVE: dict[str, SectionStart] = {
    "continuous": "continuous",
    "nextPage": "next_page",
    "evenPage": "even_page",
    "oddPage": "odd_page",
    "nextColumn": "next_column",
}
_START_TO_NATIVE = {value: key for key, value in _START_FROM_NATIVE.items()}
_VERTICAL_VALUES = {"top", "center", "both", "bottom"}
_SECTION_ORDER = [
    "headerReference",
    "footerReference",
    "footnotePr",
    "endnotePr",
    "type",
    "pgSz",
    "pgMar",
    "paperSrc",
    "pgBorders",
    "lnNumType",
    "pgNumType",
    "cols",
    "formProt",
    "vAlign",
    "noEndnote",
    "titlePg",
    "textDirection",
    "bidi",
    "rtlGutter",
    "docGrid",
    "printerSettings",
    "sectPrChange",
]
_SECTION_ORDER_INDEX = {
    local_name: index for index, local_name in enumerate(_SECTION_ORDER)
}


def _q(local: str) -> str:
    return f"{{{W}}}{local}"


def _local_name(element: ET.Element) -> str:
    return element.tag.rsplit("}", 1)[-1]


def _length_from_twips(value: str | None) -> Length | None:
    if value is None:
        return None
    try:
        return Length(value=int(value) / 20, unit="pt")
    except (TypeError, ValueError):
        return None


def _twips(value: Length) -> str:
    return str(round(value.to_points() * 20))


def _on_off(value: str | None, *, default: bool) -> bool:
    if value is None:
        return default
    return value.casefold() not in {"0", "false", "off", "no"}


def _ensure_child(section: ET.Element, local_name: str) -> ET.Element:
    existing = section.find(_q(local_name))
    if existing is not None:
        return existing
    child = ET.Element(_q(local_name))
    desired = _SECTION_ORDER_INDEX[local_name]
    for index, sibling in enumerate(list(section)):
        sibling_order = _SECTION_ORDER_INDEX.get(_local_name(sibling))
        if sibling_order is not None and sibling_order > desired:
            section.insert(index, child)
            break
    else:
        section.append(child)
    return child


def _remove_if_empty(parent: ET.Element, child: ET.Element) -> None:
    if not child.attrib and not list(child) and not (child.text or "").strip():
        parent.remove(child)


def _read_page_size(element: ET.Element | None) -> PageSize | None:
    if element is None:
        return None
    width = _length_from_twips(element.get(_q("w")))
    height = _length_from_twips(element.get(_q("h")))
    if width is None or height is None:
        return None
    width_points = width.to_points()
    height_points = height.to_points()
    orientation = (
        "landscape"
        if element.get(_q("orient")) == "landscape" or width_points > height_points
        else "portrait"
    )
    for preset in ("letter", "legal", "executive", "a3", "a4", "tabloid"):
        candidate = PageSize(preset=preset, orientation=orientation)
        expected_width, expected_height = candidate.dimensions_points()
        if (
            abs(expected_width - width_points) <= 0.11
            and abs(expected_height - height_points) <= 0.11
        ):
            return candidate
    return PageSize(
        preset="custom",
        orientation=orientation,
        width=width,
        height=height,
    )


def _read_columns(element: ET.Element | None) -> ColumnLayout | None:
    if element is None:
        return None
    equal_width = _on_off(element.get(_q("equalWidth")), default=True)
    explicit = list(element.findall(_q("col")))
    raw_count = element.get(_q("num"))
    try:
        count = int(raw_count) if raw_count is not None else max(1, len(explicit))
    except ValueError:
        count = max(1, len(explicit))
    if not equal_width and explicit:
        count = len(explicit)
    spacing = _length_from_twips(element.get(_q("space"))) or Length(
        value=36,
        unit="pt",
    )
    columns: list[SectionColumn] = []
    if not equal_width:
        for index, column in enumerate(explicit):
            width = _length_from_twips(column.get(_q("w")))
            if width is None or width.to_points() <= 0:
                return None
            space_after = _length_from_twips(column.get(_q("space")))
            columns.append(
                SectionColumn(
                    width=width,
                    space_after=(
                        space_after
                        if space_after is not None
                        else spacing
                        if index + 1 < len(explicit)
                        else Length(value=0, unit="pt")
                    ),
                )
            )
        if not columns:
            return None
    try:
        return ColumnLayout(
            count=count,
            equal_width=equal_width,
            spacing=spacing,
            separator=_on_off(element.get(_q("sep")), default=False),
            columns=columns,
        )
    except ValueError:
        return None


def read_section_layout(
    section: ET.Element,
    *,
    first: bool,
) -> SectionLayout:
    """Project supported ``w:sectPr`` values without changing native XML."""

    section_type = section.find(_q("type"))
    raw_start = section_type.get(_q("val")) if section_type is not None else None
    start_type: SectionStart | None
    if first:
        start_type = None
    elif raw_start is None:
        start_type = "next_page"
    else:
        start_type = _START_FROM_NATIVE.get(raw_start, "next_page")
    margins = section.find(_q("pgMar"))
    vertical = section.find(_q("vAlign"))
    raw_vertical = vertical.get(_q("val")) if vertical is not None else None
    title_page = section.find(_q("titlePg"))
    return SectionLayout(
        start_type=start_type,
        page_size=_read_page_size(section.find(_q("pgSz"))),
        margin_top=(
            _length_from_twips(margins.get(_q("top"))) if margins is not None else None
        ),
        margin_right=(
            _length_from_twips(margins.get(_q("right")))
            if margins is not None
            else None
        ),
        margin_bottom=(
            _length_from_twips(margins.get(_q("bottom")))
            if margins is not None
            else None
        ),
        margin_left=(
            _length_from_twips(margins.get(_q("left"))) if margins is not None else None
        ),
        gutter=(
            _length_from_twips(margins.get(_q("gutter")))
            if margins is not None
            else None
        ),
        header_distance=(
            _length_from_twips(margins.get(_q("header")))
            if margins is not None
            else None
        ),
        footer_distance=(
            _length_from_twips(margins.get(_q("footer")))
            if margins is not None
            else None
        ),
        columns=_read_columns(section.find(_q("cols"))),
        vertical_alignment=(
            cast(VerticalAlignment, raw_vertical)
            if raw_vertical in _VERTICAL_VALUES
            else None
        ),
        different_first_page=(
            _on_off(title_page.get(_q("val")), default=True)
            if title_page is not None
            else False
        ),
    )


def patch_section_layout(
    section: ET.Element,
    layout: SectionLayout,
    fields: set[str],
) -> None:
    """Patch selected supported fields while preserving all unknown XML."""

    if "start_type" in fields:
        existing = section.find(_q("type"))
        if layout.start_type is None:
            if existing is not None:
                section.remove(existing)
        else:
            element = existing if existing is not None else _ensure_child(section, "type")
            element.set(_q("val"), _START_TO_NATIVE[layout.start_type])

    if "page_size" in fields:
        existing = section.find(_q("pgSz"))
        if layout.page_size is None:
            if existing is not None:
                section.remove(existing)
        else:
            element = existing if existing is not None else _ensure_child(section, "pgSz")
            width, height = layout.page_size.dimensions_points()
            element.set(_q("w"), str(round(width * 20)))
            element.set(_q("h"), str(round(height * 20)))
            if layout.page_size.orientation == "landscape":
                element.set(_q("orient"), "landscape")
            else:
                element.attrib.pop(_q("orient"), None)

    margin_fields = {
        "margin_top": "top",
        "margin_right": "right",
        "margin_bottom": "bottom",
        "margin_left": "left",
        "gutter": "gutter",
        "header_distance": "header",
        "footer_distance": "footer",
    }
    selected_margins = fields.intersection(margin_fields)
    if selected_margins:
        margins = section.find(_q("pgMar"))
        if margins is None and any(
            getattr(layout, field_name) is not None
            for field_name in selected_margins
        ):
            margins = _ensure_child(section, "pgMar")
        if margins is not None:
            for field_name in selected_margins:
                value = getattr(layout, field_name)
                attribute = _q(margin_fields[field_name])
                if value is None:
                    margins.attrib.pop(attribute, None)
                else:
                    margins.set(attribute, _twips(value))
            _remove_if_empty(section, margins)

    if "columns" in fields:
        existing = section.find(_q("cols"))
        if layout.columns is None:
            if existing is not None:
                section.remove(existing)
        else:
            element = existing if existing is not None else _ensure_child(section, "cols")
            columns = layout.columns
            element.set(_q("num"), str(columns.count))
            element.set(_q("equalWidth"), "1" if columns.equal_width else "0")
            element.set(_q("space"), _twips(columns.spacing))
            element.set(_q("sep"), "1" if columns.separator else "0")
            for child in list(element):
                if child.tag == _q("col"):
                    element.remove(child)
            if not columns.equal_width:
                for column in columns.columns:
                    ET.SubElement(
                        element,
                        _q("col"),
                        {
                            _q("w"): _twips(column.width),
                            _q("space"): _twips(column.space_after),
                        },
                    )

    if "vertical_alignment" in fields:
        existing = section.find(_q("vAlign"))
        if layout.vertical_alignment is None:
            if existing is not None:
                section.remove(existing)
        else:
            element = existing if existing is not None else _ensure_child(section, "vAlign")
            element.set(_q("val"), layout.vertical_alignment)

    if "different_first_page" in fields:
        existing = section.find(_q("titlePg"))
        if layout.different_first_page is None:
            if existing is not None:
                section.remove(existing)
        else:
            element = (
                existing
                if existing is not None
                else _ensure_child(section, "titlePg")
            )
            element.set(_q("val"), "1" if layout.different_first_page else "0")


def apply_section_layout(section: ET.Element, layout: SectionLayout) -> None:
    """Write every explicit semantic section property."""

    patch_section_layout(
        section,
        layout,
        {
            field_name
            for field_name in SectionLayout.model_fields
            if getattr(layout, field_name) is not None
        },
    )


def native_ref_for_section(
    section: ET.Element,
    body_index: int,
    *,
    container: str,
) -> NativeRef:
    """Build a stable source reference for a body or paragraph ``w:sectPr``."""

    if container not in {"body", "paragraph"}:
        raise ValueError("Section container must be body or paragraph.")
    source_ref = native_ref_for_elements(
        [section],
        [body_index],
        native_kind=f"w:sectPr-{container}",
    )
    suffix = "" if container == "body" else "/w:pPr/w:sectPr"
    return source_ref.model_copy(
        update={
            "path_hint": (
                f"/w:document/w:body/*[{body_index + 1}]"
                f"{suffix}"
            )
        }
    )


__all__ = [
    "apply_section_layout",
    "native_ref_for_section",
    "patch_section_layout",
    "read_section_layout",
]
