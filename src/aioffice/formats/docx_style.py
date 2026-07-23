"""Shared direct-formatting projection and lowering for WordprocessingML."""

from __future__ import annotations

from collections.abc import Iterable
from xml.etree import ElementTree as ET

from aioffice.formats.docx_borders import (
    clear_border_element,
    read_border_element,
    write_border_element,
)
from aioffice.spec.models import (
    Length,
    LineSpacing,
    ParagraphBorders,
    ParagraphStyle,
    TextStyle,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_PPR_ORDER = {
    name: index
    for index, name in enumerate(
        (
            "pStyle",
            "keepNext",
            "keepLines",
            "pageBreakBefore",
            "widowControl",
            "numPr",
            "pBdr",
            "shd",
            "tabs",
            "spacing",
            "ind",
            "jc",
            "outlineLvl",
            "rPr",
            "sectPr",
            "pPrChange",
        )
    )
}
_RPR_ORDER = {
    name: index
    for index, name in enumerate(
        (
            "rStyle",
            "rFonts",
            "b",
            "i",
            "caps",
            "smallCaps",
            "strike",
            "color",
            "spacing",
            "sz",
            "szCs",
            "highlight",
            "u",
            "shd",
            "vertAlign",
            "rPrChange",
        )
    )
}
_FALSE_VALUES = {"0", "false", "off", "no", "none"}
_PARAGRAPH_BORDER_ORDER = {
    name: index
    for index, name in enumerate(
        ("top", "left", "bottom", "right", "between", "bar")
    )
}
_PARAGRAPH_BORDER_FIELDS = (
    ("top", "top"),
    ("right", "right"),
    ("bottom", "bottom"),
    ("left", "left"),
)
_SHADING_ATTRIBUTES = (
    "val",
    "color",
    "fill",
    "themeColor",
    "themeTint",
    "themeShade",
    "themeFill",
    "themeFillTint",
    "themeFillShade",
)


def _q(local: str) -> str:
    return f"{{{W}}}{local}"


def _local(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _length_from_twips(value: str | None) -> Length | None:
    if value is None:
        return None
    try:
        return Length(value=int(value) / 20.0, unit="pt")
    except (TypeError, ValueError):
        return None


def _twips(value: Length) -> str:
    return str(round(value.to_points() * 20))


def _half_points(value: Length) -> str:
    return str(round(value.to_points() * 2))


def _bool_property(parent: ET.Element, name: str) -> bool | None:
    element = parent.find(_q(name))
    if element is None:
        return None
    value = element.attrib.get(_q("val"))
    return value is None or value.lower() not in _FALSE_VALUES


def _ensure_ordered_child(
    parent: ET.Element,
    name: str,
    order: dict[str, int],
) -> ET.Element:
    existing = parent.find(_q(name))
    if existing is not None:
        return existing
    child = ET.Element(_q(name))
    rank = order.get(name, 10_000)
    for index, candidate in enumerate(list(parent)):
        if order.get(_local(candidate.tag), 10_000) > rank:
            parent.insert(index, child)
            break
    else:
        parent.append(child)
    return child


def _remove_child(parent: ET.Element, name: str) -> None:
    child = parent.find(_q(name))
    if child is not None:
        parent.remove(child)


def _set_boolean(
    parent: ET.Element,
    name: str,
    value: bool | None,
    order: dict[str, int],
) -> None:
    if value is None:
        _remove_child(parent, name)
        return
    child = _ensure_ordered_child(parent, name, order)
    if value:
        child.attrib.pop(_q("val"), None)
    else:
        child.set(_q("val"), "0")


def _remove_if_empty(parent: ET.Element, child: ET.Element) -> None:
    if not child.attrib and not list(child) and not (child.text or ""):
        parent.remove(child)


def _read_paragraph_borders(
    properties: ET.Element,
) -> ParagraphBorders | None:
    borders = properties.find(_q("pBdr"))
    if borders is None:
        return None
    payload = {
        field_name: value
        for field_name, native_name in _PARAGRAPH_BORDER_FIELDS
        if (
            value := read_border_element(
                borders.find(_q(native_name))
            )
        )
        is not None
    }
    return (
        ParagraphBorders.model_validate(payload)
        if payload
        else None
    )


def _read_paragraph_background(
    properties: ET.Element,
) -> str | None:
    shading = properties.find(_q("shd"))
    if shading is None or shading.get(_q("val")) != "clear":
        return None
    if any(
        shading.get(_q(attribute)) is not None
        for attribute in (
            "themeFill",
            "themeFillTint",
            "themeFillShade",
        )
    ):
        return None
    fill = shading.get(_q("fill"), "")
    if len(fill) != 6 or not all(
        character in "0123456789ABCDEFabcdef"
        for character in fill
    ):
        return None
    return f"#{fill.upper()}"


def _clear_paragraph_border(
    borders: ET.Element,
    native_name: str,
) -> None:
    element = borders.find(_q(native_name))
    if element is None:
        return
    clear_border_element(element)
    _remove_if_empty(borders, element)


def _patch_paragraph_borders(
    properties: ET.Element,
    value: ParagraphBorders | None,
) -> None:
    borders = properties.find(_q("pBdr"))
    has_values = value is not None and any(
        getattr(value, field_name) is not None
        for field_name, _ in _PARAGRAPH_BORDER_FIELDS
    )
    if borders is None and has_values:
        borders = _ensure_ordered_child(
            properties,
            "pBdr",
            _PPR_ORDER,
        )
    if borders is None:
        return
    for field_name, native_name in _PARAGRAPH_BORDER_FIELDS:
        border = (
            getattr(value, field_name)
            if value is not None
            else None
        )
        if border is None:
            _clear_paragraph_border(borders, native_name)
            continue
        element = borders.find(_q(native_name))
        if element is None:
            element = _ensure_ordered_child(
                borders,
                native_name,
                _PARAGRAPH_BORDER_ORDER,
            )
        write_border_element(element, border)
    _remove_if_empty(properties, borders)


def _clear_shading_attributes(shading: ET.Element) -> None:
    for attribute in _SHADING_ATTRIBUTES:
        shading.attrib.pop(_q(attribute), None)


def _patch_paragraph_background(
    properties: ET.Element,
    value: str | None,
) -> None:
    shading = properties.find(_q("shd"))
    if value is None:
        if shading is not None:
            _clear_shading_attributes(shading)
            _remove_if_empty(properties, shading)
        return
    if shading is None:
        shading = _ensure_ordered_child(
            properties,
            "shd",
            _PPR_ORDER,
        )
    _clear_shading_attributes(shading)
    shading.set(_q("val"), "clear")
    shading.set(_q("color"), "auto")
    shading.set(_q("fill"), value.removeprefix("#"))


def read_paragraph_style(paragraph: ET.Element) -> ParagraphStyle | None:
    """Project supported direct ``w:pPr`` values without resolving named styles."""

    properties = paragraph.find(_q("pPr"))
    if properties is None:
        return None
    values: dict[str, object] = {}
    background_color = _read_paragraph_background(properties)
    if background_color is not None:
        values["background_color"] = background_color
    borders = _read_paragraph_borders(properties)
    if borders is not None:
        values["borders"] = borders
    alignment = properties.find(_q("jc"))
    if alignment is not None:
        value = alignment.attrib.get(_q("val"))
        if value in {"left", "center", "right", "both", "distribute"}:
            values["alignment"] = "justify" if value == "both" else value

    spacing = properties.find(_q("spacing"))
    if spacing is not None:
        before = _length_from_twips(spacing.attrib.get(_q("before")))
        after = _length_from_twips(spacing.attrib.get(_q("after")))
        if before is not None:
            values["spacing_before"] = before
        if after is not None:
            values["spacing_after"] = after
        line = spacing.attrib.get(_q("line"))
        if line is not None:
            try:
                line_value = int(line)
            except ValueError:
                line_value = 0
            rule = spacing.attrib.get(_q("lineRule"), "auto")
            if line_value > 0 and rule == "auto":
                values["line_spacing"] = LineSpacing(
                    rule="multiple",
                    value=line_value / 240.0,
                )
            elif line_value > 0 and rule in {"exact", "atLeast"}:
                values["line_spacing"] = LineSpacing(
                    rule="exact" if rule == "exact" else "at_least",
                    value=Length(value=line_value / 20.0, unit="pt"),
                )

    indentation = properties.find(_q("ind"))
    if indentation is not None:
        for attribute, field_name in (
            ("left", "indent_left"),
            ("right", "indent_right"),
            ("firstLine", "first_line_indent"),
            ("hanging", "hanging_indent"),
        ):
            length = _length_from_twips(indentation.attrib.get(_q(attribute)))
            if length is not None:
                values[field_name] = length

    for native_name, field_name in (
        ("keepNext", "keep_with_next"),
        ("keepLines", "keep_together"),
        ("pageBreakBefore", "page_break_before"),
        ("widowControl", "widow_control"),
    ):
        value = _bool_property(properties, native_name)
        if value is not None:
            values[field_name] = value
    outline = properties.find(_q("outlineLvl"))
    if outline is not None:
        try:
            native_level = int(outline.attrib[_q("val")])
        except (KeyError, ValueError):
            native_level = -1
        if 0 <= native_level <= 8:
            values["outline_level"] = native_level + 1
    return ParagraphStyle.model_validate(values) if values else None


def read_text_style(run: ET.Element) -> TextStyle | None:
    """Project supported direct ``w:rPr`` values for one native run."""

    properties = run.find(_q("rPr"))
    if properties is None:
        return None
    values: dict[str, object] = {}
    fonts = properties.find(_q("rFonts"))
    if fonts is not None:
        latin = fonts.attrib.get(_q("ascii")) or fonts.attrib.get(_q("hAnsi"))
        east_asia = fonts.attrib.get(_q("eastAsia"))
        if latin:
            values["font_family"] = latin
        if east_asia:
            values["font_family_east_asia"] = east_asia
    size = properties.find(_q("sz"))
    if size is not None:
        try:
            values["font_size"] = Length(
                value=int(size.attrib[_q("val")]) / 2.0,
                unit="pt",
            )
        except (KeyError, ValueError):
            pass
    color = properties.find(_q("color"))
    if color is not None:
        value = color.attrib.get(_q("val"), "")
        if len(value) == 6 and all(character in "0123456789abcdefABCDEF" for character in value):
            values["color"] = f"#{value}"
    shading = properties.find(_q("shd"))
    if shading is not None:
        value = shading.attrib.get(_q("fill"), "")
        if len(value) == 6 and all(character in "0123456789abcdefABCDEF" for character in value):
            values["background_color"] = f"#{value}"
    for native_name, field_name in (
        ("b", "bold"),
        ("i", "italic"),
        ("u", "underline"),
        ("strike", "strike"),
        ("smallCaps", "small_caps"),
        ("caps", "all_caps"),
    ):
        value = _bool_property(properties, native_name)
        if value is not None:
            values[field_name] = value
    spacing = properties.find(_q("spacing"))
    if spacing is not None:
        length = _length_from_twips(spacing.attrib.get(_q("val")))
        if length is not None:
            values["letter_spacing"] = length
    vertical = properties.find(_q("vertAlign"))
    if vertical is not None:
        value = vertical.attrib.get(_q("val"))
        if value in {"baseline", "superscript", "subscript"}:
            values["baseline"] = "normal" if value == "baseline" else value
    return TextStyle.model_validate(values) if values else None


def common_text_style(paragraph: ET.Element) -> TextStyle | None:
    """Return direct properties shared by every text-bearing native run."""

    styles = [
        read_text_style(run) or TextStyle()
        for run in paragraph.iter(_q("r"))
        if any((text.text or "") for text in run.iter(_q("t")))
    ]
    if not styles:
        paragraph_properties = paragraph.find(_q("pPr"))
        return (
            read_text_style(paragraph_properties)
            if paragraph_properties is not None
            else None
        )
    values = {
        field_name: value
        for field_name in TextStyle.model_fields
        if (value := getattr(styles[0], field_name)) is not None
        and all(getattr(style, field_name) == value for style in styles[1:])
    }
    return TextStyle.model_validate(values) if values else None


def patch_paragraph_style(
    paragraph: ET.Element,
    style: ParagraphStyle | None,
    fields: Iterable[str],
) -> None:
    """Update only selected supported properties and preserve every other child."""

    properties = paragraph.find(_q("pPr"))
    if properties is None:
        properties = ET.Element(_q("pPr"))
        paragraph.insert(0, properties)
    values = style or ParagraphStyle()
    selected = set(fields)

    if "alignment" in selected:
        value = values.alignment
        if value is None:
            _remove_child(properties, "jc")
        else:
            child = _ensure_ordered_child(properties, "jc", _PPR_ORDER)
            child.set(_q("val"), "both" if value == "justify" else value)

    if "borders" in selected:
        _patch_paragraph_borders(properties, values.borders)

    if "background_color" in selected:
        _patch_paragraph_background(
            properties,
            values.background_color,
        )

    spacing_fields = {"spacing_before", "spacing_after", "line_spacing"}
    if selected & spacing_fields:
        spacing = _ensure_ordered_child(properties, "spacing", _PPR_ORDER)
        if "spacing_before" in selected:
            spacing.attrib.pop(_q("beforeLines"), None)
            value = values.spacing_before
            if value is None:
                spacing.attrib.pop(_q("before"), None)
            else:
                spacing.set(_q("before"), _twips(value))
        if "spacing_after" in selected:
            spacing.attrib.pop(_q("afterLines"), None)
            value = values.spacing_after
            if value is None:
                spacing.attrib.pop(_q("after"), None)
            else:
                spacing.set(_q("after"), _twips(value))
        if "line_spacing" in selected:
            value = values.line_spacing
            if value is None:
                spacing.attrib.pop(_q("line"), None)
                spacing.attrib.pop(_q("lineRule"), None)
            elif value.rule == "multiple":
                assert isinstance(value.value, float)
                spacing.set(_q("line"), str(round(value.value * 240)))
                spacing.set(_q("lineRule"), "auto")
            else:
                assert isinstance(value.value, Length)
                spacing.set(_q("line"), _twips(value.value))
                spacing.set(
                    _q("lineRule"),
                    "exact" if value.rule == "exact" else "atLeast",
                )
        _remove_if_empty(properties, spacing)

    indentation_fields = {
        "indent_left",
        "indent_right",
        "first_line_indent",
        "hanging_indent",
    }
    if selected & indentation_fields:
        indentation = _ensure_ordered_child(properties, "ind", _PPR_ORDER)
        for field_name, native_name in (
            ("indent_left", "left"),
            ("indent_right", "right"),
            ("first_line_indent", "firstLine"),
            ("hanging_indent", "hanging"),
        ):
            if field_name not in selected:
                continue
            value = getattr(values, field_name)
            if value is None:
                indentation.attrib.pop(_q(native_name), None)
            else:
                indentation.set(_q(native_name), _twips(value))
        _remove_if_empty(properties, indentation)

    for field_name, native_name in (
        ("keep_with_next", "keepNext"),
        ("keep_together", "keepLines"),
        ("page_break_before", "pageBreakBefore"),
        ("widow_control", "widowControl"),
    ):
        if field_name in selected:
            _set_boolean(
                properties,
                native_name,
                getattr(values, field_name),
                _PPR_ORDER,
            )
    if "outline_level" in selected:
        value = values.outline_level
        if value is None:
            _remove_child(properties, "outlineLvl")
        else:
            child = _ensure_ordered_child(properties, "outlineLvl", _PPR_ORDER)
            child.set(_q("val"), str(value - 1))
    _remove_if_empty(paragraph, properties)


def patch_paragraph_style_ref(
    paragraph: ET.Element,
    style_id: str | None,
) -> None:
    """Set or clear exactly the native ``w:pStyle`` reference."""

    properties = paragraph.find(_q("pPr"))
    if properties is None:
        if style_id is None:
            return
        properties = ET.Element(_q("pPr"))
        paragraph.insert(0, properties)
    if style_id is None:
        _remove_child(properties, "pStyle")
    else:
        child = _ensure_ordered_child(properties, "pStyle", _PPR_ORDER)
        child.set(_q("val"), style_id)
    _remove_if_empty(paragraph, properties)


def patch_text_style(
    run: ET.Element,
    style: TextStyle | None,
    fields: Iterable[str],
) -> None:
    """Update only selected ``w:rPr`` properties on one run."""

    properties = run.find(_q("rPr"))
    if properties is None:
        properties = ET.Element(_q("rPr"))
        run.insert(0, properties)
    values = style or TextStyle()
    selected = set(fields)

    if selected & {"font_family", "font_family_east_asia"}:
        fonts = _ensure_ordered_child(properties, "rFonts", _RPR_ORDER)
        if "font_family" in selected:
            value = values.font_family
            for attribute in ("ascii", "hAnsi"):
                if value is None:
                    fonts.attrib.pop(_q(attribute), None)
                else:
                    fonts.set(_q(attribute), value)
        if "font_family_east_asia" in selected:
            value = values.font_family_east_asia
            if value is None:
                fonts.attrib.pop(_q("eastAsia"), None)
            else:
                fonts.set(_q("eastAsia"), value)
        _remove_if_empty(properties, fonts)

    if "font_size" in selected:
        value = values.font_size
        for name in ("sz", "szCs"):
            if value is None:
                _remove_child(properties, name)
            else:
                child = _ensure_ordered_child(properties, name, _RPR_ORDER)
                child.set(_q("val"), _half_points(value))

    if "color" in selected:
        value = values.color
        if value is None:
            _remove_child(properties, "color")
        else:
            child = _ensure_ordered_child(properties, "color", _RPR_ORDER)
            child.set(_q("val"), value.removeprefix("#"))

    if "background_color" in selected:
        shading = _ensure_ordered_child(properties, "shd", _RPR_ORDER)
        value = values.background_color
        if value is None:
            shading.attrib.pop(_q("fill"), None)
        else:
            shading.set(_q("val"), "clear")
            shading.set(_q("color"), "auto")
            shading.set(_q("fill"), value.removeprefix("#"))
        _remove_if_empty(properties, shading)

    for field_name, native_name in (
        ("bold", "b"),
        ("italic", "i"),
        ("underline", "u"),
        ("strike", "strike"),
        ("small_caps", "smallCaps"),
        ("all_caps", "caps"),
    ):
        if field_name in selected:
            _set_boolean(
                properties,
                native_name,
                getattr(values, field_name),
                _RPR_ORDER,
            )

    if "letter_spacing" in selected:
        value = values.letter_spacing
        if value is None:
            _remove_child(properties, "spacing")
        else:
            child = _ensure_ordered_child(properties, "spacing", _RPR_ORDER)
            child.set(_q("val"), _twips(value))

    if "baseline" in selected:
        value = values.baseline
        if value is None:
            _remove_child(properties, "vertAlign")
        else:
            child = _ensure_ordered_child(properties, "vertAlign", _RPR_ORDER)
            child.set(_q("val"), "baseline" if value == "normal" else value)
    _remove_if_empty(run, properties)


def patch_paragraph_mark_text_style(
    paragraph: ET.Element,
    style: TextStyle | None,
    fields: Iterable[str],
) -> None:
    """Format the paragraph mark so empty paragraphs retain character intent."""

    properties = paragraph.find(_q("pPr"))
    if properties is None:
        properties = ET.Element(_q("pPr"))
        paragraph.insert(0, properties)
    _ensure_ordered_child(properties, "rPr", _PPR_ORDER)
    patch_text_style(properties, style, fields)
    _remove_if_empty(paragraph, properties)


def apply_paragraph_style(
    paragraph: ET.Element,
    style: ParagraphStyle | None,
) -> None:
    if style is not None:
        fields = {
            name
            for name in ParagraphStyle.model_fields
            if getattr(style, name) is not None
        }
        patch_paragraph_style(paragraph, style, fields)


def apply_text_style(run: ET.Element, style: TextStyle | None) -> None:
    if style is not None:
        fields = {
            name for name in TextStyle.model_fields if getattr(style, name) is not None
        }
        patch_text_style(run, style, fields)


def apply_paragraph_mark_text_style(
    paragraph: ET.Element,
    style: TextStyle | None,
) -> None:
    if style is not None:
        fields = {
            name for name in TextStyle.model_fields if getattr(style, name) is not None
        }
        patch_paragraph_mark_text_style(paragraph, style, fields)


__all__ = [
    "apply_paragraph_mark_text_style",
    "apply_paragraph_style",
    "apply_text_style",
    "common_text_style",
    "patch_paragraph_mark_text_style",
    "patch_paragraph_style",
    "patch_paragraph_style_ref",
    "patch_text_style",
    "read_paragraph_style",
    "read_text_style",
]
