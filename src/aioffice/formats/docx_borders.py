"""Shared loss-aware WordprocessingML border codecs."""

from __future__ import annotations

from typing import Literal, cast
from xml.etree import ElementTree as ET

from aioffice.spec.models import BorderLine, Length

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

BORDER_ATTRIBUTES = (
    "val",
    "sz",
    "space",
    "color",
    "themeColor",
    "themeTint",
    "themeShade",
)
VISIBLE_BORDER_STYLES = {
    "single",
    "double",
    "dotted",
    "dashed",
    "thick",
}


def _q(local: str) -> str:
    return f"{{{W}}}{local}"


def read_border_element(
    element: ET.Element | None,
) -> BorderLine | None:
    """Project one border only when every visible value is understood."""

    if element is None:
        return None
    raw_style = element.get(_q("val"))
    if raw_style in {"none", "nil"}:
        return BorderLine(style="none")
    if raw_style not in VISIBLE_BORDER_STYLES:
        return None
    if any(
        element.get(_q(attribute)) is not None
        for attribute in ("themeColor", "themeTint", "themeShade")
    ):
        return None
    raw_size = element.get(_q("sz"))
    try:
        size = int(raw_size) if raw_size is not None else 0
    except ValueError:
        return None
    if size < 2 or size > 96:
        return None
    raw_color = element.get(_q("color"), "auto")
    color = (
        f"#{raw_color.upper()}"
        if len(raw_color) == 6
        and all(
            character in "0123456789ABCDEFabcdef"
            for character in raw_color
        )
        else "auto"
        if raw_color == "auto"
        else None
    )
    if color is None:
        return None
    raw_space = element.get(_q("space"))
    try:
        space_value = int(raw_space) if raw_space is not None else None
    except ValueError:
        return None
    if space_value is not None and not 0 <= space_value <= 31:
        return None
    return BorderLine(
        style=cast(
            Literal[
                "single",
                "double",
                "dotted",
                "dashed",
                "thick",
            ],
            raw_style,
        ),
        width=Length(value=size / 8, unit="pt"),
        color=color,
        space=(
            Length(value=space_value, unit="pt")
            if space_value is not None
            else None
        ),
    )


def clear_border_element(element: ET.Element) -> None:
    """Remove supported border attributes while preserving extensions."""

    for attribute in BORDER_ATTRIBUTES:
        element.attrib.pop(_q(attribute), None)


def write_border_element(
    element: ET.Element,
    value: BorderLine,
) -> None:
    """Replace supported attributes on one edge and preserve unknown XML."""

    clear_border_element(element)
    element.set(_q("val"), value.style)
    if value.style == "none":
        return
    assert value.width is not None
    element.set(
        _q("sz"),
        str(round(value.width.to_points() * 8)),
    )
    element.set(
        _q("color"),
        (
            value.color.removeprefix("#")
            if value.color != "auto"
            else "auto"
        ),
    )
    if value.space is not None:
        element.set(
            _q("space"),
            str(round(value.space.to_points())),
        )


__all__ = [
    "clear_border_element",
    "read_border_element",
    "write_border_element",
]
