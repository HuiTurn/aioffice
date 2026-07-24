"""Conservative projection and verified extraction of native DOCX images."""

from __future__ import annotations

import hashlib
import posixpath
from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import PurePosixPath
from typing import Literal, TypeAlias, cast
from xml.etree import ElementTree as ET

from aioffice.core.errors import NativePackageError
from aioffice.formats.docx_header_footer import resolve_relationship_target
from aioffice.formats.docx_style import apply_paragraph_style
from aioffice.native import NativePackage
from aioffice.native.xml import (
    namespace_declarations,
    parse_xml,
    serialize_xml,
)
from aioffice.spec.models import (
    AssetRef,
    FloatingImageEffectExtent,
    FloatingImageHorizontalPosition,
    FloatingImageLayout,
    FloatingImageRelativeHeight,
    FloatingImageRelativeSize,
    FloatingImageRelativeWidth,
    FloatingImageTextDistances,
    FloatingImageTextWrap,
    FloatingImageVerticalPosition,
    FloatingImageWrapPoint,
    FloatingImageWrapPolygon,
    ImageBlock,
    ImageCrop,
    ImageInsert,
    Length,
    NativeRef,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
W14 = "http://schemas.microsoft.com/office/word/2010/wordml"
WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing"
MC = "http://schemas.openxmlformats.org/markup-compatibility/2006"
_REPLACEMENT_EXTENSIONS = {
    "image/png": "png",
    "image/jpeg": "jpg",
    "image/gif": "gif",
    "image/bmp": "bmp",
    "image/tiff": "tif",
}

IMAGE_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)
RELATIONSHIPS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-package.relationships+xml"
)
EMU_PER_POINT = 12_700
ST_COORDINATE_MIN = -27_273_042_329_600
ST_COORDINATE_MAX = 27_273_042_316_900

_VISUAL_EFFECT_NAMES = {
    "alphaBiLevel",
    "alphaCeiling",
    "alphaFloor",
    "alphaInv",
    "alphaMod",
    "alphaModFix",
    "alphaOutset",
    "alphaRepl",
    "biLevel",
    "blur",
    "clrChange",
    "clrRepl",
    "duotone",
    "effectDag",
    "effectLst",
    "fillOverlay",
    "glow",
    "grayscl",
    "hsl",
    "innerShdw",
    "lum",
    "outerShdw",
    "reflection",
    "relOff",
    "softEdge",
    "tint",
    "xfrmEffect",
}


def _q(namespace: str, local: str) -> str:
    return f"{{{namespace}}}{local}"


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _enabled(value: str | None) -> bool:
    return value not in {None, "", "0", "false", "False", "off"}


@dataclass(frozen=True, slots=True)
class SimpleNativeImage:
    """Trusted metadata for the deliberately small image projection subset."""

    relationship_id: str
    part_uri: str
    media_type: str
    filename: str
    sha256: str
    size_bytes: int
    width: Length
    height: Length
    crop: ImageCrop | None
    placement: Literal["inline", "floating"]
    floating: FloatingImageLayout | None
    name: str | None
    alt_text: str | None
    title: str | None
    native_drawing_id: str | None

    @property
    def asset_id(self) -> str:
        return f"asset_{self.sha256.lower()}"


HorizontalRelativeTo: TypeAlias = Literal[
    "character",
    "column",
    "inside_margin",
    "left_margin",
    "margin",
    "outside_margin",
    "page",
    "right_margin",
]
VerticalRelativeTo: TypeAlias = Literal[
    "bottom_margin",
    "inside_margin",
    "line",
    "margin",
    "outside_margin",
    "page",
    "paragraph",
    "top_margin",
]
RelativeWidthTo: TypeAlias = Literal[
    "inside_margin",
    "left_margin",
    "margin",
    "outside_margin",
    "page",
    "right_margin",
]
RelativeHeightTo: TypeAlias = Literal[
    "bottom_margin",
    "inside_margin",
    "margin",
    "outside_margin",
    "page",
    "top_margin",
]
WrapSide: TypeAlias = Literal["both_sides", "largest", "left", "right"]
WrapMode: TypeAlias = Literal[
    "square",
    "none",
    "top_and_bottom",
    "tight",
    "through",
]
HorizontalAlignment: TypeAlias = Literal[
    "left",
    "right",
    "center",
    "inside",
    "outside",
]
VerticalAlignment: TypeAlias = Literal[
    "top",
    "bottom",
    "center",
    "inside",
    "outside",
]

_HORIZONTAL_RELATIVE_FROM: dict[str, HorizontalRelativeTo] = {
    "character": "character",
    "column": "column",
    "insideMargin": "inside_margin",
    "leftMargin": "left_margin",
    "margin": "margin",
    "outsideMargin": "outside_margin",
    "page": "page",
    "rightMargin": "right_margin",
}
_VERTICAL_RELATIVE_FROM: dict[str, VerticalRelativeTo] = {
    "bottomMargin": "bottom_margin",
    "insideMargin": "inside_margin",
    "line": "line",
    "margin": "margin",
    "outsideMargin": "outside_margin",
    "page": "page",
    "paragraph": "paragraph",
    "topMargin": "top_margin",
}
_RELATIVE_WIDTH_FROM: dict[str, RelativeWidthTo] = {
    "insideMargin": "inside_margin",
    "leftMargin": "left_margin",
    "margin": "margin",
    "outsideMargin": "outside_margin",
    "page": "page",
    "rightMargin": "right_margin",
}
_RELATIVE_HEIGHT_FROM: dict[str, RelativeHeightTo] = {
    "bottomMargin": "bottom_margin",
    "insideMargin": "inside_margin",
    "margin": "margin",
    "outsideMargin": "outside_margin",
    "page": "page",
    "topMargin": "top_margin",
}
_WRAP_SIDES: dict[str, WrapSide] = {
    "bothSides": "both_sides",
    "largest": "largest",
    "left": "left",
    "right": "right",
}
_WRAP_TAG_TO_MODE: dict[str, WrapMode] = {
    _q(WP, "wrapSquare"): "square",
    _q(WP, "wrapNone"): "none",
    _q(WP, "wrapTopAndBottom"): "top_and_bottom",
    _q(WP, "wrapTight"): "tight",
    _q(WP, "wrapThrough"): "through",
}
_WRAP_MODE_TO_TAG = {
    value: key for key, value in _WRAP_TAG_TO_MODE.items()
}
_HORIZONTAL_ALIGNMENTS: frozenset[str] = frozenset(
    {"left", "right", "center", "inside", "outside"}
)
_VERTICAL_ALIGNMENTS: frozenset[str] = frozenset(
    {"top", "bottom", "center", "inside", "outside"}
)
_HORIZONTAL_RELATIVE_TO_NATIVE = {
    value: key for key, value in _HORIZONTAL_RELATIVE_FROM.items()
}
_VERTICAL_RELATIVE_TO_NATIVE = {
    value: key for key, value in _VERTICAL_RELATIVE_FROM.items()
}
_RELATIVE_WIDTH_TO_NATIVE = {
    value: key for key, value in _RELATIVE_WIDTH_FROM.items()
}
_RELATIVE_HEIGHT_TO_NATIVE = {
    value: key for key, value in _RELATIVE_HEIGHT_FROM.items()
}
_WRAP_SIDE_TO_NATIVE = {
    value: key for key, value in _WRAP_SIDES.items()
}
_TEXT_DISTANCE_FIELDS = (
    ("top", "distT"),
    ("right", "distR"),
    ("bottom", "distB"),
    ("left", "distL"),
)
_EFFECT_EXTENT_FIELDS = (
    ("left", "l"),
    ("top", "t"),
    ("right", "r"),
    ("bottom", "b"),
)


def _strict_boolean(value: str | None) -> bool | None:
    if value in {"1", "true"}:
        return True
    if value in {"0", "false"}:
        return False
    return None


def _ensure_wp14_compatibility(
    package: NativePackage,
    part_root: ET.Element,
    *,
    source_part: str,
) -> None:
    """Keep the Office 2010 extension prefix declared and ignorable."""

    declarations = namespace_declarations(package.get_part(source_part))
    original_prefix = next(
        (
            prefix
            for prefix, namespace in declarations.items()
            if namespace == WP14
        ),
        None,
    )
    usable_original = (
        original_prefix
        if original_prefix
        and not (
            original_prefix.startswith("ns")
            and original_prefix[2:].isdigit()
        )
        else None
    )
    active_prefix = usable_original or "wp14"
    ET.register_namespace(active_prefix, WP14)
    if not any(namespace == MC for namespace in declarations.values()):
        ET.register_namespace("mc", MC)

    ignorable_name = _q(MC, "Ignorable")
    tokens = (part_root.get(ignorable_name) or "").split()
    if (
        original_prefix
        and original_prefix != active_prefix
        and original_prefix in tokens
    ):
        tokens = [
            active_prefix if token == original_prefix else token
            for token in tokens
        ]
    if active_prefix not in tokens:
        tokens.append(active_prefix)
    part_root.set(ignorable_name, " ".join(tokens))


def _emu_length(value: int) -> Length:
    return Length(
        value=round(value / EMU_PER_POINT, 6),
        unit="pt",
    )


def _floating_text_distances(
    element: ET.Element,
    *,
    allowed_edges: frozenset[str],
) -> FloatingImageTextDistances | None:
    payload: dict[str, Length] = {}
    for field_name, attribute_name in _TEXT_DISTANCE_FIELDS:
        if attribute_name not in element.attrib:
            continue
        if field_name not in allowed_edges:
            raise ValueError("Unsupported text-distance edge.")
        native_value = int(element.attrib[attribute_name])
        if native_value < 0 or native_value > 2**32 - 1:
            raise ValueError("Text distance is outside UInt32.")
        payload[field_name] = _emu_length(native_value)
    if not payload:
        return None
    return FloatingImageTextDistances.model_validate(payload)


def _floating_effect_extent(
    element: ET.Element | None,
) -> FloatingImageEffectExtent | None:
    if element is None:
        return None
    if (
        set(element.attrib) != {"l", "t", "r", "b"}
        or len(element)
        or (element.text or "").strip()
    ):
        raise ValueError("Effect extent is not one strict native leaf.")
    values = {
        field_name: int(element.attrib[attribute_name])
        for field_name, attribute_name in _EFFECT_EXTENT_FIELDS
    }
    if any(
        value < ST_COORDINATE_MIN
        or value > ST_COORDINATE_MAX
        for value in values.values()
    ):
        raise ValueError("Effect extent is outside ST_Coordinate.")
    return FloatingImageEffectExtent.model_validate(
        {
            field_name: _emu_length(native_value)
            for field_name, native_value in values.items()
        }
    )


def _floating_wrap_point(
    element: ET.Element,
) -> FloatingImageWrapPoint:
    if (
        set(element.attrib) != {"x", "y"}
        or len(element)
        or (element.text or "").strip()
    ):
        raise ValueError("Wrap polygon point is not one strict native leaf.")
    return FloatingImageWrapPoint.model_validate(
        {
            "x": int(element.attrib["x"]),
            "y": int(element.attrib["y"]),
        }
    )


def _floating_wrap_polygon(
    element: ET.Element,
) -> FloatingImageWrapPolygon:
    if (
        any(attribute != "edited" for attribute in element.attrib)
        or (element.text or "").strip()
    ):
        raise ValueError("Wrap polygon has unsupported native metadata.")
    children = list(element)
    if (
        len(children) < 3
        or len(children) > 4097
        or children[0].tag != _q(WP, "start")
        or any(
            child.tag != _q(WP, "lineTo")
            for child in children[1:]
        )
    ):
        raise ValueError("Wrap polygon has an invalid point sequence.")
    payload: dict[str, object] = {
        "start": _floating_wrap_point(children[0]),
        "line_to": [
            _floating_wrap_point(child)
            for child in children[1:]
        ],
    }
    if "edited" in element.attrib:
        edited = _strict_boolean(element.attrib["edited"])
        if edited is None:
            raise ValueError("Wrap polygon edited flag is not boolean.")
        payload["edited"] = edited
    return FloatingImageWrapPolygon.model_validate(payload)


def _native_position_value(
    *,
    offset: Length | None,
    alignment: str | None,
    percentage_offset: float | None,
    percentage_tag: str,
    allowed_alignments: frozenset[str],
) -> tuple[str, str] | None:
    if (
        sum(
            value is not None
            for value in (offset, alignment, percentage_offset)
        )
        != 1
    ):
        return None
    if offset is not None:
        offset_emu = round(offset.to_points() * EMU_PER_POINT)
        if offset_emu < -(2**63) or offset_emu > 2**63 - 1:
            return None
        return _q(WP, "posOffset"), str(offset_emu)
    if percentage_offset is not None:
        native_percentage = round(percentage_offset * 1_000)
        if native_percentage < -(2**31) or native_percentage > 2**31 - 1:
            return None
        return _q(WP14, percentage_tag), str(native_percentage)
    if alignment not in allowed_alignments:
        return None
    return _q(WP, "align"), alignment


def _native_text_distance_attributes(
    distances: FloatingImageTextDistances | None,
    *,
    allowed_edges: frozenset[str],
) -> dict[str, str] | None:
    attributes: dict[str, str] = {}
    if distances is None:
        return attributes
    for field_name, attribute_name in _TEXT_DISTANCE_FIELDS:
        value = getattr(distances, field_name)
        if value is None:
            continue
        if field_name not in allowed_edges:
            return None
        distance_emu = round(value.to_points() * EMU_PER_POINT)
        if distance_emu < 0 or distance_emu > 2**32 - 1:
            return None
        attributes[attribute_name] = str(distance_emu)
    return attributes


def _native_effect_extent(
    effect_extent: FloatingImageEffectExtent,
) -> ET.Element | None:
    attributes: dict[str, str] = {}
    for field_name, attribute_name in _EFFECT_EXTENT_FIELDS:
        extent_emu = round(
            getattr(effect_extent, field_name).to_points()
            * EMU_PER_POINT
        )
        if (
            extent_emu < ST_COORDINATE_MIN
            or extent_emu > ST_COORDINATE_MAX
        ):
            return None
        attributes[attribute_name] = str(extent_emu)
    return ET.Element(_q(WP, "effectExtent"), attributes)


def _native_wrap_polygon(
    polygon: FloatingImageWrapPolygon,
) -> ET.Element | None:
    if len(polygon.line_to) < 2 or len(polygon.line_to) > 4096:
        return None
    attributes = (
        {}
        if polygon.edited is None
        else {"edited": "1" if polygon.edited else "0"}
    )
    element = ET.Element(_q(WP, "wrapPolygon"), attributes)

    def append_point(
        name: str,
        point: FloatingImageWrapPoint,
    ) -> bool:
        for value in (point.x, point.y):
            if (
                value < ST_COORDINATE_MIN
                or value > ST_COORDINATE_MAX
            ):
                return False
        ET.SubElement(
            element,
            _q(WP, name),
            {"x": str(point.x), "y": str(point.y)},
        )
        return True

    if not append_point("start", polygon.start):
        return None
    for point in polygon.line_to:
        if not append_point("lineTo", point):
            return None
    return element


def _native_wrap_value(
    wrap: FloatingImageTextWrap,
) -> ET.Element | None:
    tag = _WRAP_MODE_TO_TAG.get(wrap.mode)
    if tag is None:
        return None
    attributes: dict[str, str] = {}
    if wrap.mode in {"square", "tight", "through"}:
        if wrap.side is None:
            return None
        native_side = _WRAP_SIDE_TO_NATIVE.get(wrap.side)
        if native_side is None:
            return None
        attributes["wrapText"] = native_side
        allowed_edges = (
            frozenset({"top", "right", "bottom", "left"})
            if wrap.mode == "square"
            else frozenset({"right", "left"})
        )
    elif wrap.mode == "top_and_bottom":
        if wrap.side is not None:
            return None
        allowed_edges = frozenset({"top", "bottom"})
    elif (
        wrap.side is not None
        or wrap.distances is not None
        or wrap.effect_extent is not None
        or wrap.polygon is not None
    ):
        return None
    else:
        allowed_edges = frozenset()
    distance_attributes = _native_text_distance_attributes(
        wrap.distances,
        allowed_edges=allowed_edges,
    )
    if distance_attributes is None:
        return None
    attributes.update(distance_attributes)
    element = ET.Element(tag, attributes)
    if wrap.effect_extent is not None:
        effect_extent = _native_effect_extent(wrap.effect_extent)
        if (
            effect_extent is None
            or wrap.mode in {"none", "tight", "through"}
        ):
            return None
        element.append(effect_extent)
    if wrap.mode in {"tight", "through"}:
        if wrap.polygon is None:
            return None
        polygon = _native_wrap_polygon(wrap.polygon)
        if polygon is None:
            return None
        element.append(polygon)
    elif wrap.polygon is not None:
        return None
    return element


def _native_relative_size(
    relative_size: FloatingImageRelativeSize | None,
) -> list[ET.Element] | None:
    elements: list[ET.Element] = []
    if relative_size is None:
        return elements
    for axis, tag, child_tag, native_frames in (
        (
            relative_size.width,
            "sizeRelH",
            "pctWidth",
            _RELATIVE_WIDTH_TO_NATIVE,
        ),
        (
            relative_size.height,
            "sizeRelV",
            "pctHeight",
            _RELATIVE_HEIGHT_TO_NATIVE,
        ),
    ):
        if axis is None:
            continue
        native_frame = native_frames.get(axis.relative_to)
        native_percentage = round(axis.percentage * 1_000)
        if (
            native_frame is None
            or native_percentage < 0
            or native_percentage > 2**31 - 1
        ):
            return None
        element = ET.Element(
            _q(WP14, tag),
            {"relativeFrom": native_frame},
        )
        ET.SubElement(
            element,
            _q(WP14, child_tag),
        ).text = str(native_percentage)
        elements.append(element)
    return elements or None


def floating_image_layout_matches(
    left: FloatingImageLayout | None,
    right: FloatingImageLayout | None,
) -> bool:
    """Compare floating layouts by their exact native EMU semantics."""

    if left is None or right is None:
        return left is right

    def emu(length: Length) -> int:
        return round(length.to_points() * EMU_PER_POINT)

    def position_matches(
        left_offset: Length | None,
        left_alignment: str | None,
        left_percentage_offset: float | None,
        right_offset: Length | None,
        right_alignment: str | None,
        right_percentage_offset: float | None,
    ) -> bool:
        if left_alignment != right_alignment:
            return False
        if (
            left_percentage_offset is None
            or right_percentage_offset is None
        ):
            if left_percentage_offset is not right_percentage_offset:
                return False
        elif (
            round(left_percentage_offset * 1_000)
            != round(right_percentage_offset * 1_000)
        ):
            return False
        if left_offset is None or right_offset is None:
            return left_offset is right_offset
        return emu(left_offset) == emu(right_offset)

    def distances_match(
        left_distances: FloatingImageTextDistances | None,
        right_distances: FloatingImageTextDistances | None,
    ) -> bool:
        for field_name, _ in _TEXT_DISTANCE_FIELDS:
            left_value = (
                getattr(left_distances, field_name)
                if left_distances is not None
                else None
            )
            right_value = (
                getattr(right_distances, field_name)
                if right_distances is not None
                else None
            )
            if left_value is None or right_value is None:
                if left_value is not right_value:
                    return False
            elif emu(left_value) != emu(right_value):
                return False
        return True

    def effect_extents_match(
        left_extent: FloatingImageEffectExtent | None,
        right_extent: FloatingImageEffectExtent | None,
    ) -> bool:
        for field_name, _ in _EFFECT_EXTENT_FIELDS:
            left_value = (
                getattr(left_extent, field_name)
                if left_extent is not None
                else None
            )
            right_value = (
                getattr(right_extent, field_name)
                if right_extent is not None
                else None
            )
            if left_value is None or right_value is None:
                if left_value is not right_value:
                    return False
            elif emu(left_value) != emu(right_value):
                return False
        return True

    return (
        left.horizontal.relative_to == right.horizontal.relative_to
        and position_matches(
            left.horizontal.offset,
            left.horizontal.alignment,
            left.horizontal.percentage_offset,
            right.horizontal.offset,
            right.horizontal.alignment,
            right.horizontal.percentage_offset,
        )
        and left.vertical.relative_to == right.vertical.relative_to
        and position_matches(
            left.vertical.offset,
            left.vertical.alignment,
            left.vertical.percentage_offset,
            right.vertical.offset,
            right.vertical.alignment,
            right.vertical.percentage_offset,
        )
        and distances_match(
            left.anchor_distances,
            right.anchor_distances,
        )
        and effect_extents_match(
            left.anchor_effect_extent,
            right.anchor_effect_extent,
        )
        and left.wrap.mode == right.wrap.mode
        and left.wrap.side == right.wrap.side
        and distances_match(
            left.wrap.distances,
            right.wrap.distances,
        )
        and effect_extents_match(
            left.wrap.effect_extent,
            right.wrap.effect_extent,
        )
        and left.wrap.polygon == right.wrap.polygon
        and left.relative_size == right.relative_size
        and left.relative_height == right.relative_height
        and left.behind_text == right.behind_text
        and left.locked == right.locked
        and left.layout_in_cell == right.layout_in_cell
        and left.allow_overlap == right.allow_overlap
    )


def _floating_image_layout(
    anchor: ET.Element,
) -> FloatingImageLayout | None:
    allowed_attributes = {
        "distT",
        "distB",
        "distL",
        "distR",
        "simplePos",
        "relativeHeight",
        "behindDoc",
        "locked",
        "layoutInCell",
        "allowOverlap",
        _q(WP14, "anchorId"),
        _q(WP14, "editId"),
    }
    required_attributes = {
        "simplePos",
        "relativeHeight",
        "behindDoc",
        "locked",
        "layoutInCell",
        "allowOverlap",
    }
    if (
        any(attribute not in allowed_attributes for attribute in anchor.attrib)
        or not required_attributes.issubset(anchor.attrib)
    ):
        return None
    for identity_attribute in (
        _q(WP14, "anchorId"),
        _q(WP14, "editId"),
    ):
        identity = anchor.get(identity_attribute)
        if identity is None:
            continue
        try:
            int(identity, 16)
        except ValueError:
            return None
        if len(identity) != 8 or not identity.isascii():
            return None

    simple_position = anchor.find(f"./{_q(WP, 'simplePos')}")
    horizontal = anchor.find(f"./{_q(WP, 'positionH')}")
    vertical = anchor.find(f"./{_q(WP, 'positionV')}")
    wraps = [
        child
        for child in anchor
        if child.tag in _WRAP_TAG_TO_MODE
    ]
    if (
        simple_position is None
        or horizontal is None
        or vertical is None
        or len(wraps) != 1
    ):
        return None
    wrap = wraps[0]
    wrap_mode = _WRAP_TAG_TO_MODE[wrap.tag]
    optional_effect_extent = anchor.find(f"./{_q(WP, 'effectExtent')}")
    optional_frame_properties = anchor.find(
        f"./{_q(WP, 'cNvGraphicFramePr')}"
    )
    relative_width_element = anchor.find(
        f"./{_q(WP14, 'sizeRelH')}"
    )
    relative_height_element = anchor.find(
        f"./{_q(WP14, 'sizeRelV')}"
    )
    expected_children = [
        _q(WP, "simplePos"),
        _q(WP, "positionH"),
        _q(WP, "positionV"),
        _q(WP, "extent"),
        *(
            [_q(WP, "effectExtent")]
            if optional_effect_extent is not None
            else []
        ),
        wrap.tag,
        _q(WP, "docPr"),
        *(
            [_q(WP, "cNvGraphicFramePr")]
            if optional_frame_properties is not None
            else []
        ),
        _q(A, "graphic"),
        *(
            [_q(WP14, "sizeRelH")]
            if relative_width_element is not None
            else []
        ),
        *(
            [_q(WP14, "sizeRelV")]
            if relative_height_element is not None
            else []
        ),
    ]
    if [child.tag for child in anchor] != expected_children:
        return None

    if (
        _strict_boolean(anchor.attrib.get("simplePos")) is not False
        or len(simple_position)
        or set(simple_position.attrib) != {"x", "y"}
    ):
        return None
    try:
        simple_x = int(simple_position.attrib["x"])
        simple_y = int(simple_position.attrib["y"])
    except ValueError:
        return None
    if simple_x != 0 or simple_y != 0:
        return None

    def position(
        element: ET.Element,
        relative_values: Mapping[str, str],
        allowed_alignments: frozenset[str],
        percentage_tag: str,
    ) -> tuple[
        str,
        Length | None,
        str | None,
        float | None,
    ] | None:
        if set(element.attrib) != {"relativeFrom"}:
            return None
        relative_to = relative_values.get(element.attrib["relativeFrom"])
        children = list(element)
        if (
            relative_to is None
            or len(children) != 1
            or children[0].attrib
            or len(children[0])
        ):
            return None
        child = children[0]
        if child.tag == _q(WP, "align"):
            alignment = child.text or ""
            if alignment not in allowed_alignments:
                return None
            return relative_to, None, alignment, None
        if child.tag == _q(WP14, percentage_tag):
            try:
                percentage = int(child.text or "")
            except ValueError:
                return None
            if percentage < -(2**31) or percentage > 2**31 - 1:
                return None
            return relative_to, None, None, percentage / 1_000
        if child.tag != _q(WP, "posOffset"):
            return None
        try:
            offset = int(child.text or "")
        except ValueError:
            return None
        if offset < -(2**63) or offset > 2**63 - 1:
            return None
        return relative_to, _emu_length(offset), None, None

    horizontal_position = position(
        horizontal,
        _HORIZONTAL_RELATIVE_FROM,
        _HORIZONTAL_ALIGNMENTS,
        "pctPosHOffset",
    )
    vertical_position = position(
        vertical,
        _VERTICAL_RELATIVE_FROM,
        _VERTICAL_ALIGNMENTS,
        "pctPosVOffset",
    )
    if horizontal_position is None or vertical_position is None:
        return None

    def relative_axis(
        element: ET.Element | None,
        *,
        frames: Mapping[str, RelativeWidthTo | RelativeHeightTo],
        percentage_tag: str,
    ) -> tuple[RelativeWidthTo | RelativeHeightTo, float] | None:
        if element is None:
            return None
        if (
            set(element.attrib) != {"relativeFrom"}
            or (element.text or "").strip()
            or len(element) != 1
        ):
            raise ValueError("Relative-size axis has invalid structure.")
        relative_to = frames.get(element.attrib["relativeFrom"])
        child = element[0]
        if (
            relative_to is None
            or child.tag != _q(WP14, percentage_tag)
            or child.attrib
            or len(child)
        ):
            raise ValueError("Relative-size axis has invalid metadata.")
        percentage = int(child.text or "")
        if percentage < 0 or percentage > 2**31 - 1:
            raise ValueError("Relative-size percentage is outside Int32.")
        return relative_to, percentage / 1_000

    try:
        relative_width = relative_axis(
            relative_width_element,
            frames=_RELATIVE_WIDTH_FROM,
            percentage_tag="pctWidth",
        )
        relative_size_height = relative_axis(
            relative_height_element,
            frames=_RELATIVE_HEIGHT_FROM,
            percentage_tag="pctHeight",
        )
    except ValueError:
        return None

    wrap_side: WrapSide | None = None
    wrap_distances: FloatingImageTextDistances | None
    wrap_effect_extent: FloatingImageEffectExtent | None
    wrap_polygon: FloatingImageWrapPolygon | None = None
    if (wrap.text or "").strip():
        return None
    try:
        if wrap_mode == "square":
            if any(
                attribute
                not in {"wrapText", "distT", "distR", "distB", "distL"}
                for attribute in wrap.attrib
            ):
                return None
            wrap_side = _WRAP_SIDES.get(wrap.attrib.get("wrapText", ""))
            if wrap_side is None:
                return None
            wrap_distances = _floating_text_distances(
                wrap,
                allowed_edges=frozenset(
                    {"top", "right", "bottom", "left"}
                ),
            )
        elif wrap_mode == "top_and_bottom":
            if any(
                attribute not in {"distT", "distB"}
                for attribute in wrap.attrib
            ):
                return None
            wrap_distances = _floating_text_distances(
                wrap,
                allowed_edges=frozenset({"top", "bottom"}),
            )
        elif wrap_mode in {"tight", "through"}:
            if any(
                attribute
                not in {"wrapText", "distL", "distR"}
                for attribute in wrap.attrib
            ):
                return None
            wrap_side = _WRAP_SIDES.get(wrap.attrib.get("wrapText", ""))
            if wrap_side is None:
                return None
            wrap_distances = _floating_text_distances(
                wrap,
                allowed_edges=frozenset({"left", "right"}),
            )
        elif wrap.attrib or len(wrap):
            return None
        else:
            wrap_distances = None
        if wrap_mode == "none":
            wrap_effect_extent = None
        elif wrap_mode in {"tight", "through"}:
            if (
                len(wrap) != 1
                or wrap[0].tag != _q(WP, "wrapPolygon")
            ):
                return None
            wrap_effect_extent = None
            wrap_polygon = _floating_wrap_polygon(wrap[0])
        else:
            if len(wrap) > 1 or any(
                child.tag != _q(WP, "effectExtent")
                for child in wrap
            ):
                return None
            wrap_effect_extent = _floating_effect_extent(
                wrap.find(f"./{_q(WP, 'effectExtent')}")
            )
        anchor_distances = _floating_text_distances(
            anchor,
            allowed_edges=frozenset(
                {"top", "right", "bottom", "left"}
            ),
        )
        anchor_effect_extent = _floating_effect_extent(
            optional_effect_extent
        )
        relative_height = int(anchor.attrib["relativeHeight"])
    except ValueError:
        return None
    if (
        relative_height < 0
        or relative_height > 2**32 - 1
    ):
        return None

    booleans = {
        field_name: _strict_boolean(anchor.attrib[attribute])
        for field_name, attribute in (
            ("behind_text", "behindDoc"),
            ("locked", "locked"),
            ("layout_in_cell", "layoutInCell"),
            ("allow_overlap", "allowOverlap"),
        )
    }
    if any(value is None for value in booleans.values()):
        return None

    return FloatingImageLayout(
        horizontal=FloatingImageHorizontalPosition.model_validate(
            {
                "relative_to": cast(
                    HorizontalRelativeTo,
                    horizontal_position[0],
                ),
                **(
                    {"offset": horizontal_position[1]}
                    if horizontal_position[1] is not None
                    else (
                        {
                        "alignment": cast(
                            HorizontalAlignment,
                            horizontal_position[2],
                        )
                        }
                        if horizontal_position[2] is not None
                        else {
                            "percentage_offset": horizontal_position[3]
                        }
                    )
                ),
            }
        ),
        vertical=FloatingImageVerticalPosition.model_validate(
            {
                "relative_to": cast(
                    VerticalRelativeTo,
                    vertical_position[0],
                ),
                **(
                    {"offset": vertical_position[1]}
                    if vertical_position[1] is not None
                    else (
                        {
                        "alignment": cast(
                            VerticalAlignment,
                            vertical_position[2],
                        )
                        }
                        if vertical_position[2] is not None
                        else {
                            "percentage_offset": vertical_position[3]
                        }
                    )
                ),
            }
        ),
        anchor_distances=anchor_distances,
        anchor_effect_extent=anchor_effect_extent,
        wrap=FloatingImageTextWrap.model_validate(
            {
                "mode": wrap_mode,
                **(
                    {"side": cast(WrapSide, wrap_side)}
                    if wrap_side is not None
                    else {}
                ),
                **(
                    {"distances": wrap_distances}
                    if wrap_distances is not None
                    else {}
                ),
                **(
                    {"effect_extent": wrap_effect_extent}
                    if wrap_effect_extent is not None
                    else {}
                ),
                **(
                    {"polygon": wrap_polygon}
                    if wrap_polygon is not None
                    else {}
                ),
            }
        ),
        relative_size=(
            FloatingImageRelativeSize(
                width=(
                    FloatingImageRelativeWidth(
                        relative_to=cast(RelativeWidthTo, relative_width[0]),
                        percentage=relative_width[1],
                    )
                    if relative_width is not None
                    else None
                ),
                height=(
                    FloatingImageRelativeHeight(
                        relative_to=cast(
                            RelativeHeightTo,
                            relative_size_height[0],
                        ),
                        percentage=relative_size_height[1],
                    )
                    if relative_size_height is not None
                    else None
                ),
            )
            if (
                relative_width is not None
                or relative_size_height is not None
            )
            else None
        ),
        relative_height=relative_height,
        behind_text=bool(booleans["behind_text"]),
        locked=bool(booleans["locked"]),
        layout_in_cell=bool(booleans["layout_in_cell"]),
        allow_overlap=bool(booleans["allow_overlap"]),
    )


def simple_native_image(
    package: NativePackage,
    paragraph: ET.Element,
    *,
    source_part: str,
) -> SimpleNativeImage | None:
    """Return metadata only for one unambiguous embedded native picture.

    Anything that could conceal additional visible content or materially alter
    the picture geometry remains opaque. The original package is never changed.
    """

    if paragraph.tag != _q(W, "p"):
        return None
    if any(child.tag not in {_q(W, "pPr"), _q(W, "r")} for child in paragraph):
        return None

    drawings: list[ET.Element] = []
    for run in paragraph.findall(f"./{_q(W, 'r')}"):
        if any(child.tag not in {_q(W, "rPr"), _q(W, "drawing")} for child in run):
            return None
        drawings.extend(run.findall(f"./{_q(W, 'drawing')}"))
    if len(drawings) != 1:
        return None

    drawing = drawings[0]
    if len(drawing) != 1:
        return None
    drawing_container = drawing[0]
    placement: Literal["inline", "floating"]
    floating: FloatingImageLayout | None
    if drawing_container.tag == _q(WP, "inline"):
        placement = "inline"
        floating = None
        if any(
            attribute not in {"distT", "distB", "distL", "distR"}
            for attribute in drawing_container.attrib
        ):
            return None
        allowed_inline_children = {
            _q(WP, "extent"),
            _q(WP, "effectExtent"),
            _q(WP, "docPr"),
            _q(WP, "cNvGraphicFramePr"),
            _q(A, "graphic"),
        }
        if any(
            child.tag not in allowed_inline_children
            for child in drawing_container
        ):
            return None
    elif drawing_container.tag == _q(WP, "anchor"):
        placement = "floating"
        floating = _floating_image_layout(drawing_container)
        if floating is None:
            return None
    else:
        return None

    extent = drawing_container.find(f"./{_q(WP, 'extent')}")
    document_properties = drawing_container.find(f"./{_q(WP, 'docPr')}")
    graphic = drawing_container.find(f"./{_q(A, 'graphic')}")
    if (
        extent is None
        or document_properties is None
        or graphic is None
        or len(
            drawing_container.findall(f"./{_q(WP, 'extent')}")
        )
        != 1
        or len(
            drawing_container.findall(f"./{_q(WP, 'docPr')}")
        )
        != 1
        or len(drawing_container.findall(f"./{_q(A, 'graphic')}")) != 1
        or len(
            drawing_container.findall(f"./{_q(WP, 'effectExtent')}")
        )
        > 1
        or len(
            drawing_container.findall(
                f"./{_q(WP, 'cNvGraphicFramePr')}"
            )
        )
        > 1
    ):
        return None
    if _enabled(document_properties.attrib.get("hidden")):
        return None
    try:
        width_emu = int(extent.attrib.get("cx", ""))
        height_emu = int(extent.attrib.get("cy", ""))
    except ValueError:
        return None
    if width_emu <= 0 or height_emu <= 0:
        return None

    effect_extent = drawing_container.find(
        f"./{_q(WP, 'effectExtent')}"
    )
    try:
        parsed_effect_extent = _floating_effect_extent(effect_extent)
    except ValueError:
        return None
    if (
        placement == "inline"
        and parsed_effect_extent is not None
        and any(
            round(
                getattr(parsed_effect_extent, field_name).to_points()
                * EMU_PER_POINT
            )
            != 0
            for field_name, _ in _EFFECT_EXTENT_FIELDS
        )
    ):
        return None

    graphic_data = graphic.find(f"./{_q(A, 'graphicData')}")
    if (
        graphic_data is None
        or graphic_data.attrib.get("uri") != PIC
        or len(graphic_data) != 1
        or graphic_data[0].tag != _q(PIC, "pic")
    ):
        return None
    picture = graphic_data[0]
    allowed_picture_children = {
        _q(PIC, "nvPicPr"),
        _q(PIC, "blipFill"),
        _q(PIC, "spPr"),
    }
    if (
        any(
            child.tag not in allowed_picture_children
            for child in picture
        )
        or len(picture.findall(f"./{_q(PIC, 'nvPicPr')}")) != 1
        or len(picture.findall(f"./{_q(PIC, 'blipFill')}")) != 1
        or len(picture.findall(f"./{_q(PIC, 'spPr')}")) != 1
    ):
        return None

    blip_fill = picture.find(f"./{_q(PIC, 'blipFill')}")
    shape_properties = picture.find(f"./{_q(PIC, 'spPr')}")
    assert blip_fill is not None
    assert shape_properties is not None
    blips = blip_fill.findall(f"./{_q(A, 'blip')}")
    source_rectangles = blip_fill.findall(f"./{_q(A, 'srcRect')}")
    stretch = blip_fill.find(f"./{_q(A, 'stretch')}")
    expected_blip_fill_children = [
        _q(A, "blip"),
        *(
            [_q(A, "srcRect")]
            if source_rectangles
            else []
        ),
        _q(A, "stretch"),
    ]
    if (
        len(blips) != 1
        or len(source_rectangles) > 1
        or stretch is None
        or len(stretch) != 1
        or stretch[0].tag != _q(A, "fillRect")
        or [child.tag for child in blip_fill]
        != expected_blip_fill_children
    ):
        return None
    blip = blips[0]
    relationship_id = blip.attrib.get(_q(R, "embed"))
    if not relationship_id or blip.attrib.get(_q(R, "link")):
        return None

    crop: ImageCrop | None = None
    if source_rectangles:
        source_rectangle = source_rectangles[0]
        if (
            len(source_rectangle)
            or any(
                attribute not in {"l", "t", "r", "b"}
                for attribute in source_rectangle.attrib
            )
        ):
            return None
        try:
            crop_values = {
                edge: int(source_rectangle.get(attribute, "0"))
                for edge, attribute in (
                    ("left", "l"),
                    ("top", "t"),
                    ("right", "r"),
                    ("bottom", "b"),
                )
            }
        except ValueError:
            return None
        if (
            any(
                value < 0 or value >= 100_000
                for value in crop_values.values()
            )
            or crop_values["left"] + crop_values["right"]
            >= 100_000
            or crop_values["top"] + crop_values["bottom"]
            >= 100_000
        ):
            return None
        if any(crop_values.values()):
            crop = ImageCrop(
                **{
                    edge: value / 1_000
                    for edge, value in crop_values.items()
                }
            )
    if any(
        attribute != "bwMode"
        for attribute in shape_properties.attrib
    ) or shape_properties.attrib.get("bwMode") not in {None, "auto"}:
        return None
    preset_geometry = shape_properties.find(
        f"./{_q(A, 'prstGeom')}"
    )
    if (
        preset_geometry is None
        or preset_geometry.attrib.get("prst") != "rect"
        or shape_properties.find(f"./{_q(A, 'custGeom')}") is not None
    ):
        return None
    allowed_shape_children = {
        _q(A, "xfrm"),
        _q(A, "prstGeom"),
        _q(A, "noFill"),
        _q(A, "ln"),
    }
    if any(
        child.tag not in allowed_shape_children
        for child in shape_properties
    ):
        return None
    no_fills = shape_properties.findall(f"./{_q(A, 'noFill')}")
    if (
        len(no_fills) > 1
        or any(no_fill.attrib or len(no_fill) for no_fill in no_fills)
    ):
        return None
    transforms = shape_properties.findall(f"./{_q(A, 'xfrm')}")
    if len(transforms) != 1:
        return None
    transform_extent = transforms[0].find(f"./{_q(A, 'ext')}")
    if (
        transform_extent is None
        or len(transforms[0].findall(f"./{_q(A, 'ext')}")) != 1
    ):
        return None
    try:
        transform_width = int(transform_extent.attrib.get("cx", ""))
        transform_height = int(transform_extent.attrib.get("cy", ""))
    except ValueError:
        return None
    if (
        transform_width != width_emu
        or transform_height != height_emu
    ):
        return None
    for transform in transforms:
        if (
            _enabled(transform.attrib.get("flipH"))
            or _enabled(transform.attrib.get("flipV"))
        ):
            return None
        try:
            rotation = int(transform.attrib.get("rot", "0"))
        except ValueError:
            return None
        if rotation:
            return None
    for outline in picture.findall(f".//{_q(A, 'ln')}"):
        if (
            outline.attrib
            or len(outline) != 1
            or outline[0].tag != _q(A, "noFill")
        ):
            return None
    if any(
        _local_name(element.tag) in _VISUAL_EFFECT_NAMES
        for element in picture.iter()
    ):
        return None

    relationships = [
        relationship
        for relationship in package.relationships
        if relationship.source_part == source_part
        and relationship.relationship_id == relationship_id
        and relationship.relationship_type == IMAGE_RELATIONSHIP_TYPE
        and not relationship.external
    ]
    if len(relationships) != 1:
        return None
    part_uri = resolve_relationship_target(
        source_part,
        relationships[0].target,
    )
    part = next(
        (candidate for candidate in package.parts if candidate.uri == part_uri),
        None,
    )
    if (
        part is None
        or not part.content_type.casefold().startswith("image/")
        or not package.has_part(part_uri)
    ):
        return None

    def optional_attribute(name: str) -> str | None:
        value = document_properties.attrib.get(name)
        return value if value else None

    return SimpleNativeImage(
        relationship_id=relationship_id,
        part_uri=part_uri,
        media_type=part.content_type,
        filename=PurePosixPath(part_uri).name,
        sha256=part.sha256.lower(),
        size_bytes=part.size,
        width=Length(
            value=round(width_emu / EMU_PER_POINT, 6),
            unit="pt",
        ),
        height=Length(
            value=round(height_emu / EMU_PER_POINT, 6),
            unit="pt",
        ),
        crop=crop,
        placement=placement,
        floating=floating,
        name=optional_attribute("name"),
        alt_text=optional_attribute("descr"),
        title=optional_attribute("title"),
        native_drawing_id=optional_attribute("id"),
    )


def simple_native_image_from_ref(
    package: NativePackage,
    source_ref: NativeRef,
) -> SimpleNativeImage:
    """Resolve and re-verify one image through its trusted native paragraph."""

    if (
        source_ref.format != "docx"
        or source_ref.native_kind != "w:p"
        or source_ref.element_index is None
        or (
            source_ref.element_indices
            and source_ref.element_indices != [source_ref.element_index]
        )
    ):
        raise NativePackageError(
            "Image source reference is not one DOCX paragraph."
        )
    root = parse_xml(package.get_part(source_ref.part_uri))
    if source_ref.part_uri == "/word/document.xml":
        container = root.find(_q(W, "body"))
        if container is None:
            raise NativePackageError(
                "DOCX main document part has no w:body."
            )
    elif root.tag in {_q(W, "hdr"), _q(W, "ftr")}:
        container = root
    else:
        raise NativePackageError(
            "Image source reference points to an unsupported DOCX part."
        )
    elements = list(container)
    if source_ref.element_index >= len(elements):
        raise NativePackageError(
            "Image source reference points outside its DOCX container."
        )
    paragraph = elements[source_ref.element_index]
    image = simple_native_image(
        package,
        paragraph,
        source_part=source_ref.part_uri,
    )
    if image is None:
        raise NativePackageError(
            "Native image no longer matches its conservative projection."
        )
    return image


def patch_simple_native_image(
    package: NativePackage,
    paragraph: ET.Element,
    *,
    source_part: str,
    result: ImageBlock,
    fields: set[str],
) -> None:
    """Selectively update one already-proven native image occurrence."""

    original = simple_native_image(
        package,
        paragraph,
        source_part=source_part,
    )
    if original is None:
        raise NativePackageError(
            "image.update requires one supported native picture."
        )
    drawing_container = paragraph.find(
        f"./{_q(W, 'r')}/{_q(W, 'drawing')}/*"
    )
    if drawing_container is None or drawing_container.tag not in {
        _q(WP, "inline"),
        _q(WP, "anchor"),
    }:
        raise NativePackageError(
            "Supported native image has no DrawingML placement container."
        )
    document_properties = drawing_container.find(f"./{_q(WP, 'docPr')}")
    if document_properties is None:
        raise NativePackageError(
            "Supported native image has no wp:docPr element."
        )

    if fields.intersection({"width", "height"}):
        width_emu = round(result.width.to_points() * EMU_PER_POINT)
        height_emu = round(result.height.to_points() * EMU_PER_POINT)
        if (
            width_emu <= 0
            or height_emu <= 0
            or width_emu > 2**63 - 1
            or height_emu > 2**63 - 1
        ):
            raise NativePackageError(
                "image.update dimensions do not fit positive OOXML Int64 EMUs."
            )
        outer_extent = drawing_container.find(f"./{_q(WP, 'extent')}")
        inner_extent = drawing_container.find(
            f"./{_q(A, 'graphic')}/{_q(A, 'graphicData')}/"
            f"{_q(PIC, 'pic')}/{_q(PIC, 'spPr')}/"
            f"{_q(A, 'xfrm')}/{_q(A, 'ext')}"
        )
        if outer_extent is None or inner_extent is None:
            raise NativePackageError(
                "Supported native image has incomplete size geometry."
            )
        for extent in (outer_extent, inner_extent):
            extent.set("cx", str(width_emu))
            extent.set("cy", str(height_emu))

    for field_name, attribute_name in (
        ("alt_text", "descr"),
        ("title", "title"),
    ):
        if field_name not in fields:
            continue
        value = getattr(result, field_name)
        if value is None:
            document_properties.attrib.pop(attribute_name, None)
        else:
            document_properties.set(attribute_name, value)

    if "crop" in fields:
        blip_fill = drawing_container.find(
            f"./{_q(A, 'graphic')}/{_q(A, 'graphicData')}/"
            f"{_q(PIC, 'pic')}/{_q(PIC, 'blipFill')}"
        )
        if blip_fill is None:
            raise NativePackageError(
                "Supported native image has no pic:blipFill element."
            )
        source_rectangles = blip_fill.findall(f"./{_q(A, 'srcRect')}")
        if len(source_rectangles) > 1:
            raise NativePackageError(
                "Supported native image has duplicate source crop rectangles."
            )
        if result.crop is None:
            if source_rectangles:
                blip_fill.remove(source_rectangles[0])
        else:
            source_rectangle = (
                source_rectangles[0]
                if source_rectangles
                else ET.Element(_q(A, "srcRect"))
            )
            if not source_rectangles:
                stretch = blip_fill.find(f"./{_q(A, 'stretch')}")
                if stretch is None:
                    raise NativePackageError(
                        "Supported native image has no stretch fill."
                    )
                blip_fill.insert(
                    list(blip_fill).index(stretch),
                    source_rectangle,
                )
            source_rectangle.attrib.clear()
            for field_name, attribute_name in (
                ("left", "l"),
                ("top", "t"),
                ("right", "r"),
                ("bottom", "b"),
            ):
                value = round(
                    getattr(result.crop, field_name) * 1_000
                )
                if value:
                    source_rectangle.set(attribute_name, str(value))

    verified = simple_native_image(
        package,
        paragraph,
        source_part=source_part,
    )
    if (
        verified is None
        or verified.crop != result.crop
        or verified.placement != original.placement
        or not floating_image_layout_matches(
            verified.floating,
            original.floating,
        )
    ):
        raise NativePackageError(
            "image.update would leave the picture outside the supported subset."
        )


def patch_simple_native_image_anchor(
    package: NativePackage,
    part_root: ET.Element,
    paragraph: ET.Element,
    *,
    source_part: str,
    result: ImageBlock,
    fields: set[str],
) -> None:
    """Selectively update one already-proven floating image anchor."""

    original = simple_native_image(
        package,
        paragraph,
        source_part=source_part,
    )
    if (
        original is None
        or original.placement != "floating"
        or original.floating is None
        or result.placement != "floating"
        or result.floating is None
    ):
        raise NativePackageError(
            "image.anchor.update requires one supported floating picture."
        )
    anchor = paragraph.find(
        f"./{_q(W, 'r')}/{_q(W, 'drawing')}/{_q(WP, 'anchor')}"
    )
    if anchor is None:
        raise NativePackageError(
            "Supported floating picture has no wp:anchor element."
        )

    def position(
        name: str,
        *,
        relative_to: str,
        offset: Length | None,
        alignment: str | None,
        percentage_offset: float | None,
        percentage_tag: str,
        native_frames: Mapping[str, str],
        allowed_alignments: frozenset[str],
    ) -> None:
        element = anchor.find(f"./{_q(WP, name)}")
        if element is None or len(element) != 1:
            raise NativePackageError(
                f"Supported floating picture has invalid wp:{name}."
            )
        native_frame = native_frames.get(relative_to)
        native_value = _native_position_value(
            offset=offset,
            alignment=alignment,
            percentage_offset=percentage_offset,
            percentage_tag=percentage_tag,
            allowed_alignments=allowed_alignments,
        )
        if native_frame is None or native_value is None:
            raise NativePackageError(
                f"image.anchor.update {name} is outside the supported range."
            )
        element.set("relativeFrom", native_frame)
        element.remove(element[0])
        ET.SubElement(element, native_value[0]).text = native_value[1]
        if native_value[0].startswith(f"{{{WP14}}}"):
            _ensure_wp14_compatibility(
                package,
                part_root,
                source_part=source_part,
            )

    layout = result.floating
    if "horizontal" in fields:
        position(
            "positionH",
            relative_to=layout.horizontal.relative_to,
            offset=layout.horizontal.offset,
            alignment=layout.horizontal.alignment,
            percentage_offset=layout.horizontal.percentage_offset,
            percentage_tag="pctPosHOffset",
            native_frames=_HORIZONTAL_RELATIVE_TO_NATIVE,
            allowed_alignments=_HORIZONTAL_ALIGNMENTS,
        )
    if "vertical" in fields:
        position(
            "positionV",
            relative_to=layout.vertical.relative_to,
            offset=layout.vertical.offset,
            alignment=layout.vertical.alignment,
            percentage_offset=layout.vertical.percentage_offset,
            percentage_tag="pctPosVOffset",
            native_frames=_VERTICAL_RELATIVE_TO_NATIVE,
            allowed_alignments=_VERTICAL_ALIGNMENTS,
        )
    if "wrap" in fields:
        wraps = [
            child
            for child in anchor
            if child.tag in _WRAP_TAG_TO_MODE
        ]
        native_wrap = _native_wrap_value(layout.wrap)
        if len(wraps) != 1 or native_wrap is None:
            raise NativePackageError(
                "Supported floating picture has invalid text wrapping."
            )
        wrap = wraps[0]
        wrap_index = list(anchor).index(wrap)
        anchor.remove(wrap)
        anchor.insert(wrap_index, native_wrap)
    if "relative_size" in fields:
        for tag in ("sizeRelH", "sizeRelV"):
            existing = anchor.find(f"./{_q(WP14, tag)}")
            if existing is not None:
                anchor.remove(existing)
        native_relative_size = _native_relative_size(
            layout.relative_size
        )
        if (
            layout.relative_size is not None
            and native_relative_size is None
        ):
            raise NativePackageError(
                "image.anchor.update relative size is outside the "
                "supported range."
            )
        for element in native_relative_size or []:
            anchor.append(element)
        if native_relative_size:
            _ensure_wp14_compatibility(
                package,
                part_root,
                source_part=source_part,
            )
    if "anchor_distances" in fields or "wrap" in fields:
        native_distances = _native_text_distance_attributes(
            layout.anchor_distances,
            allowed_edges=frozenset(
                {"top", "right", "bottom", "left"}
            ),
        )
        if native_distances is None:
            raise NativePackageError(
                "image.anchor.update anchor distance is outside the "
                "supported range."
            )
        for _, attribute_name in _TEXT_DISTANCE_FIELDS:
            native_value = native_distances.get(attribute_name)
            if native_value is None:
                anchor.attrib.pop(attribute_name, None)
            else:
                anchor.set(attribute_name, native_value)
    if "anchor_effect_extent" in fields:
        existing_effect_extent = anchor.find(
            f"./{_q(WP, 'effectExtent')}"
        )
        if existing_effect_extent is not None:
            anchor.remove(existing_effect_extent)
        if layout.anchor_effect_extent is not None:
            native_effect_extent = _native_effect_extent(
                layout.anchor_effect_extent
            )
            extent = anchor.find(f"./{_q(WP, 'extent')}")
            if native_effect_extent is None or extent is None:
                raise NativePackageError(
                    "image.anchor.update effect extent is outside the "
                    "supported range."
                )
            anchor.insert(
                list(anchor).index(extent) + 1,
                native_effect_extent,
            )
    if "relative_height" in fields:
        anchor.set("relativeHeight", str(layout.relative_height))
    for field_name, attribute_name in (
        ("behind_text", "behindDoc"),
        ("locked", "locked"),
        ("layout_in_cell", "layoutInCell"),
        ("allow_overlap", "allowOverlap"),
    ):
        if field_name in fields:
            anchor.set(
                attribute_name,
                "1" if getattr(layout, field_name) else "0",
            )

    verified = simple_native_image(
        package,
        paragraph,
        source_part=source_part,
    )
    if (
        verified is None
        or verified.placement != "floating"
        or not floating_image_layout_matches(
            verified.floating,
            result.floating,
        )
        or verified.asset_id != original.asset_id
        or verified.width != original.width
        or verified.height != original.height
        or verified.crop != original.crop
        or verified.name != original.name
        or verified.alt_text != original.alt_text
        or verified.title != original.title
    ):
        raise NativePackageError(
            "image.anchor.update would leave the picture outside the "
            "supported subset."
        )


def _relationship_part_uri(source_part: str) -> str:
    source = PurePosixPath(source_part)
    return str(
        source.parent
        / "_rels"
        / f"{source.name}.rels"
    )


def _attach_image_asset(
    package: NativePackage,
    *,
    source_part: str,
    asset: AssetRef,
    payload: bytes,
) -> tuple[str, str]:
    """Add/reuse one content-addressed image part and add a fresh relationship."""

    if (
        asset.size_bytes is None
        or asset.size_bytes != len(payload)
        or asset.sha256 != hashlib.sha256(payload).hexdigest()
        or asset.id != f"asset_{asset.sha256}"
        or not asset.filename
    ):
        raise NativePackageError(
            "Image asset metadata does not match its binary payload."
        )
    expected_extension = _REPLACEMENT_EXTENSIONS.get(asset.media_type)
    expected_filename = (
        f"aioffice-{asset.sha256}.{expected_extension}"
        if expected_extension is not None
        else None
    )
    if asset.filename != expected_filename:
        raise NativePackageError(
            "Native image mutation requires a supported media type and canonical "
            "content-addressed native filename."
        )
    media_part_uri = f"/word/media/{asset.filename}"
    if package.has_part(media_part_uri):
        if package.get_part(media_part_uri) != payload:
            raise NativePackageError(
                "Content-addressed image part collision detected."
            )
        existing_part = next(
            part
            for part in package.parts
            if part.uri == media_part_uri
        )
        if existing_part.content_type != asset.media_type:
            raise NativePackageError(
                "Existing content-addressed image part has a different media type."
            )
    else:
        package.set_part(
            media_part_uri,
            payload,
            content_type=asset.media_type,
        )

    content_types = parse_xml(package.get_part("/[Content_Types].xml"))
    overrides = [
        element
        for element in content_types.findall(_q(CT, "Override"))
        if element.get("PartName") == media_part_uri
    ]
    if len(overrides) > 1:
        raise NativePackageError(
            "DOCX content types contain duplicate native image overrides."
        )
    if overrides:
        if overrides[0].get("ContentType") != asset.media_type:
            raise NativePackageError(
                "Native image content-type override is inconsistent."
            )
    else:
        ET.SubElement(
            content_types,
            _q(CT, "Override"),
            {
                "PartName": media_part_uri,
                "ContentType": asset.media_type,
            },
        )
        package.set_part(
            "/[Content_Types].xml",
            serialize_xml(content_types),
        )

    relationship_part_uri = _relationship_part_uri(source_part)
    relationship_part_exists = package.has_part(
        relationship_part_uri
    )
    relationships = (
        parse_xml(package.get_part(relationship_part_uri))
        if relationship_part_exists
        else ET.Element(_q(REL, "Relationships"))
    )
    if relationships.tag != _q(REL, "Relationships"):
        raise NativePackageError(
            f"DOCX relationship part {relationship_part_uri!r} has "
            "an invalid root."
        )
    relationship_ids = {
        element.get("Id", "")
        for element in relationships.findall(_q(REL, "Relationship"))
    }
    relationship_number = 1
    while f"rIdAiOfficeImage{relationship_number}" in relationship_ids:
        relationship_number += 1
    relationship_id = f"rIdAiOfficeImage{relationship_number}"
    relative_target = posixpath.relpath(
        media_part_uri,
        start=posixpath.dirname(source_part),
    )
    ET.SubElement(
        relationships,
        _q(REL, "Relationship"),
        {
            "Id": relationship_id,
            "Type": IMAGE_RELATIONSHIP_TYPE,
            "Target": relative_target,
        },
    )
    package.set_part(
        relationship_part_uri,
        serialize_xml(relationships),
        content_type=(
            None
            if relationship_part_exists
            else RELATIONSHIPS_CONTENT_TYPE
        ),
    )
    return relationship_id, media_part_uri


def replace_simple_native_image(
    package: NativePackage,
    paragraph: ET.Element,
    *,
    source_part: str,
    asset: AssetRef,
    payload: bytes,
) -> None:
    """Replace one image occurrence through a new part and relationship."""

    original = simple_native_image(
        package,
        paragraph,
        source_part=source_part,
    )
    if original is None:
        raise NativePackageError(
            "image.replace requires one supported native picture."
        )
    relationship_id, media_part_uri = _attach_image_asset(
        package,
        source_part=source_part,
        asset=asset,
        payload=payload,
    )

    blips = paragraph.findall(f".//{_q(A, 'blip')}")
    if len(blips) != 1:
        raise NativePackageError(
            "Supported native image no longer contains exactly one a:blip."
        )
    blips[0].set(_q(R, "embed"), relationship_id)

    replaced = simple_native_image(
        package,
        paragraph,
        source_part=source_part,
    )
    if (
        replaced is None
        or replaced.asset_id != asset.id
        or replaced.part_uri != media_part_uri
        or replaced.media_type != asset.media_type
        or replaced.sha256 != asset.sha256
        or replaced.size_bytes != asset.size_bytes
        or replaced.crop != original.crop
        or replaced.placement != original.placement
        or not floating_image_layout_matches(
            replaced.floating,
            original.floating,
        )
    ):
        raise NativePackageError(
            "image.replace would leave the picture outside the supported subset."
        )


def _native_paragraph_anchor(
    container: ET.Element,
    image_id: str,
) -> str:
    existing = {
        value
        for paragraph in container.findall(_q(W, "p"))
        if (value := paragraph.get(_q(W14, "paraId"))) is not None
    }
    ordinal = 0
    while True:
        candidate = hashlib.sha256(
            f"{image_id}:{ordinal}".encode()
        ).hexdigest()[:8].upper()
        if candidate == "00000000":
            candidate = "00000001"
        if candidate not in existing:
            return candidate
        ordinal += 1


def _next_drawing_id(
    package: NativePackage,
    container: ET.Element,
) -> int:
    used: set[int] = set()
    roots = [container]
    for part in package.parts:
        if (
            part.uri == "/word/document.xml"
            or part.uri.startswith("/word/header")
            or part.uri.startswith("/word/footer")
        ) and part.uri.endswith(".xml"):
            try:
                roots.append(parse_xml(package.get_part(part.uri)))
            except NativePackageError:
                continue
    for root in roots:
        for properties in root.findall(f".//{_q(WP, 'docPr')}"):
            try:
                value = int(properties.get("id", ""))
            except ValueError:
                continue
            if 0 < value <= 2**32 - 1:
                used.add(value)
    candidate = 1
    while candidate in used:
        candidate += 1
    if candidate > 2**32 - 1:
        raise NativePackageError(
            "DOCX has no available unsigned drawing property ID."
        )
    return candidate


def insert_simple_native_image_after(
    package: NativePackage,
    part_root: ET.Element,
    container: ET.Element,
    after_elements: list[ET.Element],
    *,
    source_part: str,
    image: ImageInsert,
    asset: AssetRef,
    payload: bytes,
) -> ET.Element:
    """Insert one conservative native-picture paragraph after mapped elements."""

    if (
        source_part != "/word/document.xml"
        or container.tag != _q(W, "body")
        or not after_elements
        or any(element not in list(container) for element in after_elements)
    ):
        raise NativePackageError(
            "image.insert_after requires mapped top-level document body elements."
        )
    relationship_id, _ = _attach_image_asset(
        package,
        source_part=source_part,
        asset=asset,
        payload=payload,
    )
    width_emu = round(image.width.to_points() * EMU_PER_POINT)
    height_emu = round(image.height.to_points() * EMU_PER_POINT)
    if (
        width_emu <= 0
        or height_emu <= 0
        or width_emu > 2**63 - 1
        or height_emu > 2**63 - 1
    ):
        raise NativePackageError(
            "image.insert_after dimensions do not fit positive OOXML Int64 EMUs."
        )
    paragraph = ET.Element(
        _q(W, "p"),
        {
            _q(W14, "paraId"): _native_paragraph_anchor(
                container,
                image.id,
            )
        },
    )
    apply_paragraph_style(
        paragraph,
        image.paragraph_style,
    )
    run = ET.SubElement(paragraph, _q(W, "r"))
    drawing = ET.SubElement(run, _q(W, "drawing"))
    if image.placement == "floating":
        layout = image.floating
        if layout is None:
            raise NativePackageError(
                "Floating image insertion has no validated anchor layout."
            )
        horizontal_frame = _HORIZONTAL_RELATIVE_TO_NATIVE.get(
            layout.horizontal.relative_to
        )
        vertical_frame = _VERTICAL_RELATIVE_TO_NATIVE.get(
            layout.vertical.relative_to
        )
        native_wrap = _native_wrap_value(layout.wrap)
        horizontal_value = _native_position_value(
            offset=layout.horizontal.offset,
            alignment=layout.horizontal.alignment,
            percentage_offset=layout.horizontal.percentage_offset,
            percentage_tag="pctPosHOffset",
            allowed_alignments=_HORIZONTAL_ALIGNMENTS,
        )
        vertical_value = _native_position_value(
            offset=layout.vertical.offset,
            alignment=layout.vertical.alignment,
            percentage_offset=layout.vertical.percentage_offset,
            percentage_tag="pctPosVOffset",
            allowed_alignments=_VERTICAL_ALIGNMENTS,
        )
        distances = _native_text_distance_attributes(
            layout.anchor_distances,
            allowed_edges=frozenset(
                {"top", "right", "bottom", "left"}
            ),
        )
        if (
            horizontal_frame is None
            or vertical_frame is None
            or native_wrap is None
            or horizontal_value is None
            or vertical_value is None
            or distances is None
        ):
            raise NativePackageError(
                "Floating image insertion layout is outside the supported "
                "native range."
            )
        placement = ET.SubElement(
            drawing,
            _q(WP, "anchor"),
            {
                **{
                    attribute_name: value
                    for attribute_name, value in distances.items()
                },
                "simplePos": "0",
                "relativeHeight": str(layout.relative_height),
                "behindDoc": "1" if layout.behind_text else "0",
                "locked": "1" if layout.locked else "0",
                "layoutInCell": (
                    "1" if layout.layout_in_cell else "0"
                ),
                "allowOverlap": (
                    "1" if layout.allow_overlap else "0"
                ),
            },
        )
        ET.SubElement(
            placement,
            _q(WP, "simplePos"),
            {"x": "0", "y": "0"},
        )
        horizontal = ET.SubElement(
            placement,
            _q(WP, "positionH"),
            {"relativeFrom": horizontal_frame},
        )
        ET.SubElement(
            horizontal,
            horizontal_value[0],
        ).text = horizontal_value[1]
        vertical = ET.SubElement(
            placement,
            _q(WP, "positionV"),
            {"relativeFrom": vertical_frame},
        )
        ET.SubElement(
            vertical,
            vertical_value[0],
        ).text = vertical_value[1]
        if any(
            native_value[0].startswith(f"{{{WP14}}}")
            for native_value in (horizontal_value, vertical_value)
        ):
            _ensure_wp14_compatibility(
                package,
                part_root,
                source_part=source_part,
            )
    else:
        placement = ET.SubElement(drawing, _q(WP, "inline"))
    ET.SubElement(
        placement,
        _q(WP, "extent"),
        {"cx": str(width_emu), "cy": str(height_emu)},
    )
    if image.placement == "floating":
        assert image.floating is not None
        if image.floating.anchor_effect_extent is not None:
            native_effect_extent = _native_effect_extent(
                image.floating.anchor_effect_extent
            )
            if native_effect_extent is None:
                raise NativePackageError(
                    "Floating image insertion effect extent is outside "
                    "the supported native range."
                )
            placement.append(native_effect_extent)
        native_wrap = _native_wrap_value(image.floating.wrap)
        if native_wrap is None:
            raise NativePackageError(
                "Floating image insertion has invalid text wrapping."
            )
        placement.append(native_wrap)
    else:
        ET.SubElement(
            placement,
            _q(WP, "effectExtent"),
            {"l": "0", "t": "0", "r": "0", "b": "0"},
        )
    document_properties = {
        "id": str(_next_drawing_id(package, container)),
        "name": image.name or asset.filename or "AiOffice image",
        "descr": image.alt_text,
    }
    if image.title is not None:
        document_properties["title"] = image.title
    ET.SubElement(
        placement,
        _q(WP, "docPr"),
        document_properties,
    )
    frame_properties = ET.SubElement(
        placement,
        _q(WP, "cNvGraphicFramePr"),
    )
    ET.SubElement(
        frame_properties,
        _q(A, "graphicFrameLocks"),
        {"noChangeAspect": "1"},
    )
    graphic = ET.SubElement(placement, _q(A, "graphic"))
    graphic_data = ET.SubElement(
        graphic,
        _q(A, "graphicData"),
        {"uri": PIC},
    )
    picture = ET.SubElement(graphic_data, _q(PIC, "pic"))
    non_visual = ET.SubElement(picture, _q(PIC, "nvPicPr"))
    ET.SubElement(
        non_visual,
        _q(PIC, "cNvPr"),
        {
            "id": "0",
            "name": asset.filename or "AiOffice image",
        },
    )
    ET.SubElement(non_visual, _q(PIC, "cNvPicPr"))
    blip_fill = ET.SubElement(picture, _q(PIC, "blipFill"))
    ET.SubElement(
        blip_fill,
        _q(A, "blip"),
        {_q(R, "embed"): relationship_id},
    )
    stretch = ET.SubElement(blip_fill, _q(A, "stretch"))
    ET.SubElement(stretch, _q(A, "fillRect"))
    shape = ET.SubElement(picture, _q(PIC, "spPr"))
    transform = ET.SubElement(shape, _q(A, "xfrm"))
    ET.SubElement(transform, _q(A, "off"), {"x": "0", "y": "0"})
    ET.SubElement(
        transform,
        _q(A, "ext"),
        {"cx": str(width_emu), "cy": str(height_emu)},
    )
    geometry = ET.SubElement(
        shape,
        _q(A, "prstGeom"),
        {"prst": "rect"},
    )
    ET.SubElement(geometry, _q(A, "avLst"))
    if image.placement == "floating":
        assert image.floating is not None
        native_relative_size = _native_relative_size(
            image.floating.relative_size
        )
        if (
            image.floating.relative_size is not None
            and native_relative_size is None
        ):
            raise NativePackageError(
                "Floating image insertion relative size is outside the "
                "supported range."
            )
        for element in native_relative_size or []:
            placement.append(element)
        if native_relative_size:
            _ensure_wp14_compatibility(
                package,
                part_root,
                source_part=source_part,
            )

    insert_index = max(
        list(container).index(element)
        for element in after_elements
    ) + 1
    container.insert(insert_index, paragraph)
    projected = simple_native_image(
        package,
        paragraph,
        source_part=source_part,
    )
    if (
        projected is None
        or projected.asset_id != asset.id
        or round(projected.width.to_points() * EMU_PER_POINT)
        != width_emu
        or round(projected.height.to_points() * EMU_PER_POINT)
        != height_emu
        or projected.alt_text != image.alt_text
        or projected.title != image.title
        or projected.placement != image.placement
        or not floating_image_layout_matches(
            projected.floating,
            image.floating,
        )
    ):
        raise NativePackageError(
            "image.insert_after would create an unsupported native picture."
        )
    return paragraph


__all__ = [
    "EMU_PER_POINT",
    "IMAGE_RELATIONSHIP_TYPE",
    "SimpleNativeImage",
    "floating_image_layout_matches",
    "insert_simple_native_image_after",
    "patch_simple_native_image_anchor",
    "patch_simple_native_image",
    "replace_simple_native_image",
    "simple_native_image",
    "simple_native_image_from_ref",
]
