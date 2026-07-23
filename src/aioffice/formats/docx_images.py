"""Conservative projection and verified extraction of native DOCX images."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import PurePosixPath
from xml.etree import ElementTree as ET

from aioffice.core.errors import NativePackageError
from aioffice.formats.docx_header_footer import resolve_relationship_target
from aioffice.native import NativePackage
from aioffice.native.xml import parse_xml
from aioffice.spec.models import Length, NativeRef

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"

IMAGE_RELATIONSHIP_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
)
EMU_PER_POINT = 12_700

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
class SimpleInlineImage:
    """Trusted metadata for the deliberately small image projection subset."""

    relationship_id: str
    part_uri: str
    media_type: str
    filename: str
    sha256: str
    size_bytes: int
    width: Length
    height: Length
    name: str | None
    alt_text: str | None
    title: str | None
    native_drawing_id: str | None

    @property
    def asset_id(self) -> str:
        return f"asset_{self.sha256.lower()}"


def simple_inline_image(
    package: NativePackage,
    paragraph: ET.Element,
    *,
    source_part: str,
) -> SimpleInlineImage | None:
    """Return metadata only for one unambiguous, embedded inline picture.

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
    if list(child.tag for child in drawing) != [_q(WP, "inline")]:
        return None
    inline = drawing[0]
    if any(
        attribute not in {"distT", "distB", "distL", "distR"}
        for attribute in inline.attrib
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
        for child in inline
    ):
        return None

    extent = inline.find(f"./{_q(WP, 'extent')}")
    document_properties = inline.find(f"./{_q(WP, 'docPr')}")
    graphic = inline.find(f"./{_q(A, 'graphic')}")
    if (
        extent is None
        or document_properties is None
        or graphic is None
        or len(inline.findall(f"./{_q(WP, 'extent')}")) != 1
        or len(inline.findall(f"./{_q(WP, 'docPr')}")) != 1
        or len(inline.findall(f"./{_q(A, 'graphic')}")) != 1
        or len(inline.findall(f"./{_q(WP, 'effectExtent')}")) > 1
        or len(
            inline.findall(f"./{_q(WP, 'cNvGraphicFramePr')}")
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

    effect_extent = inline.find(f"./{_q(WP, 'effectExtent')}")
    if effect_extent is not None:
        try:
            effect_values = [
                int(effect_extent.attrib.get(edge, "0"))
                for edge in ("l", "t", "r", "b")
            ]
        except ValueError:
            return None
        if any(effect_values):
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
    stretch = blip_fill.find(f"./{_q(A, 'stretch')}")
    if (
        len(blips) != 1
        or stretch is None
        or len(stretch) != 1
        or stretch[0].tag != _q(A, "fillRect")
        or len(blip_fill) != 2
    ):
        return None
    blip = blips[0]
    relationship_id = blip.attrib.get(_q(R, "embed"))
    if not relationship_id or blip.attrib.get(_q(R, "link")):
        return None

    if picture.find(f".//{_q(A, 'srcRect')}") is not None:
        return None
    if shape_properties.attrib:
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
        _q(A, "ln"),
    }
    if any(
        child.tag not in allowed_shape_children
        for child in shape_properties
    ):
        return None
    for transform in picture.findall(f".//{_q(A, 'xfrm')}"):
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

    return SimpleInlineImage(
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
        name=optional_attribute("name"),
        alt_text=optional_attribute("descr"),
        title=optional_attribute("title"),
        native_drawing_id=optional_attribute("id"),
    )


def simple_inline_image_from_ref(
    package: NativePackage,
    source_ref: NativeRef,
) -> SimpleInlineImage:
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
    image = simple_inline_image(
        package,
        paragraph,
        source_part=source_ref.part_uri,
    )
    if image is None:
        raise NativePackageError(
            "Native image no longer matches its conservative projection."
        )
    return image


__all__ = [
    "EMU_PER_POINT",
    "IMAGE_RELATIONSHIP_TYPE",
    "SimpleInlineImage",
    "simple_inline_image",
    "simple_inline_image_from_ref",
]
