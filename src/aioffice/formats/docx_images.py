"""Conservative projection and verified extraction of native DOCX images."""

from __future__ import annotations

import hashlib
import posixpath
from dataclasses import dataclass
from pathlib import PurePosixPath
from xml.etree import ElementTree as ET

from aioffice.core.errors import NativePackageError
from aioffice.formats.docx_header_footer import resolve_relationship_target
from aioffice.native import NativePackage
from aioffice.native.xml import parse_xml, serialize_xml
from aioffice.spec.models import AssetRef, ImageBlock, Length, NativeRef

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
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


def patch_simple_inline_image(
    package: NativePackage,
    paragraph: ET.Element,
    *,
    source_part: str,
    result: ImageBlock,
    fields: set[str],
) -> None:
    """Selectively update one already-proven inline image occurrence."""

    if simple_inline_image(
        package,
        paragraph,
        source_part=source_part,
    ) is None:
        raise NativePackageError(
            "image.update requires one supported native inline picture."
        )
    inline = paragraph.find(
        f"./{_q(W, 'r')}/{_q(W, 'drawing')}/{_q(WP, 'inline')}"
    )
    if inline is None:
        raise NativePackageError(
            "Supported native image has no wp:inline element."
        )
    document_properties = inline.find(f"./{_q(WP, 'docPr')}")
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
        outer_extent = inline.find(f"./{_q(WP, 'extent')}")
        inner_extent = inline.find(
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

    if simple_inline_image(
        package,
        paragraph,
        source_part=source_part,
    ) is None:
        raise NativePackageError(
            "image.update would leave the picture outside the supported subset."
        )


def _relationship_part_uri(source_part: str) -> str:
    source = PurePosixPath(source_part)
    return str(
        source.parent
        / "_rels"
        / f"{source.name}.rels"
    )


def replace_simple_inline_image(
    package: NativePackage,
    paragraph: ET.Element,
    *,
    source_part: str,
    asset: AssetRef,
    payload: bytes,
) -> None:
    """Replace one image occurrence through a new part and relationship."""

    original = simple_inline_image(
        package,
        paragraph,
        source_part=source_part,
    )
    if original is None:
        raise NativePackageError(
            "image.replace requires one supported native inline picture."
        )
    if (
        asset.size_bytes is None
        or asset.size_bytes != len(payload)
        or asset.sha256 != hashlib.sha256(payload).hexdigest()
        or asset.id != f"asset_{asset.sha256}"
        or not asset.filename
    ):
        raise NativePackageError(
            "image.replace asset metadata does not match its binary payload."
        )
    expected_extension = _REPLACEMENT_EXTENSIONS.get(asset.media_type)
    expected_filename = (
        f"aioffice-{asset.sha256}.{expected_extension}"
        if expected_extension is not None
        else None
    )
    if asset.filename != expected_filename:
        raise NativePackageError(
            "image.replace requires a supported media type and canonical "
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
            "DOCX content types contain duplicate replacement image overrides."
        )
    if overrides:
        if overrides[0].get("ContentType") != asset.media_type:
            raise NativePackageError(
                "Replacement image content-type override is inconsistent."
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
    if not package.has_part(relationship_part_uri):
        raise NativePackageError(
            f"DOCX image source has no relationship part {relationship_part_uri!r}."
        )
    relationships = parse_xml(package.get_part(relationship_part_uri))
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
    )

    blips = paragraph.findall(f".//{_q(A, 'blip')}")
    if len(blips) != 1:
        raise NativePackageError(
            "Supported native image no longer contains exactly one a:blip."
        )
    blips[0].set(_q(R, "embed"), relationship_id)

    replaced = simple_inline_image(
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
    ):
        raise NativePackageError(
            "image.replace would leave the picture outside the supported subset."
        )


__all__ = [
    "EMU_PER_POINT",
    "IMAGE_RELATIONSHIP_TYPE",
    "SimpleInlineImage",
    "patch_simple_inline_image",
    "replace_simple_inline_image",
    "simple_inline_image",
    "simple_inline_image_from_ref",
]
