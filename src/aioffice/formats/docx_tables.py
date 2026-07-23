"""WordprocessingML table geometry, projection, and selective mutation."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Literal, cast
from xml.etree import ElementTree as ET

from aioffice.native.identity import fingerprint_elements
from aioffice.spec.models import (
    Length,
    NativeRef,
    TableCellFormat,
    TableLayout,
    TableRow,
    TableWidth,
)

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"

_TABLE_PROPERTY_ORDER = [
    "tblStyle",
    "tblpPr",
    "tblOverlap",
    "bidiVisual",
    "tblStyleRowBandSize",
    "tblStyleColBandSize",
    "tblW",
    "jc",
    "tblCellSpacing",
    "tblInd",
    "tblBorders",
    "shd",
    "tblLayout",
    "tblCellMar",
    "tblLook",
    "tblCaption",
    "tblDescription",
]
_ROW_PROPERTY_ORDER = [
    "cnfStyle",
    "divId",
    "gridBefore",
    "gridAfter",
    "wBefore",
    "wAfter",
    "cantSplit",
    "trHeight",
    "tblHeader",
    "tblCellSpacing",
    "jc",
    "hidden",
]
_CELL_PROPERTY_ORDER = [
    "cnfStyle",
    "tcW",
    "gridSpan",
    "hMerge",
    "vMerge",
    "tcBorders",
    "shd",
    "noWrap",
    "tcMar",
    "textDirection",
    "tcFitText",
    "vAlign",
    "hideMark",
]
_CELL_MARGIN_ORDER = [
    "top",
    "left",
    "start",
    "bottom",
    "right",
    "end",
]


@dataclass(slots=True)
class NativeTableCellPlacement:
    """One physical ``w:tc`` mapped onto the logical table grid."""

    element: ET.Element
    row_index: int
    cell_index: int
    start_column: int
    column_span: int
    row_span: int = 1
    continuation: bool = False
    anchor_row_index: int | None = None
    anchor_cell_index: int | None = None


@dataclass(slots=True)
class TableGridAnalysis:
    """Conservative proof result for a rectangular Word table grid."""

    column_count: int
    placements: list[NativeTableCellPlacement]
    logical_grid: bool
    regular_grid: bool
    reasons: list[str]


def _q(local: str) -> str:
    return f"{{{W}}}{local}"


def _local_name(element: ET.Element) -> str:
    return element.tag.rsplit("}", 1)[-1]


def _insert_ordered(
    parent: ET.Element,
    element: ET.Element,
    order: list[str],
) -> ET.Element:
    desired = _local_name(element)
    desired_index = (
        order.index(desired)
        if desired in order
        else len(order)
    )
    insert_at = len(parent)
    for index, child in enumerate(list(parent)):
        child_name = _local_name(child)
        child_index = (
            order.index(child_name)
            if child_name in order
            else len(order)
        )
        if child_index > desired_index:
            insert_at = index
            break
    parent.insert(insert_at, element)
    return element


def _ensure_properties(
    parent: ET.Element,
    local_name: str,
) -> ET.Element:
    existing = parent.find(_q(local_name))
    if existing is not None:
        return existing
    properties = ET.Element(_q(local_name))
    parent.insert(0, properties)
    return properties


def _ensure_property(
    properties: ET.Element,
    local_name: str,
    order: list[str],
) -> ET.Element:
    existing = properties.find(_q(local_name))
    if existing is not None:
        return existing
    return _insert_ordered(
        properties,
        ET.Element(_q(local_name)),
        order,
    )


def _remove_if_empty(parent: ET.Element, element: ET.Element) -> None:
    if not element.attrib and not list(element) and not (element.text or ""):
        parent.remove(element)


def _on_off(element: ET.Element | None) -> bool | None:
    if element is None:
        return None
    value = element.get(_q("val"))
    return value is None or value.casefold() not in {
        "0",
        "false",
        "off",
        "no",
    }


def _length_from_twips(value: str | None) -> Length | None:
    if value is None:
        return None
    try:
        twips = int(value)
    except ValueError:
        return None
    if twips < 0:
        return None
    return Length(value=twips / 20, unit="pt")


def _twips(value: Length) -> str:
    return str(round(value.to_points() * 20))


def _read_table_width(element: ET.Element | None) -> TableWidth | None:
    if element is None:
        return None
    width_type = element.get(_q("type"), "dxa")
    raw_value = element.get(_q("w"))
    if width_type == "auto":
        return TableWidth(mode="auto")
    if raw_value is None:
        return None
    try:
        value = int(raw_value)
    except ValueError:
        return None
    if width_type == "pct" and value > 0:
        percent = value / 50
        if percent <= 100:
            return TableWidth(mode="percent", value=percent)
        return None
    if width_type == "dxa" and value > 0:
        return TableWidth(
            mode="exact",
            value=Length(value=value / 20, unit="pt"),
        )
    return None


def _patch_table_width(
    properties: ET.Element,
    width: TableWidth | None,
) -> None:
    existing = properties.find(_q("tblW"))
    if width is None:
        if existing is not None:
            existing.attrib.pop(_q("type"), None)
            existing.attrib.pop(_q("w"), None)
            _remove_if_empty(properties, existing)
        return
    element = (
        existing
        if existing is not None
        else _ensure_property(
            properties,
            "tblW",
            _TABLE_PROPERTY_ORDER,
        )
    )
    if width.mode == "auto":
        element.set(_q("type"), "auto")
        element.set(_q("w"), "0")
    elif width.mode == "percent":
        assert isinstance(width.value, float)
        element.set(_q("type"), "pct")
        element.set(_q("w"), str(round(width.value * 50)))
    else:
        assert isinstance(width.value, Length)
        element.set(_q("type"), "dxa")
        element.set(_q("w"), _twips(width.value))


def _read_margin(
    margins: ET.Element | None,
    *names: str,
) -> Length | None:
    if margins is None:
        return None
    for name in names:
        element = margins.find(_q(name))
        if element is None:
            continue
        if element.get(_q("type"), "dxa") != "dxa":
            return None
        return _length_from_twips(element.get(_q("w")))
    return None


def _patch_measurement(
    properties: ET.Element,
    local_name: str,
    value: Length | None,
    order: list[str],
) -> None:
    existing = properties.find(_q(local_name))
    if value is None:
        if existing is not None:
            existing.attrib.pop(_q("type"), None)
            existing.attrib.pop(_q("w"), None)
            _remove_if_empty(properties, existing)
        return
    element = (
        existing
        if existing is not None
        else _ensure_property(properties, local_name, order)
    )
    element.set(_q("type"), "dxa")
    element.set(_q("w"), _twips(value))


def _patch_cell_margin(
    properties: ET.Element,
    side: str,
    value: Length | None,
) -> None:
    margins = properties.find(_q("tblCellMar"))
    if margins is None and value is not None:
        margins = _ensure_property(
            properties,
            "tblCellMar",
            _TABLE_PROPERTY_ORDER,
        )
    if margins is None:
        return
    existing = margins.find(_q(side))
    if value is None:
        if existing is not None:
            existing.attrib.pop(_q("type"), None)
            existing.attrib.pop(_q("w"), None)
            _remove_if_empty(margins, existing)
    else:
        element = (
            existing
            if existing is not None
            else _insert_ordered(
                margins,
                ET.Element(_q(side)),
                _CELL_MARGIN_ORDER,
            )
        )
        element.set(_q("type"), "dxa")
        element.set(_q("w"), _twips(value))
    _remove_if_empty(properties, margins)


def _patch_direct_cell_margin(
    properties: ET.Element,
    side: str,
    value: Length | None,
) -> None:
    margins = properties.find(_q("tcMar"))
    if margins is None and value is not None:
        margins = _ensure_property(
            properties,
            "tcMar",
            _CELL_PROPERTY_ORDER,
        )
    if margins is None:
        return
    existing = margins.find(_q(side))
    if value is None:
        if existing is not None:
            existing.attrib.pop(_q("type"), None)
            existing.attrib.pop(_q("w"), None)
            _remove_if_empty(margins, existing)
    else:
        element = (
            existing
            if existing is not None
            else _insert_ordered(
                margins,
                ET.Element(_q(side)),
                _CELL_MARGIN_ORDER,
            )
        )
        element.set(_q("type"), "dxa")
        element.set(_q("w"), _twips(value))
    _remove_if_empty(properties, margins)


def read_table_cell_format(cell: ET.Element) -> TableCellFormat:
    """Project cell-local properties that can be selectively patched."""

    properties = cell.find(_q("tcPr"))
    if properties is None:
        return TableCellFormat()
    alignment = properties.find(_q("vAlign"))
    raw_alignment = (
        alignment.get(_q("val"))
        if alignment is not None
        else None
    )
    shading = properties.find(_q("shd"))
    raw_fill = (
        shading.get(_q("fill"))
        if shading is not None
        else None
    )
    background_color = (
        f"#{raw_fill.upper()}"
        if raw_fill is not None
        and len(raw_fill) == 6
        and all(character in "0123456789ABCDEFabcdef" for character in raw_fill)
        else None
    )
    margins = properties.find(_q("tcMar"))
    return TableCellFormat(
        vertical_alignment=(
            cast(
                Literal["top", "center", "bottom"],
                raw_alignment,
            )
            if raw_alignment in {"top", "center", "bottom"}
            else None
        ),
        no_wrap=_on_off(properties.find(_q("noWrap"))),
        fit_text=_on_off(properties.find(_q("tcFitText"))),
        background_color=background_color,
        margin_top=_read_margin(margins, "top"),
        margin_right=_read_margin(margins, "right", "end"),
        margin_bottom=_read_margin(margins, "bottom"),
        margin_left=_read_margin(margins, "left", "start"),
    )


def patch_table_cell_format(
    cell: ET.Element,
    cell_format: TableCellFormat,
    fields: set[str],
) -> None:
    """Patch selected cell properties while preserving content and unknown XML."""

    properties = cell.find(_q("tcPr"))
    if properties is None and fields:
        properties = _ensure_properties(cell, "tcPr")
    if properties is None:
        return
    if "vertical_alignment" in fields:
        existing = properties.find(_q("vAlign"))
        if cell_format.vertical_alignment is None:
            if existing is not None:
                existing.attrib.pop(_q("val"), None)
                _remove_if_empty(properties, existing)
        else:
            element = (
                existing
                if existing is not None
                else _ensure_property(
                    properties,
                    "vAlign",
                    _CELL_PROPERTY_ORDER,
                )
            )
            element.set(_q("val"), cell_format.vertical_alignment)
    for field_name, local_name in (
        ("no_wrap", "noWrap"),
        ("fit_text", "tcFitText"),
    ):
        if field_name not in fields:
            continue
        value = getattr(cell_format, field_name)
        existing = properties.find(_q(local_name))
        if value is None:
            if existing is not None:
                existing.attrib.pop(_q("val"), None)
                _remove_if_empty(properties, existing)
        else:
            element = (
                existing
                if existing is not None
                else _ensure_property(
                    properties,
                    local_name,
                    _CELL_PROPERTY_ORDER,
                )
            )
            element.set(_q("val"), "1" if value else "0")
    if "background_color" in fields:
        existing = properties.find(_q("shd"))
        if cell_format.background_color is None:
            if existing is not None:
                for attribute in ("val", "color", "fill"):
                    existing.attrib.pop(_q(attribute), None)
                _remove_if_empty(properties, existing)
        else:
            element = (
                existing
                if existing is not None
                else _ensure_property(
                    properties,
                    "shd",
                    _CELL_PROPERTY_ORDER,
                )
            )
            element.set(_q("val"), "clear")
            element.set(_q("color"), "auto")
            element.set(
                _q("fill"),
                cell_format.background_color.removeprefix("#"),
            )
    for field_name, side in (
        ("margin_top", "top"),
        ("margin_right", "right"),
        ("margin_bottom", "bottom"),
        ("margin_left", "left"),
    ):
        if field_name in fields:
            _patch_direct_cell_margin(
                properties,
                side,
                getattr(cell_format, field_name),
            )
    _remove_if_empty(cell, properties)


def apply_table_cell_format(
    cell: ET.Element,
    cell_format: TableCellFormat,
) -> None:
    fields = {
        field_name
        for field_name in TableCellFormat.model_fields
        if getattr(cell_format, field_name) is not None
    }
    patch_table_cell_format(cell, cell_format, fields)


def read_table_layout(table: ET.Element) -> TableLayout:
    """Project the conservative table-wide geometry subset."""

    properties = table.find(_q("tblPr"))
    first_row = table.find(_q("tr"))
    row_properties = (
        first_row.find(_q("trPr"))
        if first_row is not None
        else None
    )
    if properties is None:
        return TableLayout(
            repeat_header=_on_off(
                row_properties.find(_q("tblHeader"))
                if row_properties is not None
                else None
            )
        )
    style = properties.find(_q("tblStyle"))
    width = properties.find(_q("tblW"))
    justification = properties.find(_q("jc"))
    layout = properties.find(_q("tblLayout"))
    indent = properties.find(_q("tblInd"))
    spacing = properties.find(_q("tblCellSpacing"))
    margins = properties.find(_q("tblCellMar"))
    raw_alignment = (
        justification.get(_q("val"))
        if justification is not None
        else None
    )
    alignment = (
        {
            "left": "left",
            "start": "left",
            "center": "center",
            "right": "right",
            "end": "right",
        }.get(raw_alignment)
        if raw_alignment is not None
        else None
    )
    raw_algorithm = (
        layout.get(_q("type"))
        if layout is not None
        else None
    )
    return TableLayout(
        style_ref=style.get(_q("val")) if style is not None else None,
        preferred_width=_read_table_width(width),
        alignment=cast(
            Literal["left", "center", "right"] | None,
            alignment,
        ),
        algorithm=(
            cast(Literal["autofit", "fixed"], raw_algorithm)
            if raw_algorithm in {"autofit", "fixed"}
            else None
        ),
        indent=(
            _length_from_twips(indent.get(_q("w")))
            if indent is not None
            and indent.get(_q("type"), "dxa") == "dxa"
            else None
        ),
        cell_spacing=(
            _length_from_twips(spacing.get(_q("w")))
            if spacing is not None
            and spacing.get(_q("type"), "dxa") == "dxa"
            else None
        ),
        cell_margin_top=_read_margin(margins, "top"),
        cell_margin_right=_read_margin(margins, "right", "end"),
        cell_margin_bottom=_read_margin(margins, "bottom"),
        cell_margin_left=_read_margin(margins, "left", "start"),
        repeat_header=_on_off(
            row_properties.find(_q("tblHeader"))
            if row_properties is not None
            else None
        ),
    )


def patch_table_layout(
    table: ET.Element,
    layout: TableLayout,
    fields: set[str],
) -> None:
    """Patch selected known geometry while preserving unrelated table XML."""

    properties = table.find(_q("tblPr"))
    if properties is None and fields - {"repeat_header"}:
        properties = _ensure_properties(table, "tblPr")
    if properties is not None:
        if "style_ref" in fields:
            existing = properties.find(_q("tblStyle"))
            if layout.style_ref is None:
                if existing is not None:
                    existing.attrib.pop(_q("val"), None)
                    _remove_if_empty(properties, existing)
            else:
                element = (
                    existing
                    if existing is not None
                    else _ensure_property(
                        properties,
                        "tblStyle",
                        _TABLE_PROPERTY_ORDER,
                    )
                )
                element.set(_q("val"), layout.style_ref)
        if "preferred_width" in fields:
            _patch_table_width(properties, layout.preferred_width)
        if "alignment" in fields:
            existing = properties.find(_q("jc"))
            if layout.alignment is None:
                if existing is not None:
                    existing.attrib.pop(_q("val"), None)
                    _remove_if_empty(properties, existing)
            else:
                element = (
                    existing
                    if existing is not None
                    else _ensure_property(
                        properties,
                        "jc",
                        _TABLE_PROPERTY_ORDER,
                    )
                )
                element.set(_q("val"), layout.alignment)
        if "algorithm" in fields:
            existing = properties.find(_q("tblLayout"))
            if layout.algorithm is None:
                if existing is not None:
                    existing.attrib.pop(_q("type"), None)
                    _remove_if_empty(properties, existing)
            else:
                element = (
                    existing
                    if existing is not None
                    else _ensure_property(
                        properties,
                        "tblLayout",
                        _TABLE_PROPERTY_ORDER,
                    )
                )
                element.set(_q("type"), layout.algorithm)
        if "indent" in fields:
            _patch_measurement(
                properties,
                "tblInd",
                layout.indent,
                _TABLE_PROPERTY_ORDER,
            )
        if "cell_spacing" in fields:
            _patch_measurement(
                properties,
                "tblCellSpacing",
                layout.cell_spacing,
                _TABLE_PROPERTY_ORDER,
            )
        for field_name, side in (
            ("cell_margin_top", "top"),
            ("cell_margin_right", "right"),
            ("cell_margin_bottom", "bottom"),
            ("cell_margin_left", "left"),
        ):
            if field_name in fields:
                _patch_cell_margin(
                    properties,
                    side,
                    getattr(layout, field_name),
                )
        _remove_if_empty(table, properties)

    if "repeat_header" in fields:
        first_row = table.find(_q("tr"))
        if first_row is None:
            return
        row_properties = first_row.find(_q("trPr"))
        if row_properties is None and layout.repeat_header is not None:
            row_properties = _ensure_properties(first_row, "trPr")
        if row_properties is not None:
            existing = row_properties.find(_q("tblHeader"))
            if layout.repeat_header is None:
                if existing is not None:
                    row_properties.remove(existing)
            else:
                element = (
                    existing
                    if existing is not None
                    else _ensure_property(
                        row_properties,
                        "tblHeader",
                        _ROW_PROPERTY_ORDER,
                    )
                )
                element.set(
                    _q("val"),
                    "1" if layout.repeat_header else "0",
                )
            _remove_if_empty(first_row, row_properties)


def apply_table_layout(table: ET.Element, layout: TableLayout) -> None:
    fields = {
        field_name
        for field_name in TableLayout.model_fields
        if getattr(layout, field_name) is not None
    }
    patch_table_layout(table, layout, fields)


def read_table_column_widths(table: ET.Element) -> list[Length | None]:
    grid = table.find(_q("tblGrid"))
    if grid is None:
        return []
    return [
        _length_from_twips(column.get(_q("w")))
        for column in grid.findall(_q("gridCol"))
    ]


def read_table_row(
    row: ET.Element,
) -> dict[str, object]:
    properties = row.find(_q("trPr"))
    if properties is None:
        return {}
    cant_split = properties.find(_q("cantSplit"))
    height = properties.find(_q("trHeight"))
    raw_height_rule = (
        height.get(_q("hRule"))
        if height is not None
        else None
    )
    payload: dict[str, object] = {}
    cant_split_value = _on_off(cant_split)
    if cant_split_value is not None:
        payload["allow_break_across_pages"] = not cant_split_value
    projected_height = (
        _length_from_twips(height.get(_q("val")))
        if height is not None
        else None
    )
    if projected_height is not None and projected_height.to_points() > 0:
        payload["height"] = projected_height.model_dump(mode="json")
        if raw_height_rule in {"atLeast", "exact"}:
            payload["height_rule"] = (
                "at_least" if raw_height_rule == "atLeast" else "exact"
            )
    return payload


def apply_table_row(row: ET.Element, table_row: TableRow) -> None:
    fields = {
        field_name
        for field_name in (
            "allow_break_across_pages",
            "height",
            "height_rule",
        )
        if getattr(table_row, field_name) is not None
    }
    if not fields:
        return
    properties = _ensure_properties(row, "trPr")
    if "allow_break_across_pages" in fields:
        cant_split = _ensure_property(
            properties,
            "cantSplit",
            _ROW_PROPERTY_ORDER,
        )
        cant_split.set(
            _q("val"),
            "0" if table_row.allow_break_across_pages else "1",
        )
    if table_row.height is not None:
        height = _ensure_property(
            properties,
            "trHeight",
            _ROW_PROPERTY_ORDER,
        )
        height.set(_q("val"), _twips(table_row.height))
        if table_row.height_rule is not None:
            height.set(
                _q("hRule"),
                (
                    "atLeast"
                    if table_row.height_rule == "at_least"
                    else "exact"
                ),
            )


def is_regular_table_grid(table: ET.Element) -> bool:
    """Whether each native row maps one-to-one onto the declared table grid."""

    return analyze_table_grid(table, start_row=0).regular_grid


def _positive_span(cell: ET.Element) -> int | None:
    properties = cell.find(_q("tcPr"))
    span = (
        properties.find(_q("gridSpan"))
        if properties is not None
        else None
    )
    raw_value = span.get(_q("val")) if span is not None else "1"
    if raw_value is None:
        return None
    try:
        value = int(raw_value)
    except (TypeError, ValueError):
        return None
    return value if value > 0 else None


def analyze_table_grid(
    table: ET.Element,
    *,
    start_row: int = 1,
) -> TableGridAnalysis:
    """Prove a logical rectangular grid with ``gridSpan``/``vMerge`` support."""

    grid = table.find(_q("tblGrid"))
    column_count = (
        len(grid.findall(_q("gridCol")))
        if grid is not None
        else 0
    )
    reasons: list[str] = []
    placements: list[NativeTableCellPlacement] = []
    if column_count == 0:
        return TableGridAnalysis(
            column_count=0,
            placements=[],
            logical_grid=False,
            regular_grid=False,
            reasons=["missing_table_grid"],
        )
    rows = table.findall(_q("tr"))
    if start_row >= len(rows):
        return TableGridAnalysis(
            column_count=column_count,
            placements=[],
            logical_grid=True,
            regular_grid=True,
            reasons=[],
        )

    active: dict[
        tuple[int, int],
        NativeTableCellPlacement,
    ] = {}
    for row_index in range(start_row, len(rows)):
        row = rows[row_index]
        row_properties = row.find(_q("trPr"))
        if row_properties is not None and (
            row_properties.find(_q("gridBefore")) is not None
            or row_properties.find(_q("gridAfter")) is not None
        ):
            reasons.append(f"row_{row_index}_shifted_grid")
        cursor = 0
        next_active: dict[
            tuple[int, int],
            NativeTableCellPlacement,
        ] = {}
        for cell_index, cell in enumerate(row.findall(_q("tc"))):
            cell_properties = cell.find(_q("tcPr"))
            if (
                cell_properties is not None
                and cell_properties.find(_q("hMerge")) is not None
            ):
                reasons.append(
                    f"row_{row_index}_cell_{cell_index}_legacy_hmerge"
                )
            column_span = _positive_span(cell)
            if column_span is None:
                reasons.append(
                    f"row_{row_index}_cell_{cell_index}_invalid_grid_span"
                )
                column_span = 1
            if cursor + column_span > column_count:
                reasons.append(
                    f"row_{row_index}_cell_{cell_index}_outside_grid"
                )
            merge = (
                cell_properties.find(_q("vMerge"))
                if cell_properties is not None
                else None
            )
            raw_merge = (
                merge.get(_q("val"))
                if merge is not None
                else None
            )
            if merge is not None and raw_merge not in {
                None,
                "continue",
                "restart",
            }:
                reasons.append(
                    f"row_{row_index}_cell_{cell_index}_invalid_vmerge"
                )
            continuation = (
                merge is not None
                and raw_merge in {None, "continue"}
            )
            key = (cursor, column_span)
            if continuation:
                anchor = active.get(key)
                if anchor is None:
                    reasons.append(
                        f"row_{row_index}_cell_{cell_index}_orphan_vmerge"
                    )
                    placement = NativeTableCellPlacement(
                        element=cell,
                        row_index=row_index,
                        cell_index=cell_index,
                        start_column=cursor,
                        column_span=column_span,
                        continuation=True,
                    )
                else:
                    anchor.row_span += 1
                    placement = NativeTableCellPlacement(
                        element=cell,
                        row_index=row_index,
                        cell_index=cell_index,
                        start_column=cursor,
                        column_span=column_span,
                        continuation=True,
                        anchor_row_index=anchor.row_index,
                        anchor_cell_index=anchor.cell_index,
                    )
                    next_active[key] = anchor
            else:
                placement = NativeTableCellPlacement(
                    element=cell,
                    row_index=row_index,
                    cell_index=cell_index,
                    start_column=cursor,
                    column_span=column_span,
                )
                if merge is not None and raw_merge == "restart":
                    next_active[key] = placement
            placements.append(placement)
            cursor += column_span
        if cursor != column_count:
            reasons.append(f"row_{row_index}_does_not_cover_grid")
        active = next_active

    logical_grid = not reasons
    regular_grid = logical_grid and all(
        not placement.continuation
        and placement.column_span == 1
        and placement.row_span == 1
        for placement in placements
    )
    return TableGridAnalysis(
        column_count=column_count,
        placements=placements,
        logical_grid=logical_grid,
        regular_grid=regular_grid,
        reasons=reasons,
    )


def patch_table_column_width(
    table: ET.Element,
    column_index: int,
    width: Length | None,
) -> None:
    """Patch one regular grid column and its one-to-one cell preferences."""

    if not is_regular_table_grid(table):
        raise ValueError(
            "Column width editing requires a regular table grid without merged cells."
        )
    grid = table.find(_q("tblGrid"))
    assert grid is not None
    columns = grid.findall(_q("gridCol"))
    if column_index >= len(columns):
        raise ValueError("Table column index points outside the native table grid.")
    grid_column = columns[column_index]
    if width is None:
        grid_column.attrib.pop(_q("w"), None)
    else:
        grid_column.set(_q("w"), _twips(width))

    for row in table.findall(_q("tr")):
        cell = row.findall(_q("tc"))[column_index]
        properties = cell.find(_q("tcPr"))
        if properties is None and width is not None:
            properties = _ensure_properties(cell, "tcPr")
        if properties is None:
            continue
        cell_width = properties.find(_q("tcW"))
        if width is None:
            if cell_width is not None:
                cell_width.attrib.pop(_q("type"), None)
                cell_width.attrib.pop(_q("w"), None)
                _remove_if_empty(properties, cell_width)
        else:
            element = (
                cell_width
                if cell_width is not None
                else _ensure_property(
                    properties,
                    "tcW",
                    _CELL_PROPERTY_ORDER,
                )
            )
            element.set(_q("type"), "dxa")
            element.set(_q("w"), _twips(width))
        _remove_if_empty(cell, properties)


def native_ref_for_table_column(
    table: ET.Element,
    table_index: int,
    column_index: int,
) -> NativeRef:
    grid = table.find(_q("tblGrid"))
    if grid is None:
        raise ValueError("Native table has no w:tblGrid.")
    columns = grid.findall(_q("gridCol"))
    if column_index >= len(columns):
        raise ValueError("Table column index points outside w:tblGrid.")
    return NativeRef(
        format="docx",
        part_uri="/word/document.xml",
        native_kind="w:gridCol",
        element_index=table_index,
        element_indices=[table_index],
        sub_index=column_index,
        path_hint=(
            f"/w:document/w:body/*[{table_index + 1}]"
            f"/w:tblGrid/w:gridCol[{column_index + 1}]"
        ),
        fingerprint=fingerprint_elements([columns[column_index]]),
    )


def native_ref_for_table_row(
    table: ET.Element,
    table_index: int,
    row_index: int,
) -> NativeRef:
    rows = table.findall(_q("tr"))
    if row_index >= len(rows):
        raise ValueError("Table row index points outside the native table.")
    return NativeRef(
        format="docx",
        part_uri="/word/document.xml",
        native_kind="w:tr",
        element_index=table_index,
        element_indices=[table_index],
        sub_index=row_index,
        path_hint=(
            f"/w:document/w:body/*[{table_index + 1}]"
            f"/w:tr[{row_index + 1}]"
        ),
        fingerprint=fingerprint_elements([rows[row_index]]),
    )


def _flatten_table_cells(table: ET.Element) -> list[ET.Element]:
    return [
        cell
        for row in table.findall(_q("tr"))
        for cell in row.findall(_q("tc"))
    ]


def _flatten_table_cell_paragraphs(
    table: ET.Element,
) -> list[ET.Element]:
    return [
        paragraph
        for cell in _flatten_table_cells(table)
        for paragraph in cell.findall(_q("p"))
    ]


def native_ref_for_table_cell(
    table: ET.Element,
    table_index: int,
    row_index: int,
    cell_index: int,
) -> NativeRef:
    rows = table.findall(_q("tr"))
    if row_index >= len(rows):
        raise ValueError("Table cell row points outside the native table.")
    cells = rows[row_index].findall(_q("tc"))
    if cell_index >= len(cells):
        raise ValueError("Table cell points outside the native row.")
    cell = cells[cell_index]
    flattened = _flatten_table_cells(table)
    return NativeRef(
        format="docx",
        part_uri="/word/document.xml",
        native_kind="w:tc",
        element_index=table_index,
        element_indices=[table_index],
        sub_index=flattened.index(cell),
        path_hint=(
            f"/w:document/w:body/*[{table_index + 1}]"
            f"/w:tr[{row_index + 1}]/w:tc[{cell_index + 1}]"
        ),
        fingerprint=fingerprint_elements([cell]),
    )


def native_ref_for_table_cell_paragraph(
    table: ET.Element,
    table_index: int,
    row_index: int,
    cell_index: int,
    paragraph_index: int,
) -> NativeRef:
    rows = table.findall(_q("tr"))
    if row_index >= len(rows):
        raise ValueError("Table-cell paragraph row points outside the table.")
    cells = rows[row_index].findall(_q("tc"))
    if cell_index >= len(cells):
        raise ValueError("Table-cell paragraph cell points outside the row.")
    paragraphs = cells[cell_index].findall(_q("p"))
    if paragraph_index >= len(paragraphs):
        raise ValueError("Table-cell paragraph points outside the cell.")
    paragraph = paragraphs[paragraph_index]
    flattened = _flatten_table_cell_paragraphs(table)
    return NativeRef(
        format="docx",
        part_uri="/word/document.xml",
        native_kind="w:tc/w:p",
        element_index=table_index,
        element_indices=[table_index],
        sub_index=flattened.index(paragraph),
        path_hint=(
            f"/w:document/w:body/*[{table_index + 1}]"
            f"/w:tr[{row_index + 1}]/w:tc[{cell_index + 1}]"
            f"/w:p[{paragraph_index + 1}]"
        ),
        native_id=paragraph.get(
            "{http://schemas.microsoft.com/office/word/2010/wordml}paraId"
        ),
        fingerprint=fingerprint_elements([paragraph]),
    )


def table_cell_from_ref(
    table: ET.Element,
    source_ref: NativeRef,
) -> ET.Element:
    if source_ref.native_kind != "w:tc" or source_ref.sub_index is None:
        raise ValueError("Native reference does not identify a table cell.")
    cells = _flatten_table_cells(table)
    if source_ref.sub_index >= len(cells):
        raise ValueError("Table cell source reference points outside the table.")
    return cells[source_ref.sub_index]


def table_cell_paragraph_from_ref(
    table: ET.Element,
    source_ref: NativeRef,
) -> tuple[ET.Element, ET.Element]:
    if (
        source_ref.native_kind != "w:tc/w:p"
        or source_ref.sub_index is None
    ):
        raise ValueError(
            "Native reference does not identify a table-cell paragraph."
        )
    paragraphs = _flatten_table_cell_paragraphs(table)
    if source_ref.sub_index >= len(paragraphs):
        raise ValueError(
            "Table-cell paragraph source reference points outside the table."
        )
    paragraph = paragraphs[source_ref.sub_index]
    for cell in _flatten_table_cells(table):
        if paragraph in cell.findall(_q("p")):
            return cell, paragraph
    raise ValueError("Could not find the table-cell paragraph parent.")


__all__ = [
    "NativeTableCellPlacement",
    "TableGridAnalysis",
    "analyze_table_grid",
    "apply_table_cell_format",
    "apply_table_layout",
    "apply_table_row",
    "is_regular_table_grid",
    "native_ref_for_table_cell",
    "native_ref_for_table_cell_paragraph",
    "native_ref_for_table_column",
    "native_ref_for_table_row",
    "patch_table_cell_format",
    "patch_table_column_width",
    "patch_table_layout",
    "read_table_cell_format",
    "read_table_column_widths",
    "read_table_layout",
    "read_table_row",
    "table_cell_from_ref",
    "table_cell_paragraph_from_ref",
]
