"""WordprocessingML table geometry, projection, and selective mutation."""

from __future__ import annotations

from typing import Literal, cast
from xml.etree import ElementTree as ET

from aioffice.native.identity import fingerprint_elements
from aioffice.spec.models import (
    Length,
    NativeRef,
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

    grid = table.find(_q("tblGrid"))
    if grid is None:
        return False
    column_count = len(grid.findall(_q("gridCol")))
    if column_count == 0:
        return False
    rows = table.findall(_q("tr"))
    if not rows:
        return False
    for row in rows:
        properties = row.find(_q("trPr"))
        if properties is not None and (
            properties.find(_q("gridBefore")) is not None
            or properties.find(_q("gridAfter")) is not None
        ):
            return False
        cells = row.findall(_q("tc"))
        if len(cells) != column_count:
            return False
        for cell in cells:
            cell_properties = cell.find(_q("tcPr"))
            if cell_properties is not None and any(
                cell_properties.find(_q(name)) is not None
                for name in ("gridSpan", "hMerge", "vMerge")
            ):
                return False
    return True


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


__all__ = [
    "apply_table_layout",
    "apply_table_row",
    "is_regular_table_grid",
    "native_ref_for_table_column",
    "native_ref_for_table_row",
    "patch_table_column_width",
    "patch_table_layout",
    "read_table_column_widths",
    "read_table_layout",
    "read_table_row",
]
