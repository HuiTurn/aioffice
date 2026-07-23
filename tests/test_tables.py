from __future__ import annotations

import copy
import io
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from aioffice import Document, DocumentBuilder
from aioffice.native.xml import parse_xml, serialize_xml

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
FUTURE = "urn:aioffice:test:table"


def _q(local: str) -> str:
    return f"{{{W}}}{local}"


def _rewrite_document_xml(source: bytes, root: ET.Element) -> bytes:
    output = io.BytesIO()
    with (
        ZipFile(io.BytesIO(source)) as before,
        ZipFile(output, "w", compression=ZIP_DEFLATED) as after,
    ):
        for info in before.infolist():
            after.writestr(
                copy.copy(info),
                (
                    serialize_xml(root)
                    if info.filename == "word/document.xml"
                    else before.read(info.filename)
                ),
            )
    return output.getvalue()


def _geometry_document() -> Document:
    return DocumentBuilder().table(
        columns=[
            {
                "id": "metric_name",
                "key": "name",
                "title": "Metric",
                "data_type": "text",
                "width": {"value": 120, "unit": "pt"},
            },
            {
                "id": "metric_value",
                "key": "value",
                "title": "Value",
                "data_type": "number",
                "width": {"value": 180, "unit": "pt"},
            },
        ],
        rows=[
            {
                "id": "revenue_row",
                "values": {"name": "Revenue", "value": 42},
                "allow_break_across_pages": False,
                "height": {"value": 24, "unit": "pt"},
                "height_rule": "exact",
            },
            {
                "id": "margin_row",
                "values": {"name": "Margin", "value": 18.5},
            },
        ],
        id="metrics_table",
    ).build()


class TableGeometryTests(unittest.TestCase):
    def test_generation_projection_identity_inspect_and_html(self) -> None:
        document = Document.from_spec(
            {
                **_geometry_document().to_spec(),
                "content": [
                    {
                        **_geometry_document().to_spec()["content"][0],
                        "layout": {
                            "style_ref": "TableGrid",
                            "preferred_width": {
                                "mode": "percent",
                                "value": 90,
                            },
                            "alignment": "center",
                            "algorithm": "fixed",
                            "indent": {"value": 12, "unit": "pt"},
                            "cell_spacing": {"value": 2, "unit": "pt"},
                            "cell_margin_top": {
                                "value": 3,
                                "unit": "pt",
                            },
                            "cell_margin_right": {
                                "value": 5,
                                "unit": "pt",
                            },
                            "cell_margin_bottom": {
                                "value": 4,
                                "unit": "pt",
                            },
                            "cell_margin_left": {
                                "value": 6,
                                "unit": "pt",
                            },
                            "repeat_header": False,
                        },
                    }
                ],
            }
        )
        self.assertTrue(document.validate().valid, document.validate().diagnostics)

        source = document.to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
            table = root.find(f".//{_q('tbl')}")
            assert table is not None
            properties = table.find(_q("tblPr"))
            assert properties is not None
            width = properties.find(_q("tblW"))
            assert width is not None
            self.assertEqual(width.attrib[_q("type")], "pct")
            self.assertEqual(width.attrib[_q("w")], "4500")
            self.assertEqual(
                properties.find(_q("jc")).attrib[_q("val")],
                "center",
            )
            self.assertEqual(
                properties.find(_q("tblLayout")).attrib[_q("type")],
                "fixed",
            )
            self.assertEqual(
                properties.find(_q("tblInd")).attrib[_q("w")],
                "240",
            )
            self.assertEqual(
                properties.find(_q("tblCellSpacing")).attrib[_q("w")],
                "40",
            )
            margins = properties.find(_q("tblCellMar"))
            assert margins is not None
            self.assertEqual(
                [
                    (child.tag, child.attrib[_q("w")])
                    for child in list(margins)
                ],
                [
                    (_q("top"), "60"),
                    (_q("left"), "120"),
                    (_q("bottom"), "80"),
                    (_q("right"), "100"),
                ],
            )
            self.assertEqual(
                [
                    column.attrib[_q("w")]
                    for column in table.findall(
                        f"./{_q('tblGrid')}/{_q('gridCol')}"
                    )
                ],
                ["2400", "3600"],
            )
            header = table.find(_q("tr"))
            assert header is not None
            self.assertEqual(
                header.find(
                    f"./{_q('trPr')}/{_q('tblHeader')}"
                ).attrib[_q("val")],
                "0",
            )
            first_data_row = table.findall(_q("tr"))[1]
            self.assertEqual(
                first_data_row.find(
                    f"./{_q('trPr')}/{_q('cantSplit')}"
                ).attrib[_q("val")],
                "1",
            )
            row_height = first_data_row.find(
                f"./{_q('trPr')}/{_q('trHeight')}"
            )
            assert row_height is not None
            self.assertEqual(row_height.attrib[_q("val")], "480")
            self.assertEqual(row_height.attrib[_q("hRule")], "exact")

            manifest = package.read(
                "customXml/aioffice-manifest.xml"
            ).decode()
            self.assertIn('id="metric_name"', manifest)
            self.assertIn('semanticKey="name"', manifest)
            self.assertIn('semanticDataType="number"', manifest)
            self.assertIn('id="revenue_row"', manifest)

        reopened = Document.from_docx(source)
        self.assertFalse(reopened.import_diagnostics)
        table_spec = reopened.to_spec()["content"][0]
        self.assertEqual(table_spec["id"], "metrics_table")
        self.assertEqual(
            [
                (
                    column["id"],
                    column["key"],
                    column["data_type"],
                    column["source_ref"]["sub_index"],
                )
                for column in table_spec["columns"]
            ],
            [
                ("metric_name", "name", "text", 0),
                ("metric_value", "value", "number", 1),
            ],
        )
        self.assertEqual(
            [
                (
                    row["id"],
                    {
                        cell["column_key"]: (
                            cell["content"][0].get("text", "")
                            if cell.get("content")
                            else cell.get("value", "")
                        )
                        for cell in row["cells"]
                    },
                    row["source_ref"]["sub_index"],
                )
                for row in table_spec["rows"]
            ],
            [
                (
                    "revenue_row",
                    {"name": "Revenue", "value": "42"},
                    1,
                ),
                (
                    "margin_row",
                    {"name": "Margin", "value": "18.5"},
                    2,
                ),
            ],
        )
        self.assertEqual(
            table_spec["layout"]["preferred_width"],
            {"mode": "percent", "value": 90.0},
        )
        self.assertFalse(table_spec["layout"]["repeat_header"])
        self.assertEqual(reopened.to_bytes("docx"), source)

        inspection = reopened.inspect()
        compact = inspection["nodes"][0]
        self.assertTrue(compact["regular_grid"])
        self.assertEqual(
            [column["id"] for column in compact["columns"]],
            ["metric_name", "metric_value"],
        )
        html = reopened.to_bytes("html").decode()
        self.assertIn('data-column-id="metric_name"', html)
        self.assertIn('data-column-key="value"', html)
        self.assertIn("width:90%", html)
        self.assertIn("table-layout:fixed", html)
        self.assertIn("padding-left:6pt", html)
        self.assertIn("break-inside:avoid", html)

    def test_native_table_and_column_patch_preserve_unknown_xml(self) -> None:
        source_document = (
            DocumentBuilder()
            .paragraph("Before", id="before")
            .table(
                columns=[
                    {"id": "name_col", "key": "name", "title": "Name"},
                    {"id": "value_col", "key": "value", "title": "Value"},
                ],
                rows=[
                    {
                        "id": "first_row",
                        "values": {"name": "One", "value": "1"},
                    },
                    {
                        "id": "second_row",
                        "values": {"name": "Two", "value": "2"},
                    },
                ],
                id="native_table",
            )
            .paragraph("After", id="after")
            .build()
        )
        source = source_document.to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
            original_styles = package.read("word/styles.xml")
        table = root.find(f".//{_q('tbl')}")
        assert table is not None
        properties = table.find(_q("tblPr"))
        assert properties is not None
        properties.set(f"{{{FUTURE}}}tableProperty", "preserve")
        ET.SubElement(properties, f"{{{FUTURE}}}futureTable").text = "keep"
        target_grid = table.findall(
            f"./{_q('tblGrid')}/{_q('gridCol')}"
        )[1]
        target_grid.set(f"{{{FUTURE}}}gridProperty", "preserve")
        first_data_target_cell = table.findall(_q("tr"))[1].findall(
            _q("tc")
        )[1]
        cell_width = first_data_target_cell.find(
            f"./{_q('tcPr')}/{_q('tcW')}"
        )
        assert cell_width is not None
        cell_width.set(f"{{{FUTURE}}}cellWidthProperty", "preserve")
        source = _rewrite_document_xml(source, root)

        imported = Document.from_docx(source)
        self.assertFalse(imported.import_diagnostics)
        result = imported.apply(
            [
                {
                    "op": "table.format",
                    "target": "#native_table",
                    "set": {
                        "alignment": "right",
                        "algorithm": "fixed",
                        "preferred_width": {
                            "mode": "exact",
                            "value": {"value": 5.5, "unit": "in"},
                        },
                        "cell_margin_left": {
                            "value": 7,
                            "unit": "pt",
                        },
                    },
                },
                {
                    "op": "table.column.format",
                    "target": "#native_table",
                    "column": "value",
                    "set": {
                        "width": {"value": 144, "unit": "pt"}
                    },
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        assert result.fidelity is not None
        self.assertEqual(
            result.fidelity.affected_parts,
            [
                "/customXml/aioffice-manifest.xml",
                "/word/document.xml",
            ],
        )
        output = result.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as package:
            self.assertEqual(package.read("word/styles.xml"), original_styles)
            patched_root = parse_xml(package.read("word/document.xml"))
        patched = patched_root.find(f".//{_q('tbl')}")
        assert patched is not None
        patched_properties = patched.find(_q("tblPr"))
        assert patched_properties is not None
        self.assertEqual(
            patched_properties.attrib[
                f"{{{FUTURE}}}tableProperty"
            ],
            "preserve",
        )
        self.assertIsNotNone(
            patched_properties.find(f"{{{FUTURE}}}futureTable")
        )
        self.assertEqual(
            patched_properties.find(_q("jc")).attrib[_q("val")],
            "right",
        )
        self.assertEqual(
            patched_properties.find(_q("tblW")).attrib[_q("w")],
            "7920",
        )
        patched_grid = patched.findall(
            f"./{_q('tblGrid')}/{_q('gridCol')}"
        )[1]
        self.assertEqual(patched_grid.attrib[_q("w")], "2880")
        self.assertEqual(
            patched_grid.attrib[f"{{{FUTURE}}}gridProperty"],
            "preserve",
        )
        patched_cell_width = patched.findall(_q("tr"))[1].findall(
            _q("tc")
        )[1].find(f"./{_q('tcPr')}/{_q('tcW')}")
        assert patched_cell_width is not None
        self.assertEqual(patched_cell_width.attrib[_q("w")], "2880")
        self.assertEqual(
            patched_cell_width.attrib[
                f"{{{FUTURE}}}cellWidthProperty"
            ],
            "preserve",
        )

        immediate = result.document.to_spec()["content"][1]
        value_column = next(
            column
            for column in immediate["columns"]
            if column["id"] == "value_col"
        )
        self.assertEqual(
            value_column["width"],
            {"value": 144.0, "unit": "pt"},
        )
        self.assertEqual(
            Document.from_docx(output).to_spec()["content"][1]["layout"][
                "alignment"
            ],
            "right",
        )

    def test_irregular_grid_refuses_column_edit_but_allows_table_format(
        self,
    ) -> None:
        source = (
            DocumentBuilder()
            .paragraph("Before", id="before")
            .table(
                columns=[
                    {"id": "left_col", "key": "left", "title": "Left"},
                    {"id": "right_col", "key": "right", "title": "Right"},
                ],
                rows=[
                    {
                        "id": "data_row",
                        "values": {"left": "A", "right": "B"},
                    }
                ],
                id="merged_table",
            )
            .paragraph("After", id="after")
            .build()
            .to_bytes("docx")
        )
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
        table = root.find(f".//{_q('tbl')}")
        assert table is not None
        header = table.find(_q("tr"))
        assert header is not None
        header_cells = header.findall(_q("tc"))
        first_properties = header_cells[0].find(_q("tcPr"))
        assert first_properties is not None
        ET.SubElement(
            first_properties,
            _q("gridSpan"),
            {_q("val"): "2"},
        )
        header.remove(header_cells[1])
        source = _rewrite_document_xml(source, root)

        imported = Document.from_docx(source)
        table_spec = imported.to_spec()["content"][1]
        self.assertEqual(table_spec["id"], "merged_table")
        self.assertFalse(table_spec["metadata"]["regular_grid"])
        rejected = imported.apply(
            [
                {
                    "op": "table.column.format",
                    "target": "#merged_table",
                    "column": "left",
                    "set": {
                        "width": {"value": 100, "unit": "pt"}
                    },
                }
            ]
        )
        self.assertFalse(rejected.success)
        self.assertEqual(
            rejected.diagnostics[0].code,
            "NATIVE_PATCH_FAILED",
        )
        self.assertEqual(imported.to_bytes("docx"), source)

        formatted = imported.apply(
            [
                {
                    "op": "table.format",
                    "target": "#merged_table",
                    "set": {"alignment": "center"},
                }
            ]
        )
        self.assertTrue(formatted.success, formatted.model_dump())
        assert formatted.document is not None
        with ZipFile(
            io.BytesIO(formatted.document.to_bytes("docx"))
        ) as package:
            patched = parse_xml(package.read("word/document.xml"))
        self.assertIsNotNone(patched.find(f".//{_q('gridSpan')}"))
        self.assertEqual(
            patched.find(f".//{_q('tblPr')}/{_q('jc')}").attrib[
                _q("val")
            ],
            "center",
        )

    def test_duplicate_component_ids_and_invalid_table_operations(self) -> None:
        duplicate = Document.from_spec(
            {
                "content": [
                    {
                        "id": "bad_table",
                        "type": "table",
                        "columns": [
                            {
                                "id": "duplicate_component",
                                "key": "a",
                                "title": "A",
                            }
                        ],
                        "rows": [
                            {
                                "id": "duplicate_component",
                                "values": {"a": "value"},
                            }
                        ],
                    }
                ]
            }
        )
        self.assertIn(
            "INVALID_SPEC",
            {
                diagnostic.code
                for diagnostic in duplicate.validate().errors
            },
        )

        document = _geometry_document()
        invalid = document.apply(
            [
                {
                    "op": "table.column.format",
                    "target": "#metrics_table",
                    "column": "value",
                    "set": {"height": {"value": 10, "unit": "pt"}},
                }
            ]
        )
        self.assertFalse(invalid.success)
        self.assertEqual(invalid.result_revision, document.revision)
        self.assertEqual(
            invalid.diagnostics[0].code,
            "INVALID_SPEC",
        )

    def test_table_geometry_diagnostics_are_actionable(self) -> None:
        document = DocumentBuilder().table(
            columns=[
                {
                    "id": "wide_left",
                    "key": "left",
                    "title": "Left",
                    "width": {"value": 300, "unit": "pt"},
                },
                {
                    "id": "wide_right",
                    "key": "right",
                    "title": "Right",
                    "width": {"value": 300, "unit": "pt"},
                },
            ],
            rows=[{"left": "A", "right": "B"}],
            layout={"algorithm": "fixed"},
            id="wide_table",
        ).table(
            columns=[
                {
                    "id": "known_width",
                    "key": "known",
                    "title": "Known",
                    "width": {"value": 100, "unit": "pt"},
                },
                {
                    "id": "missing_width",
                    "key": "missing",
                    "title": "Missing",
                },
            ],
            rows=[{"known": "A", "missing": "B"}],
            layout={"algorithm": "fixed"},
            id="incomplete_table",
        ).build()

        warnings = {
            diagnostic.code: diagnostic
            for diagnostic in document.validate().warnings
        }
        self.assertIn("TABLE_WIDTH_OVERFLOW", warnings)
        self.assertEqual(
            warnings["TABLE_WIDTH_OVERFLOW"].suggested_actions[0][
                "action"
            ],
            "resize_table_columns",
        )
        self.assertIn("TABLE_COLUMN_WIDTH_INCOMPLETE", warnings)
        self.assertIn(
            "missing_width",
            warnings["TABLE_COLUMN_WIDTH_INCOMPLETE"].node_ids,
        )


if __name__ == "__main__":
    unittest.main()
