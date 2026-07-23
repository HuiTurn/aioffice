from __future__ import annotations

import copy
import io
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from aioffice import Document
from aioffice.core.errors import SpecValidationError
from aioffice.native.xml import parse_xml, serialize_xml

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
FUTURE = "urn:aioffice:test:border"


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


def _border_document() -> Document:
    return Document.from_spec(
        {
            "content": [
                {
                    "id": "border_table",
                    "type": "table",
                    "columns": [
                        {
                            "id": "column_left",
                            "key": "left",
                            "title": "Left",
                        },
                        {
                            "id": "column_right",
                            "key": "right",
                            "title": "Right",
                        },
                    ],
                    "layout": {
                        "borders": {
                            "top": {
                                "style": "single",
                                "width": {"value": 1.5, "unit": "pt"},
                                "color": "#1F4E78",
                                "space": {"value": 1, "unit": "pt"},
                            },
                            "right": {
                                "style": "single",
                                "width": {"value": 1, "unit": "pt"},
                                "color": "#1F4E78",
                            },
                            "bottom": {
                                "style": "double",
                                "width": {"value": 2, "unit": "pt"},
                                "color": "#1F4E78",
                            },
                            "left": {
                                "style": "single",
                                "width": {"value": 1, "unit": "pt"},
                                "color": "#1F4E78",
                            },
                            "inside_horizontal": {
                                "style": "dotted",
                                "width": {"value": 0.5, "unit": "pt"},
                                "color": "#A6A6A6",
                            },
                            "inside_vertical": {"style": "none"},
                        }
                    },
                    "rows": [
                        {
                            "id": "border_row",
                            "cells": [
                                {
                                    "id": "border_cell",
                                    "column_key": "left",
                                    "value": "Emphasis",
                                    "format": {
                                        "borders": {
                                            "bottom": {
                                                "style": "double",
                                                "width": {
                                                    "value": 3,
                                                    "unit": "pt",
                                                },
                                                "color": "#C00000",
                                            },
                                            "left": {"style": "none"},
                                        }
                                    },
                                },
                                {
                                    "id": "plain_cell",
                                    "column_key": "right",
                                    "value": "Plain",
                                },
                            ],
                        }
                    ],
                }
            ]
        }
    )


class TableBorderTests(unittest.TestCase):
    def test_border_models_generation_projection_and_html(self) -> None:
        with self.assertRaises(SpecValidationError):
            Document.from_spec(
                {
                    "content": [
                        {
                            "type": "table",
                            "columns": [
                                {"key": "a", "title": "A"},
                            ],
                            "layout": {
                                "borders": {
                                    "top": {"style": "single"}
                                }
                            },
                            "rows": [{"values": {"a": "x"}}],
                        }
                    ]
                }
            )
        with self.assertRaises(SpecValidationError):
            Document.from_spec(
                {
                    "content": [
                        {
                            "type": "table",
                            "columns": [
                                {"key": "a", "title": "A"},
                            ],
                            "layout": {
                                "borders": {
                                    "top": {
                                        "style": "none",
                                        "width": {
                                            "value": 1,
                                            "unit": "pt",
                                        },
                                    }
                                }
                            },
                            "rows": [{"values": {"a": "x"}}],
                        }
                    ]
                }
            )

        document = _border_document()
        source = document.to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
        table = root.find(f".//{_q('tbl')}")
        assert table is not None
        table_borders = table.find(
            f"./{_q('tblPr')}/{_q('tblBorders')}"
        )
        assert table_borders is not None
        self.assertEqual(
            table_borders.find(_q("top")).get(_q("sz")),
            "12",
        )
        self.assertEqual(
            table_borders.find(_q("top")).get(_q("space")),
            "1",
        )
        self.assertEqual(
            table_borders.find(_q("bottom")).get(_q("val")),
            "double",
        )
        self.assertEqual(
            table_borders.find(_q("insideH")).get(_q("val")),
            "dotted",
        )
        self.assertEqual(
            table_borders.find(_q("insideV")).get(_q("val")),
            "none",
        )
        data_cell = table.findall(_q("tr"))[1].findall(_q("tc"))[0]
        cell_borders = data_cell.find(
            f"./{_q('tcPr')}/{_q('tcBorders')}"
        )
        assert cell_borders is not None
        self.assertEqual(
            cell_borders.find(_q("bottom")).get(_q("sz")),
            "24",
        )
        self.assertEqual(
            cell_borders.find(_q("left")).get(_q("val")),
            "none",
        )

        reopened = Document.from_docx(source)
        table_spec = next(
            node
            for node in reopened.spec.content
            if node.type == "table"
        )
        assert table_spec.layout.borders is not None
        assert table_spec.layout.borders.top is not None
        self.assertEqual(
            table_spec.layout.borders.top.width.to_points(),
            1.5,
        )
        self.assertEqual(
            table_spec.layout.borders.inside_vertical.style,
            "none",
        )
        assert table_spec.rows[0].cells[0].format.borders is not None
        self.assertEqual(
            table_spec.rows[0].cells[0].format.borders.bottom.style,
            "double",
        )
        self.assertEqual(reopened.to_bytes("docx"), source)

        html = document.to_bytes("html").decode()
        self.assertIn(
            "border-top:1.5pt solid #1F4E78",
            html,
        )
        self.assertIn("border:none", html)
        self.assertIn(
            "border-bottom:2pt double #1F4E78",
            html,
        )
        self.assertIn(
            "border-bottom:3pt double #C00000",
            html,
        )
        self.assertIn("border-left:none", html)

    def test_native_border_patch_preserves_unknown_xml(self) -> None:
        source = _border_document().to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
            styles = package.read("word/styles.xml")
        table = root.find(f".//{_q('tbl')}")
        assert table is not None
        top = table.find(
            f"./{_q('tblPr')}/{_q('tblBorders')}/{_q('top')}"
        )
        assert top is not None
        top.set(f"{{{FUTURE}}}attribute", "keep")
        ET.SubElement(top, f"{{{FUTURE}}}edgeData").text = "keep"
        data_cell = table.findall(_q("tr"))[1].findall(_q("tc"))[0]
        bottom = data_cell.find(
            f"./{_q('tcPr')}/{_q('tcBorders')}/{_q('bottom')}"
        )
        assert bottom is not None
        bottom.set(f"{{{FUTURE}}}attribute", "keep")
        ET.SubElement(
            bottom,
            f"{{{FUTURE}}}cellEdgeData",
        ).text = "keep"
        source = _rewrite_document_xml(source, root)

        imported = Document.from_docx(source)
        imported_table = next(
            node
            for node in imported.spec.content
            if node.type == "table"
        )
        result = imported.apply(
            [
                {
                    "op": "table.format",
                    "target": f"#{imported_table.id}",
                    "set": {
                        "borders": {
                            "top": {
                                "style": "thick",
                                "width": {
                                    "value": 4,
                                    "unit": "pt",
                                },
                                "color": "#548235",
                            }
                        }
                    },
                },
                {
                    "op": "table.cell.format",
                    "target": f"#{imported_table.id}",
                    "cell": "#border_cell",
                    "set": {
                        "borders": {
                            "bottom": {
                                "style": "dashed",
                                "width": {
                                    "value": 1,
                                    "unit": "pt",
                                },
                                "color": "#7030A0",
                            }
                        }
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
            self.assertEqual(package.read("word/styles.xml"), styles)
            patched_root = parse_xml(
                package.read("word/document.xml")
            )
        patched_table = patched_root.find(f".//{_q('tbl')}")
        assert patched_table is not None
        patched_top = patched_table.find(
            f"./{_q('tblPr')}/{_q('tblBorders')}/{_q('top')}"
        )
        assert patched_top is not None
        self.assertEqual(patched_top.get(_q("val")), "thick")
        self.assertEqual(patched_top.get(_q("sz")), "32")
        self.assertEqual(patched_top.get(_q("color")), "548235")
        self.assertEqual(
            patched_top.get(f"{{{FUTURE}}}attribute"),
            "keep",
        )
        self.assertIsNotNone(
            patched_top.find(f"{{{FUTURE}}}edgeData")
        )
        patched_cell = patched_table.findall(_q("tr"))[1].findall(
            _q("tc")
        )[0]
        patched_bottom = patched_cell.find(
            f"./{_q('tcPr')}/{_q('tcBorders')}/{_q('bottom')}"
        )
        assert patched_bottom is not None
        self.assertEqual(patched_bottom.get(_q("val")), "dashed")
        self.assertEqual(patched_bottom.get(_q("sz")), "8")
        self.assertEqual(patched_bottom.get(_q("color")), "7030A0")
        self.assertEqual(
            patched_bottom.get(f"{{{FUTURE}}}attribute"),
            "keep",
        )
        self.assertIsNotNone(
            patched_bottom.find(f"{{{FUTURE}}}cellEdgeData")
        )

        cleared = result.document.apply(
            [
                {
                    "op": "table.format",
                    "target": f"#{imported_table.id}",
                    "clear": ["borders"],
                },
                {
                    "op": "table.cell.format",
                    "target": f"#{imported_table.id}",
                    "cell": "#border_cell",
                    "clear": ["borders"],
                },
            ]
        )
        self.assertTrue(cleared.success, cleared.model_dump())
        assert cleared.document is not None
        cleared_output = cleared.document.to_bytes("docx")
        with ZipFile(io.BytesIO(cleared_output)) as package:
            cleared_root = parse_xml(
                package.read("word/document.xml")
            )
        cleared_table = cleared_root.find(f".//{_q('tbl')}")
        assert cleared_table is not None
        cleared_top = cleared_table.find(
            f"./{_q('tblPr')}/{_q('tblBorders')}/{_q('top')}"
        )
        assert cleared_top is not None
        self.assertIsNone(cleared_top.get(_q("val")))
        self.assertIsNone(cleared_top.get(_q("sz")))
        self.assertEqual(
            cleared_top.get(f"{{{FUTURE}}}attribute"),
            "keep",
        )
        self.assertIsNotNone(
            cleared_top.find(f"{{{FUTURE}}}edgeData")
        )
        cleared_cell = cleared_table.findall(_q("tr"))[1].findall(
            _q("tc")
        )[0]
        cleared_bottom = cleared_cell.find(
            f"./{_q('tcPr')}/{_q('tcBorders')}/{_q('bottom')}"
        )
        assert cleared_bottom is not None
        self.assertIsNone(cleared_bottom.get(_q("val")))
        self.assertIsNone(cleared_bottom.get(_q("sz")))
        self.assertEqual(
            cleared_bottom.get(f"{{{FUTURE}}}attribute"),
            "keep",
        )
        self.assertIsNotNone(
            cleared_bottom.find(f"{{{FUTURE}}}cellEdgeData")
        )
        cleared_projection = Document.from_docx(cleared_output)
        cleared_table_spec = next(
            node
            for node in cleared_projection.spec.content
            if node.type == "table"
        )
        self.assertIsNone(cleared_table_spec.layout.borders)
        self.assertIsNone(
            cleared_table_spec.rows[0].cells[0].format.borders
        )


if __name__ == "__main__":
    unittest.main()
