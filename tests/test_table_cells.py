from __future__ import annotations

import copy
import io
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from aioffice import Document, DocumentBuilder
from aioffice.native.xml import parse_xml, serialize_xml

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
FUTURE = "urn:aioffice:test:cell"


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


def _merged_rich_document() -> Document:
    return Document.from_spec(
        {
            "content": [
                {
                    "id": "before",
                    "type": "paragraph",
                    "text": "Before",
                },
                {
                    "id": "merged_table",
                    "type": "table",
                    "columns": [
                        {
                            "id": "left_column",
                            "key": "left",
                            "title": "Left",
                            "width": {"value": 110, "unit": "pt"},
                        },
                        {
                            "id": "middle_column",
                            "key": "middle",
                            "title": "Middle",
                            "width": {"value": 130, "unit": "pt"},
                        },
                        {
                            "id": "right_column",
                            "key": "right",
                            "title": "Right",
                            "width": {"value": 100, "unit": "pt"},
                        },
                    ],
                    "rows": [
                        {
                            "id": "first_row",
                            "cells": [
                                {
                                    "id": "merged_cell",
                                    "column_key": "left",
                                    "column_span": 2,
                                    "row_span": 2,
                                    "format": {
                                        "vertical_alignment": "center",
                                        "no_wrap": True,
                                        "background_color": "#EAF2F8",
                                        "margin_left": {
                                            "value": 8,
                                            "unit": "pt",
                                        },
                                    },
                                    "content": [
                                        {
                                            "id": "merged_title",
                                            "type": "paragraph",
                                            "content": [
                                                {
                                                    "type": "text",
                                                    "text": "Merged",
                                                    "marks": ["strong"],
                                                },
                                                {
                                                    "type": "text",
                                                    "text": " summary",
                                                },
                                            ],
                                            "paragraph_style": {
                                                "alignment": "center"
                                            },
                                        },
                                        {
                                            "id": "merged_note",
                                            "type": "paragraph",
                                            "text": "Second line",
                                        },
                                    ],
                                },
                                {
                                    "id": "right_first",
                                    "column_key": "right",
                                    "value": 100,
                                },
                            ],
                        },
                        {
                            "id": "second_row",
                            "cells": [
                                {
                                    "id": "right_second",
                                    "column_key": "right",
                                    "value": 200,
                                }
                            ],
                        },
                        {
                            "id": "third_row",
                            "cells": [
                                {
                                    "id": "left_third",
                                    "column_key": "left",
                                    "value": "A",
                                },
                                {
                                    "id": "middle_third",
                                    "column_key": "middle",
                                    "value": "B",
                                },
                                {
                                    "id": "right_third",
                                    "column_key": "right",
                                    "value": "C",
                                },
                            ],
                        },
                    ],
                    "layout": {
                        "algorithm": "fixed",
                        "repeat_header": True,
                    },
                },
                {
                    "id": "after",
                    "type": "paragraph",
                    "text": "After",
                },
            ]
        }
    )


class TableCellTests(unittest.TestCase):
    def test_legacy_values_migrate_to_deterministic_cells(self) -> None:
        payload = {
            "artifact": {
                "id": "deterministic_doc",
                "kind": "document",
                "revision": 1,
            },
            "content": [
                {
                    "id": "deterministic_table",
                    "type": "table",
                    "columns": [
                        {
                            "id": "deterministic_column",
                            "key": "value",
                            "title": "Value",
                        }
                    ],
                    "rows": [
                        {
                            "id": "deterministic_row",
                            "values": {"value": "stable"},
                        }
                    ],
                }
            ],
        }
        first = Document.from_spec(payload)
        second = Document.from_spec(payload)
        first_row = first.to_spec()["content"][0]["rows"][0]
        second_row = second.to_spec()["content"][0]["rows"][0]
        self.assertNotIn("values", first_row)
        self.assertEqual(first_row["cells"], second_row["cells"])
        self.assertEqual(
            first_row["cells"][0]["id"],
            "cell_f118fff3b9b9aa4843efba66",
        )
        self.assertEqual(first.to_bytes("docx"), second.to_bytes("docx"))

    def test_merged_rich_generation_projection_identity_and_html(
        self,
    ) -> None:
        document = _merged_rich_document()
        self.assertTrue(
            document.validate().valid,
            document.validate().model_dump(),
        )
        source = document.to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
            manifest = package.read(
                "customXml/aioffice-manifest.xml"
            ).decode()
        table = root.find(f".//{_q('tbl')}")
        assert table is not None
        rows = table.findall(_q("tr"))
        self.assertEqual(len(rows), 4)
        merged = rows[1].findall(_q("tc"))[0]
        merged_properties = merged.find(_q("tcPr"))
        assert merged_properties is not None
        self.assertEqual(
            merged_properties.find(_q("gridSpan")).get(_q("val")),
            "2",
        )
        self.assertEqual(
            merged_properties.find(_q("vMerge")).get(_q("val")),
            "restart",
        )
        self.assertEqual(
            merged_properties.find(_q("vAlign")).get(_q("val")),
            "center",
        )
        self.assertEqual(
            merged_properties.find(_q("noWrap")).get(_q("val")),
            "1",
        )
        self.assertEqual(
            merged_properties.find(_q("shd")).get(_q("fill")),
            "EAF2F8",
        )
        continuation = rows[2].findall(_q("tc"))[0]
        continuation_properties = continuation.find(_q("tcPr"))
        assert continuation_properties is not None
        self.assertEqual(
            continuation_properties.find(_q("gridSpan")).get(_q("val")),
            "2",
        )
        self.assertEqual(
            continuation_properties.find(_q("vMerge")).get(_q("val")),
            "continue",
        )
        self.assertIn('id="merged_cell"', manifest)
        self.assertIn('id="merged_title"', manifest)
        self.assertIn('semanticKey="left"', manifest)

        reopened = Document.from_docx(source)
        self.assertFalse(
            reopened.import_diagnostics,
            reopened.import_diagnostics,
        )
        reopened_table = reopened.to_spec()["content"][1]
        self.assertTrue(reopened_table["metadata"]["logical_grid"])
        self.assertFalse(reopened_table["metadata"]["regular_grid"])
        projected_cell = reopened_table["rows"][0]["cells"][0]
        self.assertEqual(projected_cell["id"], "merged_cell")
        self.assertEqual(projected_cell["column_span"], 2)
        self.assertEqual(projected_cell["row_span"], 2)
        self.assertEqual(
            [
                paragraph["id"]
                for paragraph in projected_cell["content"]
            ],
            ["merged_title", "merged_note"],
        )
        self.assertEqual(reopened.to_bytes("docx"), source)

        html = reopened.to_bytes("html").decode()
        self.assertIn('data-cell-id="merged_cell"', html)
        self.assertIn('colspan="2"', html)
        self.assertIn('rowspan="2"', html)
        self.assertIn("background-color:#EAF2F8", html)
        self.assertIn('id="merged_title"', html)
        self.assertIn("font-weight:700", html)
        self.assertIn(">Merged</span>", html)

    def test_native_cell_and_rich_paragraph_patch_preserves_unknown_xml(
        self,
    ) -> None:
        source = _merged_rich_document().to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
            original_styles = package.read("word/styles.xml")
        table = root.find(f".//{_q('tbl')}")
        assert table is not None
        target_cell = table.findall(_q("tr"))[1].findall(_q("tc"))[0]
        target_properties = target_cell.find(_q("tcPr"))
        assert target_properties is not None
        target_properties.set(f"{{{FUTURE}}}attribute", "preserve")
        ET.SubElement(
            target_properties,
            f"{{{FUTURE}}}futureCellProperty",
        ).text = "keep"
        source = _rewrite_document_xml(source, root)

        imported = Document.from_docx(source)
        self.assertFalse(
            imported.import_diagnostics,
            imported.import_diagnostics,
        )
        result = imported.apply(
            [
                {
                    "op": "text.replace",
                    "target": "#merged_title",
                    "search": "Merged",
                    "replacement": "Combined",
                },
                {
                    "op": "table.cell.format",
                    "target": "#merged_table",
                    "cell": "#merged_cell",
                    "set": {
                        "vertical_alignment": "bottom",
                        "background_color": "#FFF2CC",
                        "margin_right": {
                            "value": 9,
                            "unit": "pt",
                        },
                    },
                    "clear": ["no_wrap"],
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
            self.assertEqual(
                package.read("word/styles.xml"),
                original_styles,
            )
            patched_root = parse_xml(
                package.read("word/document.xml")
            )
        patched_table = patched_root.find(f".//{_q('tbl')}")
        assert patched_table is not None
        patched_cell = patched_table.findall(_q("tr"))[1].findall(
            _q("tc")
        )[0]
        patched_properties = patched_cell.find(_q("tcPr"))
        assert patched_properties is not None
        self.assertEqual(
            patched_properties.get(f"{{{FUTURE}}}attribute"),
            "preserve",
        )
        self.assertIsNotNone(
            patched_properties.find(
                f"{{{FUTURE}}}futureCellProperty"
            )
        )
        self.assertEqual(
            patched_properties.find(_q("vAlign")).get(_q("val")),
            "bottom",
        )
        self.assertEqual(
            patched_properties.find(_q("shd")).get(_q("fill")),
            "FFF2CC",
        )
        self.assertIsNone(patched_properties.find(_q("noWrap")))
        self.assertEqual(
            patched_properties.find(
                f"./{_q('tcMar')}/{_q('right')}"
            ).get(_q("w")),
            "180",
        )
        self.assertIn(
            "Combined summary",
            "".join(
                text.text or ""
                for text in patched_cell.iter(_q("t"))
            ),
        )
        reopened = Document.from_docx(output)
        self.assertFalse(
            reopened.import_diagnostics,
            reopened.import_diagnostics,
        )
        title = (
            reopened.to_spec()["content"][1]["rows"][0]["cells"][0][
                "content"
            ][0]
        )
        self.assertEqual(
            "".join(
                span["text"]
                for span in title["content"]
                if span["type"] == "text"
            ),
            "Combined summary",
        )

    def test_grid_validation_and_opaque_cell_boundary(self) -> None:
        invalid = Document.from_spec(
            {
                "content": [
                    {
                        "id": "invalid_table",
                        "type": "table",
                        "columns": [
                            {"key": "a", "title": "A"},
                            {"key": "b", "title": "B"},
                        ],
                        "rows": [
                            {
                                "cells": [
                                    {
                                        "id": "overlap_one",
                                        "column_key": "a",
                                        "column_span": 2,
                                        "value": "wide",
                                    },
                                    {
                                        "id": "overlap_two",
                                        "column_key": "b",
                                        "value": "overlap",
                                    },
                                ]
                            }
                        ],
                    }
                ]
            }
        )
        codes = {
            diagnostic.code
            for diagnostic in invalid.validate().errors
        }
        self.assertIn("TABLE_CELL_OVERLAP", codes)

        incomplete = Document.from_spec(
            {
                "content": [
                    {
                        "id": "incomplete_table",
                        "type": "table",
                        "columns": [
                            {"key": "a", "title": "A"},
                            {"key": "b", "title": "B"},
                        ],
                        "rows": [
                            {
                                "cells": [
                                    {
                                        "column_key": "a",
                                        "value": "only one",
                                    }
                                ]
                            }
                        ],
                    }
                ]
            }
        )
        self.assertIn(
            "TABLE_GRID_INCOMPLETE",
            {
                diagnostic.code
                for diagnostic in incomplete.validate().errors
            },
        )

        source = (
            DocumentBuilder()
            .table(
                id="opaque_table",
                columns=[
                    {
                        "id": "opaque_column",
                        "key": "value",
                        "title": "Value",
                    }
                ],
                rows=[
                    {
                        "id": "opaque_row",
                        "values": {"value": "Visible"},
                    }
                ],
            )
            .build()
            .to_bytes("docx")
        )
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
        table = root.find(f".//{_q('tbl')}")
        assert table is not None
        data_cell = table.findall(_q("tr"))[1].find(_q("tc"))
        assert data_cell is not None
        run = data_cell.find(f"./{_q('p')}/{_q('r')}")
        assert run is not None
        ET.SubElement(run, _q("drawing"))
        source = _rewrite_document_xml(source, root)

        imported = Document.from_docx(source)
        imported_table = imported.to_spec()["content"][0]
        cell = imported_table["rows"][0]["cells"][0]
        self.assertFalse(cell["metadata"]["content_editable"])
        self.assertEqual(
            cell["metadata"]["content_projection"],
            "plain_text_read_only",
        )
        self.assertEqual(cell["content"], [])
        formatted = imported.apply(
            [
                {
                    "op": "table.cell.format",
                    "target": f"#{imported_table['id']}",
                    "cell": f"#{cell['id']}",
                    "set": {"vertical_alignment": "center"},
                }
            ]
        )
        self.assertTrue(formatted.success, formatted.model_dump())
        assert formatted.document is not None
        with ZipFile(
            io.BytesIO(formatted.document.to_bytes("docx"))
        ) as package:
            patched = parse_xml(package.read("word/document.xml"))
        self.assertIsNotNone(patched.find(f".//{_q('drawing')}"))


if __name__ == "__main__":
    unittest.main()
