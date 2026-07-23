from __future__ import annotations

import json
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
from pathlib import Path

from aioffice.cli import main
from aioffice.cli.main import _parse_page_numbers
from aioffice.documents import DocumentBuilder


class CliTests(unittest.TestCase):
    def test_page_selection_parser_is_bounded_and_one_based(self) -> None:
        self.assertEqual(
            _parse_page_numbers("1,3-5", max_pages=5),
            [1, 3, 4, 5],
        )
        self.assertIsNone(_parse_page_numbers(None, max_pages=5))
        with self.assertRaises(ValueError):
            _parse_page_numbers("0", max_pages=5)
        with self.assertRaises(ValueError):
            _parse_page_numbers("5-3", max_pages=5)
        with self.assertRaises(ValueError):
            _parse_page_numbers("1-10", max_pages=5)
        with self.assertRaises(ValueError):
            _parse_page_numbers("1-3,3", max_pages=5)

    def test_schema_exposes_text_selector_contracts(self) -> None:
        stdout = StringIO()
        with redirect_stdout(stdout):
            self.assertEqual(main(["schema", "--kind", "text-range"]), 0)
        schema = json.loads(stdout.getvalue())
        self.assertEqual(
            schema["properties"]["unit"]["const"],
            "unicode_codepoint",
        )
        self.assertEqual(schema["properties"]["start"]["minimum"], 0)

        stdout = StringIO()
        with redirect_stdout(stdout):
            self.assertEqual(main(["schema", "--kind", "named-style"]), 0)
        named_style_schema = json.loads(stdout.getvalue())
        self.assertFalse(named_style_schema["additionalProperties"])
        self.assertIn("semantic_role", named_style_schema["properties"])
        self.assertIn("based_on", named_style_schema["properties"])

        stdout = StringIO()
        with redirect_stdout(stdout):
            self.assertEqual(main(["schema", "--kind", "table-layout"]), 0)
        table_layout_schema = json.loads(stdout.getvalue())
        self.assertFalse(table_layout_schema["additionalProperties"])
        self.assertIn("preferred_width", table_layout_schema["properties"])
        self.assertIn("repeat_header", table_layout_schema["properties"])

        stdout = StringIO()
        with redirect_stdout(stdout):
            self.assertEqual(main(["schema", "--kind", "table-column"]), 0)
        table_column_schema = json.loads(stdout.getvalue())
        self.assertFalse(table_column_schema["additionalProperties"])
        self.assertIn("width", table_column_schema["properties"])
        self.assertIn("data_type", table_column_schema["properties"])

        stdout = StringIO()
        with redirect_stdout(stdout):
            self.assertEqual(main(["schema", "--kind", "table-width"]), 0)
        table_width_schema = json.loads(stdout.getvalue())
        self.assertFalse(table_width_schema["additionalProperties"])
        self.assertIn("mode", table_width_schema["properties"])
        self.assertIn("value", table_width_schema["properties"])

        stdout = StringIO()
        with redirect_stdout(stdout):
            self.assertEqual(
                main(["schema", "--kind", "table-cell-format"]),
                0,
            )
        table_cell_format_schema = json.loads(stdout.getvalue())
        self.assertFalse(
            table_cell_format_schema["additionalProperties"]
        )
        self.assertIn(
            "vertical_alignment",
            table_cell_format_schema["properties"],
        )
        self.assertIn(
            "background_color",
            table_cell_format_schema["properties"],
        )
        self.assertIn(
            "borders",
            table_cell_format_schema["properties"],
        )

        stdout = StringIO()
        with redirect_stdout(stdout):
            self.assertEqual(main(["schema", "--kind", "border-line"]), 0)
        border_line_schema = json.loads(stdout.getvalue())
        self.assertFalse(border_line_schema["additionalProperties"])
        self.assertIn("style", border_line_schema["properties"])
        self.assertIn("width", border_line_schema["properties"])

        stdout = StringIO()
        with redirect_stdout(stdout):
            self.assertEqual(main(["schema", "--kind", "table-borders"]), 0)
        table_borders_schema = json.loads(stdout.getvalue())
        self.assertFalse(table_borders_schema["additionalProperties"])
        self.assertIn(
            "inside_horizontal",
            table_borders_schema["properties"],
        )
        self.assertIn(
            "inside_vertical",
            table_borders_schema["properties"],
        )

        stdout = StringIO()
        with redirect_stdout(stdout):
            self.assertEqual(
                main(["schema", "--kind", "table-cell-borders"]),
                0,
            )
        table_cell_borders_schema = json.loads(stdout.getvalue())
        self.assertFalse(
            table_cell_borders_schema["additionalProperties"]
        )
        self.assertEqual(
            set(table_cell_borders_schema["properties"]),
            {"top", "right", "bottom", "left"},
        )

        stdout = StringIO()
        with redirect_stdout(stdout):
            self.assertEqual(
                main(["schema", "--kind", "paragraph-style"]),
                0,
            )
        paragraph_style_schema = json.loads(stdout.getvalue())
        self.assertFalse(
            paragraph_style_schema["additionalProperties"]
        )
        self.assertIn(
            "background_color",
            paragraph_style_schema["properties"],
        )
        self.assertIn(
            "borders",
            paragraph_style_schema["properties"],
        )

        stdout = StringIO()
        with redirect_stdout(stdout):
            self.assertEqual(
                main(["schema", "--kind", "paragraph-borders"]),
                0,
            )
        paragraph_borders_schema = json.loads(stdout.getvalue())
        self.assertFalse(
            paragraph_borders_schema["additionalProperties"]
        )
        self.assertEqual(
            set(paragraph_borders_schema["properties"]),
            {"top", "right", "bottom", "left"},
        )

        stdout = StringIO()
        with redirect_stdout(stdout):
            self.assertEqual(main(["schema", "--kind", "table-cell"]), 0)
        table_cell_schema = json.loads(stdout.getvalue())
        self.assertFalse(table_cell_schema["additionalProperties"])
        self.assertIn("column_span", table_cell_schema["properties"])
        self.assertIn("row_span", table_cell_schema["properties"])
        self.assertIn("content", table_cell_schema["properties"])

    def test_build_validate_inspect_and_apply(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            source = root / "report.json"
            DocumentBuilder(title="Report").paragraph("Old", id="status").build().export(source)

            stdout = StringIO()
            with redirect_stdout(stdout):
                self.assertEqual(main(["validate", str(source)]), 0)
            self.assertIn("VALID", stdout.getvalue())

            stdout = StringIO()
            with redirect_stdout(stdout):
                self.assertEqual(main(["inspect", str(source)]), 0)
            inspection = json.loads(stdout.getvalue())
            self.assertEqual(inspection["node_count"], 1)

            target = root / "report.docx"
            with redirect_stdout(StringIO()):
                self.assertEqual(main(["build", str(source), "--output", str(target)]), 0)
            self.assertTrue(target.exists())

            stdout = StringIO()
            with redirect_stdout(stdout):
                self.assertEqual(main(["capabilities", str(target)]), 0)
            capabilities = json.loads(stdout.getvalue())
            self.assertEqual(capabilities["origin"], "native")
            self.assertEqual(
                capabilities["operations"],
                [
                    "text.replace",
                    "paragraph.format",
                    "text.format",
                    "node.remove",
                    "style.apply",
                    "style.define",
                "style.format",
                "section.format",
                "field.update",
                "image.insert_after",
                "table.format",
                    "table.column.format",
                    "table.cell.format",
                ],
            )
            self.assertEqual(
                capabilities["formatting"]["text_scopes"],
                ["whole_node", "range", "match"],
            )
            self.assertIn(
                "alignment",
                capabilities["formatting"]["paragraph_properties"],
            )
            border_contract = capabilities["formatting"][
                "table_contract"
            ]["border_contract"]
            self.assertEqual(
                border_contract["width_range_points"],
                [0.25, 12],
            )
            self.assertTrue(
                border_contract["direct_cell_precedence"]
            )
            self.assertTrue(
                border_contract[
                    "unsupported_theme_colors_preserved"
                ]
            )
            paragraph_surface = capabilities["formatting"][
                "paragraph_surface_contract"
            ]
            self.assertEqual(
                paragraph_surface["background"],
                "solid_srgb_fill",
            )
            self.assertTrue(
                paragraph_surface["native_style_inheritance"]
            )
            render_providers = {
                provider["name"]: provider
                for provider in capabilities["render"]["providers"]
            }
            self.assertIn("semantic-html", render_providers)
            self.assertIn("libreoffice", render_providers)
            self.assertTrue(capabilities["roundtrip"]["noop_exact"])

            rendered_html = root / "rendered.html"
            stdout = StringIO()
            with redirect_stdout(stdout):
                self.assertEqual(
                    main(
                        [
                            "render",
                            str(source),
                            "--format",
                            "html",
                            "--output",
                            str(rendered_html),
                        ]
                    ),
                    0,
                )
            render_summary = json.loads(stdout.getvalue())
            self.assertEqual(render_summary["provider"], "semantic-html")
            self.assertEqual(render_summary["output"], str(rendered_html))
            self.assertNotIn("content", render_summary)
            self.assertTrue(rendered_html.read_bytes().startswith(b"<!doctype html>"))

            patch = root / "patch.json"
            patch.write_text(
                json.dumps(
                    {
                        "base_revision": 1,
                        "operations": [
                            {
                                "op": "text.replace",
                                "target": "#status",
                                "search": "Old",
                                "replacement": "New",
                            }
                        ],
                    }
                ),
                encoding="utf-8",
            )
            updated = root / "updated.json"
            with redirect_stdout(StringIO()), redirect_stderr(StringIO()):
                self.assertEqual(
                    main(["apply", str(source), str(patch), "--output", str(updated)]),
                    0,
                )
            self.assertIn("New", updated.read_text(encoding="utf-8"))
            self.assertIn('"revision": 2', updated.read_text(encoding="utf-8"))

            workspace_root = root / "workspace"
            stdout = StringIO()
            with redirect_stdout(stdout):
                self.assertEqual(
                    main(["workspace", "init", str(workspace_root)]),
                    0,
                )
            self.assertIn("workspace_id", json.loads(stdout.getvalue()))

            stdout = StringIO()
            with redirect_stdout(stdout):
                self.assertEqual(
                    main(
                        [
                            "workspace",
                            "import",
                            str(target),
                            "--root",
                            str(workspace_root),
                        ]
                    ),
                    0,
                )
            imported = json.loads(stdout.getvalue())
            artifact_id = imported["artifact"]["artifact_id"]

            stdout = StringIO()
            with redirect_stdout(stdout):
                self.assertEqual(
                    main(
                        [
                            "workspace",
                            "list",
                            "--root",
                            str(workspace_root),
                        ]
                    ),
                    0,
                )
            listing = json.loads(stdout.getvalue())
            self.assertEqual(listing["artifacts"][0]["artifact_id"], artifact_id)

            stdout = StringIO()
            with redirect_stdout(stdout):
                self.assertEqual(
                    main(
                        [
                            "workspace",
                            "capabilities",
                            artifact_id,
                            "--root",
                            str(workspace_root),
                        ]
                    ),
                    0,
                )
            workspace_capabilities = json.loads(stdout.getvalue())
            self.assertTrue(workspace_capabilities["revision_store"])
            self.assertEqual(
                workspace_capabilities["artifact"]["artifact_id"],
                artifact_id,
            )


if __name__ == "__main__":
    unittest.main()
