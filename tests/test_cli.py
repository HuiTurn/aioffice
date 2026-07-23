from __future__ import annotations

import json
import tempfile
import unittest
from contextlib import redirect_stderr, redirect_stdout
from io import StringIO
from pathlib import Path

from aioffice.cli import main
from aioffice.documents import DocumentBuilder


class CliTests(unittest.TestCase):
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
            self.assertTrue(capabilities["roundtrip"]["noop_exact"])

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
