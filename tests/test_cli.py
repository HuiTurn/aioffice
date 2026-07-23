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
                ],
            )
            self.assertEqual(
                capabilities["formatting"]["text_scope"],
                "whole_node",
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
