from __future__ import annotations

import unittest

from aioffice.documents import DocumentBuilder


class PatchTests(unittest.TestCase):
    def setUp(self) -> None:
        self.document = (
            DocumentBuilder(title="Report")
            .heading("Report", id="title")
            .paragraph("Phase one is complete.", id="status")
            .build()
        )

    def test_replace_returns_next_revision_without_mutating_source(self) -> None:
        result = self.document.apply(
            [
                {
                    "op": "text.replace",
                    "target": "#status",
                    "search": "one",
                    "replacement": "two",
                }
            ],
            base_revision=1,
            dry_run=True,
        )
        self.assertTrue(result.success)
        self.assertEqual(result.result_revision, 2)
        self.assertIsNotNone(result.document)
        self.assertEqual(self.document.revision, 1)
        self.assertIn("Phase one", self.document.to_json())
        assert result.document is not None
        self.assertIn("Phase two", result.document.to_json())

    def test_patch_is_atomic_when_later_operation_fails(self) -> None:
        result = self.document.apply(
            [
                {
                    "op": "text.replace",
                    "target": "#status",
                    "search": "one",
                    "replacement": "two",
                },
                {"op": "node.remove", "target": "#missing"},
            ]
        )
        self.assertFalse(result.success)
        self.assertIsNone(result.document)
        self.assertEqual(self.document.revision, 1)
        self.assertIn("Phase one", self.document.to_json())

    def test_revision_conflict_is_structured(self) -> None:
        result = self.document.apply(
            [{"op": "node.remove", "target": "#status"}],
            base_revision=9,
        )
        self.assertFalse(result.success)
        self.assertEqual(result.diagnostics[0].code, "REVISION_CONFLICT")

    def test_append_and_insert_after_generate_valid_nodes(self) -> None:
        result = self.document.apply(
            [
                {
                    "op": "node.insert_after",
                    "target": "#title",
                    "content": {"type": "paragraph", "text": "Summary"},
                },
                {
                    "op": "node.append",
                    "target": "$",
                    "content": {"type": "page_break"},
                },
            ]
        )
        self.assertTrue(result.success)
        assert result.document is not None
        self.assertEqual(result.document.revision, 2)
        self.assertEqual(len(result.document.to_spec()["content"]), 4)
        self.assertNotEqual(result.changes[0]["created_nodes"], ["<generated>"])
        self.assertNotEqual(result.changes[1]["created_nodes"], ["<generated>"])


if __name__ == "__main__":
    unittest.main()
