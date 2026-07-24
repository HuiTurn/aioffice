from __future__ import annotations

import json
import unittest

from aioffice.core.errors import SpecValidationError
from aioffice.documents import Document, DocumentBuilder


class DocumentTests(unittest.TestCase):
    def test_builder_creates_valid_stable_spec(self) -> None:
        document = (
            DocumentBuilder(title="Status")
            .heading("Status", id="title")
            .paragraph("Ready", id="status")
            .build()
        )
        self.assertTrue(document.validate().valid)
        self.assertEqual(document.revision, 1)
        self.assertEqual(
            [node["id"] for node in document.to_spec()["content"]], ["title", "status"]
        )

        loaded = Document.from_json(document.to_json())
        self.assertEqual(loaded.id, document.id)
        self.assertEqual(loaded.to_spec(), document.to_spec())

    def test_unknown_fields_are_rejected(self) -> None:
        with self.assertRaises(SpecValidationError) as caught:
            Document.from_spec(
                {
                    "content": [
                        {
                            "type": "paragraph",
                            "text": "Hello",
                            "invented_by_model": True,
                        }
                    ]
                }
            )
        self.assertEqual(caught.exception.diagnostics[0].code, "INVALID_SPEC")

    def test_duplicate_ids_and_unknown_table_columns_are_diagnostics(self) -> None:
        document = Document.from_spec(
            {
                "content": [
                    {"id": "same", "type": "paragraph", "text": "A"},
                    {"id": "same", "type": "paragraph", "text": "B"},
                    {
                        "id": "table",
                        "type": "table",
                        "columns": [{"key": "known", "title": "Known"}],
                        "rows": [{"id": "row", "values": {"unknown": "value"}}],
                    },
                ]
            }
        )
        result = document.validate()
        self.assertFalse(result.valid)
        self.assertGreaterEqual(len(result.errors), 2)

    def test_json_root_must_be_object(self) -> None:
        with self.assertRaises(SpecValidationError):
            Document.from_json(json.dumps([]))

    def test_legacy_draft_spec_is_migrated(self) -> None:
        document = Document.from_spec(
            {
                "$schema": "https://schemas.aioffice.dev/spec/1.0/document.json",
                "spec_version": "1.0",
                "content": [{"type": "paragraph", "text": "Legacy"}],
            }
        )
        self.assertEqual(document.spec_version, "0.2-draft.39")
        self.assertEqual(
            document.to_spec()["$schema"],
            "https://schemas.aioffice.dev/spec/draft/0.2/document.json",
        )


if __name__ == "__main__":
    unittest.main()
