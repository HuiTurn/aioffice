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
        invalid = self.document.apply(
            [
                {
                    "op": "node.append",
                    "target": "$",
                    "content": {
                        "id": "invalid_append",
                        "type": "paragraph",
                        "text": "Invalid",
                    },
                    "index": 0,
                }
            ]
        )
        self.assertFalse(invalid.success)
        self.assertEqual(
            invalid.diagnostics[0].code,
            "INVALID_SPEC",
        )
        self.assertIn(
            "unknown fields: index",
            invalid.diagnostics[0].message,
        )

    def test_insert_before_prepends_and_rebinds_section_start(self) -> None:
        document = (
            DocumentBuilder(
                sections=[
                    {"id": "front", "start_at": None},
                    {
                        "id": "body_section",
                        "start_at": "c",
                        "layout": {"start_type": "next_page"},
                    },
                ]
            )
            .paragraph("A", id="a")
            .paragraph("B", id="b")
            .paragraph("C", id="c")
            .paragraph("D", id="d")
            .build()
        )
        result = document.apply(
            [
                {
                    "op": "node.insert_before",
                    "target": "#a",
                    "content": {
                        "id": "prelude",
                        "type": "paragraph",
                        "text": "Prelude",
                    },
                },
                {
                    "op": "node.insert_before",
                    "target": "#c",
                    "content": {
                        "id": "section_intro",
                        "type": "heading",
                        "level": 2,
                        "text": "Section intro",
                    },
                },
                {
                    "op": "node.insert_before",
                    "target": "#section_intro",
                    "content": {
                        "id": "section_label",
                        "type": "paragraph",
                        "text": "Section label",
                    },
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        self.assertEqual(
            [
                node["id"]
                for node in result.document.to_spec()["content"]
            ],
            [
                "prelude",
                "a",
                "b",
                "section_label",
                "section_intro",
                "c",
                "d",
            ],
        )
        self.assertEqual(
            result.document.to_spec()["sections"][1]["start_at"],
            "section_label",
        )
        self.assertEqual(
            result.changes[1]["section_start_updated"],
            {
                "section_id": "body_section",
                "from": "c",
                "to": "section_intro",
            },
        )
        self.assertEqual(
            result.changes[2]["section_start_updated"],
            {
                "section_id": "body_section",
                "from": "section_intro",
                "to": "section_label",
            },
        )
        invalid = document.apply(
            [
                {
                    "op": "node.insert_before",
                    "target": "#a",
                    "content": {
                        "id": "invalid",
                        "type": "paragraph",
                        "text": "Invalid",
                    },
                    "index": 0,
                }
            ]
        )
        self.assertFalse(invalid.success)
        self.assertEqual(
            invalid.diagnostics[0].code,
            "INVALID_SPEC",
        )
        self.assertIn(
            "unknown fields: index",
            invalid.diagnostics[0].message,
        )

    def test_move_after_uses_stable_ids_and_preserves_sections(self) -> None:
        document = (
            DocumentBuilder(
                sections=[
                    {"id": "front", "start_at": None},
                    {
                        "id": "body_section",
                        "start_at": "c",
                        "layout": {"start_type": "next_page"},
                    },
                ]
            )
            .paragraph("A", id="a")
            .paragraph("B", id="b")
            .paragraph("C", id="c")
            .paragraph("D", id="d")
            .paragraph("E", id="e")
            .paragraph("F", id="f")
            .build()
        )
        structural_capabilities = document.capabilities()[
            "structural_editing"
        ]
        self.assertTrue(structural_capabilities["available"])
        self.assertEqual(
            structural_capabilities["append_operation"],
            "node.append",
        )
        self.assertEqual(
            structural_capabilities["append_position"],
            "before_terminal_body_sectPr",
        )
        self.assertEqual(
            structural_capabilities["append_section"],
            "last_semantic_section",
        )
        self.assertTrue(
            structural_capabilities["append_empty_document"]
        )
        self.assertEqual(
            structural_capabilities["insert_operations"],
            {
                "after": "node.insert_after",
                "before": "node.insert_before",
            },
        )
        self.assertEqual(
            structural_capabilities["move_operations"],
            {
                "after": "node.move_after",
                "before": "node.move_before",
            },
        )
        self.assertEqual(
            structural_capabilities["prepend_to_section"],
            "rebind_section_start_at",
        )
        self.assertEqual(
            structural_capabilities["remove_operation"],
            "node.remove",
        )
        self.assertEqual(
            structural_capabilities["native_remove_orphan_policy"],
            "preserve_unreferenced_relationships_and_parts",
        )
        result = document.apply(
            [
                {
                    "op": "node.move_after",
                    "target": "#d",
                    "after": "#f",
                },
                {
                    "op": "node.move_after",
                    "target": "#f",
                    "after": "#c",
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        self.assertEqual(
            [
                node["id"]
                for node in result.document.to_spec()["content"]
            ],
            ["a", "b", "c", "f", "e", "d"],
        )
        self.assertEqual(
            result.changes,
            [
                {
                    "operation": "node.move_after",
                    "moved_nodes": ["d"],
                    "from_after": "c",
                    "after": "f",
                    "section_index": 1,
                },
                {
                    "operation": "node.move_after",
                    "moved_nodes": ["f"],
                    "from_after": "e",
                    "after": "c",
                    "section_index": 1,
                },
            ],
        )
        self.assertEqual(
            [node["id"] for node in document.to_spec()["content"]],
            ["a", "b", "c", "d", "e", "f"],
        )
        self.assertEqual(
            [
                section.get("start_at")
                for section in result.document.to_spec()["sections"]
            ],
            [None, "c"],
        )
        before_result = document.apply(
            [
                {
                    "op": "node.move_before",
                    "target": "#f",
                    "before": "#c",
                }
            ]
        )
        self.assertTrue(
            before_result.success,
            before_result.model_dump(),
        )
        assert before_result.document is not None
        self.assertEqual(
            [
                node["id"]
                for node in before_result.document.to_spec()["content"]
            ],
            ["a", "b", "f", "c", "d", "e"],
        )
        self.assertEqual(
            [
                section.get("start_at")
                for section in before_result.document.to_spec()["sections"]
            ],
            [None, "f"],
        )
        self.assertEqual(
            before_result.changes,
            [
                {
                    "operation": "node.move_before",
                    "moved_nodes": ["f"],
                    "from_after": "e",
                    "section_index": 1,
                    "before": "c",
                    "section_start_updated": {
                        "section_id": "body_section",
                        "from": "c",
                        "to": "f",
                    },
                }
            ],
        )

        invalid_operations = [
            {
                "op": "node.move_after",
                "target": "#a",
                "after": "#c",
            },
            {
                "op": "node.move_after",
                "target": "#c",
                "after": "#d",
            },
            {
                "op": "node.move_after",
                "target": "#b",
                "after": "#a",
            },
            {
                "op": "node.move_after",
                "target": "#d",
                "after": "#d",
            },
            {
                "op": "node.move_after",
                "target": "#d",
                "after": "#f",
                "index": 2,
            },
            {
                "op": "node.move_before",
                "target": "#a",
                "before": "#c",
            },
            {
                "op": "node.move_before",
                "target": "#c",
                "before": "#e",
            },
            {
                "op": "node.move_before",
                "target": "#a",
                "before": "#b",
            },
            {
                "op": "node.move_before",
                "target": "#d",
                "before": "#d",
            },
            {
                "op": "node.move_before",
                "target": "#d",
                "before": "#f",
                "index": 2,
            },
        ]
        expected_codes = [
            "CROSS_SECTION_MOVE_UNSUPPORTED",
            "UNSUPPORTED_FEATURE",
            "NO_CHANGES",
            "INVALID_SPEC",
            "INVALID_SPEC",
            "CROSS_SECTION_MOVE_UNSUPPORTED",
            "UNSUPPORTED_FEATURE",
            "NO_CHANGES",
            "INVALID_SPEC",
            "INVALID_SPEC",
        ]
        for operation, expected_code in zip(
            invalid_operations,
            expected_codes,
            strict=True,
        ):
            with self.subTest(operation=operation):
                invalid = document.apply([operation])
                self.assertFalse(invalid.success)
                self.assertEqual(
                    invalid.diagnostics[0].code,
                    expected_code,
                )
        self.assertEqual(
            [node["id"] for node in document.to_spec()["content"]],
            ["a", "b", "c", "d", "e", "f"],
        )


if __name__ == "__main__":
    unittest.main()
