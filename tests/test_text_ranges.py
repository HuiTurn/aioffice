from __future__ import annotations

import unittest

from aioffice.documents import DocumentBuilder
from aioffice.operations import TextMatch, TextRange


def _plain_text(node: dict[str, object]) -> str:
    text = node.get("text")
    if isinstance(text, str):
        return text
    content = node.get("content", [])
    assert isinstance(content, list)
    return "".join(str(span.get("text", "")) for span in content if isinstance(span, dict))


class TextRangeSemanticTests(unittest.TestCase):
    def test_unicode_codepoint_range_splits_exactly(self) -> None:
        document = DocumentBuilder().paragraph("A😀BC中文D", id="body").build()
        result = document.apply(
            [
                {
                    "op": "text.format",
                    "target": "#body",
                    "range": {
                        "start": 1,
                        "end": 4,
                        "unit": "unicode_codepoint",
                    },
                    "set": {"color": "#ff0000", "bold": True},
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(result.changes[0]["selected_text"], "😀BC")
        assert result.document is not None
        node = result.document.to_spec()["content"][0]
        self.assertEqual(
            [span["text"] for span in node["content"]],
            ["A", "😀BC", "中文D"],
        )
        self.assertEqual(
            node["content"][1]["style"],
            {"color": "#FF0000", "bold": True},
        )

    def test_exact_match_uses_one_based_occurrence(self) -> None:
        document = DocumentBuilder().paragraph("one two one", id="body").build()
        result = document.apply(
            [
                {
                    "op": "text.format",
                    "target": "#body",
                    "match": {"text": "one", "occurrence": 2},
                    "set": {"underline": True},
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        self.assertEqual(
            result.changes[0]["selection"],
            {"start": 8, "end": 11, "unit": "unicode_codepoint"},
        )
        assert result.document is not None
        spans = result.document.to_spec()["content"][0]["content"]
        self.assertEqual([span["text"] for span in spans], ["one two ", "one"])
        self.assertEqual(spans[1]["style"], {"underline": True})

    def test_whole_node_format_overrides_conflicting_span_property(self) -> None:
        document = (
            DocumentBuilder()
            .rich_paragraph(
                [
                    {
                        "text": "Red",
                        "style": {"color": "#C00000", "italic": True},
                    },
                    {"text": " plain"},
                ],
                id="body",
            )
            .build()
        )
        result = document.apply(
            [
                {
                    "op": "text.format",
                    "target": "#body",
                    "set": {"color": "#1F4E78"},
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        node = result.document.to_spec()["content"][0]
        self.assertEqual(node["text_style"]["color"], "#1F4E78")
        self.assertEqual(node["content"][0]["style"], {"italic": True})

    def test_partial_clear_materializes_common_style_without_visual_drift(self) -> None:
        document = (
            DocumentBuilder()
            .paragraph(
                "ABCD",
                id="body",
                text_style={"bold": True, "color": "#1F4E78"},
            )
            .build()
        )
        result = document.apply(
            [
                {
                    "op": "text.format",
                    "target": "#body",
                    "range": {"start": 1, "end": 3},
                    "clear": ["bold"],
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        node = result.document.to_spec()["content"][0]
        self.assertNotIn("text_style", node)
        self.assertEqual(
            [span["style"] for span in node["content"]],
            [
                {"color": "#1F4E78", "bold": True},
                {"color": "#1F4E78"},
                {"color": "#1F4E78", "bold": True},
            ],
        )

    def test_replace_crosses_semantic_span_boundaries(self) -> None:
        document = (
            DocumentBuilder()
            .rich_paragraph(
                [
                    {"text": "Alpha ", "marks": ["strong"]},
                    {"text": "Beta", "marks": ["emphasis"]},
                    {"text": " Gamma"},
                ],
                id="body",
            )
            .build()
        )
        result = document.apply(
            [
                {
                    "op": "text.replace",
                    "target": "#body",
                    "search": "ha Be",
                    "replacement": "HA-BE",
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        node = result.document.to_spec()["content"][0]
        self.assertEqual(_plain_text(node), "AlpHA-BEta Gamma")
        self.assertEqual(node["content"][0]["marks"], ["strong"])
        self.assertEqual(node["content"][1]["marks"], ["emphasis"])

    def test_range_format_supports_rich_heading(self) -> None:
        document = (
            DocumentBuilder()
            .rich_heading(
                [
                    {"text": "Quarterly "},
                    {"text": "Review", "marks": ["emphasis"]},
                ],
                id="title",
            )
            .build()
        )
        result = document.apply(
            [
                {
                    "op": "text.format",
                    "target": "#title",
                    "match": {"text": "Review"},
                    "set": {"color": "#1F4E78"},
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        heading = result.document.spec.content[0]
        self.assertEqual(heading.plain_text, "Quarterly Review")
        self.assertIn("#1F4E78", result.document.to_bytes("html").decode())

    def test_invalid_or_ambiguous_selection_is_atomic(self) -> None:
        document = DocumentBuilder().paragraph("short", id="body").build()
        invalid = document.apply(
            [
                {
                    "op": "text.format",
                    "target": "#body",
                    "range": {"start": 2, "end": 99},
                    "set": {"bold": True},
                }
            ]
        )
        self.assertFalse(invalid.success)
        self.assertEqual(invalid.diagnostics[0].code, "INVALID_TEXT_SELECTION")
        self.assertIsNone(invalid.document)
        ambiguous = document.apply(
            [
                {
                    "op": "text.format",
                    "target": "#body",
                    "range": {"start": 0, "end": 2},
                    "match": {"text": "sh"},
                    "set": {"bold": True},
                }
            ]
        )
        self.assertFalse(ambiguous.success)
        self.assertEqual(ambiguous.diagnostics[0].code, "INVALID_TEXT_SELECTION")

    def test_public_selector_models_reject_ambiguous_units_and_indexes(self) -> None:
        self.assertEqual(TextRange(start=0, end=1).unit, "unicode_codepoint")
        self.assertEqual(TextMatch(text="x").occurrence, 1)
        with self.assertRaises(ValueError):
            TextRange(start=True, end=2)

    def test_text_replace_rejects_coerced_boolean_and_unknown_fields(self) -> None:
        document = DocumentBuilder().paragraph("one one", id="body").build()
        coerced = document.apply(
            [
                {
                    "op": "text.replace",
                    "target": "#body",
                    "search": "one",
                    "replacement": "two",
                    "replace_all": "false",
                }
            ]
        )
        self.assertFalse(coerced.success)
        self.assertEqual(coerced.diagnostics[0].code, "INVALID_SPEC")
        unknown = document.apply(
            [
                {
                    "op": "text.replace",
                    "target": "#body",
                    "search": "one",
                    "replacement": "two",
                    "occurrence": 1,
                }
            ]
        )
        self.assertFalse(unknown.success)
        self.assertIn("unknown fields", unknown.diagnostics[0].message)


if __name__ == "__main__":
    unittest.main()
