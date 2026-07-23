from __future__ import annotations

import io
import unittest
from importlib.util import find_spec

from aioffice.core.errors import SpecValidationError
from aioffice.documents import DocumentBuilder
from aioffice.rendering import compare_raster_images


class StyleAndRenderingTests(unittest.TestCase):
    def test_style_values_are_strict_and_export_to_html(self) -> None:
        document = DocumentBuilder(title="Style").paragraph(
            "Centered",
            id="body",
            paragraph_style={
                "alignment": "center",
                "spacing_before": {"value": 12, "unit": "pt"},
                "line_spacing": {"rule": "multiple", "value": 1.25},
            },
            text_style={
                "font_family": 'Aptos "Display"',
                "font_size": {"value": 11, "unit": "pt"},
                "color": "#1f4e78",
                "bold": True,
            },
        ).build()

        node = document.to_spec()["content"][0]
        self.assertEqual(node["text_style"]["color"], "#1F4E78")
        output = document.to_bytes("html").decode()
        self.assertIn("text-align:center", output)
        self.assertIn("margin-top:12pt", output)
        self.assertIn("font-size:11pt", output)
        self.assertIn("#1F4E78", output)
        self.assertIn("&quot;Display\\&quot;", output)

        with self.assertRaises(SpecValidationError):
            DocumentBuilder().paragraph(
                "Invalid",
                text_style={"font_size": 12},
            ).build()

    def test_format_patch_returns_property_changes_and_semantic_diff(self) -> None:
        document = DocumentBuilder().paragraph("Body", id="body").build()
        result = document.apply(
            [
                {
                    "op": "paragraph.format",
                    "target": "#body",
                    "set": {
                        "alignment": "justify",
                        "spacing_after": {"value": 8, "unit": "pt"},
                    },
                },
                {
                    "op": "text.format",
                    "target": "#body",
                    "set": {
                        "font_size": {"value": 10.5, "unit": "pt"},
                        "color": "#336699",
                    },
                },
            ]
        )
        self.assertTrue(result.success)
        self.assertEqual(
            result.changes[0]["property_changes"][0]["path"],
            "paragraph_style.alignment",
        )
        self.assertIsNotNone(result.diff)
        assert result.diff is not None
        paths = {entry.path for entry in result.diff.entries}
        self.assertIn("content.#body.paragraph_style", paths)
        self.assertIn("content.#body.text_style", paths)
        self.assertEqual(result.diff.summary["added"], 2)

        assert result.document is not None
        cleared = result.document.apply(
            [
                {
                    "op": "paragraph.format",
                    "target": "#body",
                    "clear": ["alignment"],
                }
            ]
        )
        self.assertTrue(cleared.success)
        assert cleared.document is not None
        node = cleared.document.to_spec()["content"][0]
        self.assertNotIn("alignment", node["paragraph_style"])

    def test_render_contract_marks_semantic_html_as_preview_only(self) -> None:
        document = DocumentBuilder().paragraph("Preview", id="body").build()
        first = document.render()
        second = document.render()
        self.assertEqual(first.content, second.content)
        self.assertEqual(first.cache_key, second.cache_key)
        self.assertEqual(first.fidelity, "approximate")
        self.assertEqual(first.verification_status, "preview_only")
        self.assertEqual(first.diagnostics[0].code, "APPROXIMATE_RENDER")
        self.assertNotIn("content", first.summary())
        flow = document.render(
            options={
                "page_view": False,
                "include_document_metadata": False,
                "locale": "zh-CN",
            }
        )
        self.assertNotEqual(flow.content, first.content)
        self.assertNotEqual(flow.cache_key, first.cache_key)
        self.assertIn(b'<html lang="zh-CN">', flow.content)
        self.assertIn(b"box-shadow:none", flow.content)

    @unittest.skipUnless(find_spec("PIL") is not None, "Pillow is optional")
    def test_raster_visual_regression_metrics(self) -> None:
        from PIL import Image

        def png(color: tuple[int, int, int]) -> bytes:
            output = io.BytesIO()
            Image.new("RGB", (2, 2), color).save(output, format="PNG")
            return output.getvalue()

        same = compare_raster_images(png((255, 255, 255)), png((255, 255, 255)))
        self.assertTrue(same.passed)
        changed = compare_raster_images(
            png((255, 255, 255)),
            png((0, 0, 0)),
        )
        self.assertFalse(changed.passed)
        self.assertEqual(changed.changed_pixel_ratio, 1.0)
        self.assertEqual(changed.diagnostics[0].code, "VISUAL_REGRESSION")


if __name__ == "__main__":
    unittest.main()
