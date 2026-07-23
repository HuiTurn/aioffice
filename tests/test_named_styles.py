from __future__ import annotations

import copy
import io
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile, ZipInfo

from aioffice import Document, DocumentBuilder, resolve_named_style
from aioffice.formats.docx_native import W
from aioffice.native.xml import parse_xml, serialize_xml


def _q(local: str) -> str:
    return f"{{{W}}}{local}"


def _rewrite_part(source: bytes, name: str, payload: bytes) -> bytes:
    output = io.BytesIO()
    with ZipFile(io.BytesIO(source)) as before, ZipFile(output, "w") as after:
        for info in before.infolist():
            data = payload if info.filename == name else before.read(info.filename)
            replacement = ZipInfo(info.filename, date_time=info.date_time)
            replacement.compress_type = info.compress_type or ZIP_DEFLATED
            replacement.external_attr = info.external_attr
            replacement.comment = info.comment
            replacement.extra = info.extra
            after.writestr(replacement, data)
    return output.getvalue()


class NamedStyleTests(unittest.TestCase):
    def test_theme_style_inheritance_is_visible_to_ai_and_html(self) -> None:
        document = (
            DocumentBuilder()
            .paragraph("A considered quotation.", id="quote", style_ref="Quote")
            .build()
        )
        self.assertTrue(document.validate().valid)
        inspection = document.inspect()
        quote_style = next(style for style in inspection["styles"] if style["id"] == "Quote")
        self.assertEqual(quote_style["semantic_role"], "quote")
        self.assertEqual(quote_style["usage_count"], 1)

        resolved = resolve_named_style(document.spec, "Quote")
        assert resolved.paragraph_style is not None
        assert resolved.text_style is not None
        self.assertEqual(resolved.inheritance_chain, ("Normal", "Quote"))
        self.assertEqual(resolved.paragraph_style.indent_left.to_points(), 24)
        self.assertTrue(resolved.text_style.italic)

        html = document.to_bytes("html").decode()
        self.assertIn('data-aioffice-style="Quote"', html)
        self.assertIn("margin-left:24pt", html)
        self.assertIn("font-style:italic", html)

    def test_missing_style_cycle_and_semantic_mismatch_are_diagnostics(self) -> None:
        missing = Document.from_spec(
            {
                "content": [
                    {
                        "type": "paragraph",
                        "id": "body",
                        "text": "Body",
                        "style_ref": "Missing",
                    }
                ]
            }
        )
        self.assertIn("STYLE_NOT_FOUND", {item.code for item in missing.validate().errors})

        cycle = Document.from_spec(
            {
                "styles": [
                    {
                        "id": "CycleA",
                        "name": "Cycle A",
                        "based_on": "CycleB",
                    },
                    {
                        "id": "CycleB",
                        "name": "Cycle B",
                        "based_on": "CycleA",
                    },
                ],
                "content": [{"type": "paragraph", "text": "Body"}],
            }
        )
        self.assertIn(
            "STYLE_INHERITANCE_CYCLE",
            {item.code for item in cycle.validate().errors},
        )

        mismatch = (
            DocumentBuilder()
            .heading("Not a quote", style_ref="Quote")
            .build()
        )
        self.assertIn(
            "STYLE_SEMANTIC_MISMATCH",
            {item.code for item in mismatch.validate().errors},
        )

    def test_custom_heading_style_round_trips_with_role_and_level(self) -> None:
        document = (
            DocumentBuilder()
            .define_style(
                {
                    "id": "BoardHeading",
                    "name": "Board Heading",
                    "semantic_role": "heading",
                    "heading_level": 2,
                    "based_on": "Heading2",
                    "next_style": "Normal",
                    "quick_style": True,
                    "text_style": {"color": "#7A1F5B"},
                }
            )
            .heading(
                "Board summary",
                id="board",
                level=2,
                style_ref="BoardHeading",
            )
            .build()
        )
        self.assertTrue(document.validate().valid)
        reopened = Document.from_docx(document.to_bytes("docx"))
        node = reopened.to_spec()["content"][0]
        self.assertEqual(node["type"], "heading")
        self.assertEqual(node["level"], 2)
        self.assertEqual(node["style_ref"], "BoardHeading")
        style = next(
            style for style in reopened.to_spec()["styles"] if style["id"] == "BoardHeading"
        )
        self.assertEqual(style["semantic_role"], "heading")
        self.assertEqual(style["heading_level"], 2)
        self.assertEqual(style["text_style"]["color"], "#7A1F5B")

    def test_style_apply_changes_semantic_role_deterministically(self) -> None:
        document = DocumentBuilder().paragraph("Decision", id="decision").build()
        promoted = document.apply(
            [
                {
                    "op": "style.apply",
                    "target": "#decision",
                    "style_ref": "Heading2",
                }
            ]
        )
        self.assertTrue(promoted.success, promoted.model_dump())
        assert promoted.document is not None
        promoted_node = promoted.document.to_spec()["content"][0]
        self.assertEqual(promoted_node["type"], "heading")
        self.assertEqual(promoted_node["level"], 2)
        self.assertNotIn("style_ref", promoted_node)

        demoted = promoted.document.apply(
            [
                {
                    "op": "style.apply",
                    "target": "#decision",
                    "style_ref": "Quote",
                }
            ]
        )
        self.assertTrue(demoted.success, demoted.model_dump())
        assert demoted.document is not None
        demoted_node = demoted.document.to_spec()["content"][0]
        self.assertEqual(demoted_node["type"], "paragraph")
        self.assertEqual(demoted_node["style_ref"], "Quote")
        self.assertNotIn("level", demoted_node)

    def test_clearing_custom_heading_style_uses_implicit_heading_style_natively(self) -> None:
        source = (
            DocumentBuilder()
            .define_style(
                {
                    "id": "CustomHeading2",
                    "name": "Custom Heading 2",
                    "semantic_role": "heading",
                    "heading_level": 2,
                    "based_on": "Heading2",
                }
            )
            .heading(
                "Summary",
                id="summary",
                level=2,
                style_ref="CustomHeading2",
            )
            .build()
            .to_bytes("docx")
        )
        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "style.apply",
                    "target": "#summary",
                    "style_ref": None,
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        semantic = result.document.to_spec()["content"][0]
        self.assertEqual(semantic["type"], "heading")
        self.assertEqual(semantic["level"], 2)
        self.assertNotIn("style_ref", semantic)

        reopened = Document.from_docx(result.document.to_bytes("docx"))
        node = reopened.to_spec()["content"][0]
        self.assertEqual(node["type"], "heading")
        self.assertEqual(node["level"], 2)
        self.assertNotIn("style_ref", node)

    def test_native_style_patch_preserves_unknown_style_xml_and_other_parts(self) -> None:
        source = DocumentBuilder().paragraph("Body", id="body").build().to_bytes("docx")
        styles_root = parse_xml(ZipFile(io.BytesIO(source)).read("word/styles.xml"))
        quote = next(
            element
            for element in styles_root.findall(_q("style"))
            if element.attrib.get(_q("styleId")) == "Quote"
        )
        future = ET.SubElement(quote, "{urn:aioffice:test}futureStyle")
        future.set("mode", "preserve")
        source = _rewrite_part(source, "word/styles.xml", serialize_xml(styles_root))

        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "style.format",
                    "target": "@Quote",
                    "paragraph": {
                        "set": {
                            "spacing_after": {"value": 15, "unit": "pt"},
                        }
                    },
                    "text": {
                        "set": {"color": "#A02040"},
                        "clear": ["italic"],
                    },
                },
                {
                    "op": "style.apply",
                    "target": "#body",
                    "style_ref": "Quote",
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
                "/word/styles.xml",
            ],
        )
        output = result.document.to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as before, ZipFile(io.BytesIO(output)) as after:
            for name in before.namelist():
                if name not in {
                    "customXml/aioffice-manifest.xml",
                    "word/document.xml",
                    "word/styles.xml",
                }:
                    self.assertEqual(before.read(name), after.read(name), name)
            patched_styles = parse_xml(after.read("word/styles.xml"))
        patched_quote = next(
            element
            for element in patched_styles.findall(_q("style"))
            if element.attrib.get(_q("styleId")) == "Quote"
        )
        self.assertIsNotNone(patched_quote.find("{urn:aioffice:test}futureStyle"))

        reopened = Document.from_docx(output)
        node = reopened.to_spec()["content"][0]
        self.assertEqual(node["style_ref"], "Quote")
        style = next(
            style for style in reopened.to_spec()["styles"] if style["id"] == "Quote"
        )
        self.assertEqual(style["paragraph_style"]["spacing_after"]["value"], 15.0)
        self.assertEqual(style["text_style"]["color"], "#A02040")
        self.assertNotIn("italic", style["text_style"])

    def test_native_style_define_and_apply_are_atomic_in_one_patch(self) -> None:
        source = DocumentBuilder().paragraph("Important", id="body").build().to_bytes("docx")
        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "style.define",
                    "style": {
                        "id": "Executive",
                        "name": "Executive",
                        "based_on": "Normal",
                        "semantic_role": "custom",
                        "paragraph_style": {
                            "keep_together": True,
                            "spacing_after": {"value": 14, "unit": "pt"},
                        },
                        "text_style": {
                            "font_size": {"value": 13, "unit": "pt"},
                            "bold": True,
                        },
                    },
                },
                {
                    "op": "style.apply",
                    "target": "#body",
                    "style_ref": "Executive",
                },
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        reopened = Document.from_docx(result.document.to_bytes("docx"))
        self.assertEqual(reopened.to_spec()["content"][0]["style_ref"], "Executive")
        executive = next(
            style for style in reopened.to_spec()["styles"] if style["id"] == "Executive"
        )
        self.assertTrue(executive["text_style"]["bold"])
        self.assertTrue(executive["metadata"]["native_custom_style"])

    def test_style_only_native_patch_does_not_rewrite_document_xml(self) -> None:
        source = DocumentBuilder().paragraph("Body", id="body").build().to_bytes("docx")
        document = Document.from_docx(source)
        result = document.apply(
            [
                {
                    "op": "style.format",
                    "target": "Normal",
                    "text": {"set": {"color": "#334455"}},
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        assert result.fidelity is not None
        self.assertEqual(
            result.fidelity.affected_parts,
            ["/customXml/aioffice-manifest.xml", "/word/styles.xml"],
        )
        with ZipFile(io.BytesIO(source)) as before, ZipFile(
            io.BytesIO(result.document.to_bytes("docx"))
        ) as after:
            self.assertEqual(
                before.read("word/document.xml"),
                after.read("word/document.xml"),
            )

    def test_style_format_rejects_unknown_property_atomically(self) -> None:
        document = DocumentBuilder().paragraph("Body", id="body").build()
        result = document.apply(
            [
                {
                    "op": "style.format",
                    "target": "Normal",
                    "paragraph": {"set": {"imaginary_spacing": 12}},
                }
            ]
        )
        self.assertFalse(result.success)
        self.assertEqual(result.diagnostics[0].code, "INVALID_SPEC")
        self.assertEqual(document.revision, 1)
        self.assertEqual(document.to_spec()["styles"], [])

    def test_duplicate_native_style_ids_remain_noop_lossless_but_are_not_edited(self) -> None:
        source = DocumentBuilder().paragraph("Body", id="body").build().to_bytes("docx")
        styles_root = parse_xml(ZipFile(io.BytesIO(source)).read("word/styles.xml"))
        quote = next(
            element
            for element in styles_root.findall(_q("style"))
            if element.attrib.get(_q("styleId")) == "Quote"
        )
        styles_root.append(copy.deepcopy(quote))
        source = _rewrite_part(source, "word/styles.xml", serialize_xml(styles_root))

        document = Document.from_docx(source)
        self.assertTrue(document.validate().valid)
        self.assertEqual(document.to_bytes("docx"), source)
        self.assertIn(
            "STYLE_PROJECTION_AMBIGUOUS",
            {item.code for item in document.import_diagnostics},
        )
        result = document.apply(
            [
                {
                    "op": "style.format",
                    "target": "Quote",
                    "text": {"set": {"bold": True}},
                }
            ]
        )
        self.assertFalse(result.success)
        self.assertEqual(result.diagnostics[0].code, "NATIVE_PATCH_FAILED")
        self.assertIn("ambiguous", result.diagnostics[0].message)
        self.assertEqual(document.to_bytes("docx"), source)

    def test_document_defaults_round_trip_as_supported_properties(self) -> None:
        document = (
            DocumentBuilder(
                defaults={
                    "paragraph_style": {
                        "line_spacing": {"rule": "multiple", "value": 1.4},
                        "widow_control": False,
                    },
                    "text_style": {
                        "font_family": "Arial",
                        "font_family_east_asia": "SimSun",
                        "font_size": {"value": 10.5, "unit": "pt"},
                    },
                }
            )
            .paragraph("Defaults")
            .build()
        )
        reopened = Document.from_docx(document.to_bytes("docx"))
        defaults = reopened.to_spec()["defaults"]
        self.assertEqual(defaults["paragraph_style"]["line_spacing"]["value"], 1.4)
        self.assertFalse(defaults["paragraph_style"]["widow_control"])
        self.assertEqual(defaults["text_style"]["font_family"], "Arial")
        self.assertEqual(defaults["text_style"]["font_family_east_asia"], "SimSun")
        self.assertEqual(defaults["text_style"]["font_size"]["value"], 10.5)


if __name__ == "__main__":
    unittest.main()
