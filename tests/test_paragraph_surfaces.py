from __future__ import annotations

import copy
import io
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZIP_DEFLATED, ZipFile

from aioffice import Document
from aioffice.core.errors import SpecValidationError
from aioffice.native.xml import parse_xml, serialize_xml

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
FUTURE = "urn:aioffice:test:paragraph-surface"


def _q(local: str) -> str:
    return f"{{{W}}}{local}"


def _rewrite_part(
    source: bytes,
    name: str,
    payload: bytes,
) -> bytes:
    output = io.BytesIO()
    with (
        ZipFile(io.BytesIO(source)) as before,
        ZipFile(output, "w", compression=ZIP_DEFLATED) as after,
    ):
        for info in before.infolist():
            after.writestr(
                copy.copy(info),
                payload if info.filename == name else before.read(info.filename),
            )
    return output.getvalue()


def _surface_document() -> Document:
    edge = {
        "style": "single",
        "width": {"value": 1, "unit": "pt"},
        "color": "#1F4E78",
        "space": {"value": 4, "unit": "pt"},
    }
    return Document.from_spec(
        {
            "styles": [
                {
                    "id": "Callout",
                    "name": "Callout",
                    "semantic_role": "custom",
                    "based_on": "Normal",
                    "paragraph_style": {
                        "background_color": "#EAF2F8",
                        "borders": {
                            "top": edge,
                            "right": edge,
                            "bottom": edge,
                            "left": edge,
                        },
                        "spacing_before": {
                            "value": 8,
                            "unit": "pt",
                        },
                        "spacing_after": {
                            "value": 8,
                            "unit": "pt",
                        },
                    },
                    "text_style": {
                        "color": "#17365D",
                        "bold": True,
                    },
                }
            ],
            "content": [
                {
                    "id": "callout",
                    "type": "paragraph",
                    "style_ref": "Callout",
                    "text": "Decision approved",
                    "paragraph_style": {
                        "background_color": "#FFF2CC",
                        "borders": {
                            "bottom": {
                                "style": "double",
                                "width": {
                                    "value": 2,
                                    "unit": "pt",
                                },
                                "color": "#548235",
                                "space": {
                                    "value": 3,
                                    "unit": "pt",
                                },
                            }
                        },
                    },
                }
            ],
        }
    )


def _paragraph_with_text(
    root: ET.Element,
    text: str,
) -> ET.Element:
    return next(
        paragraph
        for paragraph in root.iter(_q("p"))
        if "".join(
            node.text or ""
            for node in paragraph.iter(_q("t"))
        )
        == text
    )


class ParagraphSurfaceTests(unittest.TestCase):
    def test_generation_projection_inheritance_and_html(self) -> None:
        with self.assertRaises(SpecValidationError):
            Document.from_spec(
                {
                    "content": [
                        {
                            "type": "paragraph",
                            "text": "Invalid",
                            "paragraph_style": {
                                "background_color": "blue"
                            },
                        }
                    ]
                }
            )

        document = _surface_document()
        source = document.to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            document_root = parse_xml(
                package.read("word/document.xml")
            )
            styles_root = parse_xml(package.read("word/styles.xml"))

        paragraph = _paragraph_with_text(
            document_root,
            "Decision approved",
        )
        direct_properties = paragraph.find(_q("pPr"))
        assert direct_properties is not None
        direct_shading = direct_properties.find(_q("shd"))
        assert direct_shading is not None
        self.assertEqual(
            direct_shading.get(_q("fill")),
            "FFF2CC",
        )
        direct_borders = direct_properties.find(_q("pBdr"))
        assert direct_borders is not None
        self.assertIsNone(direct_borders.find(_q("top")))
        direct_bottom = direct_borders.find(_q("bottom"))
        assert direct_bottom is not None
        self.assertEqual(direct_bottom.get(_q("val")), "double")
        self.assertEqual(direct_bottom.get(_q("sz")), "16")
        self.assertEqual(direct_bottom.get(_q("space")), "3")

        callout_style = next(
            style
            for style in styles_root.findall(_q("style"))
            if style.get(_q("styleId")) == "Callout"
        )
        style_properties = callout_style.find(_q("pPr"))
        assert style_properties is not None
        style_shading = style_properties.find(_q("shd"))
        style_borders = style_properties.find(_q("pBdr"))
        assert style_shading is not None
        assert style_borders is not None
        self.assertEqual(style_shading.get(_q("fill")), "EAF2F8")
        self.assertEqual(
            {
                child.tag
                for child in style_borders
                if child.tag
                in {
                    _q("top"),
                    _q("right"),
                    _q("bottom"),
                    _q("left"),
                }
            },
            {
                _q("top"),
                _q("right"),
                _q("bottom"),
                _q("left"),
            },
        )

        reopened = Document.from_docx(source)
        self.assertEqual(reopened.to_bytes("docx"), source)
        spec = reopened.to_spec()
        node = next(
            item
            for item in spec["content"]
            if item["id"] == "callout"
        )
        self.assertEqual(
            node["paragraph_style"]["background_color"],
            "#FFF2CC",
        )
        self.assertEqual(
            set(node["paragraph_style"]["borders"]),
            {"bottom"},
        )
        style = next(
            item
            for item in spec["styles"]
            if item["id"] == "Callout"
        )
        self.assertEqual(
            style["paragraph_style"]["background_color"],
            "#EAF2F8",
        )

        html = reopened.to_bytes("html").decode()
        self.assertIn("background-color:#FFF2CC", html)
        self.assertIn(
            "border-top:1pt solid #1F4E78",
            html,
        )
        self.assertIn(
            "border-bottom:2pt double #548235",
            html,
        )
        self.assertIn("padding-top:4pt", html)
        self.assertIn("padding-bottom:3pt", html)

    def test_native_patch_and_clear_preserve_unknown_xml(self) -> None:
        source = _surface_document().to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
            styles = package.read("word/styles.xml")
        paragraph = _paragraph_with_text(root, "Decision approved")
        properties = paragraph.find(_q("pPr"))
        assert properties is not None
        borders = properties.find(_q("pBdr"))
        shading = properties.find(_q("shd"))
        assert borders is not None
        assert shading is not None
        borders.set(f"{{{FUTURE}}}attribute", "keep")
        top = ET.SubElement(
            borders,
            _q("top"),
            {
                _q("val"): "single",
                _q("sz"): "8",
                _q("color"): "C00000",
                f"{{{FUTURE}}}attribute": "keep",
            },
        )
        ET.SubElement(top, f"{{{FUTURE}}}edgeData").text = "keep"
        ET.SubElement(
            borders,
            _q("between"),
            {
                _q("val"): "dashed",
                _q("sz"): "8",
                _q("color"): "7030A0",
            },
        )
        shading.set(f"{{{FUTURE}}}attribute", "keep")
        ET.SubElement(
            shading,
            f"{{{FUTURE}}}shadingData",
        ).text = "keep"
        source = _rewrite_part(
            source,
            "word/document.xml",
            serialize_xml(root),
        )

        imported = Document.from_docx(source)
        self.assertEqual(imported.to_bytes("docx"), source)
        imported_node = next(
            node
            for node in imported.spec.content
            if node.plain_text == "Decision approved"
        )
        result = imported.apply(
            [
                {
                    "op": "paragraph.format",
                    "target": f"#{imported_node.id}",
                    "set": {
                        "background_color": "#D9EAD3",
                        "borders": {
                            "left": {
                                "style": "thick",
                                "width": {
                                    "value": 4,
                                    "unit": "pt",
                                },
                                "color": "#548235",
                                "space": {
                                    "value": 6,
                                    "unit": "pt",
                                },
                            },
                            "bottom": {"style": "none"},
                        },
                    },
                }
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
            self.assertEqual(package.read("word/styles.xml"), styles)
            patched_root = parse_xml(
                package.read("word/document.xml")
            )
        patched = _paragraph_with_text(
            patched_root,
            "Decision approved",
        )
        patched_properties = patched.find(_q("pPr"))
        assert patched_properties is not None
        patched_borders = patched_properties.find(_q("pBdr"))
        patched_shading = patched_properties.find(_q("shd"))
        assert patched_borders is not None
        assert patched_shading is not None
        patched_left = patched_borders.find(_q("left"))
        patched_bottom = patched_borders.find(_q("bottom"))
        patched_top = patched_borders.find(_q("top"))
        assert patched_left is not None
        assert patched_bottom is not None
        assert patched_top is not None
        self.assertEqual(patched_left.get(_q("val")), "thick")
        self.assertEqual(patched_left.get(_q("sz")), "32")
        self.assertEqual(patched_left.get(_q("space")), "6")
        self.assertEqual(patched_bottom.get(_q("val")), "none")
        self.assertIsNone(patched_top.get(_q("val")))
        self.assertEqual(
            patched_top.get(f"{{{FUTURE}}}attribute"),
            "keep",
        )
        self.assertIsNotNone(
            patched_top.find(f"{{{FUTURE}}}edgeData")
        )
        between = patched_borders.find(_q("between"))
        assert between is not None
        self.assertEqual(between.get(_q("val")), "dashed")
        self.assertEqual(
            patched_borders.get(f"{{{FUTURE}}}attribute"),
            "keep",
        )
        self.assertEqual(patched_shading.get(_q("fill")), "D9EAD3")
        self.assertEqual(
            patched_shading.get(f"{{{FUTURE}}}attribute"),
            "keep",
        )
        self.assertIsNotNone(
            patched_shading.find(f"{{{FUTURE}}}shadingData")
        )

        reopened = Document.from_docx(output)
        reopened_node = next(
            node
            for node in reopened.spec.content
            if node.plain_text == "Decision approved"
        )
        assert reopened_node.paragraph_style is not None
        self.assertEqual(
            reopened_node.paragraph_style.background_color,
            "#D9EAD3",
        )
        assert reopened_node.paragraph_style.borders is not None
        self.assertEqual(
            reopened_node.paragraph_style.borders.left.style,
            "thick",
        )
        self.assertEqual(
            reopened_node.paragraph_style.borders.bottom.style,
            "none",
        )

        cleared = result.document.apply(
            [
                {
                    "op": "paragraph.format",
                    "target": f"#{imported_node.id}",
                    "clear": ["background_color", "borders"],
                }
            ]
        )
        self.assertTrue(cleared.success, cleared.model_dump())
        assert cleared.document is not None
        cleared_output = cleared.document.to_bytes("docx")
        with ZipFile(io.BytesIO(cleared_output)) as package:
            cleared_root = parse_xml(
                package.read("word/document.xml")
            )
        cleared_paragraph = _paragraph_with_text(
            cleared_root,
            "Decision approved",
        )
        cleared_properties = cleared_paragraph.find(_q("pPr"))
        assert cleared_properties is not None
        cleared_borders = cleared_properties.find(_q("pBdr"))
        cleared_shading = cleared_properties.find(_q("shd"))
        assert cleared_borders is not None
        assert cleared_shading is not None
        cleared_top = cleared_borders.find(_q("top"))
        assert cleared_top is not None
        self.assertIsNone(cleared_top.get(_q("val")))
        self.assertIsNotNone(
            cleared_top.find(f"{{{FUTURE}}}edgeData")
        )
        self.assertIsNotNone(cleared_borders.find(_q("between")))
        self.assertIsNone(cleared_shading.get(_q("fill")))
        self.assertIsNotNone(
            cleared_shading.find(f"{{{FUTURE}}}shadingData")
        )
        cleared_projection = Document.from_docx(cleared_output)
        cleared_node = next(
            node
            for node in cleared_projection.spec.content
            if node.plain_text == "Decision approved"
        )
        self.assertIsNone(cleared_node.paragraph_style)
        inherited_html = cleared_projection.to_bytes("html").decode()
        self.assertIn("background-color:#EAF2F8", inherited_html)
        self.assertIn(
            "border-bottom:1pt solid #1F4E78",
            inherited_html,
        )

    def test_native_named_style_surface_patch_is_part_scoped(self) -> None:
        source = _surface_document().to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            original_document = package.read("word/document.xml")
        imported = Document.from_docx(source)
        result = imported.apply(
            [
                {
                    "op": "style.format",
                    "target": "@Callout",
                    "paragraph": {
                        "set": {
                            "background_color": "#E4DFEC",
                            "borders": {
                                "top": {
                                    "style": "dotted",
                                    "width": {
                                        "value": 1.5,
                                        "unit": "pt",
                                    },
                                    "color": "#7030A0",
                                }
                            },
                        }
                    },
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        assert result.fidelity is not None
        self.assertEqual(
            result.fidelity.affected_parts,
            [
                "/customXml/aioffice-manifest.xml",
                "/word/styles.xml",
            ],
        )
        output = result.document.to_bytes("docx")
        with ZipFile(io.BytesIO(output)) as package:
            self.assertEqual(
                package.read("word/document.xml"),
                original_document,
            )
        reopened = Document.from_docx(output)
        style = next(
            item
            for item in reopened.spec.styles
            if item.id == "Callout"
        )
        assert style.paragraph_style is not None
        self.assertEqual(
            style.paragraph_style.background_color,
            "#E4DFEC",
        )
        assert style.paragraph_style.borders is not None
        self.assertEqual(
            style.paragraph_style.borders.top.style,
            "dotted",
        )
        self.assertIsNone(style.paragraph_style.borders.bottom)

    def test_theme_pattern_shading_stays_native_and_untouched(self) -> None:
        source = _surface_document().to_bytes("docx")
        with ZipFile(io.BytesIO(source)) as package:
            root = parse_xml(package.read("word/document.xml"))
        paragraph = _paragraph_with_text(root, "Decision approved")
        shading = paragraph.find(f"./{_q('pPr')}/{_q('shd')}")
        assert shading is not None
        shading.set(_q("val"), "pct20")
        shading.set(_q("color"), "C00000")
        shading.set(_q("themeColor"), "accent6")
        shading.set(_q("themeFill"), "accent3")
        shading.set(_q("themeFillTint"), "66")
        bottom = paragraph.find(
            f"./{_q('pPr')}/{_q('pBdr')}/{_q('bottom')}"
        )
        assert bottom is not None
        bottom.set(_q("themeColor"), "accent4")
        bottom.set(_q("themeTint"), "33")
        source = _rewrite_part(
            source,
            "word/document.xml",
            serialize_xml(root),
        )

        imported = Document.from_docx(source)
        self.assertEqual(imported.to_bytes("docx"), source)
        imported_node = next(
            node
            for node in imported.spec.content
            if node.plain_text == "Decision approved"
        )
        self.assertIsNone(imported_node.paragraph_style)
        original_attributes = dict(shading.attrib)
        original_border_attributes = dict(bottom.attrib)

        result = imported.apply(
            [
                {
                    "op": "paragraph.format",
                    "target": f"#{imported_node.id}",
                    "set": {"alignment": "center"},
                }
            ]
        )
        self.assertTrue(result.success, result.model_dump())
        assert result.document is not None
        with ZipFile(
            io.BytesIO(result.document.to_bytes("docx"))
        ) as package:
            patched_root = parse_xml(
                package.read("word/document.xml")
            )
        patched_paragraph = _paragraph_with_text(
            patched_root,
            "Decision approved",
        )
        patched_shading = patched_paragraph.find(
            f"./{_q('pPr')}/{_q('shd')}"
        )
        assert patched_shading is not None
        self.assertEqual(
            dict(patched_shading.attrib),
            original_attributes,
        )
        patched_bottom = patched_paragraph.find(
            f"./{_q('pPr')}/{_q('pBdr')}/{_q('bottom')}"
        )
        assert patched_bottom is not None
        self.assertEqual(
            dict(patched_bottom.attrib),
            original_border_attributes,
        )


if __name__ == "__main__":
    unittest.main()
