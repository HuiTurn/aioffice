from __future__ import annotations

import io
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from aioffice.documents import Document, DocumentBuilder
from aioffice.formats.docx import compile_docx
from aioffice.formats.html import export_html
from aioffice.formats.markdown import export_markdown, import_markdown
from aioffice.native import MANIFEST_PART_URI, MANIFEST_RELATIONSHIP_TYPE


class FormatTests(unittest.TestCase):
    def test_markdown_table_import_and_export(self) -> None:
        source = (
            "# Report\n\n"
            "| Risk | Level |\n"
            "| --- | --- |\n"
            "| Scope | Medium |\n\n"
            "- Validate\n"
            "- Export\n"
        )
        spec = import_markdown(source)
        self.assertEqual([node.type for node in spec.content], ["heading", "table", "bullet_list"])
        output = export_markdown(spec)
        self.assertIn("| Risk | Level |", output)
        self.assertIn("- Validate", output)

    def test_html_escapes_user_content_and_keeps_semantics(self) -> None:
        document = DocumentBuilder(title="<Report>").heading("<Risk>", id="risk").build()
        output = export_html(document.spec)
        self.assertIn("<article", output)
        self.assertIn("<h1", output)
        self.assertIn("&lt;Risk&gt;", output)
        self.assertNotIn("<Risk>", output)

    def test_docx_package_is_complete_parseable_and_deterministic(self) -> None:
        document = (
            DocumentBuilder(title="Project", author="AiOffice")
            .heading("Project", id="title")
            .rich_paragraph(
                [
                    {"text": "See "},
                    {
                        "text": "documentation",
                        "marks": ["link"],
                        "href": "https://example.com",
                    },
                ],
                id="intro",
            )
            .bullet_list(["One", "Two"])
            .table(
                [{"key": "name", "title": "Name"}],
                [{"name": "AiOffice"}],
                id="table",
            )
            .build()
        )
        first = compile_docx(document.spec)
        second = compile_docx(document.spec)
        self.assertEqual(first, second)

        required = {
            "[Content_Types].xml",
            "_rels/.rels",
            MANIFEST_PART_URI.lstrip("/"),
            "word/document.xml",
            "word/_rels/document.xml.rels",
            "word/styles.xml",
            "word/numbering.xml",
            "docProps/core.xml",
            "docProps/app.xml",
        }
        with ZipFile(io.BytesIO(first)) as archive:
            self.assertTrue(required.issubset(archive.namelist()))
            for name in archive.namelist():
                if name.endswith(".xml") or name.endswith(".rels"):
                    ET.fromstring(archive.read(name))
            document_xml = archive.read("word/document.xml").decode("utf-8")
            relationships = archive.read("word/_rels/document.xml.rels").decode("utf-8")
            self.assertIn("AiOffice", document_xml)
            self.assertIn("https://example.com", relationships)
            self.assertIn(
                '<Relationships xmlns="http://schemas.openxmlformats.org/'
                'package/2006/relationships">',
                relationships,
            )
            self.assertNotIn("ns0:Relationships", relationships)
            root_relationships = archive.read("_rels/.rels").decode("utf-8")
            self.assertIn(MANIFEST_RELATIONSHIP_TYPE, root_relationships)
            content_types = archive.read("[Content_Types].xml").decode("utf-8")
            self.assertIn(
                '<Types xmlns="http://schemas.openxmlformats.org/package/'
                '2006/content-types">',
                content_types,
            )
            self.assertNotIn("ns0:Types", content_types)

    def test_generated_docx_restores_embedded_semantic_identity(self) -> None:
        document = (
            DocumentBuilder(title="Identity")
            .heading("Identity", id="stable_heading")
            .paragraph("Body", id="stable_paragraph")
            .bullet_list(["One", "Two"], id="stable_list")
            .page_break(id="stable_break")
            .table(
                [{"key": "name", "title": "Name"}],
                [{"name": "AiOffice"}],
                id="stable_table",
            )
            .build()
        )
        reopened = Document.from_docx(compile_docx(document.spec))
        self.assertEqual(reopened.id, document.id)
        self.assertEqual(reopened.revision, document.revision)
        self.assertEqual(
            [(node["id"], node["type"]) for node in reopened.to_spec()["content"]],
            [
                ("stable_heading", "heading"),
                ("stable_paragraph", "paragraph"),
                ("stable_list", "bullet_list"),
                ("stable_break", "page_break"),
                ("stable_table", "table"),
            ],
        )
        self.assertEqual(reopened.import_diagnostics, [])
        self.assertEqual(
            reopened.capabilities()["identity"]["source"],
            "embedded",
        )
        self.assertTrue(
            reopened.capabilities()["identity"]["safe_to_commit"],
        )


if __name__ == "__main__":
    unittest.main()
