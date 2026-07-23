from __future__ import annotations

import io
import unittest
from xml.etree import ElementTree as ET
from zipfile import ZipFile

from aioffice.documents import Document, DocumentBuilder
from aioffice.formats.docx import compile_docx
from aioffice.formats.html import export_html
from aioffice.formats.markdown import export_markdown, import_markdown


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


if __name__ == "__main__":
    unittest.main()
