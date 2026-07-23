"""Semantic HTML exporter for Document IR."""

from __future__ import annotations

from html import escape

from aioffice.spec.models import (
    AiOfficeDocumentSpec,
    BulletList,
    Heading,
    OrderedList,
    PageBreak,
    Paragraph,
    Table,
    TextSpan,
)
from aioffice.themes import get_theme


def _span_html(span: TextSpan) -> str:
    value = escape(span.text)
    wrappers = {
        "strong": ("<strong>", "</strong>"),
        "emphasis": ("<em>", "</em>"),
        "underline": ('<span class="underline">', "</span>"),
        "strike": ("<s>", "</s>"),
        "code": ("<code>", "</code>"),
        "subscript": ("<sub>", "</sub>"),
        "superscript": ("<sup>", "</sup>"),
        "highlight": ("<mark>", "</mark>"),
    }
    for mark in span.marks:
        if mark == "link" and span.href:
            value = f'<a href="{escape(span.href, quote=True)}">{value}</a>'
        elif mark in wrappers:
            before, after = wrappers[mark]
            value = f"{before}{value}{after}"
    return value


def export_html(spec: AiOfficeDocumentSpec) -> str:
    """Render a document as standalone, semantic HTML."""

    theme = get_theme(spec.theme.ref) or get_theme("business-clean") or {}
    tokens = theme.get("tokens", {})
    primary = tokens.get("color.primary", "#1F4E78")
    text_color = tokens.get("color.text", "#222222")
    body_font = tokens.get("font.body.latin", "sans-serif")
    heading_font = tokens.get("font.heading.latin", body_font)
    title = escape(spec.metadata.title or "AiOffice Document")

    lines = [
        "<!doctype html>",
        '<html lang="en">',
        "<head>",
        '<meta charset="utf-8">',
        '<meta name="viewport" content="width=device-width, initial-scale=1">',
        f"<title>{title}</title>",
        "<style>",
        (
            ":root{"
            f"--aio-primary:{primary};--aio-text:{text_color};"
            f"--aio-body:{body_font!r};--aio-heading:{heading_font!r}"
            "}"
        ),
        (
            "body{color:var(--aio-text);font-family:var(--aio-body);"
            "line-height:1.6;margin:0;background:#f5f7f9}"
        ),
        (
            "article{box-sizing:border-box;max-width:840px;margin:32px auto;"
            "padding:64px;background:white;box-shadow:0 2px 12px #00000014}"
        ),
        "h1,h2,h3,h4,h5,h6{font-family:var(--aio-heading);color:var(--aio-primary)}",
        "table{border-collapse:collapse;width:100%;margin:1em 0}",
        "th,td{border:1px solid #c8d0d8;padding:.5em .65em;text-align:left}",
        "th{background:#eef3f7}",
        ".underline{text-decoration:underline}",
        ".page-break{break-after:page;border:0;border-top:1px dashed #bbb;margin:2em 0}",
        "</style>",
        "</head>",
        "<body>",
        f'<article data-aioffice-id="{escape(spec.artifact.id, quote=True)}">',
    ]

    for block in spec.content:
        block_id = escape(block.id, quote=True)
        if isinstance(block, Heading):
            lines.append(
                f'<h{block.level} id="{block_id}">{escape(block.text)}</h{block.level}>'
            )
        elif isinstance(block, Paragraph):
            value = (
                escape(block.text)
                if block.text is not None
                else "".join(_span_html(span) for span in block.content)
            )
            lines.append(f'<p id="{block_id}">{value}</p>')
        elif isinstance(block, (BulletList, OrderedList)):
            tag = "ul" if isinstance(block, BulletList) else "ol"
            lines.append(f'<{tag} id="{block_id}">')
            lines.extend(f"<li>{escape(item)}</li>" for item in block.items)
            lines.append(f"</{tag}>")
        elif isinstance(block, Table):
            lines.append(f'<table id="{block_id}">')
            lines.append("<thead><tr>")
            lines.extend(f"<th>{escape(column.title)}</th>" for column in block.columns)
            lines.append("</tr></thead>")
            lines.append("<tbody>")
            for row in block.rows:
                lines.append(f'<tr data-row-id="{escape(row.id, quote=True)}">')
                lines.extend(
                    f"<td>{escape(str(row.values.get(column.key, '')))}</td>"
                    for column in block.columns
                )
                lines.append("</tr>")
            lines.append("</tbody></table>")
        elif isinstance(block, PageBreak):
            lines.append(f'<hr id="{block_id}" class="page-break">')

    lines.extend(["</article>", "</body>", "</html>", ""])
    return "\n".join(lines)
