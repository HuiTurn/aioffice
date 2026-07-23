"""Semantic HTML exporter for Document IR."""

from __future__ import annotations

from html import escape

from aioffice.spec.models import (
    AiOfficeDocumentSpec,
    BulletList,
    Heading,
    Length,
    OrderedList,
    PageBreak,
    Paragraph,
    ParagraphStyle,
    Table,
    TextSpan,
    TextStyle,
)
from aioffice.styles import resolve_node_styles
from aioffice.themes import get_theme


def _merge_text_style(
    base: TextStyle | None,
    override: TextStyle | None,
) -> TextStyle | None:
    if base is None:
        return override
    if override is None:
        return base
    return TextStyle.model_validate(
        {
            **base.model_dump(mode="json", exclude_none=True),
            **override.model_dump(mode="json", exclude_none=True),
        }
    )


def _paragraph_css(style: ParagraphStyle | None) -> str:
    if style is None:
        return ""
    values: list[str] = []
    if style.alignment is not None:
        values.append(f"text-align:{style.alignment}")
    if style.spacing_before is not None:
        values.append(f"margin-top:{style.spacing_before.to_css()}")
    if style.spacing_after is not None:
        values.append(f"margin-bottom:{style.spacing_after.to_css()}")
    if style.line_spacing is not None:
        line_spacing = style.line_spacing
        if line_spacing.rule == "multiple":
            values.append(f"line-height:{line_spacing.value}")
        else:
            assert isinstance(line_spacing.value, Length)
            values.append(f"line-height:{line_spacing.value.to_css()}")
    if style.indent_left is not None:
        values.append(f"margin-left:{style.indent_left.to_css()}")
    if style.indent_right is not None:
        values.append(f"margin-right:{style.indent_right.to_css()}")
    if style.first_line_indent is not None:
        values.append(f"text-indent:{style.first_line_indent.to_css()}")
    if style.hanging_indent is not None:
        values.append(f"text-indent:-{style.hanging_indent.to_css()}")
        values.append(f"padding-left:{style.hanging_indent.to_css()}")
    if style.keep_with_next:
        values.append("break-after:avoid")
    if style.keep_together:
        values.append("break-inside:avoid")
    if style.page_break_before:
        values.append("break-before:page")
    if style.widow_control is not None:
        values.append(f"widows:{2 if style.widow_control else 1}")
        values.append(f"orphans:{2 if style.widow_control else 1}")
    return ";".join(values)


def _text_css(style: TextStyle | None) -> str:
    if style is None:
        return ""
    values: list[str] = []
    if style.font_family is not None:
        family = style.font_family.replace("\\", "\\\\").replace('"', '\\"')
        values.append(f'font-family:"{family}"')
    if style.font_size is not None:
        values.append(f"font-size:{style.font_size.to_css()}")
    if style.color is not None:
        values.append(f"color:{style.color}")
    if style.background_color is not None:
        values.append(f"background-color:{style.background_color}")
    if style.bold is not None:
        values.append(f"font-weight:{700 if style.bold else 400}")
    if style.italic is not None:
        values.append(f"font-style:{'italic' if style.italic else 'normal'}")
    decorations: list[str] = []
    if style.underline:
        decorations.append("underline")
    if style.strike:
        decorations.append("line-through")
    if style.underline is False and style.strike is False:
        decorations.append("none")
    if decorations:
        values.append(f"text-decoration:{' '.join(decorations)}")
    if style.small_caps is not None:
        values.append(f"font-variant-caps:{'small-caps' if style.small_caps else 'normal'}")
    if style.all_caps is not None:
        values.append(f"text-transform:{'uppercase' if style.all_caps else 'none'}")
    if style.letter_spacing is not None:
        values.append(f"letter-spacing:{style.letter_spacing.to_css()}")
    if style.baseline in {"superscript", "subscript"}:
        values.append(f"vertical-align:{'super' if style.baseline == 'superscript' else 'sub'}")
    elif style.baseline == "normal":
        values.append("vertical-align:baseline")
    return ";".join(values)


def _style_attribute(css: str) -> str:
    return f' style="{escape(css, quote=True)}"' if css else ""


def _span_html(span: TextSpan, inherited_style: TextStyle | None = None) -> str:
    value = escape(span.text)
    css = _text_css(_merge_text_style(inherited_style, span.style))
    if css:
        value = f"<span{_style_attribute(css)}>{value}</span>"
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


def export_html(
    spec: AiOfficeDocumentSpec,
    *,
    page_view: bool = True,
    include_document_metadata: bool = True,
    locale: str | None = None,
) -> str:
    """Render a document as standalone, semantic HTML."""

    theme = get_theme(spec.theme.ref) or get_theme("business-clean") or {}
    tokens = theme.get("tokens", {})
    primary = tokens.get("color.primary", "#1F4E78")
    text_color = tokens.get("color.text", "#222222")
    body_font = tokens.get("font.body.latin", "sans-serif")
    heading_font = tokens.get("font.heading.latin", body_font)
    title = escape(
        spec.metadata.title
        if include_document_metadata and spec.metadata.title
        else "AiOffice Document"
    )
    language = escape(locale or spec.metadata.language or "en", quote=True)
    body_css = (
        "body{color:var(--aio-text);font-family:var(--aio-body);"
        "line-height:1.6;margin:0;background:#f5f7f9}"
        if page_view
        else (
            "body{color:var(--aio-text);font-family:var(--aio-body);"
            "line-height:1.6;margin:0;background:white}"
        )
    )
    article_css = (
        "article{box-sizing:border-box;max-width:840px;margin:32px auto;"
        "padding:64px;background:white;box-shadow:0 2px 12px #00000014}"
        if page_view
        else (
            "article{box-sizing:border-box;max-width:none;margin:0;"
            "padding:0;background:white;box-shadow:none}"
        )
    )

    lines = [
        "<!doctype html>",
        f'<html lang="{language}">',
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
        body_css,
        article_css,
        "h1,h2,h3,h4,h5,h6{font-family:var(--aio-heading);color:var(--aio-primary)}",
        "table{border-collapse:collapse;width:100%;margin:1em 0}",
        "th,td{border:1px solid #c8d0d8;padding:.5em .65em;text-align:left}",
        "th{background:#eef3f7}",
        ".underline{text-decoration:underline}",
        ".page-break{break-after:page;border:0;border-top:1px dashed #bbb;margin:2em 0}",
        "</style>",
    ]
    if include_document_metadata:
        if spec.metadata.author:
            lines.append(
                f'<meta name="author" content="{escape(spec.metadata.author, quote=True)}">'
            )
        if spec.metadata.subject:
            lines.append(
                f'<meta name="description" content="{escape(spec.metadata.subject, quote=True)}">'
            )
        if spec.metadata.keywords:
            lines.append(
                f'<meta name="keywords" content="'
                f'{escape(", ".join(spec.metadata.keywords), quote=True)}">'
            )
    lines.extend(
        [
            "</head>",
            "<body>",
            f'<article data-aioffice-id="{escape(spec.artifact.id, quote=True)}">',
        ]
    )

    for block in spec.content:
        block_id = escape(block.id, quote=True)
        if isinstance(block, Heading):
            named_style_ref = block.style_ref or f"Heading{block.level}"
            resolved_paragraph, resolved_text = resolve_node_styles(
                spec,
                style_ref=named_style_ref,
                paragraph_style=block.paragraph_style,
                text_style=block.text_style,
            )
            paragraph_css = _paragraph_css(resolved_paragraph)
            if block.text is not None:
                heading_value = (
                    f"<span{_style_attribute(_text_css(resolved_text))}>"
                    f"{escape(block.text)}</span>"
                    if resolved_text is not None
                    else escape(block.text)
                )
            else:
                heading_value = "".join(
                    _span_html(span, resolved_text) for span in block.content
                )
            style_data = (
                f' data-aioffice-style="{escape(named_style_ref, quote=True)}"'
            )
            lines.append(
                f'<h{block.level} id="{block_id}"{style_data}'
                f"{_style_attribute(paragraph_css)}>"
                f"{heading_value}</h{block.level}>"
            )
        elif isinstance(block, Paragraph):
            resolved_paragraph, resolved_text = resolve_node_styles(
                spec,
                style_ref=block.style_ref,
                paragraph_style=block.paragraph_style,
                text_style=block.text_style,
            )
            paragraph_css = _paragraph_css(resolved_paragraph)
            value = (
                (
                    f"<span{_style_attribute(_text_css(resolved_text))}>"
                    f"{escape(block.text)}</span>"
                    if resolved_text is not None
                    else escape(block.text)
                )
                if block.text is not None
                else "".join(_span_html(span, resolved_text) for span in block.content)
            )
            style_data = (
                f' data-aioffice-style="{escape(block.style_ref, quote=True)}"'
                if block.style_ref is not None
                else ""
            )
            lines.append(
                f'<p id="{block_id}"{style_data}'
                f"{_style_attribute(paragraph_css)}>{value}</p>"
            )
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
