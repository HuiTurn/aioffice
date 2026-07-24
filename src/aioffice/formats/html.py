"""Semantic HTML exporter for Document IR."""

from __future__ import annotations

from html import escape

from aioffice.spec.models import (
    AiOfficeDocumentSpec,
    BorderLine,
    BulletList,
    DocumentField,
    DocumentSection,
    Heading,
    HeaderFooterPart,
    ImageBlock,
    Length,
    OrderedList,
    OpaqueBlock,
    PageBreak,
    Paragraph,
    ParagraphStyle,
    Table,
    TableCell,
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
    if style.background_color is not None:
        values.append(
            f"background-color:{style.background_color}"
        )
    if style.borders is not None:
        for side in ("top", "right", "bottom", "left"):
            border = getattr(style.borders, side)
            if border is None:
                continue
            values.append(f"border-{side}:{_border_css(border)}")
            if border.space is not None:
                values.append(
                    f"padding-{side}:{border.space.to_css()}"
                )
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


def _border_css(border: BorderLine) -> str:
    if border.style == "none":
        return "none"
    assert border.width is not None
    css_style = {
        "single": "solid",
        "double": "double",
        "dotted": "dotted",
        "dashed": "dashed",
        "thick": "solid",
    }[border.style]
    color = (
        "currentColor"
        if border.color == "auto"
        else border.color
    )
    return f"{border.width.to_css()} {css_style} {color}"


def _table_css(table: Table) -> str:
    layout = table.layout
    values: list[str] = []
    if layout.preferred_width is not None:
        width = layout.preferred_width
        if width.mode == "auto":
            values.append("width:auto")
        elif width.mode == "percent":
            assert isinstance(width.value, float)
            rendered = f"{width.value:.6f}".rstrip("0").rstrip(".")
            values.append(f"width:{rendered}%")
        else:
            assert isinstance(width.value, Length)
            values.append(f"width:{width.value.to_css()}")
    if layout.algorithm is not None:
        values.append(
            "table-layout:"
            f"{'fixed' if layout.algorithm == 'fixed' else 'auto'}"
        )
    if layout.cell_spacing is not None:
        values.append("border-collapse:separate")
        values.append(
            f"border-spacing:{layout.cell_spacing.to_css()}"
        )
    elif layout.borders is not None or any(
        cell.format.borders is not None
        for row in table.rows
        for cell in row.cells
    ):
        values.append("border-collapse:collapse")
    if layout.borders is not None:
        for side in ("top", "right", "bottom", "left"):
            border = getattr(layout.borders, side)
            if border is not None:
                values.append(
                    f"border-{side}:{_border_css(border)}"
                )
    if layout.alignment == "center":
        values.extend(["margin-left:auto", "margin-right:auto"])
    elif layout.alignment == "right":
        values.extend(["margin-left:auto", "margin-right:0"])
    elif layout.alignment == "left":
        values.extend(["margin-left:0", "margin-right:auto"])
    if layout.indent is not None:
        values.append(f"margin-left:{layout.indent.to_css()}")
    return ";".join(values)


def _table_cell_css(
    table: Table,
    cell: TableCell | None = None,
    *,
    inside_right: bool = True,
    inside_bottom: bool = True,
) -> str:
    layout = table.layout
    values: list[str] = []
    if layout.borders is not None or (
        cell is not None and cell.format.borders is not None
    ):
        values.append("border:none")
    for property_name, field_name in (
        ("padding-top", "cell_margin_top"),
        ("padding-right", "cell_margin_right"),
        ("padding-bottom", "cell_margin_bottom"),
        ("padding-left", "cell_margin_left"),
    ):
        cell_value = (
            getattr(cell.format, field_name.removeprefix("cell_"))
            if cell is not None
            else None
        )
        value = (
            cell_value
            if cell_value is not None
            else getattr(layout, field_name)
        )
        if value is not None:
            values.append(f"{property_name}:{value.to_css()}")
    if cell is not None:
        if cell.format.vertical_alignment is not None:
            values.append(
                f"vertical-align:{cell.format.vertical_alignment}"
            )
        if cell.format.no_wrap is not None:
            values.append(
                "white-space:"
                f"{'nowrap' if cell.format.no_wrap else 'normal'}"
            )
        if cell.format.background_color is not None:
            values.append(
                f"background-color:{cell.format.background_color}"
            )
        if cell.format.borders is not None:
            for side in ("top", "right", "bottom", "left"):
                border = getattr(cell.format.borders, side)
                if border is not None:
                    values.append(
                        f"border-{side}:{_border_css(border)}"
                    )
    if (
        inside_right
        and
        layout.borders is not None
        and layout.borders.inside_vertical is not None
        and (
            cell is None
            or cell.format.borders is None
            or cell.format.borders.right is None
        )
    ):
        values.append(
            "border-right:"
            f"{_border_css(layout.borders.inside_vertical)}"
        )
    if (
        inside_bottom
        and
        layout.borders is not None
        and layout.borders.inside_horizontal is not None
        and (
            cell is None
            or cell.format.borders is None
            or cell.format.borders.bottom is None
        )
    ):
        values.append(
            "border-bottom:"
            f"{_border_css(layout.borders.inside_horizontal)}"
        )
    return ";".join(values)


def _style_attribute(css: str) -> str:
    return f' style="{escape(css, quote=True)}"' if css else ""


def _section_css(section: DocumentSection, *, page_view: bool) -> str:
    if not page_view:
        return ""
    layout = section.layout
    values = [
        "box-sizing:border-box",
        "margin:32px auto",
        "background:white",
        "box-shadow:0 2px 12px #00000014",
    ]
    if layout.page_size is not None:
        width, height = layout.page_size.dimensions_points()
        values.extend((f"width:{width:g}pt", f"min-height:{height:g}pt"))
    if all(
        value is not None
        for value in (
            layout.margin_top,
            layout.margin_right,
            layout.margin_bottom,
            layout.margin_left,
        )
    ):
        assert layout.margin_top is not None
        assert layout.margin_right is not None
        assert layout.margin_bottom is not None
        assert layout.margin_left is not None
        left = layout.margin_left.to_points() + (
            layout.gutter.to_points() if layout.gutter is not None else 0
        )
        values.append(
            "padding:"
            f"{layout.margin_top.to_points():g}pt "
            f"{layout.margin_right.to_points():g}pt "
            f"{layout.margin_bottom.to_points():g}pt "
            f"{left:g}pt"
        )
    if layout.columns is not None and layout.columns.count > 1:
        values.append(f"column-count:{layout.columns.count}")
        values.append(f"column-gap:{layout.columns.spacing.to_points():g}pt")
        if layout.columns.separator:
            values.append("column-rule:1px solid #b8c2cc")
    if layout.start_type in {"next_page", "even_page", "odd_page"}:
        values.append(
            "break-before:"
            + {
                "next_page": "page",
                "even_page": "left",
                "odd_page": "right",
            }[layout.start_type]
        )
    elif layout.start_type == "next_column":
        values.append("break-before:column")
    return ";".join(values)


def _section_open(
    section: DocumentSection,
    *,
    page_view: bool,
) -> str:
    return (
        f'<section class="document-section" '
        f'data-aioffice-section="{escape(section.id, quote=True)}"'
        f"{_style_attribute(_section_css(section, page_view=page_view))}>"
    )


def _effective_header_footers(
    spec: AiOfficeDocumentSpec,
) -> dict[str, dict[str, HeaderFooterPart]]:
    parts = {part.id: part for part in spec.header_footers}
    inherited: dict[str, HeaderFooterPart] = {}
    result: dict[str, dict[str, HeaderFooterPart]] = {}
    for section in spec.sections:
        if section.header_footer is not None:
            for field_name, part_id in section.header_footer.model_dump(
                mode="python",
                exclude_none=True,
            ).items():
                part = parts.get(part_id)
                if part is not None:
                    inherited[field_name] = part
        result[section.id] = dict(inherited)
    return result


def _header_footer_html(
    spec: AiOfficeDocumentSpec,
    part: HeaderFooterPart | None,
    *,
    kind: str,
    variant: str,
) -> list[str]:
    if part is None:
        return []
    lines = [
        f'<{kind} class="document-{kind}" '
        f'data-aioffice-header-footer="{escape(part.id, quote=True)}" '
        f'data-aioffice-variant="{variant}">'
    ]
    for block in part.content:
        block_id = escape(block.id, quote=True)
        if isinstance(block, Paragraph):
            resolved_paragraph, resolved_text = resolve_node_styles(
                spec,
                style_ref=block.style_ref,
                paragraph_style=block.paragraph_style,
                text_style=block.text_style,
            )
            value = (
                (
                    f"<span"
                    f"{_style_attribute(_text_css(resolved_text))}>"
                    f"{escape(block.text)}</span>"
                    if resolved_text is not None
                    else escape(block.text)
                )
                if block.text is not None
                else "".join(
                    _span_html(inline, resolved_text)
                    for inline in block.content
                )
            )
            style_data = (
                f' data-aioffice-style="'
                f'{escape(block.style_ref, quote=True)}"'
                if block.style_ref is not None
                else ""
            )
            lines.append(
                f'<p id="{block_id}"{style_data}'
                f"{_style_attribute(_paragraph_css(resolved_paragraph))}>"
                f"{value}</p>"
            )
        elif isinstance(block, ImageBlock):
            asset = next(
                (
                    candidate
                    for candidate in spec.assets
                    if candidate.id == block.asset_id
                ),
                None,
            )
            label = (
                block.alt_text
                or block.title
                or block.name
                or f"Native {kind} image"
            )
            media_type = (
                asset.media_type if asset is not None else ""
            )
            figure_style = _style_attribute(
                f"width:{block.width.to_css()}"
            )
            placeholder_style = _style_attribute(
                f"width:{block.width.to_css()};"
                f"height:{block.height.to_css()}"
            )
            lines.append(
                f'<figure id="{block_id}" class="native-image" '
                f'data-aioffice-asset-id="'
                f'{escape(block.asset_id, quote=True)}" '
                f'data-aioffice-media-type="'
                f'{escape(media_type, quote=True)}" '
                f"{figure_style}>"
                '<div class="native-image-placeholder" role="img" '
                f'aria-label="{escape(label, quote=True)}" '
                f"{placeholder_style}>"
                "Native image — use native rendering or extract the asset"
                "</div>"
                f"<figcaption>{escape(label)}</figcaption></figure>"
            )
        elif isinstance(block, OpaqueBlock):
            lines.append(
                f'<div id="{block_id}" class="opaque-header-footer">'
                f"{escape(block.summary)}</div>"
            )
    lines.append(f"</{kind}>")
    return lines


def _span_html(
    span: TextSpan | DocumentField,
    inherited_style: TextStyle | None = None,
) -> str:
    if isinstance(span, DocumentField):
        value = escape(span.display_text)
        css = _text_css(_merge_text_style(inherited_style, span.style))
        editable = "true" if span.editable else "false"
        title = (
            span.instruction or "native field"
            if span.kind == "native"
            else span.kind.replace("_", " ")
        )
        return (
            '<span class="document-field" '
            f'id="{escape(span.id, quote=True)}" '
            f'data-aioffice-field-kind="{escape(span.kind, quote=True)}" '
            f'data-aioffice-field-editable="{editable}" '
            f'title="{escape(title, quote=True)}"'
            f"{_style_attribute(css)}>{value}</span>"
        )
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
        "article{box-sizing:border-box;max-width:none;margin:0;padding:0}"
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
        (
            ".document-header{border-bottom:1px solid #d8dee5;margin-bottom:18pt;"
            "padding-bottom:6pt;color:#59636e}"
        ),
        (
            ".document-footer{border-top:1px solid #d8dee5;margin-top:18pt;"
            "padding-top:6pt;color:#59636e}"
        ),
        ".opaque-header-footer{font-style:italic;color:#7a828a}",
        (
            ".opaque-native-block{border:1px dashed #a8b1ba;background:#f7f8f9;"
            "color:#697580;padding:10px 12px;font-style:italic}"
        ),
        (
            ".native-image{box-sizing:border-box;max-width:100%;margin-left:0;"
            "margin-right:0}.native-image-placeholder{box-sizing:border-box;"
            "display:flex;align-items:center;justify-content:center;"
            "min-height:72px;max-width:100%;border:1px dashed #9aa6b2;"
            "background:#f3f6f8;color:#59636e;text-align:center;padding:12px}"
        ),
        (
            ".native-image figcaption{font-size:.85em;color:#697580;"
            "margin-top:4px}"
        ),
        (
            ".document-field{display:inline-block;min-width:.7em;"
            "border-bottom:1px dotted #74808c}"
        ),
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

    first_section = spec.sections[0]
    active_section = first_section
    effective_header_footers = _effective_header_footers(spec)
    opened_section_ids = {first_section.id}
    section_starts = {
        section.start_at: section
        for section in spec.sections[1:]
        if section.start_at is not None
    }
    lines.append(_section_open(first_section, page_view=page_view))
    active_parts = effective_header_footers[first_section.id]
    header_variant = (
        "first"
        if first_section.layout.different_first_page is True
        and "header_first" in active_parts
        else "default"
    )
    lines.extend(
        _header_footer_html(
            spec,
            active_parts.get(f"header_{header_variant}"),
            kind="header",
            variant=header_variant,
        )
    )
    for block in spec.content:
        starting_section = section_starts.get(block.id)
        if starting_section is not None:
            footer_variant = (
                "first"
                if active_section.layout.different_first_page is True
                and "footer_first" in active_parts
                else "default"
            )
            lines.extend(
                _header_footer_html(
                    spec,
                    active_parts.get(f"footer_{footer_variant}"),
                    kind="footer",
                    variant=footer_variant,
                )
            )
            lines.append("</section>")
            lines.append(_section_open(starting_section, page_view=page_view))
            active_section = starting_section
            active_parts = effective_header_footers[starting_section.id]
            header_variant = (
                "first"
                if starting_section.layout.different_first_page is True
                and "header_first" in active_parts
                else "default"
            )
            lines.extend(
                _header_footer_html(
                    spec,
                    active_parts.get(f"header_{header_variant}"),
                    kind="header",
                    variant=header_variant,
                )
            )
            opened_section_ids.add(starting_section.id)
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
            table_style = (
                f' data-aioffice-table-style="'
                f'{escape(block.layout.style_ref, quote=True)}"'
                if block.layout.style_ref is not None
                else ""
            )
            lines.append(
                f'<table id="{block_id}"{table_style}'
                f"{_style_attribute(_table_css(block))}>"
            )
            lines.append("<colgroup>")
            for column in block.columns:
                column_css = (
                    f"width:{column.width.to_css()}"
                    if column.width is not None
                    else ""
                )
                lines.append(
                    "<col "
                    f'data-column-id="{escape(column.id, quote=True)}" '
                    f'data-column-key="{escape(column.key, quote=True)}"'
                    f"{_style_attribute(column_css)}>"
                )
            lines.append("</colgroup>")
            lines.append("<thead><tr>")
            for index, column in enumerate(block.columns):
                header_cell_css = _table_cell_css(
                    block,
                    inside_right=index < len(block.columns) - 1,
                )
                lines.append(
                    f'<th data-column-id="{escape(column.id, quote=True)}"'
                    f"{_style_attribute(header_cell_css)}>"
                    f"{escape(column.title)}</th>"
                )
            lines.append("</tr></thead>")
            lines.append("<tbody>")
            for row_index, row in enumerate(block.rows):
                row_css: list[str] = []
                if row.allow_break_across_pages is False:
                    row_css.append("break-inside:avoid")
                if row.height is not None:
                    row_css.append(
                        (
                            "height:"
                            if row.height_rule == "exact"
                            else "min-height:"
                        )
                        + row.height.to_css()
                    )
                lines.append(
                    f'<tr data-row-id="{escape(row.id, quote=True)}"'
                    f"{_style_attribute(';'.join(row_css))}>"
                )
                column_by_key = {
                    column.key: column
                    for column in block.columns
                }
                column_positions = {
                    column.key: index
                    for index, column in enumerate(block.columns)
                }
                column_count = len(block.columns)
                for cell in sorted(
                    row.cells,
                    key=lambda item: column_positions.get(
                        item.column_key,
                        column_count,
                    ),
                ):
                    column = column_by_key.get(cell.column_key)
                    if column is None:
                        continue
                    span_attributes = (
                        (
                            f' colspan="{cell.column_span}"'
                            if cell.column_span > 1
                            else ""
                        )
                        + (
                            f' rowspan="{cell.row_span}"'
                            if cell.row_span > 1
                            else ""
                        )
                    )
                    value = escape(cell.plain_text)
                    if cell.content:
                        rendered_paragraphs: list[str] = []
                        for paragraph in cell.content:
                            resolved_paragraph, resolved_text = (
                                resolve_node_styles(
                                    spec,
                                    style_ref=paragraph.style_ref,
                                    paragraph_style=(
                                        paragraph.paragraph_style
                                    ),
                                    text_style=paragraph.text_style,
                                )
                            )
                            paragraph_value = (
                                (
                                    f"<span"
                                    f"{_style_attribute(_text_css(resolved_text))}>"
                                    f"{escape(paragraph.text)}</span>"
                                    if resolved_text is not None
                                    else escape(paragraph.text)
                                )
                                if paragraph.text is not None
                                else "".join(
                                    _span_html(span, resolved_text)
                                    for span in paragraph.content
                                )
                            )
                            rendered_paragraphs.append(
                                f'<p id="{escape(paragraph.id, quote=True)}"'
                                f"{_style_attribute(_paragraph_css(resolved_paragraph))}>"
                                f"{paragraph_value}</p>"
                            )
                        value = "".join(rendered_paragraphs)
                    data_cell_css = _table_cell_css(
                        block,
                        cell,
                        inside_right=(
                            column_positions[column.key]
                            + cell.column_span
                            < column_count
                        ),
                        inside_bottom=(
                            row_index + cell.row_span
                            < len(block.rows)
                        ),
                    )
                    lines.append(
                        f'<td data-cell-id="'
                        f'{escape(cell.id, quote=True)}" '
                        f'data-column-id="{escape(column.id, quote=True)}"'
                        f"{span_attributes}"
                        f"{_style_attribute(data_cell_css)}>"
                        f"{value}</td>"
                    )
                lines.append("</tr>")
            lines.append("</tbody></table>")
        elif isinstance(block, ImageBlock):
            asset = next(
                (
                    candidate
                    for candidate in spec.assets
                    if candidate.id == block.asset_id
                ),
                None,
            )
            resolved_paragraph, _ = resolve_node_styles(
                spec,
                style_ref=block.style_ref,
                paragraph_style=block.paragraph_style,
                text_style=None,
            )
            label = (
                block.alt_text
                or block.title
                or block.name
                or "Native document image"
            )
            media_type = asset.media_type if asset is not None else ""
            figure_css = ";".join(
                value
                for value in (
                    _paragraph_css(resolved_paragraph),
                    f"width:{block.width.to_css()}",
                )
                if value
            )
            placeholder_css = (
                f"width:{block.width.to_css()};"
                f"height:{block.height.to_css()}"
            )
            lines.append(
                f'<figure id="{block_id}" class="native-image" '
                f'data-aioffice-asset-id="'
                f'{escape(block.asset_id, quote=True)}" '
                f'data-aioffice-media-type="'
                f'{escape(media_type, quote=True)}"'
                f"{_style_attribute(figure_css)}>"
            )
            lines.append(
                '<div class="native-image-placeholder" role="img" '
                f'aria-label="{escape(label, quote=True)}" '
                f"{_style_attribute(placeholder_css)}>"
                "Native image — use native rendering or extract the asset"
                "</div>"
            )
            lines.append(
                f"<figcaption>{escape(label)}</figcaption></figure>"
            )
        elif isinstance(block, OpaqueBlock):
            lines.append(
                f'<div id="{block_id}" class="opaque-native-block" '
                'data-aioffice-native-opaque="true">'
                f"{escape(block.summary)}</div>"
            )
        elif isinstance(block, PageBreak):
            lines.append(f'<hr id="{block_id}" class="page-break">')

    footer_variant = (
        "first"
        if active_section.layout.different_first_page is True
        and "footer_first" in active_parts
        else "default"
    )
    lines.extend(
        _header_footer_html(
            spec,
            active_parts.get(f"footer_{footer_variant}"),
            kind="footer",
            variant=footer_variant,
        )
    )
    lines.append("</section>")
    for section in spec.sections:
        if section.id not in opened_section_ids:
            lines.extend(
                [
                    _section_open(section, page_view=page_view),
                    "</section>",
                ]
            )
    lines.extend(["</article>", "</body>", "</html>", ""])
    return "\n".join(lines)
