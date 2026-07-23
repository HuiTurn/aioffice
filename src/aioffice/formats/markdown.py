"""CommonMark-oriented Markdown import and export."""

from __future__ import annotations

import re
from typing import Any

from aioffice.spec.models import (
    AiOfficeDocumentSpec,
    BulletList,
    DocumentField,
    Heading,
    ImageBlock,
    OpaqueBlock,
    OrderedList,
    PageBreak,
    Paragraph,
    Table,
)

_HEADING_RE = re.compile(r"^(#{1,6})\s+(.+?)\s*$")
_BULLET_RE = re.compile(r"^\s*[-+*]\s+(.+?)\s*$")
_ORDERED_RE = re.compile(r"^\s*\d+[.)]\s+(.+?)\s*$")
_TABLE_SEPARATOR_RE = re.compile(r"^\s*\|?(?:\s*:?-+:?\s*\|)+\s*:?-+:?\s*\|?\s*$")


def _escape_cell(value: Any) -> str:
    return str(value).replace("\\", "\\\\").replace("|", "\\|").replace("\n", "<br>")


def _escape_link_text(value: str) -> str:
    return (
        value.replace("\\", "\\\\")
        .replace("[", "\\[")
        .replace("]", "\\]")
        .replace("\n", " ")
    )


def _rich_text_markdown(text: str | None, content: list[Any]) -> str:
    if text is not None:
        return text
    rendered: list[str] = []
    for span in content:
        if isinstance(span, DocumentField):
            rendered.append(
                (
                    "{{native_field}}"
                    if span.kind == "native"
                    else "{{" + span.kind + "}}"
                )
            )
            continue
        value = span.text
        if "code" in span.marks:
            value = f"`{value}`"
        if "strong" in span.marks:
            value = f"**{value}**"
        if "emphasis" in span.marks:
            value = f"*{value}*"
        if "strike" in span.marks:
            value = f"~~{value}~~"
        if "link" in span.marks and span.href:
            value = f"[{value}]({span.href})"
        rendered.append(value)
    return "".join(rendered)


def export_markdown(spec: AiOfficeDocumentSpec) -> str:
    lines: list[str] = []
    for block in spec.content:
        if isinstance(block, Heading):
            lines.append(f"{'#' * block.level} {_rich_text_markdown(block.text, block.content)}")
        elif isinstance(block, Paragraph):
            lines.append(_rich_text_markdown(block.text, block.content))
        elif isinstance(block, BulletList):
            lines.extend(f"- {item}" for item in block.items)
        elif isinstance(block, OrderedList):
            lines.extend(f"{index}. {item}" for index, item in enumerate(block.items, start=1))
        elif isinstance(block, Table):
            titles = [_escape_cell(column.title) for column in block.columns]
            lines.append("| " + " | ".join(titles) + " |")
            lines.append("| " + " | ".join("---" for _ in titles) + " |")
            for row in block.rows:
                values_by_key = {
                    cell.column_key: cell.plain_text
                    for cell in row.cells
                }
                values = [
                    _escape_cell(values_by_key.get(column.key, ""))
                    for column in block.columns
                ]
                lines.append("| " + " | ".join(values) + " |")
        elif isinstance(block, ImageBlock):
            alt_text = (
                block.alt_text
                or block.title
                or block.name
                or "Native document image"
            )
            lines.append(
                f"![{_escape_link_text(alt_text)}]"
                f"(aioffice-asset:{block.asset_id})"
            )
        elif isinstance(block, OpaqueBlock):
            summary = block.summary.replace("--", "—").replace(
                "\n",
                " ",
            )
            lines.append(
                f"<!-- aioffice opaque native content: {summary} -->"
            )
        elif isinstance(block, PageBreak):
            lines.append("<!-- pagebreak -->")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def _split_table_row(line: str) -> list[str]:
    stripped = line.strip().strip("|")
    cells = re.split(r"(?<!\\)\|", stripped)
    return [cell.strip().replace("\\|", "|").replace("\\\\", "\\") for cell in cells]


def import_markdown(text: str, *, title: str | None = None) -> AiOfficeDocumentSpec:
    """Parse a practical CommonMark subset into a strict Document Spec."""

    source_lines = text.replace("\r\n", "\n").replace("\r", "\n").split("\n")
    content: list[dict[str, Any]] = []
    index = 0
    while index < len(source_lines):
        line = source_lines[index]
        if not line.strip():
            index += 1
            continue

        heading = _HEADING_RE.match(line)
        if heading:
            content.append(
                {"type": "heading", "level": len(heading.group(1)), "text": heading.group(2)}
            )
            index += 1
            continue

        if line.strip().lower() == "<!-- pagebreak -->":
            content.append({"type": "page_break"})
            index += 1
            continue

        bullet = _BULLET_RE.match(line)
        if bullet:
            items: list[str] = []
            while index < len(source_lines):
                match = _BULLET_RE.match(source_lines[index])
                if not match:
                    break
                items.append(match.group(1))
                index += 1
            content.append({"type": "bullet_list", "items": items})
            continue

        ordered = _ORDERED_RE.match(line)
        if ordered:
            items = []
            while index < len(source_lines):
                match = _ORDERED_RE.match(source_lines[index])
                if not match:
                    break
                items.append(match.group(1))
                index += 1
            content.append({"type": "ordered_list", "items": items})
            continue

        if (
            "|" in line
            and index + 1 < len(source_lines)
            and _TABLE_SEPARATOR_RE.match(source_lines[index + 1])
        ):
            headers = _split_table_row(line)
            columns = [
                {"key": f"column_{column_index + 1}", "title": header}
                for column_index, header in enumerate(headers)
            ]
            rows: list[dict[str, Any]] = []
            index += 2
            while index < len(source_lines) and "|" in source_lines[index]:
                values = _split_table_row(source_lines[index])
                rows.append(
                    {
                        "values": {
                            column["key"]: values[cell_index] if cell_index < len(values) else ""
                            for cell_index, column in enumerate(columns)
                        }
                    }
                )
                index += 1
            content.append({"type": "table", "columns": columns, "rows": rows})
            continue

        paragraph_lines = [line.strip()]
        index += 1
        while index < len(source_lines):
            candidate = source_lines[index]
            if not candidate.strip():
                break
            if (
                _HEADING_RE.match(candidate)
                or _BULLET_RE.match(candidate)
                or _ORDERED_RE.match(candidate)
                or candidate.strip().lower() == "<!-- pagebreak -->"
            ):
                break
            paragraph_lines.append(candidate.strip())
            index += 1
        content.append({"type": "paragraph", "text": " ".join(paragraph_lines)})

    inferred_title = title
    if inferred_title is None:
        inferred_title = next(
            (
                block["text"]
                for block in content
                if block["type"] == "heading" and block["level"] == 1
            ),
            None,
        )
    return AiOfficeDocumentSpec.model_validate(
        {"metadata": {"title": inferred_title}, "content": content}
    )
