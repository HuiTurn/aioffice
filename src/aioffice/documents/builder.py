"""Convenience builder that emits the same strict Document Spec."""

from __future__ import annotations

from copy import deepcopy
from typing import Any, Iterable, Mapping

from .document import Document


class DocumentBuilder:
    def __init__(
        self,
        *,
        title: str | None = None,
        author: str | None = None,
        theme: str = "business-clean",
    ) -> None:
        self._metadata: dict[str, Any] = {"title": title, "author": author}
        self._theme = theme
        self._content: list[dict[str, Any]] = []

    def heading(
        self,
        text: str,
        *,
        level: int = 1,
        id: str | None = None,
        tags: Iterable[str] = (),
        paragraph_style: Mapping[str, Any] | None = None,
        text_style: Mapping[str, Any] | None = None,
    ) -> "DocumentBuilder":
        node: dict[str, Any] = {
            "type": "heading",
            "text": text,
            "level": level,
            "tags": list(tags),
        }
        if paragraph_style is not None:
            node["paragraph_style"] = deepcopy(dict(paragraph_style))
        if text_style is not None:
            node["text_style"] = deepcopy(dict(text_style))
        if id is not None:
            node["id"] = id
        self._content.append(node)
        return self

    def paragraph(
        self,
        text: str,
        *,
        id: str | None = None,
        tags: Iterable[str] = (),
        paragraph_style: Mapping[str, Any] | None = None,
        text_style: Mapping[str, Any] | None = None,
    ) -> "DocumentBuilder":
        node: dict[str, Any] = {
            "type": "paragraph",
            "text": text,
            "tags": list(tags),
        }
        if paragraph_style is not None:
            node["paragraph_style"] = deepcopy(dict(paragraph_style))
        if text_style is not None:
            node["text_style"] = deepcopy(dict(text_style))
        if id is not None:
            node["id"] = id
        self._content.append(node)
        return self

    def rich_paragraph(
        self,
        spans: Iterable[Mapping[str, Any]],
        *,
        id: str | None = None,
        tags: Iterable[str] = (),
        paragraph_style: Mapping[str, Any] | None = None,
        text_style: Mapping[str, Any] | None = None,
    ) -> "DocumentBuilder":
        node: dict[str, Any] = {
            "type": "paragraph",
            "content": [deepcopy(dict(span)) for span in spans],
            "tags": list(tags),
        }
        if paragraph_style is not None:
            node["paragraph_style"] = deepcopy(dict(paragraph_style))
        if text_style is not None:
            node["text_style"] = deepcopy(dict(text_style))
        if id is not None:
            node["id"] = id
        self._content.append(node)
        return self

    def bullet_list(
        self,
        items: Iterable[str],
        *,
        id: str | None = None,
        tags: Iterable[str] = (),
    ) -> "DocumentBuilder":
        node: dict[str, Any] = {
            "type": "bullet_list",
            "items": list(items),
            "tags": list(tags),
        }
        if id is not None:
            node["id"] = id
        self._content.append(node)
        return self

    def ordered_list(
        self,
        items: Iterable[str],
        *,
        id: str | None = None,
        tags: Iterable[str] = (),
    ) -> "DocumentBuilder":
        node: dict[str, Any] = {
            "type": "ordered_list",
            "items": list(items),
            "tags": list(tags),
        }
        if id is not None:
            node["id"] = id
        self._content.append(node)
        return self

    def table(
        self,
        columns: Iterable[Mapping[str, Any]],
        rows: Iterable[Mapping[str, Any]],
        *,
        id: str | None = None,
        tags: Iterable[str] = (),
    ) -> "DocumentBuilder":
        normalized_rows: list[dict[str, Any]] = []
        for row in rows:
            value = deepcopy(dict(row))
            normalized_rows.append(value if "values" in value else {"values": value})
        node: dict[str, Any] = {
            "type": "table",
            "columns": [deepcopy(dict(column)) for column in columns],
            "rows": normalized_rows,
            "tags": list(tags),
        }
        if id is not None:
            node["id"] = id
        self._content.append(node)
        return self

    def page_break(self, *, id: str | None = None) -> "DocumentBuilder":
        node: dict[str, Any] = {"type": "page_break"}
        if id is not None:
            node["id"] = id
        self._content.append(node)
        return self

    def build(self) -> Document:
        metadata = {key: value for key, value in self._metadata.items() if value is not None}
        return Document.from_spec(
            {
                "metadata": metadata,
                "theme": {"ref": self._theme},
                "content": deepcopy(self._content),
            }
        )


__all__ = ["DocumentBuilder"]
