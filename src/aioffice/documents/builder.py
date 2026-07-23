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
        defaults: Mapping[str, Any] | None = None,
        settings: Mapping[str, Any] | None = None,
        styles: Iterable[Mapping[str, Any]] = (),
        sections: Iterable[Mapping[str, Any]] | None = None,
        header_footers: Iterable[Mapping[str, Any]] = (),
    ) -> None:
        self._metadata: dict[str, Any] = {"title": title, "author": author}
        self._theme = theme
        self._defaults = deepcopy(dict(defaults)) if defaults is not None else {}
        self._settings = (
            deepcopy(dict(settings)) if settings is not None else None
        )
        self._styles = [deepcopy(dict(style)) for style in styles]
        self._sections = (
            [deepcopy(dict(section)) for section in sections]
            if sections is not None
            else None
        )
        self._header_footers = [
            deepcopy(dict(part)) for part in header_footers
        ]
        self._content: list[dict[str, Any]] = []

    def define_style(self, style: Mapping[str, Any]) -> "DocumentBuilder":
        """Add a strict document-local named style definition."""

        self._styles.append(deepcopy(dict(style)))
        return self

    def define_header_footer(
        self,
        part: Mapping[str, Any],
    ) -> "DocumentBuilder":
        """Add a reusable header/footer part referenced by section bindings."""

        self._header_footers.append(deepcopy(dict(part)))
        return self

    def section(
        self,
        layout: Mapping[str, Any],
        *,
        start_at: str | None = None,
        id: str | None = None,
    ) -> "DocumentBuilder":
        """Add an ordered page/section definition anchored at a content node."""

        if self._sections is None:
            self._sections = []
        section: dict[str, Any] = {
            "type": "section",
            "start_at": start_at,
            "layout": deepcopy(dict(layout)),
        }
        if id is not None:
            section["id"] = id
        self._sections.append(section)
        return self

    def heading(
        self,
        text: str,
        *,
        level: int = 1,
        id: str | None = None,
        tags: Iterable[str] = (),
        style_ref: str | None = None,
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
        if style_ref is not None:
            node["style_ref"] = style_ref
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
        style_ref: str | None = None,
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
        if style_ref is not None:
            node["style_ref"] = style_ref
        if text_style is not None:
            node["text_style"] = deepcopy(dict(text_style))
        if id is not None:
            node["id"] = id
        self._content.append(node)
        return self

    def rich_heading(
        self,
        spans: Iterable[Mapping[str, Any]],
        *,
        level: int = 1,
        id: str | None = None,
        tags: Iterable[str] = (),
        style_ref: str | None = None,
        paragraph_style: Mapping[str, Any] | None = None,
        text_style: Mapping[str, Any] | None = None,
    ) -> "DocumentBuilder":
        node: dict[str, Any] = {
            "type": "heading",
            "content": [deepcopy(dict(span)) for span in spans],
            "level": level,
            "tags": list(tags),
        }
        if paragraph_style is not None:
            node["paragraph_style"] = deepcopy(dict(paragraph_style))
        if style_ref is not None:
            node["style_ref"] = style_ref
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
        style_ref: str | None = None,
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
        if style_ref is not None:
            node["style_ref"] = style_ref
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
        payload = {
            "metadata": metadata,
            "theme": {"ref": self._theme},
            "defaults": deepcopy(self._defaults),
            "styles": deepcopy(self._styles),
            "header_footers": deepcopy(self._header_footers),
            "content": deepcopy(self._content),
        }
        if self._settings is not None:
            payload["settings"] = deepcopy(self._settings)
        if self._sections is not None:
            payload["sections"] = deepcopy(self._sections)
        return Document.from_spec(payload)


__all__ = ["DocumentBuilder"]
