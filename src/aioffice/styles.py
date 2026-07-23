"""Named-style catalog and deterministic inheritance resolution."""

from __future__ import annotations

from dataclasses import dataclass
from typing import TypeVar

from pydantic import BaseModel

from aioffice.spec.models import (
    AiOfficeDocumentSpec,
    DocumentDefaults,
    NamedStyle,
    ParagraphStyle,
    TextStyle,
)
from aioffice.themes import get_theme

StyleModel = TypeVar("StyleModel", bound=BaseModel)


def _merge_model(
    model_type: type[StyleModel],
    base: StyleModel | None,
    override: StyleModel | None,
) -> StyleModel | None:
    values = {
        **(base.model_dump(mode="json", exclude_none=True) if base is not None else {}),
        **(
            override.model_dump(mode="json", exclude_none=True)
            if override is not None
            else {}
        ),
    }
    return model_type.model_validate(values) if values else None


def merge_paragraph_styles(
    base: ParagraphStyle | None,
    override: ParagraphStyle | None,
) -> ParagraphStyle | None:
    return _merge_model(ParagraphStyle, base, override)


def merge_text_styles(
    base: TextStyle | None,
    override: TextStyle | None,
) -> TextStyle | None:
    return _merge_model(TextStyle, base, override)


def theme_named_styles(theme_name: str) -> list[NamedStyle]:
    theme = get_theme(theme_name) or {}
    return [NamedStyle.model_validate(value) for value in theme.get("styles", [])]


def style_catalog(spec: AiOfficeDocumentSpec) -> dict[str, NamedStyle]:
    """Return theme styles with document-local definitions taking precedence."""

    result = {style.id: style for style in theme_named_styles(spec.theme.ref)}
    result.update({style.id: style for style in spec.styles})
    return result


def resolve_document_defaults(spec: AiOfficeDocumentSpec) -> DocumentDefaults:
    theme = get_theme(spec.theme.ref) or {}
    theme_defaults = DocumentDefaults.model_validate(theme.get("defaults", {}))
    return DocumentDefaults(
        paragraph_style=merge_paragraph_styles(
            theme_defaults.paragraph_style,
            spec.defaults.paragraph_style,
        ),
        text_style=merge_text_styles(
            theme_defaults.text_style,
            spec.defaults.text_style,
        ),
    )


@dataclass(frozen=True, slots=True)
class ResolvedNamedStyle:
    style: NamedStyle
    paragraph_style: ParagraphStyle | None
    text_style: TextStyle | None
    inheritance_chain: tuple[str, ...]


def resolve_named_style(
    spec: AiOfficeDocumentSpec,
    style_id: str,
) -> ResolvedNamedStyle:
    """Resolve one style from document defaults through its based-on chain."""

    catalog = style_catalog(spec)
    if style_id not in catalog:
        raise KeyError(style_id)
    defaults = resolve_document_defaults(spec)
    visiting: set[str] = set()
    chain: list[NamedStyle] = []
    current_id: str | None = style_id
    while current_id is not None:
        if current_id in visiting:
            raise ValueError(f"Named style inheritance cycle includes {current_id!r}.")
        visiting.add(current_id)
        try:
            current = catalog[current_id]
        except KeyError as error:
            raise KeyError(
                f"Named style {chain[-1].id!r} is based on missing style {current_id!r}."
            ) from error
        chain.append(current)
        current_id = current.based_on

    paragraph_style = defaults.paragraph_style
    text_style = defaults.text_style
    for style in reversed(chain):
        paragraph_style = merge_paragraph_styles(paragraph_style, style.paragraph_style)
        text_style = merge_text_styles(text_style, style.text_style)
    return ResolvedNamedStyle(
        style=chain[0],
        paragraph_style=paragraph_style,
        text_style=text_style,
        inheritance_chain=tuple(style.id for style in reversed(chain)),
    )


def resolve_node_styles(
    spec: AiOfficeDocumentSpec,
    *,
    style_ref: str | None,
    paragraph_style: ParagraphStyle | None,
    text_style: TextStyle | None,
) -> tuple[ParagraphStyle | None, TextStyle | None]:
    defaults = resolve_document_defaults(spec)
    resolved_paragraph = defaults.paragraph_style
    resolved_text = defaults.text_style
    if style_ref is not None:
        named = resolve_named_style(spec, style_ref)
        resolved_paragraph = named.paragraph_style
        resolved_text = named.text_style
    return (
        merge_paragraph_styles(resolved_paragraph, paragraph_style),
        merge_text_styles(resolved_text, text_style),
    )


__all__ = [
    "ResolvedNamedStyle",
    "merge_paragraph_styles",
    "merge_text_styles",
    "resolve_document_defaults",
    "resolve_named_style",
    "resolve_node_styles",
    "style_catalog",
    "theme_named_styles",
]
