"""Built-in design tokens."""

from __future__ import annotations

from copy import deepcopy
from typing import Any

_BUILTIN_THEMES: dict[str, dict[str, Any]] = {
    "business-clean": {
        "name": "business-clean",
        "tokens": {
            "color.primary": "#1F4E78",
            "color.text": "#222222",
            "font.body.latin": "Aptos",
            "font.body.east_asia": "Microsoft YaHei",
            "font.heading.latin": "Aptos Display",
            "font.heading.east_asia": "Microsoft YaHei",
            "spacing.base": "4pt",
        },
        "document": {
            "page.margin.top": "72pt",
            "page.margin.right": "72pt",
            "page.margin.bottom": "72pt",
            "page.margin.left": "72pt",
        },
    }
}


def get_theme(name: str) -> dict[str, Any] | None:
    theme = _BUILTIN_THEMES.get(name)
    return deepcopy(theme) if theme is not None else None


def list_themes() -> list[str]:
    return sorted(_BUILTIN_THEMES)


__all__ = ["get_theme", "list_themes"]
