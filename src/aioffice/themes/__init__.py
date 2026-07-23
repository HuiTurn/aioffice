"""Built-in design tokens."""

from __future__ import annotations

from copy import deepcopy
from typing import Any

_BUILTIN_THEMES: dict[str, dict[str, Any]] = {
    "native-docx": {
        "name": "native-docx",
        "tokens": {},
        "defaults": {},
        "styles": [],
        "document": {},
    },
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
        "defaults": {
            "paragraph_style": {
                "spacing_after": {"value": 8, "unit": "pt"},
                "line_spacing": {"rule": "multiple", "value": 1.15},
                "widow_control": True,
            },
            "text_style": {
                "font_family": "Aptos",
                "font_family_east_asia": "Microsoft YaHei",
                "font_size": {"value": 11, "unit": "pt"},
                "color": "#222222",
            },
        },
        "styles": [
            {
                "id": "Normal",
                "name": "Normal",
                "semantic_role": "body",
                "quick_style": True,
            },
            {
                "id": "Title",
                "name": "Title",
                "semantic_role": "title",
                "based_on": "Normal",
                "next_style": "Subtitle",
                "quick_style": True,
                "paragraph_style": {
                    "alignment": "center",
                    "spacing_before": {"value": 18, "unit": "pt"},
                    "spacing_after": {"value": 8, "unit": "pt"},
                    "keep_with_next": True,
                },
                "text_style": {
                    "font_family": "Aptos Display",
                    "font_family_east_asia": "Microsoft YaHei",
                    "font_size": {"value": 28, "unit": "pt"},
                    "color": "#1F4E78",
                    "bold": True,
                },
            },
            {
                "id": "Subtitle",
                "name": "Subtitle",
                "semantic_role": "subtitle",
                "based_on": "Normal",
                "next_style": "Normal",
                "quick_style": True,
                "paragraph_style": {
                    "alignment": "center",
                    "spacing_after": {"value": 18, "unit": "pt"},
                    "keep_with_next": True,
                },
                "text_style": {
                    "font_size": {"value": 12, "unit": "pt"},
                    "color": "#5B6573",
                    "italic": True,
                },
            },
            *[
                {
                    "id": f"Heading{level}",
                    "name": f"heading {level}",
                    "semantic_role": "heading",
                    "heading_level": level,
                    "based_on": "Normal",
                    "next_style": "Normal",
                    "quick_style": True,
                    "paragraph_style": {
                        "spacing_before": {
                            "value": 12 if level <= 2 else 8,
                            "unit": "pt",
                        },
                        "spacing_after": {"value": 6, "unit": "pt"},
                        "keep_with_next": True,
                        "outline_level": level,
                    },
                    "text_style": {
                        "font_family": "Aptos Display",
                        "font_family_east_asia": "Microsoft YaHei",
                        "font_size": {
                            "value": {
                                1: 18,
                                2: 16,
                                3: 14,
                                4: 12,
                                5: 11,
                                6: 10,
                            }[level],
                            "unit": "pt",
                        },
                        "color": "#1F4E78",
                        "bold": True,
                    },
                }
                for level in range(1, 7)
            ],
            {
                "id": "Quote",
                "name": "Quote",
                "semantic_role": "quote",
                "based_on": "Normal",
                "next_style": "Normal",
                "quick_style": True,
                "paragraph_style": {
                    "indent_left": {"value": 24, "unit": "pt"},
                    "indent_right": {"value": 24, "unit": "pt"},
                    "spacing_before": {"value": 6, "unit": "pt"},
                    "spacing_after": {"value": 6, "unit": "pt"},
                    "keep_together": True,
                },
                "text_style": {
                    "color": "#4A5563",
                    "italic": True,
                },
            },
            {
                "id": "Caption",
                "name": "Caption",
                "semantic_role": "caption",
                "based_on": "Normal",
                "next_style": "Normal",
                "quick_style": True,
                "paragraph_style": {
                    "alignment": "center",
                    "spacing_before": {"value": 4, "unit": "pt"},
                    "spacing_after": {"value": 8, "unit": "pt"},
                    "keep_with_next": True,
                },
                "text_style": {
                    "font_size": {"value": 9, "unit": "pt"},
                    "color": "#5B6573",
                },
            },
            {
                "id": "CodeBlock",
                "name": "Code Block",
                "semantic_role": "code",
                "based_on": "Normal",
                "next_style": "Normal",
                "paragraph_style": {
                    "indent_left": {"value": 12, "unit": "pt"},
                    "indent_right": {"value": 12, "unit": "pt"},
                    "keep_together": True,
                },
                "text_style": {
                    "font_family": "Consolas",
                    "font_size": {"value": 9, "unit": "pt"},
                    "background_color": "#F3F5F7",
                },
            },
        ],
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
