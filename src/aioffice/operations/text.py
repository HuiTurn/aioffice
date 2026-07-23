"""Shared semantic text selection, formatting, and replacement primitives."""

from __future__ import annotations

from copy import deepcopy
from typing import Any, Mapping, Sequence

from aioffice.spec.models import TextSpan, TextStyle

from .models import TextMatch, TextRange


def node_plain_text(node: Mapping[str, Any]) -> str:
    text = node.get("text")
    if isinstance(text, str):
        return text
    return "".join(
        (
            str(span.get("text", ""))
            if span.get("type") == "text"
            else str(span.get("cached_result") or "")
            if span.get("type") == "field"
            else ""
        )
        for span in node.get("content", [])
        if isinstance(span, Mapping)
    )


def resolve_text_selection(
    text: str,
    *,
    range_value: Any = None,
    match_value: Any = None,
) -> TextRange | None:
    """Resolve one optional exact selector against the current node text."""

    if range_value is not None and match_value is not None:
        raise ValueError("text.format accepts either range or match, not both.")
    if range_value is not None:
        selection = TextRange.model_validate(range_value)
        if selection.end > len(text):
            raise ValueError(
                f"Text range [{selection.start}, {selection.end}) exceeds "
                f"node length {len(text)}."
            )
        return selection
    if match_value is not None:
        match = TextMatch.model_validate(match_value)
        cursor = 0
        found = -1
        for _ in range(match.occurrence):
            found = text.find(match.text, cursor)
            if found < 0:
                raise ValueError(
                    f"Text occurrence {match.occurrence} of {match.text!r} was not found."
                )
            cursor = found + len(match.text)
        return TextRange(start=found, end=found + len(match.text))
    return None


def _normalized_style(value: Mapping[str, Any]) -> dict[str, Any]:
    return TextStyle.model_validate(value).model_dump(mode="json", exclude_none=True)


def _merge_style(
    base: Mapping[str, Any] | None,
    override: Mapping[str, Any] | None,
) -> dict[str, Any]:
    return _normalized_style({**(base or {}), **(override or {})})


def _update_style(
    current: Mapping[str, Any] | None,
    set_values: Mapping[str, Any],
    clear_values: Sequence[str],
) -> dict[str, Any]:
    candidate = {**(current or {}), **deepcopy(dict(set_values))}
    for field_name in clear_values:
        candidate.pop(field_name, None)
    return _normalized_style(candidate)


def _same_span_format(left: Mapping[str, Any], right: Mapping[str, Any]) -> bool:
    if left.get("type", "text") != "text" or right.get("type", "text") != "text":
        return False
    keys = (set(left) | set(right)) - {"text"}
    return all(left.get(key) == right.get(key) for key in keys)


def _merge_adjacent_spans(spans: Sequence[Mapping[str, Any]]) -> list[dict[str, Any]]:
    merged: list[dict[str, Any]] = []
    for raw_span in spans:
        span = deepcopy(dict(raw_span))
        span.setdefault("type", "text")
        if span["type"] == "text" and not span.get("text") and merged:
            continue
        if merged and _same_span_format(merged[-1], span):
            merged[-1]["text"] += span.get("text", "")
        else:
            merged.append(span)
    return merged or [{"type": "text", "text": ""}]


def _node_spans(
    node: Mapping[str, Any],
    *,
    materialize_node_style: bool,
) -> list[dict[str, Any]]:
    if isinstance(node.get("text"), str):
        spans = [{"type": "text", "text": node["text"]}]
    else:
        spans = [deepcopy(dict(span)) for span in node.get("content", [])]
    if materialize_node_style:
        base_style = node.get("text_style")
        if isinstance(base_style, Mapping):
            for span in spans:
                style = _merge_style(
                    base_style,
                    span.get("style") if isinstance(span.get("style"), Mapping) else None,
                )
                if style:
                    span["style"] = style
                else:
                    span.pop("style", None)
    return spans


def format_text_range(
    node: dict[str, Any],
    selection: TextRange,
    *,
    set_values: Mapping[str, Any],
    clear_values: Sequence[str],
) -> None:
    """Apply direct format to exactly one semantic character range."""

    spans = _node_spans(node, materialize_node_style=True)
    output: list[dict[str, Any]] = []
    cursor = 0
    for span in spans:
        text = str(span.get("text", ""))
        span_start = cursor
        span_end = cursor + len(text)
        cursor = span_end
        boundaries = {
            0,
            len(text),
            max(0, min(len(text), selection.start - span_start)),
            max(0, min(len(text), selection.end - span_start)),
        }
        ordered = sorted(boundaries)
        for left, right in zip(ordered, ordered[1:]):
            if left == right:
                continue
            piece = deepcopy(span)
            piece["text"] = text[left:right]
            piece_start = span_start + left
            piece_end = span_start + right
            selected = (
                piece_start >= selection.start
                and piece_end <= selection.end
                and piece_start < piece_end
            )
            if selected:
                updated_style = _update_style(
                    piece.get("style")
                    if isinstance(piece.get("style"), Mapping)
                    else None,
                    set_values,
                    clear_values,
                )
                if updated_style:
                    piece["style"] = updated_style
                else:
                    piece.pop("style", None)
            output.append(piece)
    node.pop("text", None)
    node.pop("text_style", None)
    node["content"] = _merge_adjacent_spans(output)


def format_entire_text(
    node: dict[str, Any],
    *,
    set_values: Mapping[str, Any],
    clear_values: Sequence[str],
) -> tuple[dict[str, Any], dict[str, Any]]:
    """Apply uniform direct formatting without allowing span overrides."""

    before_style = deepcopy(node.get("text_style", {}))
    normalized = _update_style(before_style, set_values, clear_values)
    if normalized:
        node["text_style"] = normalized
    else:
        node.pop("text_style", None)
    changed_fields = set(set_values) | set(clear_values)
    for span in node.get("content", []):
        style = deepcopy(span.get("style", {}))
        for field_name in changed_fields:
            style.pop(field_name, None)
        normalized_span = _normalized_style(style)
        if normalized_span:
            span["style"] = normalized_span
        else:
            span.pop("style", None)
    return before_style, normalized


def _occurrences(text: str, search: str, replace_all: bool) -> list[TextRange]:
    result: list[TextRange] = []
    cursor = 0
    while True:
        index = text.find(search, cursor)
        if index < 0:
            break
        result.append(TextRange(start=index, end=index + len(search)))
        if not replace_all:
            break
        cursor = index + len(search)
    return result


def _replace_span_range(
    spans: Sequence[Mapping[str, Any]],
    selection: TextRange,
    replacement: str,
) -> list[dict[str, Any]]:
    before: list[dict[str, Any]] = []
    after: list[dict[str, Any]] = []
    replacement_format: dict[str, Any] | None = None
    cursor = 0
    for raw_span in spans:
        span = deepcopy(dict(raw_span))
        text = str(span.get("text", ""))
        span_start = cursor
        span_end = cursor + len(text)
        cursor = span_end
        if span_end <= selection.start:
            before.append(span)
            continue
        if span_start >= selection.end:
            after.append(span)
            continue
        left_count = max(0, selection.start - span_start)
        right_offset = max(0, selection.end - span_start)
        if left_count:
            left = deepcopy(span)
            left["text"] = text[:left_count]
            before.append(left)
        if replacement_format is None:
            replacement_format = deepcopy(span)
            replacement_format["text"] = replacement
        if right_offset < len(text):
            right = deepcopy(span)
            right["text"] = text[right_offset:]
            after.append(right)
    if replacement_format is not None and replacement:
        before.append(replacement_format)
    before.extend(after)
    return _merge_adjacent_spans(before)


def replace_node_text(
    node: dict[str, Any],
    *,
    search: str,
    replacement: str,
    replace_all: bool,
) -> int:
    """Replace text across semantic span boundaries using first-run formatting."""

    text = node_plain_text(node)
    matches = _occurrences(text, search, replace_all)
    if not matches:
        return 0
    if isinstance(node.get("text"), str):
        node["text"] = text.replace(search, replacement, -1 if replace_all else 1)
        return len(matches)
    spans = _node_spans(node, materialize_node_style=False)
    for selection in reversed(matches):
        spans = _replace_span_range(spans, selection, replacement)
    node["content"] = [
        TextSpan.model_validate(span).model_dump(mode="json", exclude_none=True)
        for span in spans
    ]
    return len(matches)


__all__ = [
    "format_entire_text",
    "format_text_range",
    "node_plain_text",
    "replace_node_text",
    "resolve_text_selection",
]
