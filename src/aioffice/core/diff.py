"""Stable, machine-readable semantic document diffs."""

from __future__ import annotations

from collections import Counter
from typing import TYPE_CHECKING, Any, Literal

from pydantic import BaseModel, ConfigDict, Field

if TYPE_CHECKING:
    from aioffice.spec.models import AiOfficeDocumentSpec


class DiffEntry(BaseModel):
    model_config = ConfigDict(extra="forbid")

    path: str
    kind: Literal["added", "removed", "changed", "moved"]
    node_id: str | None = None
    before: Any = None
    after: Any = None


class DocumentDiff(BaseModel):
    model_config = ConfigDict(extra="forbid")

    artifact_id: str
    base_revision: int
    result_revision: int
    entries: list[DiffEntry] = Field(default_factory=list)

    @property
    def changed(self) -> bool:
        return bool(self.entries)

    @property
    def summary(self) -> dict[str, int]:
        counts = Counter(entry.kind for entry in self.entries)
        return {
            "added": counts["added"],
            "removed": counts["removed"],
            "changed": counts["changed"],
            "moved": counts["moved"],
            "total": len(self.entries),
        }


def _walk(
    before: Any,
    after: Any,
    *,
    path: str,
    entries: list[DiffEntry],
    node_id: str | None = None,
) -> None:
    if isinstance(before, dict) and isinstance(after, dict):
        for key in sorted(set(before) | set(after)):
            child_path = f"{path}.{key}" if path else key
            if key not in before:
                entries.append(
                    DiffEntry(
                        path=child_path,
                        kind="added",
                        node_id=node_id,
                        after=after[key],
                    )
                )
            elif key not in after:
                entries.append(
                    DiffEntry(
                        path=child_path,
                        kind="removed",
                        node_id=node_id,
                        before=before[key],
                    )
                )
            else:
                _walk(
                    before[key],
                    after[key],
                    path=child_path,
                    entries=entries,
                    node_id=node_id,
                )
        return
    if before != after:
        entries.append(
            DiffEntry(
                path=path,
                kind="changed",
                node_id=node_id,
                before=before,
                after=after,
            )
        )


def _semantic_value(value: Any, *, include_native: bool) -> Any:
    if isinstance(value, dict):
        return {
            key: _semantic_value(item, include_native=include_native)
            for key, item in value.items()
            if key not in {"revision_added", "revision_updated"}
            and (include_native or key != "source_ref")
        }
    if isinstance(value, list):
        return [
            _semantic_value(item, include_native=include_native)
            for item in value
        ]
    return value


def _semantic_node(node: dict[str, Any], *, include_native: bool) -> dict[str, Any]:
    return _semantic_value(node, include_native=include_native)


def _semantic_header_footer(
    part: dict[str, Any],
    *,
    include_native: bool,
) -> dict[str, Any]:
    value = _semantic_node(part, include_native=include_native)
    value["content"] = [
        _semantic_node(block, include_native=include_native)
        for block in part.get("content", [])
    ]
    return value


def _relative_order_changed(
    before: list[str],
    after: list[str],
) -> bool:
    """Return whether identities present on both sides changed relative order."""

    common = set(before).intersection(after)
    return (
        [node_id for node_id in before if node_id in common]
        != [node_id for node_id in after if node_id in common]
    )


def compute_document_diff(
    before: AiOfficeDocumentSpec,
    after: AiOfficeDocumentSpec,
    *,
    include_native: bool = False,
) -> DocumentDiff:
    """Compare two revisions using node identities instead of array positions."""

    entries: list[DiffEntry] = []
    before_payload = before.model_dump(mode="json", by_alias=True, exclude_none=True)
    after_payload = after.model_dump(mode="json", by_alias=True, exclude_none=True)

    for payload in (before_payload, after_payload):
        payload.pop("content", None)
        payload.pop("sections", None)
        payload.pop("header_footers", None)
        payload.pop("engine_version", None)
        artifact = payload.get("artifact", {})
        artifact.pop("revision", None)
        if not include_native:
            extensions = payload.get("extensions", {})
            extensions.pop("dev.aioffice.native", None)

    _walk(before_payload, after_payload, path="", entries=entries)

    before_nodes = {
        node.id: _semantic_node(
            node.model_dump(mode="json", exclude_none=True),
            include_native=include_native,
        )
        for node in before.content
    }
    after_nodes = {
        node.id: _semantic_node(
            node.model_dump(mode="json", exclude_none=True),
            include_native=include_native,
        )
        for node in after.content
    }
    before_order = [node.id for node in before.content]
    after_order = [node.id for node in after.content]
    if _relative_order_changed(before_order, after_order):
        entries.append(
            DiffEntry(
                path="content.order",
                kind="moved",
                before=before_order,
                after=after_order,
            )
        )
    for node_id in before_order:
        if node_id not in after_nodes:
            entries.append(
                DiffEntry(
                    path=f"content.#{node_id}",
                    kind="removed",
                    node_id=node_id,
                    before=before_nodes[node_id],
                )
            )
    for node_id in after_order:
        if node_id not in before_nodes:
            entries.append(
                DiffEntry(
                    path=f"content.#{node_id}",
                    kind="added",
                    node_id=node_id,
                    after=after_nodes[node_id],
                )
            )
        else:
            _walk(
                before_nodes[node_id],
                after_nodes[node_id],
                path=f"content.#{node_id}",
                entries=entries,
                node_id=node_id,
            )

    before_sections = {
        section.id: _semantic_node(
            section.model_dump(mode="json", exclude_none=True),
            include_native=include_native,
        )
        for section in before.sections
    }
    after_sections = {
        section.id: _semantic_node(
            section.model_dump(mode="json", exclude_none=True),
            include_native=include_native,
        )
        for section in after.sections
    }
    before_section_order = [section.id for section in before.sections]
    after_section_order = [section.id for section in after.sections]
    if _relative_order_changed(
        before_section_order,
        after_section_order,
    ):
        entries.append(
            DiffEntry(
                path="sections.order",
                kind="moved",
                before=before_section_order,
                after=after_section_order,
            )
        )
    for section_id in before_section_order:
        if section_id not in after_sections:
            entries.append(
                DiffEntry(
                    path=f"sections.#{section_id}",
                    kind="removed",
                    node_id=section_id,
                    before=before_sections[section_id],
                )
            )
    for section_id in after_section_order:
        if section_id not in before_sections:
            entries.append(
                DiffEntry(
                    path=f"sections.#{section_id}",
                    kind="added",
                    node_id=section_id,
                    after=after_sections[section_id],
                )
            )
        else:
            _walk(
                before_sections[section_id],
                after_sections[section_id],
                path=f"sections.#{section_id}",
                entries=entries,
                node_id=section_id,
            )

    before_parts = {
        part.id: _semantic_header_footer(
            part.model_dump(mode="json", exclude_none=True),
            include_native=include_native,
        )
        for part in before.header_footers
    }
    after_parts = {
        part.id: _semantic_header_footer(
            part.model_dump(mode="json", exclude_none=True),
            include_native=include_native,
        )
        for part in after.header_footers
    }
    before_part_order = [part.id for part in before.header_footers]
    after_part_order = [part.id for part in after.header_footers]
    if _relative_order_changed(before_part_order, after_part_order):
        entries.append(
            DiffEntry(
                path="header_footers.order",
                kind="moved",
                before=before_part_order,
                after=after_part_order,
            )
        )
    for part_id in before_part_order:
        if part_id not in after_parts:
            entries.append(
                DiffEntry(
                    path=f"header_footers.#{part_id}",
                    kind="removed",
                    node_id=part_id,
                    before=before_parts[part_id],
                )
            )
    for part_id in after_part_order:
        if part_id not in before_parts:
            entries.append(
                DiffEntry(
                    path=f"header_footers.#{part_id}",
                    kind="added",
                    node_id=part_id,
                    after=after_parts[part_id],
                )
            )
        else:
            _walk(
                before_parts[part_id],
                after_parts[part_id],
                path=f"header_footers.#{part_id}",
                entries=entries,
                node_id=part_id,
            )
    return DocumentDiff(
        artifact_id=after.artifact.id,
        base_revision=before.artifact.revision,
        result_revision=after.artifact.revision,
        entries=entries,
    )


__all__ = ["DiffEntry", "DocumentDiff", "compute_document_diff"]
