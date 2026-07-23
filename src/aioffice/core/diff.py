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


def _semantic_node(node: dict[str, Any], *, include_native: bool) -> dict[str, Any]:
    value = dict(node)
    value.pop("revision_added", None)
    value.pop("revision_updated", None)
    if not include_native:
        value.pop("source_ref", None)
    return value


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
    if before_order != after_order:
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
    return DocumentDiff(
        artifact_id=after.artifact.id,
        base_revision=before.artifact.revision,
        result_revision=after.artifact.revision,
        entries=entries,
    )


__all__ = ["DiffEntry", "DocumentDiff", "compute_document_diff"]
