"""Document artifact, validation, export, inspection, and atomic patching."""

from __future__ import annotations

import json
from copy import deepcopy
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence

from pydantic import TypeAdapter, ValidationError

from aioffice.core.diagnostics import Diagnostic, Severity, ValidationResult
from aioffice.core.errors import ExportError, SpecValidationError, UnsupportedFormatError
from aioffice.formats.docx import export_docx
from aioffice.formats.html import export_html
from aioffice.formats.markdown import export_markdown, import_markdown
from aioffice.spec.models import (
    AiOfficeDocumentSpec,
    Block,
    Heading,
    Paragraph,
    Table,
)
from aioffice.themes import get_theme


def _validation_error_diagnostics(error: ValidationError) -> list[Diagnostic]:
    diagnostics: list[Diagnostic] = []
    for item in error.errors(include_url=False):
        path = ".".join(str(part) for part in item["loc"])
        diagnostics.append(
            Diagnostic(
                severity=Severity.ERROR,
                code="INVALID_SPEC",
                message=item["msg"],
                path=path or None,
                recoverable=True,
                suggested_actions=[{"action": "fix_field", "path": path}],
            )
        )
    return diagnostics


def _parse_spec(value: AiOfficeDocumentSpec | Mapping[str, Any]) -> AiOfficeDocumentSpec:
    if isinstance(value, AiOfficeDocumentSpec):
        return value.model_copy(deep=True)
    payload = deepcopy(dict(value))
    try:
        return AiOfficeDocumentSpec.model_validate(payload)
    except ValidationError as error:
        diagnostics = _validation_error_diagnostics(error)
        summary = "; ".join(
            f"{item.path or '<root>'}: {item.message}" for item in diagnostics[:3]
        )
        raise SpecValidationError(f"Invalid AiOffice Document Spec: {summary}", diagnostics) from error


@dataclass(slots=True)
class PatchResult:
    success: bool
    base_revision: int
    result_revision: int
    dry_run: bool
    document: "Document | None" = None
    changes: list[dict[str, Any]] = field(default_factory=list)
    diagnostics: list[Diagnostic] = field(default_factory=list)
    idempotency_key: str | None = None

    def model_dump(self) -> dict[str, Any]:
        return {
            "success": self.success,
            "base_revision": self.base_revision,
            "result_revision": self.result_revision,
            "dry_run": self.dry_run,
            "changes": deepcopy(self.changes),
            "diagnostics": [item.model_dump(mode="json") for item in self.diagnostics],
            "idempotency_key": self.idempotency_key,
        }


class _PatchFailure(Exception):
    def __init__(self, diagnostic: Diagnostic) -> None:
        super().__init__(diagnostic.message)
        self.diagnostic = diagnostic


class Document:
    """A validated, logically immutable Document artifact."""

    def __init__(self, spec: AiOfficeDocumentSpec) -> None:
        self._spec = spec.model_copy(deep=True)

    @classmethod
    def from_spec(cls, spec: AiOfficeDocumentSpec | Mapping[str, Any]) -> "Document":
        return cls(_parse_spec(spec))

    @classmethod
    def from_json(cls, source: str | bytes | Path) -> "Document":
        if isinstance(source, Path):
            raw = source.read_text(encoding="utf-8")
        elif isinstance(source, bytes):
            raw = source.decode("utf-8")
        elif source.lstrip().startswith(("{", "[")):
            raw = source
        else:
            raw = Path(source).read_text(encoding="utf-8")
        try:
            payload = json.loads(raw)
        except json.JSONDecodeError as error:
            diagnostic = Diagnostic(
                severity=Severity.ERROR,
                code="INVALID_SPEC",
                message=f"Invalid JSON at line {error.lineno}, column {error.colno}: {error.msg}",
                path=f"line:{error.lineno}:column:{error.colno}",
            )
            raise SpecValidationError(diagnostic.message, [diagnostic]) from error
        if not isinstance(payload, dict):
            raise SpecValidationError("The root of an AiOffice Document Spec must be an object.")
        return cls.from_spec(payload)

    @classmethod
    def from_markdown(cls, source: str | Path, *, title: str | None = None) -> "Document":
        if isinstance(source, Path):
            text = source.read_text(encoding="utf-8")
        elif "\n" not in source and Path(source).suffix.lower() in {".md", ".markdown"}:
            text = Path(source).read_text(encoding="utf-8")
        else:
            text = source
        return cls(import_markdown(text, title=title))

    @property
    def id(self) -> str:
        return self._spec.artifact.id

    @property
    def kind(self) -> str:
        return self._spec.artifact.kind

    @property
    def revision(self) -> int:
        return self._spec.artifact.revision

    @property
    def spec_version(self) -> str:
        return self._spec.spec_version

    @property
    def spec(self) -> AiOfficeDocumentSpec:
        return self._spec.model_copy(deep=True)

    def to_spec(self) -> dict[str, Any]:
        return self._spec.model_dump(mode="json", by_alias=True, exclude_none=True)

    def to_json(self, *, indent: int = 2) -> str:
        return json.dumps(self.to_spec(), ensure_ascii=False, indent=indent) + "\n"

    def inspect(self, *, response_format: str = "compact") -> dict[str, Any]:
        if response_format not in {"compact", "expanded", "summary"}:
            raise ValueError("response_format must be compact, expanded, or summary.")
        counts: dict[str, int] = {}
        for node in self._spec.content:
            counts[node.type] = counts.get(node.type, 0) + 1
        result: dict[str, Any] = {
            "artifact_id": self.id,
            "kind": self.kind,
            "revision": self.revision,
            "spec_version": self.spec_version,
            "title": self._spec.metadata.title,
            "node_count": len(self._spec.content),
            "node_types": counts,
        }
        if response_format == "compact":
            result["nodes"] = [
                {
                    "id": node.id,
                    "type": node.type,
                    **(
                        {"text": node.text, "level": node.level}
                        if isinstance(node, Heading)
                        else {}
                    ),
                }
                for node in self._spec.content
            ]
        elif response_format == "expanded":
            result["nodes"] = [
                node.model_dump(mode="json", exclude_none=True) for node in self._spec.content
            ]
        return result

    def validate(self) -> ValidationResult:
        diagnostics: list[Diagnostic] = []
        seen: dict[str, str] = {self.id: "artifact"}
        previous_heading_level: int | None = None

        if not self._spec.content:
            diagnostics.append(
                Diagnostic(
                    severity=Severity.WARNING,
                    code="EMPTY_DOCUMENT",
                    message="The document has no content.",
                    node_ids=[self.id],
                    suggested_actions=[{"action": "add_content"}],
                )
            )

        if get_theme(self._spec.theme.ref) is None:
            diagnostics.append(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="UNSUPPORTED_FEATURE",
                    message=f"Theme {self._spec.theme.ref!r} is not registered.",
                    recoverable=True,
                    suggested_actions=[{"action": "use_theme", "name": "business-clean"}],
                )
            )

        for index, node in enumerate(self._spec.content):
            if node.id in seen:
                diagnostics.append(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=f"Duplicate node ID {node.id!r}.",
                        node_ids=[node.id],
                        path=f"content.{index}.id",
                        suggested_actions=[{"action": "assign_unique_id"}],
                    )
                )
            else:
                seen[node.id] = f"content.{index}"

            if node.revision_added > self.revision or node.revision_updated > self.revision:
                diagnostics.append(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=(
                            f"Node {node.id!r} references a revision newer than "
                            f"artifact revision {self.revision}."
                        ),
                        node_ids=[node.id],
                        path=f"content.{index}",
                    )
                )

            if isinstance(node, Heading):
                if (
                    previous_heading_level is not None
                    and node.level > previous_heading_level + 1
                ):
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.WARNING,
                            code="HEADING_LEVEL_JUMP",
                            message=(
                                f"Heading {node.id!r} jumps from level "
                                f"{previous_heading_level} to {node.level}."
                            ),
                            node_ids=[node.id],
                            path=f"content.{index}.level",
                            suggested_actions=[
                                {"action": "set_heading_level", "maximum": previous_heading_level + 1}
                            ],
                        )
                    )
                previous_heading_level = node.level

            if isinstance(node, Table):
                keys = [column.key for column in node.columns]
                if len(keys) != len(set(keys)):
                    diagnostics.append(
                        Diagnostic(
                            severity=Severity.ERROR,
                            code="INVALID_SPEC",
                            message=f"Table {node.id!r} has duplicate column keys.",
                            node_ids=[node.id],
                            path=f"content.{index}.columns",
                        )
                    )
                known_keys = set(keys)
                for row_index, row in enumerate(node.rows):
                    if row.id in seen:
                        diagnostics.append(
                            Diagnostic(
                                severity=Severity.ERROR,
                                code="INVALID_SPEC",
                                message=f"Duplicate node ID {row.id!r}.",
                                node_ids=[row.id],
                                path=f"content.{index}.rows.{row_index}.id",
                            )
                        )
                    else:
                        seen[row.id] = f"content.{index}.rows.{row_index}"
                    unknown = sorted(set(row.values) - known_keys)
                    if unknown:
                        diagnostics.append(
                            Diagnostic(
                                severity=Severity.ERROR,
                                code="INVALID_SPEC",
                                message=(
                                    f"Table row {row.id!r} uses unknown columns: "
                                    f"{', '.join(unknown)}."
                                ),
                                node_ids=[node.id, row.id],
                                path=f"content.{index}.rows.{row_index}.values",
                            )
                        )
                    missing = sorted(known_keys - set(row.values))
                    if missing:
                        diagnostics.append(
                            Diagnostic(
                                severity=Severity.WARNING,
                                code="TABLE_VALUE_MISSING",
                                message=(
                                    f"Table row {row.id!r} has no values for: "
                                    f"{', '.join(missing)}."
                                ),
                                node_ids=[node.id, row.id],
                                path=f"content.{index}.rows.{row_index}.values",
                                suggested_actions=[{"action": "fill_table_values", "keys": missing}],
                            )
                        )
        return ValidationResult(diagnostics=diagnostics)

    def export(self, target: str | Path) -> Path:
        validation = self.validate()
        if not validation.valid:
            summary = "; ".join(item.message for item in validation.errors)
            raise ExportError(f"Document validation failed: {summary}")
        path = Path(target)
        suffix = path.suffix.lower()
        path.parent.mkdir(parents=True, exist_ok=True)
        if suffix == ".json":
            path.write_text(self.to_json(), encoding="utf-8")
        elif suffix in {".md", ".markdown"}:
            path.write_text(export_markdown(self._spec), encoding="utf-8")
        elif suffix in {".html", ".htm"}:
            path.write_text(export_html(self._spec), encoding="utf-8")
        elif suffix == ".docx":
            export_docx(self._spec, path)
        else:
            raise UnsupportedFormatError(
                f"Unsupported export format {suffix or '<none>'!r}; "
                "use .json, .md, .html, or .docx."
            )
        return path

    def apply(
        self,
        operations: Sequence[Mapping[str, Any]],
        *,
        dry_run: bool = False,
        base_revision: int | None = None,
        idempotency_key: str | None = None,
    ) -> PatchResult:
        expected_revision = self.revision if base_revision is None else base_revision
        if expected_revision != self.revision:
            diagnostic = Diagnostic(
                severity=Severity.ERROR,
                code="REVISION_CONFLICT",
                message=(
                    f"Patch targets revision {expected_revision}, but the document is at "
                    f"revision {self.revision}."
                ),
                node_ids=[self.id],
                suggested_actions=[{"action": "refresh_artifact"}],
            )
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=[diagnostic],
                idempotency_key=idempotency_key,
            )
        if not operations:
            diagnostic = Diagnostic(
                severity=Severity.ERROR,
                code="INVALID_SPEC",
                message="A patch must contain at least one operation.",
                suggested_actions=[{"action": "add_operation"}],
            )
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=[diagnostic],
                idempotency_key=idempotency_key,
            )

        payload = self.to_spec()
        next_revision = self.revision + 1
        changes: list[dict[str, Any]] = []
        try:
            for operation in operations:
                changes.append(self._apply_operation(payload, dict(operation), next_revision))
            payload["artifact"]["revision"] = next_revision
            updated = Document.from_spec(payload)
            validation = updated.validate()
            if not validation.valid:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="VALIDATION_FAILED",
                        message="The patch result failed document validation.",
                        node_ids=[self.id],
                        suggested_actions=[
                            {
                                "action": "inspect_diagnostics",
                                "diagnostics": [
                                    item.model_dump(mode="json") for item in validation.errors
                                ],
                            }
                        ],
                    )
                )
        except _PatchFailure as error:
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=[error.diagnostic],
                idempotency_key=idempotency_key,
            )
        except SpecValidationError as error:
            return PatchResult(
                success=False,
                base_revision=self.revision,
                result_revision=self.revision,
                dry_run=dry_run,
                diagnostics=error.diagnostics,
                idempotency_key=idempotency_key,
            )

        return PatchResult(
            success=True,
            base_revision=self.revision,
            result_revision=next_revision,
            dry_run=dry_run,
            document=updated,
            changes=changes,
            diagnostics=validation.warnings,
            idempotency_key=idempotency_key,
        )

    @staticmethod
    def _normalize_new_block(candidate: dict[str, Any], next_revision: int) -> dict[str, Any]:
        candidate = deepcopy(candidate)
        candidate.setdefault("revision_added", next_revision)
        candidate.setdefault("revision_updated", next_revision)
        try:
            block = TypeAdapter(Block).validate_python(candidate)
        except ValidationError as error:
            details = _validation_error_diagnostics(error)
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="INVALID_SPEC",
                    message="Patch content is not a valid document block.",
                    suggested_actions=[
                        {
                            "action": "fix_content",
                            "diagnostics": [
                                item.model_dump(mode="json") for item in details
                            ],
                        }
                    ],
                )
            ) from error
        return block.model_dump(mode="json", exclude_none=True)

    @staticmethod
    def _target_id(target: Any) -> str:
        if not isinstance(target, str) or not target:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="INVALID_SPEC",
                    message="Patch target must be a node ID or #node_id selector.",
                )
            )
        return target[1:] if target.startswith("#") else target

    @staticmethod
    def _find_node(payload: dict[str, Any], target: Any) -> tuple[int, dict[str, Any]]:
        target_id = Document._target_id(target)
        matches = [
            (index, node)
            for index, node in enumerate(payload["content"])
            if node.get("id") == target_id
        ]
        if not matches:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="TARGET_NOT_FOUND",
                    message=f"No node matched #{target_id}.",
                    suggested_actions=[{"action": "inspect_nodes"}],
                )
            )
        if len(matches) > 1:
            raise _PatchFailure(
                Diagnostic(
                    severity=Severity.ERROR,
                    code="AMBIGUOUS_SELECTOR",
                    message=f"Multiple nodes matched #{target_id}.",
                    node_ids=[target_id],
                    suggested_actions=[{"action": "repair_duplicate_ids"}],
                )
            )
        return matches[0]

    @staticmethod
    def _apply_operation(
        payload: dict[str, Any], operation: dict[str, Any], next_revision: int
    ) -> dict[str, Any]:
        operation_name = operation.get("op")
        if operation_name == "text.replace":
            _, node = Document._find_node(payload, operation.get("target"))
            search = operation.get("search")
            replacement = operation.get("replacement")
            replace_all = operation.get("replace_all", False)
            if not isinstance(search, str) or not search:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="text.replace requires a non-empty search string.",
                    )
                )
            if not isinstance(replacement, str):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="text.replace requires a string replacement.",
                    )
                )
            count = 0
            if node["type"] == "heading" or node.get("text") is not None:
                old_text = node.get("text", "")
                count = old_text.count(search) if replace_all else int(search in old_text)
                node["text"] = old_text.replace(search, replacement, -1 if replace_all else 1)
            elif node["type"] == "paragraph":
                for span in node.get("content", []):
                    old_text = span["text"]
                    span_count = old_text.count(search) if replace_all else int(search in old_text)
                    if span_count:
                        span["text"] = old_text.replace(
                            search, replacement, -1 if replace_all else 1
                        )
                        count += span_count
                        if not replace_all:
                            break
            else:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="UNSUPPORTED_FEATURE",
                        message=f"text.replace does not support node type {node['type']!r}.",
                        node_ids=[node["id"]],
                    )
                )
            if count == 0:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="TARGET_NOT_FOUND",
                        message=f"Search text {search!r} was not found in node {node['id']!r}.",
                        node_ids=[node["id"]],
                        suggested_actions=[{"action": "inspect_node", "node_id": node["id"]}],
                    )
                )
            node["revision_updated"] = next_revision
            return {
                "operation": "text.replace",
                "node_ids": [node["id"]],
                "replacement_count": count,
            }

        if operation_name == "node.append":
            target = operation.get("target", "$")
            if target not in {"$", payload["artifact"]["id"], f"#{payload['artifact']['id']}"}:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="UNSUPPORTED_FEATURE",
                        message="V0.1 node.append supports only the document root target '$'.",
                    )
                )
            content = operation.get("content")
            if not isinstance(content, dict):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="node.append requires an object in content.",
                    )
                )
            candidate = Document._normalize_new_block(content, next_revision)
            payload["content"].append(candidate)
            return {
                "operation": "node.append",
                "created_nodes": [candidate["id"]],
            }

        if operation_name == "node.insert_after":
            index, node = Document._find_node(payload, operation.get("target"))
            content = operation.get("content")
            if not isinstance(content, dict):
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="node.insert_after requires an object in content.",
                    )
                )
            candidate = Document._normalize_new_block(content, next_revision)
            payload["content"].insert(index + 1, candidate)
            return {
                "operation": "node.insert_after",
                "after": node["id"],
                "created_nodes": [candidate["id"]],
            }

        if operation_name == "node.remove":
            index, node = Document._find_node(payload, operation.get("target"))
            payload["content"].pop(index)
            return {"operation": "node.remove", "removed_nodes": [node["id"]]}

        if operation_name == "node.update":
            _, node = Document._find_node(payload, operation.get("target"))
            changes = operation.get("changes")
            if not isinstance(changes, dict) or not changes:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message="node.update requires a non-empty changes object.",
                    )
                )
            immutable = {"id", "type", "revision_added", "revision_updated"}
            forbidden = sorted(immutable.intersection(changes))
            if forbidden:
                raise _PatchFailure(
                    Diagnostic(
                        severity=Severity.ERROR,
                        code="INVALID_SPEC",
                        message=f"node.update cannot change: {', '.join(forbidden)}.",
                        node_ids=[node["id"]],
                    )
                )
            node.update(deepcopy(changes))
            node["revision_updated"] = next_revision
            return {
                "operation": "node.update",
                "node_ids": [node["id"]],
                "fields": sorted(changes),
            }

        raise _PatchFailure(
            Diagnostic(
                severity=Severity.ERROR,
                code="UNSUPPORTED_FEATURE",
                message=(
                    f"Unsupported operation {operation_name!r}; V0.1 supports text.replace, "
                    "node.append, node.insert_after, node.remove, and node.update."
                ),
                suggested_actions=[{"action": "use_supported_operation"}],
            )
        )

    def __repr__(self) -> str:
        return (
            f"Document(id={self.id!r}, revision={self.revision}, "
            f"nodes={len(self._spec.content)})"
        )


def open_artifact(source: str | Path) -> Document:
    path = Path(source)
    suffix = path.suffix.lower()
    if suffix == ".json":
        return Document.from_json(path)
    if suffix in {".md", ".markdown"}:
        return Document.from_markdown(path)
    if suffix in {".docx", ".html", ".htm"}:
        raise UnsupportedFormatError(
            f"Importing {suffix} is planned for a later release; V0.1 imports .json and .md."
        )
    raise UnsupportedFormatError(
        f"Unsupported source format {suffix or '<none>'!r}; use .json or .md."
    )


__all__ = ["Document", "PatchResult", "open_artifact"]
