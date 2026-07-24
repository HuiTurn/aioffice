"""Local artifact and revision store for safe, repeatable AI editing sessions."""

from __future__ import annotations

import hashlib
import json
import os
import tempfile
from collections.abc import Mapping, Sequence
from copy import deepcopy
from pathlib import Path
from typing import Any

from pydantic import ValidationError

from aioffice.core.diagnostics import Diagnostic
from aioffice.core.diff import DocumentDiff
from aioffice.core.errors import SecurityError, WorkspaceError
from aioffice.core.ids import new_id
from aioffice.documents.assets import ImageAsset
from aioffice.documents.document import Document, PatchResult
from aioffice.native import (
    FidelityPolicy,
    FidelityReport,
    IdentityManifest,
    build_identity_manifest,
)
from aioffice.security import SecurityPolicy
from aioffice.spec.models import (
    FloatingImageLayout,
    ImageOutline,
    ImageTransform,
    Length,
    ParagraphStyle,
)

from .models import ArtifactEntry, PatchRecord, WorkspaceIndex

_MAX_METADATA_BYTES = 16 * 1024 * 1024


def _json_bytes(value: Any) -> bytes:
    return (
        json.dumps(
            value,
            ensure_ascii=False,
            indent=2,
            sort_keys=True,
        )
        + "\n"
    ).encode("utf-8")


def _atomic_write(path: Path, payload: bytes) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{path.name}.",
        suffix=".tmp",
        dir=path.parent,
    )
    temporary = Path(temporary_name)
    try:
        with os.fdopen(descriptor, "wb") as handle:
            handle.write(payload)
            handle.flush()
            os.fsync(handle.fileno())
        os.replace(temporary, path)
    except Exception:
        temporary.unlink(missing_ok=True)
        raise


def _read_json(path: Path) -> Any:
    try:
        size = path.stat().st_size
    except FileNotFoundError as error:
        raise WorkspaceError(f"Workspace metadata {path} does not exist.") from error
    if size > _MAX_METADATA_BYTES:
        raise WorkspaceError(f"Workspace metadata {path} exceeds 16 MB.")
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except (UnicodeDecodeError, json.JSONDecodeError) as error:
        raise WorkspaceError(f"Workspace metadata {path} is invalid JSON.") from error


class Workspace:
    """A bounded local store containing native revisions, specs, manifests, and patches."""

    def __init__(self, root: Path, index: WorkspaceIndex) -> None:
        self.root = root.resolve()
        self.state_dir = self.root / ".aioffice"
        self._index = index

    @classmethod
    def init(cls, root: str | Path = ".") -> "Workspace":
        project_root = Path(root).resolve()
        project_root.mkdir(parents=True, exist_ok=True)
        state_dir = project_root / ".aioffice"
        index_path = state_dir / "workspace.json"
        if index_path.exists():
            return cls.open(project_root)
        if state_dir.exists() and any(state_dir.iterdir()):
            raise WorkspaceError(
                f"Refusing to initialize non-empty workspace directory {state_dir}."
            )
        state_dir.mkdir(parents=True, exist_ok=True)
        index = WorkspaceIndex(workspace_id=new_id("workspace"))
        _atomic_write(index_path, _json_bytes(index.model_dump(mode="json")))
        return cls(project_root, index)

    @classmethod
    def open(cls, root: str | Path = ".") -> "Workspace":
        project_root = Path(root).resolve()
        state_dir = project_root / ".aioffice"
        if not state_dir.resolve().is_relative_to(project_root):
            raise WorkspaceError("Workspace state directory escapes the project root.")
        payload = _read_json(state_dir / "workspace.json")
        try:
            index = WorkspaceIndex.model_validate(payload)
        except ValidationError as error:
            raise WorkspaceError(f"Invalid workspace index: {error}") from error
        for artifact_id, entry in index.artifacts.items():
            if artifact_id != entry.artifact_id:
                raise WorkspaceError(
                    f"Workspace artifact key {artifact_id!r} does not match its entry."
                )
        return cls(project_root, index)

    @property
    def id(self) -> str:
        return self._index.workspace_id

    def list_artifacts(self) -> list[dict[str, Any]]:
        return [entry.model_dump(mode="json") for _, entry in sorted(self._index.artifacts.items())]

    def capabilities(self, artifact_id: str | None = None) -> dict[str, Any]:
        result: dict[str, Any] = {
            "workspace_id": self.id,
            "formats": ["docx"],
            "operations": [
                "import",
                "inspect",
                "apply",
                "insert_image_after",
                "replace_image",
                "checkout",
                "reconcile",
                "export",
            ],
            "patch_operations": [
                "text.replace",
                "paragraph.format",
                "text.format",
                "node.append",
                "node.insert_after",
                "node.insert_before",
                "node.move_after",
                "node.move_before",
                "node.remove",
                "style.apply",
                "style.define",
                "style.format",
                "header_footer.create",
                "header_footer.clone",
                "section.header_footer.bind",
                "section.insert_before",
                "section.format",
                "field.update",
                "image.anchor.update",
                "image.update",
                "table.format",
                "table.column.format",
                "table.cell.format",
            ],
            "binary_operations": {
                "image.replace": {
                    "api": "Workspace.replace_image",
                    "transport": "out_of_band",
                    "recorded_binary": False,
                },
                "image.insert_after": {
                    "api": "Workspace.insert_image_after",
                    "transport": "out_of_band",
                    "recorded_binary": False,
                    "placements": [
                        "inline",
                        (
                            "floating_offset_alignment_or_percentage_"
                            "supported_wrap"
                        ),
                    ],
                    "default_placement": "inline",
                }
            },
            "revision_store": True,
            "idempotency": True,
            "external_reconciliation": True,
            "overwrite_by_default": False,
        }
        if artifact_id is not None:
            entry = self._entry(artifact_id)
            result["artifact"] = entry.model_dump(mode="json")
        return result

    def import_document(
        self,
        source: str | Path,
        *,
        roundtrip: FidelityPolicy | str = FidelityPolicy.PRESERVE_UNKNOWN,
        security_policy: SecurityPolicy | None = None,
    ) -> Document:
        source_path = Path(source).resolve()
        if source_path.suffix.lower() != ".docx":
            raise WorkspaceError("V0.2 workspaces currently import DOCX files only.")
        active_security_policy = security_policy or SecurityPolicy()
        source_size = source_path.stat().st_size
        if source_size > active_security_policy.max_file_size_bytes:
            raise SecurityError(
                f"Input package exceeds {active_security_policy.max_file_size_mb} MB."
            )
        source_bytes = source_path.read_bytes()
        document = Document.from_docx(
            source_bytes,
            roundtrip=roundtrip,
            security_policy=active_security_policy,
        )
        existing = self._index.artifacts.get(document.id)
        if existing is not None:
            existing_manifest = self._load_identity_manifest(
                document.id,
                existing.latest_revision,
            )
            source_sha256 = hashlib.sha256(source_bytes).hexdigest()
            if existing_manifest.package_sha256 == source_sha256:
                return self.open_document(document.id)
            raise WorkspaceError(
                f"Artifact {document.id!r} already exists with different native content."
            )
        self._persist_revision(
            document,
            source_bytes,
            source_name=source_path.name,
            patch=None,
            expected_base_revision=None,
        )
        return self.open_document(document.id)

    def open_document(
        self,
        artifact_id: str,
        *,
        revision: int | None = None,
        roundtrip: FidelityPolicy | str = FidelityPolicy.PRESERVE_UNKNOWN,
        security_policy: SecurityPolicy | None = None,
    ) -> Document:
        entry = self._entry(artifact_id)
        selected_revision = entry.latest_revision if revision is None else revision
        if selected_revision < entry.first_revision or selected_revision > entry.latest_revision:
            raise WorkspaceError(
                f"Revision {selected_revision} is outside artifact {artifact_id!r} history."
            )
        native_path = self._revision_path(artifact_id, selected_revision)
        identity = self._load_identity_manifest(artifact_id, selected_revision)
        document = Document.from_docx(
            native_path,
            roundtrip=roundtrip,
            security_policy=security_policy,
            identity_manifest=identity,
        )
        if document.id != artifact_id or document.revision != selected_revision:
            raise WorkspaceError(
                "Workspace identity manifest did not restore the requested revision."
            )
        return document

    checkout = open_document

    def apply(
        self,
        artifact_id: str,
        operations: Sequence[Mapping[str, Any]],
        *,
        dry_run: bool = False,
        base_revision: int | None = None,
        idempotency_key: str | None = None,
    ) -> PatchResult:
        normalized_operations = [deepcopy(dict(operation)) for operation in operations]
        if idempotency_key == "":
            raise WorkspaceError("Idempotency keys cannot be empty.")
        if idempotency_key is not None:
            replay = self._idempotent_result(
                artifact_id,
                idempotency_key,
                normalized_operations,
            )
            if replay is not None:
                return replay
        entry = self._entry(artifact_id)
        expected_revision = entry.latest_revision if base_revision is None else base_revision
        document = self.open_document(artifact_id)
        result = document.apply(
            normalized_operations,
            dry_run=dry_run,
            base_revision=expected_revision,
            idempotency_key=idempotency_key,
        )
        if not result.success or dry_run:
            return result
        assert result.document is not None
        patch = PatchRecord(
            base_revision=result.base_revision,
            result_revision=result.result_revision,
            idempotency_key=idempotency_key,
            operations=normalized_operations,
            changes=deepcopy(result.changes),
            diagnostics=[diagnostic.model_dump(mode="json") for diagnostic in result.diagnostics],
            fidelity=(
                result.fidelity.model_dump(mode="json") if result.fidelity is not None else None
            ),
            diff=(result.diff.model_dump(mode="json") if result.diff is not None else None),
        )
        self._persist_revision(
            result.document,
            result.document.to_bytes("docx"),
            source_name=entry.source_name,
            patch=patch,
            expected_base_revision=result.base_revision,
        )
        return result

    def replace_image(
        self,
        artifact_id: str,
        image_id: str,
        source: bytes | bytearray | memoryview | str | Path | ImageAsset,
        *,
        media_type: str | None = None,
        dry_run: bool = False,
        base_revision: int | None = None,
    ) -> PatchResult:
        """Replace one image and persist only verified asset metadata in the log."""

        entry = self._entry(artifact_id)
        expected_revision = (
            entry.latest_revision
            if base_revision is None
            else base_revision
        )
        document = self.open_document(artifact_id)
        result = document.replace_image(
            image_id,
            source,
            media_type=media_type,
            dry_run=dry_run,
            base_revision=expected_revision,
        )
        if not result.success or dry_run:
            return result
        assert result.document is not None
        asset_change = result.changes[0].get("asset_change", {})
        replacement_asset = asset_change.get("after")
        if not isinstance(replacement_asset, dict):
            raise WorkspaceError(
                "Image replacement did not return verified asset metadata."
            )
        operation = {
            "op": "image.replace",
            "target": image_id,
            "asset": deepcopy(replacement_asset),
        }
        patch = PatchRecord(
            base_revision=result.base_revision,
            result_revision=result.result_revision,
            operations=[operation],
            changes=deepcopy(result.changes),
            diagnostics=[
                diagnostic.model_dump(mode="json")
                for diagnostic in result.diagnostics
            ],
            fidelity=(
                result.fidelity.model_dump(mode="json")
                if result.fidelity is not None
                else None
            ),
            diff=(
                result.diff.model_dump(mode="json")
                if result.diff is not None
                else None
            ),
        )
        self._persist_revision(
            result.document,
            result.document.to_bytes("docx"),
            source_name=entry.source_name,
            patch=patch,
            expected_base_revision=result.base_revision,
        )
        return result

    def insert_image_after(
        self,
        artifact_id: str,
        target: str,
        source: bytes | bytearray | memoryview | str | Path | ImageAsset,
        *,
        width: Length | Mapping[str, Any],
        height: Length | Mapping[str, Any],
        alt_text: str,
        media_type: str | None = None,
        image_id: str | None = None,
        name: str | None = None,
        title: str | None = None,
        transform: ImageTransform | Mapping[str, Any] | None = None,
        outline: ImageOutline | Mapping[str, Any] | None = None,
        floating: FloatingImageLayout | Mapping[str, Any] | None = None,
        paragraph_style: ParagraphStyle | Mapping[str, Any] | None = None,
        dry_run: bool = False,
        base_revision: int | None = None,
    ) -> PatchResult:
        """Insert one image and persist only verified metadata in the log."""

        entry = self._entry(artifact_id)
        expected_revision = (
            entry.latest_revision
            if base_revision is None
            else base_revision
        )
        document = self.open_document(artifact_id)
        result = document.insert_image_after(
            target,
            source,
            width=width,
            height=height,
            alt_text=alt_text,
            media_type=media_type,
            image_id=image_id,
            name=name,
            title=title,
            transform=transform,
            outline=outline,
            floating=floating,
            paragraph_style=paragraph_style,
            dry_run=dry_run,
            base_revision=expected_revision,
        )
        if not result.success or dry_run:
            return result
        assert result.document is not None
        created_ids = result.changes[0].get("created_nodes", [])
        if (
            not isinstance(created_ids, list)
            or len(created_ids) != 1
            or not isinstance(created_ids[0], str)
        ):
            raise WorkspaceError(
                "Image insertion did not return one created image ID."
            )
        created_id = created_ids[0]
        result_spec = result.document.to_spec()
        inserted = next(
            (
                node
                for node in result_spec["content"]
                if node.get("id") == created_id
                and node.get("type") == "image"
            ),
            None,
        )
        if inserted is None:
            raise WorkspaceError(
                "Image insertion result does not contain its created node."
            )
        asset = next(
            (
                candidate
                for candidate in result_spec.get("assets", [])
                if candidate.get("id") == inserted.get("asset_id")
            ),
            None,
        )
        if asset is None:
            raise WorkspaceError(
                "Image insertion result does not contain its verified asset."
            )
        image_metadata = {
            field_name: deepcopy(inserted[field_name])
            for field_name in (
                "id",
                "placement",
                "floating",
                "width",
                "height",
                "transform",
                "outline",
                "name",
                "alt_text",
                "title",
                "paragraph_style",
            )
            if field_name in inserted
        }
        operation = {
            "op": "image.insert_after",
            "target": target,
            "image": image_metadata,
            "asset": deepcopy(asset),
        }
        patch = PatchRecord(
            base_revision=result.base_revision,
            result_revision=result.result_revision,
            operations=[operation],
            changes=deepcopy(result.changes),
            diagnostics=[
                diagnostic.model_dump(mode="json")
                for diagnostic in result.diagnostics
            ],
            fidelity=(
                result.fidelity.model_dump(mode="json")
                if result.fidelity is not None
                else None
            ),
            diff=(
                result.diff.model_dump(mode="json")
                if result.diff is not None
                else None
            ),
        )
        self._persist_revision(
            result.document,
            result.document.to_bytes("docx"),
            source_name=entry.source_name,
            patch=patch,
            expected_base_revision=result.base_revision,
        )
        return result

    def reconcile_document(
        self,
        artifact_id: str,
        source: str | Path,
        *,
        commit: bool = False,
        roundtrip: FidelityPolicy | str = FidelityPolicy.PRESERVE_UNKNOWN,
        security_policy: SecurityPolicy | None = None,
    ) -> Document:
        entry = self._entry(artifact_id)
        source_path = Path(source).resolve()
        if source_path.suffix.lower() != ".docx":
            raise WorkspaceError("V0.2 workspaces currently reconcile DOCX files only.")
        active_security_policy = security_policy or SecurityPolicy()
        source_size = source_path.stat().st_size
        if source_size > active_security_policy.max_file_size_bytes:
            raise SecurityError(
                f"Input package exceeds {active_security_policy.max_file_size_mb} MB."
            )
        source_bytes = source_path.read_bytes()
        source_sha256 = hashlib.sha256(source_bytes).hexdigest()
        current_identity = self._load_identity_manifest(
            artifact_id,
            entry.latest_revision,
        )
        if current_identity.package_sha256 == source_sha256:
            return self.open_document(artifact_id)

        matching_identity = current_identity.model_copy(
            update={"revision": entry.latest_revision + 1},
            deep=True,
        )
        preview = Document.from_docx(
            source_bytes,
            roundtrip=roundtrip,
            security_policy=active_security_policy,
            identity_manifest=matching_identity,
        )
        ambiguous = [
            diagnostic
            for diagnostic in preview.import_diagnostics
            if diagnostic.code == "IDENTITY_AMBIGUOUS"
        ]
        if commit and ambiguous:
            node_ids = sorted(
                {node_id for diagnostic in ambiguous for node_id in diagnostic.node_ids}
            )
            raise WorkspaceError(
                "External edit has ambiguous native identities; refusing to commit "
                f"nodes: {', '.join(node_ids)}."
            )
        if not commit:
            return preview

        synchronized = preview.synchronize_identity_manifest()
        patch = PatchRecord(
            base_revision=entry.latest_revision,
            result_revision=entry.latest_revision + 1,
            operations=[
                {
                    "op": "native.reconcile",
                    "source_sha256": source_sha256,
                }
            ],
            changes=[
                {
                    "operation": "native.reconcile",
                    "source": source_path.name,
                }
            ],
            diagnostics=[
                diagnostic.model_dump(mode="json") for diagnostic in synchronized.import_diagnostics
            ],
            fidelity=(
                synchronized.fidelity.model_dump(mode="json")
                if synchronized.fidelity is not None
                else None
            ),
        )
        self._persist_revision(
            synchronized,
            synchronized.to_bytes("docx"),
            source_name=source_path.name,
            patch=patch,
            expected_base_revision=entry.latest_revision,
        )
        return self.open_document(artifact_id)

    def export_document(
        self,
        artifact_id: str,
        target: str | Path,
        *,
        revision: int | None = None,
        overwrite: bool = False,
    ) -> Path:
        path = Path(target)
        if path.exists() and not overwrite:
            raise WorkspaceError(
                f"Refusing to overwrite existing export {path}; pass overwrite=True."
            )
        document = self.open_document(artifact_id, revision=revision)
        _atomic_write(path, document.to_bytes(path.suffix or "docx"))
        return path

    def _entry(self, artifact_id: str) -> ArtifactEntry:
        try:
            return self._index.artifacts[artifact_id]
        except KeyError as error:
            raise WorkspaceError(f"Artifact {artifact_id!r} is not in this workspace.") from error

    def _artifact_dir(self, artifact_id: str) -> Path:
        if "/" in artifact_id or "\\" in artifact_id or "\x00" in artifact_id:
            raise WorkspaceError("Unsafe artifact ID in workspace path.")
        path = self.state_dir / "artifacts" / artifact_id
        if not path.resolve().is_relative_to(self.state_dir.resolve()):
            raise WorkspaceError("Artifact path escapes the workspace.")
        return path

    def _revision_path(self, artifact_id: str, revision: int) -> Path:
        return self._artifact_dir(artifact_id) / "revisions" / f"{revision:08d}.docx"

    def _manifest_path(self, artifact_id: str, revision: int) -> Path:
        return self._artifact_dir(artifact_id) / "manifests" / f"{revision:08d}.json"

    def _snapshot_path(self, artifact_id: str, revision: int) -> Path:
        return self._artifact_dir(artifact_id) / "snapshots" / f"{revision:08d}.json"

    def _patch_path(self, artifact_id: str, revision: int) -> Path:
        return self._artifact_dir(artifact_id) / "patches" / f"{revision:08d}.json"

    def _load_identity_manifest(
        self,
        artifact_id: str,
        revision: int,
    ) -> IdentityManifest:
        payload = _read_json(self._manifest_path(artifact_id, revision))
        try:
            return IdentityManifest.model_validate(payload)
        except ValidationError as error:
            raise WorkspaceError(f"Invalid artifact identity manifest: {error}") from error

    def _persist_revision(
        self,
        document: Document,
        native_bytes: bytes,
        *,
        source_name: str,
        patch: PatchRecord | None,
        expected_base_revision: int | None,
    ) -> None:
        revision = document.revision
        current = self._index.artifacts.get(document.id)
        if expected_base_revision is None:
            disk_index = Workspace.open(self.root)._index
            self._index = disk_index
            current = disk_index.artifacts.get(document.id)
            if current is not None:
                raise WorkspaceError(f"Artifact {document.id!r} already exists.")
        else:
            disk_index = Workspace.open(self.root)._index
            disk_entry = disk_index.artifacts.get(document.id)
            if (
                disk_entry is None
                or disk_entry.latest_revision != expected_base_revision
                or revision != expected_base_revision + 1
            ):
                raise WorkspaceError(
                    "Workspace revision changed concurrently; refresh before committing."
                )
            self._index = disk_index
            current = disk_entry

        package_sha256 = hashlib.sha256(native_bytes).hexdigest()
        identity = build_identity_manifest(
            document.spec,
            package_sha256=package_sha256,
        )
        _atomic_write(self._revision_path(document.id, revision), native_bytes)
        _atomic_write(
            self._manifest_path(document.id, revision),
            _json_bytes(identity.model_dump(mode="json", exclude_none=True)),
        )
        _atomic_write(
            self._snapshot_path(document.id, revision),
            document.to_bytes("json"),
        )
        if patch is not None:
            _atomic_write(
                self._patch_path(document.id, revision),
                _json_bytes(patch.model_dump(mode="json", exclude_none=True)),
            )
        artifact_dir = self._artifact_dir(document.id)
        _atomic_write(
            artifact_dir / "manifest.json",
            _json_bytes(identity.model_dump(mode="json", exclude_none=True)),
        )

        self._index.artifacts[document.id] = ArtifactEntry(
            artifact_id=document.id,
            first_revision=(current.first_revision if current is not None else revision),
            latest_revision=revision,
            source_name=source_name,
        )
        _atomic_write(
            self.state_dir / "workspace.json",
            _json_bytes(self._index.model_dump(mode="json")),
        )

    def _idempotent_result(
        self,
        artifact_id: str,
        idempotency_key: str,
        operations: list[dict[str, Any]],
    ) -> PatchResult | None:
        entry = self._entry(artifact_id)
        for revision in range(entry.first_revision + 1, entry.latest_revision + 1):
            path = self._patch_path(artifact_id, revision)
            if not path.exists():
                continue
            try:
                payload = PatchRecord.model_validate(_read_json(path))
            except ValidationError as error:
                raise WorkspaceError(f"Invalid workspace patch record: {error}") from error
            if payload.idempotency_key != idempotency_key:
                continue
            if payload.operations != operations:
                raise WorkspaceError(
                    f"Idempotency key {idempotency_key!r} was already used for a different patch."
                )
            diagnostics = [Diagnostic.model_validate(value) for value in payload.diagnostics]
            fidelity_value = payload.fidelity
            fidelity = (
                FidelityReport.model_validate(fidelity_value)
                if fidelity_value is not None
                else None
            )
            diff_value = payload.diff
            diff = DocumentDiff.model_validate(diff_value) if diff_value is not None else None
            return PatchResult(
                success=True,
                base_revision=payload.base_revision,
                result_revision=payload.result_revision,
                dry_run=False,
                document=self.open_document(artifact_id, revision=revision),
                changes=deepcopy(payload.changes),
                diagnostics=diagnostics,
                idempotency_key=idempotency_key,
                fidelity=fidelity,
                diff=diff,
            )
        return None


__all__ = ["Workspace"]
