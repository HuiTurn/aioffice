"""Persistent semantic-to-native identities shared by embedded and workspace manifests."""

from __future__ import annotations

import hashlib
from collections.abc import Mapping, Sequence
from typing import Annotated, Any, Literal
from xml.etree import ElementTree as ET

from pydantic import (
    BaseModel,
    ConfigDict,
    Field,
    StringConstraints,
    ValidationError,
    model_validator,
)

from aioffice.core.diagnostics import Diagnostic, Severity
from aioffice.core.errors import NativePackageError
from aioffice.core.ids import new_id
from aioffice.native.xml import parse_xml
from aioffice.spec.models import AiOfficeDocumentSpec, NativeRef, NodeId

MANIFEST_VERSION = "0.1"
MANIFEST_NAMESPACE = "https://schemas.aioffice.dev/manifest/0.1"
MANIFEST_PART_URI = "/customXml/aioffice-manifest.xml"
MANIFEST_RELATIONSHIP_TYPE = (
    "https://schemas.aioffice.dev/relationships/aioffice-manifest"
)

ET.register_namespace("ao", MANIFEST_NAMESPACE)


def _q(local: str) -> str:
    return f"{{{MANIFEST_NAMESPACE}}}{local}"


def _fingerprint_payload(payloads: Sequence[bytes]) -> str:
    digest = hashlib.sha256()
    for payload in payloads:
        digest.update(len(payload).to_bytes(8, "big"))
        digest.update(payload)
    return "sha256:" + digest.hexdigest()


def fingerprint_elements(elements: Sequence[ET.Element]) -> str:
    return _fingerprint_payload(
        [ET.tostring(element, encoding="utf-8") for element in elements]
    )


def native_ref_for_elements(
    elements: Sequence[ET.Element],
    indices: Sequence[int],
    *,
    native_kind: str,
    native_id: str | None = None,
) -> NativeRef:
    if not elements or len(elements) != len(indices):
        raise ValueError("A native reference requires matching elements and indices.")
    normalized_indices = list(indices)
    if normalized_indices != sorted(set(normalized_indices)):
        raise ValueError("Native reference indices must be sorted and unique.")
    if len(normalized_indices) == 1:
        path_hint = f"/w:document/w:body/*[{normalized_indices[0] + 1}]"
    else:
        path_hint = (
            f"/w:document/w:body/*[{normalized_indices[0] + 1}"
            f"..{normalized_indices[-1] + 1}]"
        )
    return NativeRef(
        format="docx",
        part_uri="/word/document.xml",
        native_kind=native_kind,
        element_index=normalized_indices[0],
        element_indices=normalized_indices,
        path_hint=path_hint,
        native_id=native_id,
        fingerprint=fingerprint_elements(elements),
    )


class IdentityNode(BaseModel):
    model_config = ConfigDict(extra="forbid")

    node_id: NodeId
    node_type: str
    source_ref: NativeRef
    previous_fingerprint: Annotated[
        str,
        StringConstraints(pattern=r"^sha256:[a-fA-F0-9]{64}$"),
    ] | None = None
    next_fingerprint: Annotated[
        str,
        StringConstraints(pattern=r"^sha256:[a-fA-F0-9]{64}$"),
    ] | None = None


class IdentityManifest(BaseModel):
    model_config = ConfigDict(extra="forbid")

    manifest_version: Literal["0.1"] = MANIFEST_VERSION
    artifact_id: NodeId
    revision: int = Field(ge=1)
    spec_version: str
    format: Literal["docx", "xlsx", "pptx"]
    package_sha256: Annotated[
        str,
        StringConstraints(pattern=r"^[a-fA-F0-9]{64}$"),
    ] | None = None
    nodes: list[IdentityNode] = Field(default_factory=list)

    @model_validator(mode="after")
    def validate_nodes(self) -> "IdentityManifest":
        node_ids = [node.node_id for node in self.nodes]
        if len(node_ids) != len(set(node_ids)):
            raise ValueError("Identity manifest node IDs must be unique.")
        if any(node.source_ref.format != self.format for node in self.nodes):
            raise ValueError("Identity node format must match the manifest format.")
        return self


def build_identity_manifest(
    spec: AiOfficeDocumentSpec,
    *,
    refs: Mapping[str, NativeRef] | None = None,
    package_sha256: str | None = None,
) -> IdentityManifest:
    records: list[IdentityNode] = []
    for nodes in (spec.content, spec.sections):
        group: list[IdentityNode] = []
        for node in nodes:
            source_ref = refs.get(node.id) if refs is not None else node.source_ref
            if not isinstance(source_ref, NativeRef):
                continue
            group.append(
                IdentityNode(
                    node_id=node.id,
                    node_type=node.type,
                    source_ref=source_ref,
                )
            )
        for index, record in enumerate(group):
            if index:
                record.previous_fingerprint = group[index - 1].source_ref.fingerprint
            if index + 1 < len(group):
                record.next_fingerprint = group[index + 1].source_ref.fingerprint
        records.extend(group)
    return IdentityManifest(
        artifact_id=spec.artifact.id,
        revision=spec.artifact.revision,
        spec_version=spec.spec_version,
        format="docx",
        package_sha256=package_sha256,
        nodes=records,
    )


def serialize_identity_manifest(manifest: IdentityManifest) -> bytes:
    attributes = {
        "version": manifest.manifest_version,
        "artifactId": manifest.artifact_id,
        "revision": str(manifest.revision),
        "specVersion": manifest.spec_version,
        "format": manifest.format,
    }
    if manifest.package_sha256 is not None:
        attributes["packageSha256"] = manifest.package_sha256
    root = ET.Element(_q("manifest"), attributes)
    for record in manifest.nodes:
        source_ref = record.source_ref
        values = {
            "id": record.node_id,
            "type": record.node_type,
            "partUri": source_ref.part_uri,
            "nativeKind": source_ref.native_kind,
        }
        if source_ref.element_index is not None:
            values["elementIndex"] = str(source_ref.element_index)
        if source_ref.element_indices:
            values["elementIndices"] = " ".join(
                str(index) for index in source_ref.element_indices
            )
        if source_ref.path_hint is not None:
            values["pathHint"] = source_ref.path_hint
        if source_ref.native_id is not None:
            values["nativeId"] = source_ref.native_id
        if source_ref.fingerprint is not None:
            values["fingerprint"] = source_ref.fingerprint
        if record.previous_fingerprint is not None:
            values["previousFingerprint"] = record.previous_fingerprint
        if record.next_fingerprint is not None:
            values["nextFingerprint"] = record.next_fingerprint
        ET.SubElement(root, _q("node"), values)
    return ET.tostring(root, encoding="utf-8", xml_declaration=True)


def parse_identity_manifest(payload: bytes) -> IdentityManifest:
    root = parse_xml(payload)
    if root.tag != _q("manifest"):
        raise NativePackageError("AiOffice identity manifest has an invalid root element.")
    nodes: list[dict[str, Any]] = []
    for element in root:
        if element.tag != _q("node"):
            continue
        raw_indices = element.attrib.get("elementIndices", "")
        try:
            indices = [int(value) for value in raw_indices.split()] if raw_indices else []
            element_index = (
                int(element.attrib["elementIndex"])
                if "elementIndex" in element.attrib
                else None
            )
        except ValueError as error:
            raise NativePackageError(
                "AiOffice identity manifest contains an invalid element index."
            ) from error
        nodes.append(
            {
                "node_id": element.attrib.get("id", ""),
                "node_type": element.attrib.get("type", ""),
                "source_ref": {
                    "format": root.attrib.get("format", "docx"),
                    "part_uri": element.attrib.get("partUri", ""),
                    "native_kind": element.attrib.get("nativeKind", ""),
                    "element_index": element_index,
                    "element_indices": indices,
                    "path_hint": element.attrib.get("pathHint"),
                    "native_id": element.attrib.get("nativeId"),
                    "fingerprint": element.attrib.get("fingerprint"),
                },
                "previous_fingerprint": element.attrib.get("previousFingerprint"),
                "next_fingerprint": element.attrib.get("nextFingerprint"),
            }
        )
    try:
        return IdentityManifest.model_validate(
            {
                "manifest_version": root.attrib.get("version"),
                "artifact_id": root.attrib.get("artifactId", ""),
                "revision": root.attrib.get("revision"),
                "spec_version": root.attrib.get("specVersion", ""),
                "format": root.attrib.get("format"),
                "package_sha256": root.attrib.get("packageSha256"),
                "nodes": nodes,
            }
        )
    except ValidationError as error:
        raise NativePackageError(f"Invalid AiOffice identity manifest: {error}") from error


def _candidate_ref(candidate: Mapping[str, Any]) -> NativeRef | None:
    value = candidate.get("source_ref")
    try:
        return NativeRef.model_validate(value) if isinstance(value, Mapping) else None
    except ValidationError:
        return None


def apply_identity_manifest(
    content: list[dict[str, Any]],
    manifest: IdentityManifest,
    *,
    package_sha256: str,
    sections: list[dict[str, Any]] | None = None,
) -> list[Diagnostic]:
    diagnostics: list[Diagnostic] = []
    projected = [*content, *(sections or [])]
    candidates = [
        (index, candidate, _candidate_ref(candidate))
        for index, candidate in enumerate(projected)
    ]
    candidates.sort(
        key=lambda item: (
            (
                item[2].element_index
                if item[2] is not None and item[2].element_index is not None
                else 2**31
            ),
            item[2].native_kind if item[2] is not None else "",
            item[0],
        )
    )
    candidates = [
        (index, candidate, source_ref)
        for index, (_, candidate, source_ref) in enumerate(candidates)
    ]
    used: set[int] = set()
    bound_records: dict[int, int] = {}
    unresolved: list[tuple[int, IdentityNode, int]] = []
    exact_package = manifest.package_sha256 == package_sha256

    for record_index, record in enumerate(manifest.nodes):
        source_ref = record.source_ref
        matches: list[tuple[int, dict[str, Any], NativeRef]] = []

        if exact_package:
            matches = [
                (index, candidate, candidate_ref)
                for index, candidate, candidate_ref in candidates
                if candidate_ref is not None
                and index not in used
                and candidate_ref.native_kind == source_ref.native_kind
                and candidate_ref.element_index == source_ref.element_index
                and candidate_ref.element_indices == source_ref.element_indices
            ]

        if not matches and source_ref.native_id:
            matches = [
                (index, candidate, candidate_ref)
                for index, candidate, candidate_ref in candidates
                if candidate_ref is not None
                and index not in used
                and candidate_ref.native_kind == source_ref.native_kind
                and candidate_ref.native_id == source_ref.native_id
            ]

        if not matches and source_ref.fingerprint:
            matches = [
                (index, candidate, candidate_ref)
                for index, candidate, candidate_ref in candidates
                if candidate_ref is not None
                and index not in used
                and candidate_ref.native_kind == source_ref.native_kind
                and candidate_ref.fingerprint == source_ref.fingerprint
            ]

        if not matches and source_ref.element_index is not None:
            path_candidates = [
                (index, candidate, candidate_ref)
                for index, candidate, candidate_ref in candidates
                if candidate_ref is not None
                and index not in used
                and candidate_ref.native_kind == source_ref.native_kind
                and candidate_ref.element_index == source_ref.element_index
            ]
            for match in path_candidates:
                index = match[0]
                previous_ref = candidates[index - 1][2] if index > 0 else None
                next_ref = (
                    candidates[index + 1][2]
                    if index + 1 < len(candidates)
                    else None
                )
                previous_matches = (
                    record.previous_fingerprint is not None
                    and previous_ref is not None
                    and previous_ref.fingerprint == record.previous_fingerprint
                )
                next_matches = (
                    record.next_fingerprint is not None
                    and next_ref is not None
                    and next_ref.fingerprint == record.next_fingerprint
                )
                has_previous = record.previous_fingerprint is not None
                has_next = record.next_fingerprint is not None
                neighbors_confirm = (
                    has_previous
                    and has_next
                    and previous_matches
                    and next_matches
                ) or (
                    has_previous
                    and not has_next
                    and previous_matches
                ) or (
                    has_next
                    and not has_previous
                    and next_matches
                )
                if neighbors_confirm:
                    matches.append(match)

        if len(matches) == 1:
            index, candidate, _ = matches[0]
            candidate["id"] = record.node_id
            used.add(index)
            bound_records[record_index] = index
            continue

        unresolved.append((record_index, record, len(matches)))

    for record_index, record, match_count in unresolved:
        previous_bound = next(
            (
                bound_records[index]
                for index in range(record_index - 1, -1, -1)
                if index in bound_records
            ),
            None,
        )
        next_bound = next(
            (
                bound_records[index]
                for index in range(record_index + 1, len(manifest.nodes))
                if index in bound_records
            ),
            None,
        )
        safely_deleted = match_count == 0 and (
            (
                previous_bound is not None
                and next_bound is not None
                and next_bound == previous_bound + 1
            )
            or (
                previous_bound is None
                and next_bound == 0
            )
            or (
                next_bound is None
                and previous_bound == len(candidates) - 1
            )
        )
        if safely_deleted:
            continue

        diagnostics.append(
            Diagnostic(
                severity=Severity.WARNING,
                code="IDENTITY_AMBIGUOUS",
                message=(
                    f"Could not safely bind native content to semantic node "
                    f"{record.node_id!r}; {match_count} candidates matched."
                ),
                node_ids=[record.node_id],
                recoverable=True,
                suggested_actions=[
                    {"action": "inspect_identity_candidates"},
                    {"action": "assign_new_node_id"},
                ],
            )
        )
        for index, candidate, _ in candidates:
            if index not in used and candidate.get("id") == record.node_id:
                candidate["id"] = new_id(
                    "section" if candidate.get("type") == "section" else "node"
                )

    id_owners: dict[str, list[int]] = {}
    for index, candidate, _ in candidates:
        candidate_id = candidate.get("id")
        if isinstance(candidate_id, str):
            id_owners.setdefault(candidate_id, []).append(index)
    for candidate_id, owners in id_owners.items():
        if len(owners) < 2:
            continue
        authoritative = [index for index in owners if index in used]
        if len(authoritative) > 1:
            raise NativePackageError(
                f"Identity manifest bound multiple native nodes to {candidate_id!r}."
            )
        preserved = authoritative[0] if authoritative else owners[0]
        for index in owners:
            if index != preserved:
                candidates[index][1]["id"] = new_id(
                    "section"
                    if candidates[index][1].get("type") == "section"
                    else "node"
                )
    return diagnostics


__all__ = [
    "IdentityManifest",
    "IdentityNode",
    "MANIFEST_NAMESPACE",
    "MANIFEST_PART_URI",
    "MANIFEST_RELATIONSHIP_TYPE",
    "MANIFEST_VERSION",
    "apply_identity_manifest",
    "build_identity_manifest",
    "fingerprint_elements",
    "native_ref_for_elements",
    "parse_identity_manifest",
    "serialize_identity_manifest",
]
