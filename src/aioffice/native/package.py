"""Safe, copy-on-write Open Packaging Conventions container."""

from __future__ import annotations

import copy
import hashlib
from io import BytesIO
from pathlib import Path, PurePosixPath
from zipfile import BadZipFile, ZipFile

from aioffice.core.errors import NativePackageError, SecurityError
from aioffice.security import SecurityPolicy

from .fidelity import FidelityLevel, FidelityPolicy, FidelityReport
from .types import NativePart, NativeRelationship
from .xml import parse_xml

def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def _part_uri(name: str) -> str:
    return "/" + name.lstrip("/")


def _safe_member_name(name: str) -> bool:
    if not name or "\x00" in name or "\\" in name:
        return False
    path = PurePosixPath(name)
    return not path.is_absolute() and ".." not in path.parts


def _relationship_source(path: str) -> str:
    if path == "_rels/.rels":
        return "/"
    parts = list(PurePosixPath(path).parts)
    try:
        marker = len(parts) - 1 - parts[::-1].index("_rels")
    except ValueError:
        return _part_uri(path)
    if marker + 1 >= len(parts):
        return _part_uri(path)
    filename = parts[marker + 1]
    if not filename.endswith(".rels"):
        return _part_uri(path)
    source_parts = parts[:marker] + [filename[: -len(".rels")]]
    return "/" + "/".join(source_parts)


class NativePackage:
    """An immutable-base OPC package with copy-on-write part overrides."""

    def __init__(
        self,
        source_bytes: bytes,
        *,
        format_name: str,
        policy: FidelityPolicy,
        security_policy: SecurityPolicy,
    ) -> None:
        self._source_bytes = source_bytes
        self.format_name = format_name
        self.policy = policy
        self.security_policy = security_policy
        self._overrides: dict[str, bytes] = {}
        self._deleted: set[str] = set()
        self._parts: dict[str, NativePart] = {}
        self._relationships: list[NativeRelationship] = []
        self._entry_names: list[str] = []
        self._inspect_package()

    @classmethod
    def open(
        cls,
        source: str | Path | bytes,
        *,
        format_name: str,
        policy: FidelityPolicy | str = FidelityPolicy.PRESERVE_UNKNOWN,
        security_policy: SecurityPolicy | None = None,
    ) -> "NativePackage":
        active_security_policy = security_policy or SecurityPolicy()
        if isinstance(source, bytes):
            source_bytes = source
        else:
            source_path = Path(source)
            source_size = source_path.stat().st_size
            if source_size > active_security_policy.max_file_size_bytes:
                raise SecurityError(
                    f"Package size {source_size} exceeds "
                    f"{active_security_policy.max_file_size_mb} MB."
                )
            source_bytes = source_path.read_bytes()
        if len(source_bytes) > active_security_policy.max_file_size_bytes:
            raise SecurityError(
                f"Package size {len(source_bytes)} exceeds "
                f"{active_security_policy.max_file_size_mb} MB."
            )
        return cls(
            source_bytes,
            format_name=format_name,
            policy=FidelityPolicy(policy),
            security_policy=active_security_policy,
        )

    @property
    def source_sha256(self) -> str:
        return hashlib.sha256(self._source_bytes).hexdigest()

    @property
    def parts(self) -> tuple[NativePart, ...]:
        return tuple(self._parts[uri] for uri in sorted(self._parts))

    @property
    def relationships(self) -> tuple[NativeRelationship, ...]:
        return tuple(self._relationships)

    @property
    def affected_parts(self) -> tuple[str, ...]:
        return tuple(sorted(set(self._overrides) | self._deleted))

    def has_part(self, uri: str) -> bool:
        normalized = _part_uri(uri)
        return normalized not in self._deleted and (
            normalized in self._overrides or normalized in self._parts
        )

    def clone(self) -> "NativePackage":
        cloned = object.__new__(NativePackage)
        cloned._source_bytes = self._source_bytes
        cloned.format_name = self.format_name
        cloned.policy = self.policy
        cloned.security_policy = self.security_policy
        cloned._overrides = dict(self._overrides)
        cloned._deleted = set(self._deleted)
        cloned._parts = dict(self._parts)
        cloned._relationships = list(self._relationships)
        cloned._entry_names = list(self._entry_names)
        return cloned

    def get_part(self, uri: str) -> bytes:
        normalized = _part_uri(uri)
        if normalized in self._deleted:
            raise NativePackageError(f"Part {normalized!r} was deleted.")
        if normalized in self._overrides:
            return self._overrides[normalized]
        name = normalized.lstrip("/")
        try:
            with ZipFile(BytesIO(self._source_bytes)) as archive:
                return archive.read(name)
        except KeyError as error:
            raise NativePackageError(f"Part {normalized!r} does not exist.") from error

    def set_part(
        self,
        uri: str,
        payload: bytes,
        *,
        content_type: str | None = None,
    ) -> None:
        normalized = _part_uri(uri)
        name = normalized.lstrip("/")
        if not _safe_member_name(name):
            raise NativePackageError(f"Unsafe part URI {uri!r}.")
        if not isinstance(payload, bytes):
            raise TypeError("Native part payload must be bytes.")
        if content_type is not None and (
            not content_type
            or "/" not in content_type
            or any(character.isspace() for character in content_type)
        ):
            raise NativePackageError(
                f"Invalid content type {content_type!r} for part {normalized!r}."
            )
        if name.lower().endswith((".xml", ".rels")) and len(
            payload
        ) > self.security_policy.max_xml_part_size_bytes:
            raise SecurityError(
                f"XML part {normalized!r} exceeds "
                f"{self.security_policy.max_xml_part_size_mb} MB."
            )
        existing = self._parts.get(normalized)
        active_part_count = sum(
            uri not in self._deleted
            for uri in self._parts
        )
        if (
            (
                existing is None
                or normalized in self._deleted
            )
            and normalized not in self._overrides
            and active_part_count >= self.security_policy.max_package_parts
        ):
            raise SecurityError(
                "Adding a native part would exceed the package part limit "
                f"of {self.security_policy.max_package_parts}."
            )
        prospective_uncompressed = sum(
            part.size
            for part_uri, part in self._parts.items()
            if part_uri not in self._deleted and part_uri != normalized
        ) + len(payload)
        if (
            prospective_uncompressed
            > self.security_policy.max_uncompressed_size_bytes
        ):
            raise SecurityError(
                "Adding the native part would exceed the package uncompressed "
                f"size limit of {self.security_policy.max_uncompressed_size_mb} MB."
            )
        replacement_relationships: list[NativeRelationship] | None = None
        relationship_source: str | None = None
        if name.endswith(".rels"):
            replacement_relationships = self._read_relationships(
                payload,
                name,
            )
            relationship_source = _relationship_source(name)

        self._overrides[normalized] = payload
        self._deleted.discard(normalized)
        self._parts[normalized] = NativePart(
            uri=normalized,
            content_type=(
                content_type
                or (
                    existing.content_type
                    if existing is not None
                    else "application/octet-stream"
                )
            ),
            sha256=hashlib.sha256(payload).hexdigest(),
            size=len(payload),
            compressed_size=0,
            state="modified" if existing is not None else "created",
        )
        if (
            replacement_relationships is not None
            and relationship_source is not None
        ):
            self._relationships = [
                relationship
                for relationship in self._relationships
                if relationship.source_part != relationship_source
            ]
            self._relationships.extend(replacement_relationships)

    def delete_part(self, uri: str) -> None:
        normalized = _part_uri(uri)
        if normalized not in self._parts and normalized not in self._overrides:
            raise NativePackageError(f"Part {normalized!r} does not exist.")
        self._deleted.add(normalized)
        self._overrides.pop(normalized, None)
        existing = self._parts.get(normalized)
        if existing is not None:
            self._parts[normalized] = NativePart(
                uri=existing.uri,
                content_type=existing.content_type,
                sha256=existing.sha256,
                size=existing.size,
                compressed_size=existing.compressed_size,
                state="deleted",
            )
        name = normalized.lstrip("/")
        if name.endswith(".rels"):
            relationship_source = _relationship_source(name)
            self._relationships = [
                relationship
                for relationship in self._relationships
                if relationship.source_part != relationship_source
            ]

    def export_bytes(self) -> bytes:
        if not self._overrides and not self._deleted:
            return self._source_bytes
        output = BytesIO()
        with ZipFile(BytesIO(self._source_bytes)) as source_archive, ZipFile(
            output, mode="w"
        ) as target_archive:
            written: set[str] = set()
            for source_info in source_archive.infolist():
                name = source_info.filename
                uri = _part_uri(name)
                if uri in self._deleted:
                    continue
                payload = self._overrides.get(uri)
                if payload is None:
                    payload = source_archive.read(name) if not source_info.is_dir() else b""
                target_info = copy.copy(source_info)
                target_archive.writestr(target_info, payload)
                written.add(uri)
            for uri, payload in sorted(self._overrides.items()):
                if uri in written or uri in self._deleted:
                    continue
                target_archive.writestr(uri.lstrip("/"), payload)
        return output.getvalue()

    def write(self, target: str | Path) -> Path:
        path = Path(target)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(self.export_bytes())
        return path

    def fidelity_report(self) -> FidelityReport:
        result = self.export_bytes()
        exact = not self._overrides and not self._deleted
        return FidelityReport(
            policy=self.policy,
            level=FidelityLevel.EXACT_PACKAGE if exact else FidelityLevel.STRUCTURAL,
            source_sha256=self.source_sha256,
            result_sha256=hashlib.sha256(result).hexdigest(),
            affected_parts=list(self.affected_parts),
            untouched_parts=max(0, len(self._parts) - len(self.affected_parts)),
            visual_verification_required=not exact,
        )

    def _inspect_package(self) -> None:
        try:
            archive = ZipFile(BytesIO(self._source_bytes))
        except BadZipFile as error:
            raise NativePackageError("Input is not a valid ZIP/OPC package.") from error
        with archive:
            infos = archive.infolist()
            if len(infos) > self.security_policy.max_package_parts:
                raise SecurityError(
                    f"Package has {len(infos)} entries; limit is "
                    f"{self.security_policy.max_package_parts}."
                )
            seen: set[str] = set()
            total_uncompressed = 0
            for info in infos:
                name = info.filename
                if not _safe_member_name(name):
                    raise SecurityError(f"Unsafe package member path {name!r}.")
                if name in seen:
                    raise SecurityError(f"Duplicate package member {name!r}.")
                seen.add(name)
                if info.flag_bits & 0x1:
                    raise SecurityError(f"Encrypted package member {name!r} is not allowed.")
                total_uncompressed += info.file_size
                if total_uncompressed > self.security_policy.max_uncompressed_size_bytes:
                    raise SecurityError(
                        "Package uncompressed size exceeds "
                        f"{self.security_policy.max_uncompressed_size_mb} MB."
                    )
                if info.file_size and not info.is_dir():
                    ratio = info.file_size / max(1, info.compress_size)
                    if ratio > self.security_policy.max_compression_ratio:
                        raise SecurityError(
                            f"Package member {name!r} has suspicious compression ratio "
                            f"{ratio:.1f}."
                        )
            self._entry_names = [info.filename for info in infos]
            required = {"[Content_Types].xml", "_rels/.rels"}
            if self.format_name == "docx":
                required.add("word/document.xml")
            missing = sorted(required - seen)
            if missing:
                raise NativePackageError(
                    f"Package is missing required parts: {', '.join(missing)}."
                )

            content_types = self._read_content_types(archive)
            for info in infos:
                if info.is_dir():
                    continue
                uri = _part_uri(info.filename)
                extension = PurePosixPath(info.filename).suffix.lstrip(".").lower()
                content_type = content_types["overrides"].get(
                    uri, content_types["defaults"].get(extension, "application/octet-stream")
                )
                lowered_type = content_type.lower()
                if not self.security_policy.allow_macros and (
                    "macroenabled" in lowered_type
                    or "vbaproject" in lowered_type
                    or info.filename.lower().endswith("vbaproject.bin")
                ):
                    raise SecurityError("Macro-enabled Office packages are not allowed.")
                payload = archive.read(info.filename)
                if info.filename.lower().endswith((".xml", ".rels")) and len(
                    payload
                ) > self.security_policy.max_xml_part_size_bytes:
                    raise SecurityError(
                        f"XML part {uri!r} exceeds "
                        f"{self.security_policy.max_xml_part_size_mb} MB."
                    )
                self._parts[uri] = NativePart(
                    uri=uri,
                    content_type=content_type,
                    sha256=hashlib.sha256(payload).hexdigest(),
                    size=info.file_size,
                    compressed_size=info.compress_size,
                )
                if info.filename.endswith(".rels"):
                    self._relationships.extend(
                        self._read_relationships(payload, info.filename)
                    )

    def _read_content_types(self, archive: ZipFile) -> dict[str, dict[str, str]]:
        root = parse_xml(archive.read("[Content_Types].xml"))
        defaults: dict[str, str] = {}
        overrides: dict[str, str] = {}
        for child in root:
            name = _local_name(child.tag)
            if name == "Default":
                extension = child.attrib.get("Extension", "").lower()
                defaults[extension] = child.attrib.get(
                    "ContentType", "application/octet-stream"
                )
            elif name == "Override":
                uri = _part_uri(child.attrib.get("PartName", ""))
                overrides[uri] = child.attrib.get(
                    "ContentType", "application/octet-stream"
                )
        return {"defaults": defaults, "overrides": overrides}

    def _read_relationships(
        self, payload: bytes, relationship_path: str
    ) -> list[NativeRelationship]:
        root = parse_xml(payload)
        source_part = _relationship_source(relationship_path)
        relationships: list[NativeRelationship] = []
        for child in root:
            if _local_name(child.tag) != "Relationship":
                continue
            external = child.attrib.get("TargetMode") == "External"
            if external and not self.security_policy.allow_external_relationships:
                raise SecurityError(
                    f"External relationship {child.attrib.get('Target')!r} is not allowed."
                )
            relationships.append(
                NativeRelationship(
                    source_part=source_part,
                    relationship_id=child.attrib.get("Id", ""),
                    relationship_type=child.attrib.get("Type", ""),
                    target=child.attrib.get("Target", ""),
                    external=external,
                )
            )
        return relationships
