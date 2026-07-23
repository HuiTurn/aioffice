"""Typed records for an OPC package graph."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True, slots=True)
class NativeRelationship:
    source_part: str
    relationship_id: str
    relationship_type: str
    target: str
    external: bool = False


@dataclass(frozen=True, slots=True)
class NativePart:
    uri: str
    content_type: str
    sha256: str
    size: int
    compressed_size: int
    state: str = "untouched"
