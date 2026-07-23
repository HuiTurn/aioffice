"""Lossless native Office package primitives."""

from .fidelity import FidelityLevel, FidelityPolicy, FidelityReport
from .identity import (
    IdentityManifest,
    IdentityNode,
    MANIFEST_PART_URI,
    MANIFEST_RELATIONSHIP_TYPE,
    apply_identity_manifest,
    build_identity_manifest,
    native_ref_for_elements,
    native_ref_for_part_elements,
    parse_identity_manifest,
    serialize_identity_manifest,
)
from .package import NativePackage
from .types import NativePart, NativeRelationship

__all__ = [
    "FidelityLevel",
    "FidelityPolicy",
    "FidelityReport",
    "IdentityManifest",
    "IdentityNode",
    "MANIFEST_PART_URI",
    "MANIFEST_RELATIONSHIP_TYPE",
    "NativePackage",
    "NativePart",
    "NativeRelationship",
    "apply_identity_manifest",
    "build_identity_manifest",
    "native_ref_for_elements",
    "native_ref_for_part_elements",
    "parse_identity_manifest",
    "serialize_identity_manifest",
]
