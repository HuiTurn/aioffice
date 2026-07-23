"""Lossless native Office package primitives."""

from .fidelity import FidelityLevel, FidelityPolicy, FidelityReport
from .package import NativePackage
from .types import NativePart, NativeRelationship

__all__ = [
    "FidelityLevel",
    "FidelityPolicy",
    "FidelityReport",
    "NativePackage",
    "NativePart",
    "NativeRelationship",
]
