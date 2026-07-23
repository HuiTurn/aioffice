"""AiOffice public Python API."""

from __future__ import annotations

from ._version import __version__
from .core import (
    AiOfficeError,
    Diagnostic,
    ExportError,
    NativePackageError,
    SecurityError,
    Severity,
    SpecValidationError,
    UnsupportedFormatError,
    ValidationResult,
    WorkspaceError,
    new_id,
)
from .documents import Document, DocumentBuilder, PatchResult, open_artifact
from .native import FidelityLevel, FidelityPolicy, FidelityReport
from .security import SecurityPolicy
from .workspace import Workspace

# The product specification intentionally exposes ``aioffice.open(...)``.
open = open_artifact

__all__ = [
    "AiOfficeError",
    "Diagnostic",
    "Document",
    "DocumentBuilder",
    "ExportError",
    "FidelityLevel",
    "FidelityPolicy",
    "FidelityReport",
    "NativePackageError",
    "PatchResult",
    "Severity",
    "SecurityError",
    "SecurityPolicy",
    "SpecValidationError",
    "UnsupportedFormatError",
    "ValidationResult",
    "Workspace",
    "WorkspaceError",
    "__version__",
    "new_id",
    "open",
    "open_artifact",
]
