"""Shared primitives used by AiOffice engines."""

from .diagnostics import Diagnostic, Severity, ValidationResult
from .diff import DiffEntry, DocumentDiff
from .errors import (
    AiOfficeError,
    ExportError,
    NativePackageError,
    RenderingError,
    SecurityError,
    SpecValidationError,
    UnsupportedFormatError,
    WorkspaceError,
)
from .ids import new_id

__all__ = [
    "AiOfficeError",
    "Diagnostic",
    "DiffEntry",
    "DocumentDiff",
    "ExportError",
    "NativePackageError",
    "RenderingError",
    "SecurityError",
    "Severity",
    "SpecValidationError",
    "UnsupportedFormatError",
    "ValidationResult",
    "WorkspaceError",
    "new_id",
]
