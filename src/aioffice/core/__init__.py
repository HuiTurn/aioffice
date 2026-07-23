"""Shared primitives used by AiOffice engines."""

from .diagnostics import Diagnostic, Severity, ValidationResult
from .errors import (
    AiOfficeError,
    ExportError,
    NativePackageError,
    SecurityError,
    SpecValidationError,
    UnsupportedFormatError,
    WorkspaceError,
)
from .ids import new_id

__all__ = [
    "AiOfficeError",
    "Diagnostic",
    "ExportError",
    "NativePackageError",
    "SecurityError",
    "Severity",
    "SpecValidationError",
    "UnsupportedFormatError",
    "ValidationResult",
    "WorkspaceError",
    "new_id",
]
