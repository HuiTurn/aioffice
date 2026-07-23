"""Shared primitives used by AiOffice engines."""

from .diagnostics import Diagnostic, Severity, ValidationResult
from .errors import AiOfficeError, ExportError, SpecValidationError, UnsupportedFormatError
from .ids import new_id

__all__ = [
    "AiOfficeError",
    "Diagnostic",
    "ExportError",
    "Severity",
    "SpecValidationError",
    "UnsupportedFormatError",
    "ValidationResult",
    "new_id",
]
