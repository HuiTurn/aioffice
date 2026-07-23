"""AiOffice public Python API."""

from __future__ import annotations

from ._version import __version__
from .core import (
    AiOfficeError,
    Diagnostic,
    ExportError,
    Severity,
    SpecValidationError,
    UnsupportedFormatError,
    ValidationResult,
    new_id,
)
from .documents import Document, DocumentBuilder, PatchResult, open_artifact

# The product specification intentionally exposes ``aioffice.open(...)``.
open = open_artifact

__all__ = [
    "AiOfficeError",
    "Diagnostic",
    "Document",
    "DocumentBuilder",
    "ExportError",
    "PatchResult",
    "Severity",
    "SpecValidationError",
    "UnsupportedFormatError",
    "ValidationResult",
    "__version__",
    "new_id",
    "open",
    "open_artifact",
]
