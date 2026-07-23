"""AiOffice public Python API."""

from __future__ import annotations

from ._version import __version__
from .core import (
    AiOfficeError,
    Diagnostic,
    DiffEntry,
    DocumentDiff,
    ExportError,
    NativePackageError,
    RenderingError,
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
from .operations import TextMatch, TextRange
from .security import SecurityPolicy
from .rendering import (
    RenderOptions,
    RenderResult,
    VisualComparison,
    compare_raster_images,
)
from .workspace import Workspace

# The product specification intentionally exposes ``aioffice.open(...)``.
open = open_artifact

__all__ = [
    "AiOfficeError",
    "Diagnostic",
    "DiffEntry",
    "DocumentDiff",
    "Document",
    "DocumentBuilder",
    "ExportError",
    "FidelityLevel",
    "FidelityPolicy",
    "FidelityReport",
    "NativePackageError",
    "PatchResult",
    "RenderOptions",
    "RenderResult",
    "RenderingError",
    "Severity",
    "SecurityError",
    "SecurityPolicy",
    "SpecValidationError",
    "TextMatch",
    "TextRange",
    "UnsupportedFormatError",
    "ValidationResult",
    "VisualComparison",
    "Workspace",
    "WorkspaceError",
    "__version__",
    "compare_raster_images",
    "new_id",
    "open",
    "open_artifact",
]
