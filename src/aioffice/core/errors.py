"""Public exception hierarchy."""

from __future__ import annotations

from .diagnostics import Diagnostic


class AiOfficeError(Exception):
    """Base error raised by AiOffice."""


class SpecValidationError(AiOfficeError):
    """Raised when an AiOffice Spec cannot be parsed."""

    def __init__(self, message: str, diagnostics: list[Diagnostic] | None = None) -> None:
        super().__init__(message)
        self.diagnostics = diagnostics or []


class UnsupportedFormatError(AiOfficeError):
    """Raised when a source or target format is not supported by this release."""


class ExportError(AiOfficeError):
    """Raised when a valid artifact cannot be exported."""


class SecurityError(AiOfficeError):
    """Raised when an untrusted input violates the active security policy."""


class NativePackageError(AiOfficeError):
    """Raised when an Office package is malformed or cannot be patched safely."""


class WorkspaceError(AiOfficeError):
    """Raised when a workspace is invalid or a revision cannot be committed safely."""
