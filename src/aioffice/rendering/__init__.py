"""Rendering contracts, providers, and visual regression helpers."""

from .compare import compare_raster_images
from .libreoffice import (
    LIBREOFFICE_PROVIDER,
    libreoffice_render_capabilities,
    render_docx_libreoffice,
)
from .models import RenderOptions, RenderResult, VisualComparison
from .providers import render_semantic_html

__all__ = [
    "LIBREOFFICE_PROVIDER",
    "RenderOptions",
    "RenderResult",
    "VisualComparison",
    "compare_raster_images",
    "libreoffice_render_capabilities",
    "render_docx_libreoffice",
    "render_semantic_html",
]
