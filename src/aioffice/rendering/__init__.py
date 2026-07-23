"""Rendering contracts, providers, and visual regression helpers."""

from .analysis import analyze_raster_page
from .compare import compare_raster_images
from .libreoffice import (
    LIBREOFFICE_PROVIDER,
    libreoffice_render_capabilities,
    render_docx_libreoffice,
    render_docx_pages_libreoffice,
)
from .models import (
    PageVisualAnalysis,
    PaginatedRenderResult,
    RenderedPage,
    RenderOptions,
    RenderResult,
    VisualComparison,
)
from .providers import render_semantic_html

__all__ = [
    "LIBREOFFICE_PROVIDER",
    "PageVisualAnalysis",
    "PaginatedRenderResult",
    "RenderedPage",
    "RenderOptions",
    "RenderResult",
    "VisualComparison",
    "analyze_raster_page",
    "compare_raster_images",
    "libreoffice_render_capabilities",
    "render_docx_libreoffice",
    "render_docx_pages_libreoffice",
    "render_semantic_html",
]
