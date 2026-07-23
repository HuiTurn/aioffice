"""Rendering contracts, providers, and visual regression helpers."""

from .compare import compare_raster_images
from .models import RenderOptions, RenderResult, VisualComparison
from .providers import render_semantic_html

__all__ = [
    "RenderOptions",
    "RenderResult",
    "VisualComparison",
    "compare_raster_images",
    "render_semantic_html",
]
