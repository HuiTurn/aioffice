"""Pixel-level visual regression primitives for raster render providers."""

from __future__ import annotations

import io
from collections.abc import Iterable
from pathlib import Path
from typing import cast

from aioffice.core.diagnostics import Diagnostic, Severity
from aioffice.core.errors import RenderingError

from .models import RenderResult, VisualComparison


def _image_bytes(value: bytes | str | Path | RenderResult) -> bytes:
    if isinstance(value, RenderResult):
        if value.format != "png":
            raise RenderingError(f"Visual comparison requires PNG renders, not {value.format!r}.")
        return value.content
    if isinstance(value, bytes):
        return value
    return Path(value).read_bytes()


def compare_raster_images(
    baseline: bytes | str | Path | RenderResult,
    candidate: bytes | str | Path | RenderResult,
    *,
    max_mean_absolute_error: float = 0.01,
    max_changed_pixel_ratio: float = 0.01,
    pixel_tolerance: int = 8,
) -> VisualComparison:
    """Compare PNG evidence without resizing or concealing page-size changes."""

    if not 0 <= max_mean_absolute_error <= 1:
        raise ValueError("max_mean_absolute_error must be between 0 and 1.")
    if not 0 <= max_changed_pixel_ratio <= 1:
        raise ValueError("max_changed_pixel_ratio must be between 0 and 1.")
    if not 0 <= pixel_tolerance <= 255:
        raise ValueError("pixel_tolerance must be between 0 and 255.")
    try:
        from PIL import Image, ImageChops
    except ImportError as error:
        raise RenderingError(
            "Raster comparison requires the optional render dependency; install aioffice[render]."
        ) from error

    try:
        baseline_image = Image.open(io.BytesIO(_image_bytes(baseline))).convert("RGBA")
        candidate_image = Image.open(io.BytesIO(_image_bytes(candidate))).convert("RGBA")
    except Exception as error:
        raise RenderingError(f"Could not decode raster render: {error}") from error

    baseline_size = baseline_image.size
    candidate_size = candidate_image.size
    diagnostics: list[Diagnostic] = []
    if baseline_size != candidate_size:
        diagnostics.append(
            Diagnostic(
                severity=Severity.ERROR,
                code="RENDER_SIZE_CHANGED",
                message=(
                    f"Raster size changed from {baseline_size[0]}x{baseline_size[1]} "
                    f"to {candidate_size[0]}x{candidate_size[1]}."
                ),
                recoverable=True,
                suggested_actions=[{"action": "inspect_pagination"}],
            )
        )
        return VisualComparison(
            passed=False,
            baseline_size=baseline_size,
            candidate_size=candidate_size,
            mean_absolute_error=1.0,
            changed_pixel_ratio=1.0,
            max_mean_absolute_error=max_mean_absolute_error,
            max_changed_pixel_ratio=max_changed_pixel_ratio,
            pixel_tolerance=pixel_tolerance,
            diagnostics=diagnostics,
        )

    difference = ImageChops.difference(baseline_image, candidate_image)
    histogram = difference.histogram()
    channel_count = 4
    pixel_count = baseline_size[0] * baseline_size[1]
    absolute_total = sum(
        value * count
        for channel in range(channel_count)
        for value, count in enumerate(histogram[channel * 256 : (channel + 1) * 256])
    )
    mean_absolute_error = absolute_total / (pixel_count * channel_count * 255)
    pixel_data = cast(
        Iterable[tuple[int, int, int, int]],
        (
            difference.get_flattened_data()
            if hasattr(difference, "get_flattened_data")
            else difference.getdata()
        ),
    )
    changed_pixels = sum(
        1 for pixel in pixel_data if any(channel > pixel_tolerance for channel in pixel)
    )
    changed_pixel_ratio = changed_pixels / pixel_count
    passed = (
        mean_absolute_error <= max_mean_absolute_error
        and changed_pixel_ratio <= max_changed_pixel_ratio
    )
    if not passed:
        diagnostics.append(
            Diagnostic(
                severity=Severity.ERROR,
                code="VISUAL_REGRESSION",
                message=("Raster difference exceeded the configured visual regression thresholds."),
                recoverable=True,
                suggested_actions=[{"action": "inspect_render_difference"}],
            )
        )
    return VisualComparison(
        passed=passed,
        baseline_size=baseline_size,
        candidate_size=candidate_size,
        mean_absolute_error=mean_absolute_error,
        changed_pixel_ratio=changed_pixel_ratio,
        max_mean_absolute_error=max_mean_absolute_error,
        max_changed_pixel_ratio=max_changed_pixel_ratio,
        pixel_tolerance=pixel_tolerance,
        diagnostics=diagnostics,
    )


__all__ = ["compare_raster_images"]
