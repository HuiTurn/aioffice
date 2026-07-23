"""Conservative page-image measurements for AI layout review."""

from __future__ import annotations

import io
from pathlib import Path
from typing import Literal, cast

from aioffice.core.diagnostics import Diagnostic, Severity
from aioffice.core.errors import RenderingError

from .models import PageVisualAnalysis, RenderResult, RenderedPage


def _png_bytes(
    value: bytes | str | Path | RenderResult | RenderedPage,
) -> bytes:
    if isinstance(value, RenderResult):
        if value.format != "png":
            raise RenderingError(
                f"Page analysis requires PNG evidence, not {value.format!r}."
            )
        return value.content
    if isinstance(value, RenderedPage):
        return value.content
    if isinstance(value, bytes):
        return value
    return Path(value).read_bytes()


def analyze_raster_page(
    value: bytes | str | Path | RenderResult | RenderedPage,
    *,
    page_number: int | None = None,
    background_tolerance: int = 12,
    blank_ink_ratio: float = 0.0001,
    edge_margin_ratio: float = 0.01,
) -> PageVisualAnalysis:
    """Measure page occupancy without claiming subjective aesthetic approval."""

    if isinstance(page_number, bool) or page_number is not None and page_number < 1:
        raise ValueError("page_number must be a positive one-based integer.")
    if not 0 <= background_tolerance <= 255:
        raise ValueError("background_tolerance must be between 0 and 255.")
    if not 0 <= blank_ink_ratio <= 1:
        raise ValueError("blank_ink_ratio must be between 0 and 1.")
    if not 0 <= edge_margin_ratio <= 0.25:
        raise ValueError("edge_margin_ratio must be between 0 and 0.25.")
    try:
        from PIL import Image, ImageChops
    except ImportError as error:
        raise RenderingError(
            "Page analysis requires Pillow; install aioffice[render]."
        ) from error

    try:
        image = Image.open(io.BytesIO(_png_bytes(value))).convert("RGB")
        image.load()
    except Exception as error:
        raise RenderingError(
            f"Could not decode PNG page evidence: {error}"
        ) from error
    width, height = image.size
    corners = [
        cast(tuple[int, int, int], image.getpixel((0, 0))),
        cast(tuple[int, int, int], image.getpixel((width - 1, 0))),
        cast(tuple[int, int, int], image.getpixel((0, height - 1))),
        cast(
            tuple[int, int, int],
            image.getpixel((width - 1, height - 1)),
        ),
    ]
    background = tuple(
        int(round(sum(pixel[channel] for pixel in corners) / len(corners)))
        for channel in range(3)
    )
    background_image = Image.new("RGB", image.size, background)
    difference = ImageChops.difference(image, background_image).convert("L")
    def threshold(pixel: int) -> int:
        return 255 if pixel > background_tolerance else 0

    mask = difference.point(threshold, mode="1")
    histogram = mask.histogram()
    ink_pixels = histogram[255] if len(histogram) > 255 else 0
    pixel_count = width * height
    ink_pixel_ratio = ink_pixels / pixel_count
    appears_blank = ink_pixel_ratio <= blank_ink_ratio
    bbox = mask.getbbox()

    if bbox is None:
        top = right = bottom = left = 1.0
    else:
        x0, y0, x1, y1 = bbox
        top = y0 / height
        right = (width - x1) / width
        bottom = (height - y1) / height
        left = x0 / width

    edge_contact: list[
        Literal["top", "right", "bottom", "left"]
    ] = []
    if bbox is not None and not appears_blank:
        if top <= edge_margin_ratio:
            edge_contact.append("top")
        if right <= edge_margin_ratio:
            edge_contact.append("right")
        if bottom <= edge_margin_ratio:
            edge_contact.append("bottom")
        if left <= edge_margin_ratio:
            edge_contact.append("left")

    inferred_page_number = (
        value.page_number
        if isinstance(value, RenderedPage)
        else (
            value.metadata.get("page_number")
            if isinstance(value, RenderResult)
            else None
        )
    )
    active_page_number = (
        page_number
        if page_number is not None
        else (
            inferred_page_number
            if isinstance(inferred_page_number, int)
            and not isinstance(inferred_page_number, bool)
            and inferred_page_number >= 1
            else 1
        )
    )
    diagnostics: list[Diagnostic] = []
    if appears_blank:
        diagnostics.append(
            Diagnostic(
                severity=Severity.WARNING,
                code="PAGE_APPEARS_BLANK",
                message=(
                    f"Page {active_page_number} has an ink ratio of "
                    f"{ink_pixel_ratio:.6f} and appears blank."
                ),
                recoverable=True,
                suggested_actions=[
                    {
                        "action": "inspect_unexpected_blank_page",
                        "page_number": active_page_number,
                    }
                ],
            )
        )
    if edge_contact:
        diagnostics.append(
            Diagnostic(
                severity=Severity.WARNING,
                code="PAGE_CONTENT_NEAR_EDGE",
                message=(
                    f"Page {active_page_number} has visible content within "
                    f"{edge_margin_ratio:.1%} of the "
                    f"{', '.join(edge_contact)} edge"
                    f"{'s' if len(edge_contact) != 1 else ''}."
                ),
                recoverable=True,
                suggested_actions=[
                    {
                        "action": "inspect_page_edge",
                        "page_number": active_page_number,
                        "edges": edge_contact,
                    }
                ],
            )
        )
    background_color = "#" + "".join(
        f"{channel:02X}" for channel in background
    )
    return PageVisualAnalysis(
        page_number=active_page_number,
        width_pixels=width,
        height_pixels=height,
        background_color=background_color,
        ink_pixel_ratio=ink_pixel_ratio,
        appears_blank=appears_blank,
        content_bbox=bbox,
        whitespace_top_ratio=top,
        whitespace_right_ratio=right,
        whitespace_bottom_ratio=bottom,
        whitespace_left_ratio=left,
        edge_contact=edge_contact,
        diagnostics=diagnostics,
    )


__all__ = ["analyze_raster_page"]
