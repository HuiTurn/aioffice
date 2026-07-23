"""Built-in render providers."""

from __future__ import annotations

import json

from aioffice.core.diagnostics import Diagnostic, Severity
from aioffice.formats.html import export_html
from aioffice.spec.models import AiOfficeDocumentSpec

from .models import RenderOptions, RenderResult

SEMANTIC_HTML_PROVIDER = "semantic-html"
SEMANTIC_HTML_PROVIDER_VERSION = "1"


def render_semantic_html(
    spec: AiOfficeDocumentSpec,
    options: RenderOptions | None = None,
) -> RenderResult:
    """Create an inspectable HTML preview, explicitly not a Word layout verdict."""

    active_options = options or RenderOptions()
    content = export_html(
        spec,
        page_view=active_options.page_view,
        include_document_metadata=active_options.include_document_metadata,
        locale=active_options.locale,
    ).encode("utf-8")
    cache_payload = {
        "artifact_id": spec.artifact.id,
        "revision": spec.artifact.revision,
        "theme": spec.theme.model_dump(mode="json"),
        "provider": SEMANTIC_HTML_PROVIDER,
        "provider_version": SEMANTIC_HTML_PROVIDER_VERSION,
        "options": active_options.model_dump(mode="json"),
    }
    return RenderResult.create(
        format="html",
        media_type="text/html; charset=utf-8",
        provider=SEMANTIC_HTML_PROVIDER,
        provider_version=SEMANTIC_HTML_PROVIDER_VERSION,
        fidelity="approximate",
        verification_status="preview_only",
        content=content,
        cache_material=json.dumps(
            cache_payload,
            ensure_ascii=False,
            sort_keys=True,
            separators=(",", ":"),
        ).encode("utf-8"),
        diagnostics=[
            Diagnostic(
                severity=Severity.WARNING,
                code="APPROXIMATE_RENDER",
                message=(
                    "Semantic HTML is an AI inspection preview; it is not evidence "
                    "of final pagination or native Word layout."
                ),
                node_ids=[spec.artifact.id],
                recoverable=True,
                suggested_actions=[
                    {
                        "action": "render_with_native_provider",
                        "format": "png",
                    }
                ],
            )
        ],
        metadata={
            "layout_authority": "semantic",
            "page_view": active_options.page_view,
            "native_visual_verification_required": True,
        },
    )


__all__ = [
    "SEMANTIC_HTML_PROVIDER",
    "SEMANTIC_HTML_PROVIDER_VERSION",
    "render_semantic_html",
]
