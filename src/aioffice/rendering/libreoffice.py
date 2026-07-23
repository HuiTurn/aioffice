"""Native-compatible DOCX rendering through isolated LibreOffice and Poppler jobs."""

from __future__ import annotations

import hashlib
import json
import os
import platform
import re
import shutil
import signal
import struct
import subprocess
import tempfile
from dataclasses import dataclass
from pathlib import Path
from typing import Literal, Sequence

from aioffice.core.diagnostics import Diagnostic, Severity
from aioffice.core.errors import RenderingError

from .models import RenderOptions, RenderResult

LIBREOFFICE_PROVIDER = "libreoffice"
_PDF_SIGNATURE = b"%PDF-"
_PNG_SIGNATURE = b"\x89PNG\r\n\x1a\n"
_MAX_COMMAND_OUTPUT = 2_000


@dataclass(frozen=True, slots=True)
class _CommandOutput:
    stdout: str
    stderr: str


def _resolve_tool(name: str) -> str | None:
    resolved = shutil.which(name)
    if resolved is not None:
        return resolved
    candidates: dict[str, tuple[str, ...]] = {
        "soffice": (
            "/Applications/LibreOffice.app/Contents/MacOS/soffice",
            "/opt/libreoffice/program/soffice",
            "C:/Program Files/LibreOffice/program/soffice.com",
        ),
        "pdfinfo": (),
        "pdftoppm": (),
        "fc-list": (),
    }
    for candidate in candidates.get(name, ()):
        if Path(candidate).is_file():
            return candidate
    return None


def _bounded_output(value: str) -> str:
    normalized = " ".join(value.split())
    if len(normalized) <= _MAX_COMMAND_OUTPUT:
        return normalized
    return normalized[:_MAX_COMMAND_OUTPUT] + "…"


def _run_command(
    command: Sequence[str],
    *,
    timeout_seconds: float,
    cwd: Path | None = None,
) -> _CommandOutput:
    kwargs: dict[str, object] = {
        "cwd": cwd,
        "stdout": subprocess.PIPE,
        "stderr": subprocess.PIPE,
        "text": True,
        "encoding": "utf-8",
        "errors": "replace",
    }
    if os.name == "posix":
        kwargs["start_new_session"] = True
    elif os.name == "nt":
        kwargs["creationflags"] = subprocess.CREATE_NEW_PROCESS_GROUP

    process = subprocess.Popen(list(command), **kwargs)  # type: ignore[arg-type]
    try:
        stdout, stderr = process.communicate(timeout=timeout_seconds)
    except subprocess.TimeoutExpired as error:
        if os.name == "posix":
            os.killpg(process.pid, signal.SIGKILL)
        else:
            process.kill()
        stdout, stderr = process.communicate()
        detail = _bounded_output(stderr or stdout)
        suffix = f" Last output: {detail}" if detail else ""
        raise RenderingError(
            f"Native render command timed out after {timeout_seconds:g} seconds.{suffix}"
        ) from error
    if process.returncode != 0:
        detail = _bounded_output(stderr or stdout)
        suffix = f" Output: {detail}" if detail else ""
        raise RenderingError(
            f"Native render command failed with exit code {process.returncode}.{suffix}"
        )
    return _CommandOutput(stdout=stdout, stderr=stderr)


def _tool_version(
    executable: str,
    *,
    arguments: Sequence[str],
    timeout_seconds: float,
) -> str:
    output = _run_command(
        [executable, *arguments],
        timeout_seconds=min(timeout_seconds, 15.0),
    )
    text = output.stdout.strip() or output.stderr.strip()
    line = text.splitlines()[0].strip() if text else "unknown"
    return line


def _required_tool(name: str, purpose: str) -> str:
    executable = _resolve_tool(name)
    if executable is None:
        raise RenderingError(
            f"Native rendering requires {purpose} ({name!r}) on PATH. "
            "Install LibreOffice and Poppler, then retry."
        )
    return executable


def libreoffice_render_capabilities() -> dict[str, object]:
    """Report tool discovery without launching the external renderers."""

    discovered = {
        "soffice": _resolve_tool("soffice"),
        "pdfinfo": _resolve_tool("pdfinfo"),
        "pdftoppm": _resolve_tool("pdftoppm"),
        "fc-list": _resolve_tool("fc-list"),
    }
    formats: list[str] = []
    if discovered["soffice"] and discovered["pdfinfo"]:
        formats.append("pdf")
        if discovered["pdftoppm"]:
            formats.append("png")
    return {
        "name": LIBREOFFICE_PROVIDER,
        "formats": formats,
        "available": bool(formats),
        "fidelity": "native",
        "verification_status": "unverified",
        "missing_tools": sorted(
            name
            for name in ("soffice", "pdfinfo", "pdftoppm")
            if discovered[name] is None
        ),
        "font_inventory_available": discovered["fc-list"] is not None,
        "isolated_user_profile": True,
    }


def _font_environment(
    options: RenderOptions,
) -> tuple[str | None, str, int | None, list[Diagnostic]]:
    if options.font_environment_hash is not None:
        return options.font_environment_hash, "caller", None, []
    fc_list = _resolve_tool("fc-list")
    if fc_list is None:
        return (
            None,
            "unavailable",
            None,
            [
                Diagnostic(
                    severity=Severity.WARNING,
                    code="FONT_ENVIRONMENT_UNVERIFIED",
                    message=(
                        "The render succeeded, but no fontconfig inventory was available. "
                        "Cross-machine visual comparisons may include font substitutions."
                    ),
                    recoverable=True,
                    suggested_actions=[
                        {"action": "provide_font_environment_hash"},
                    ],
                )
            ],
        )
    try:
        output = _run_command(
            [
                fc_list,
                "--format",
                "%{file}\t%{family}\t%{style}\n",
            ],
            timeout_seconds=min(options.timeout_seconds, 15.0),
        )
    except RenderingError:
        return (
            None,
            "unavailable",
            None,
            [
                Diagnostic(
                    severity=Severity.WARNING,
                    code="FONT_ENVIRONMENT_UNVERIFIED",
                    message=(
                        "The render succeeded, but the font inventory could not be captured. "
                        "Cross-machine visual comparisons may include font substitutions."
                    ),
                    recoverable=True,
                    suggested_actions=[
                        {"action": "provide_font_environment_hash"},
                    ],
                )
            ],
        )
    inventory = sorted(line for line in output.stdout.splitlines() if line.strip())
    material = "\n".join(inventory).encode("utf-8")
    return hashlib.sha256(material).hexdigest(), "fontconfig", len(inventory), []


def _pdf_page_count(
    pdfinfo: str,
    pdf_path: Path,
    *,
    timeout_seconds: float,
) -> int:
    output = _run_command(
        [pdfinfo, str(pdf_path)],
        timeout_seconds=timeout_seconds,
    )
    match = re.search(r"^Pages:\s*(\d+)\s*$", output.stdout, flags=re.MULTILINE)
    if match is None:
        raise RenderingError("Poppler did not report a page count for the rendered PDF.")
    page_count = int(match.group(1))
    if page_count < 1:
        raise RenderingError("The native renderer produced a PDF with no pages.")
    return page_count


def _png_size(content: bytes) -> tuple[int, int]:
    if len(content) < 24 or not content.startswith(_PNG_SIGNATURE):
        raise RenderingError("Poppler produced an invalid PNG page render.")
    if content[12:16] != b"IHDR":
        raise RenderingError("Poppler produced a PNG without an IHDR header.")
    width, height = struct.unpack(">II", content[16:24])
    if width < 1 or height < 1:
        raise RenderingError("Poppler produced an empty PNG page render.")
    return width, height


def render_docx_libreoffice(
    docx_content: bytes,
    *,
    format: Literal["pdf", "png"],
    options: RenderOptions | None = None,
) -> RenderResult:
    """Render DOCX bytes through LibreOffice with deterministic evidence metadata."""

    if not docx_content:
        raise RenderingError("Cannot render an empty DOCX artifact.")
    active_options = options or RenderOptions()
    if format == "pdf" and active_options.page_number is not None:
        raise RenderingError(
            "page_number selects one raster page and is valid only for format='png'."
        )
    soffice = _required_tool("soffice", "LibreOffice")
    pdfinfo = _required_tool("pdfinfo", "Poppler pdfinfo")
    pdftoppm = (
        _required_tool("pdftoppm", "Poppler pdftoppm")
        if format == "png"
        else None
    )
    provider_version = _tool_version(
        soffice,
        arguments=("--version",),
        timeout_seconds=active_options.timeout_seconds,
    )
    poppler_version = _tool_version(
        pdfinfo,
        arguments=("-v",),
        timeout_seconds=active_options.timeout_seconds,
    )
    source_sha256 = hashlib.sha256(docx_content).hexdigest()

    with tempfile.TemporaryDirectory(prefix="aioffice-render-") as directory:
        workspace = Path(directory)
        source_path = workspace / "document.docx"
        profile_path = workspace / "libreoffice-profile"
        profile_path.mkdir()
        source_path.write_bytes(docx_content)
        _run_command(
            [
                soffice,
                f"-env:UserInstallation={profile_path.as_uri()}",
                "--headless",
                "--nologo",
                "--nodefault",
                "--norestore",
                "--convert-to",
                "pdf:writer_pdf_Export",
                "--outdir",
                str(workspace),
                str(source_path),
            ],
            timeout_seconds=active_options.timeout_seconds,
            cwd=workspace,
        )
        pdf_path = workspace / "document.pdf"
        if not pdf_path.is_file():
            raise RenderingError("LibreOffice completed without producing document.pdf.")
        pdf_content = pdf_path.read_bytes()
        if not pdf_content.startswith(_PDF_SIGNATURE):
            raise RenderingError("LibreOffice produced an invalid PDF render.")
        page_count = _pdf_page_count(
            pdfinfo,
            pdf_path,
            timeout_seconds=active_options.timeout_seconds,
        )
        pdf_sha256 = hashlib.sha256(pdf_content).hexdigest()

        page_number: int | None = None
        pixel_size: tuple[int, int] | None = None
        rasterizer_version: str | None = None
        if format == "pdf":
            content = pdf_content
            media_type = "application/pdf"
        else:
            assert pdftoppm is not None
            page_number = active_options.page_number or 1
            if page_number > page_count:
                raise RenderingError(
                    f"Requested page {page_number}, but the rendered PDF has "
                    f"{page_count} page{'s' if page_count != 1 else ''}."
                )
            rasterizer_version = _tool_version(
                pdftoppm,
                arguments=("-v",),
                timeout_seconds=active_options.timeout_seconds,
            )
            page_prefix = workspace / "page"
            _run_command(
                [
                    pdftoppm,
                    "-f",
                    str(page_number),
                    "-l",
                    str(page_number),
                    "-singlefile",
                    "-r",
                    str(active_options.dpi),
                    "-png",
                    str(pdf_path),
                    str(page_prefix),
                ],
                timeout_seconds=active_options.timeout_seconds,
                cwd=workspace,
            )
            page_path = workspace / "page.png"
            if not page_path.is_file():
                raise RenderingError("Poppler completed without producing page.png.")
            content = page_path.read_bytes()
            pixel_size = _png_size(content)
            media_type = "image/png"

        (
            font_environment_hash,
            font_environment_source,
            font_count,
            font_diagnostics,
        ) = _font_environment(active_options)
        cache_payload = {
            "source_docx_sha256": source_sha256,
            "provider": LIBREOFFICE_PROVIDER,
            "provider_version": provider_version,
            "poppler_version": poppler_version,
            "format": format,
            "page_number": page_number,
            "dpi": active_options.dpi if format == "png" else None,
            "font_environment_hash": font_environment_hash,
            "platform": platform.system(),
            "machine": platform.machine(),
        }
        metadata: dict[str, object] = {
            "layout_authority": "libreoffice",
            "source_docx_sha256": source_sha256,
            "pdf_sha256": pdf_sha256,
            "page_count": page_count,
            "page_number": page_number,
            "dpi": active_options.dpi if format == "png" else None,
            "pixel_size": list(pixel_size) if pixel_size is not None else None,
            "font_environment_hash": font_environment_hash,
            "font_environment_source": font_environment_source,
            "font_count": font_count,
            "rasterizer_version": rasterizer_version,
            "pdf_inspector_version": poppler_version,
            "platform": platform.system(),
            "machine": platform.machine(),
            "isolated_user_profile": True,
            "aesthetic_review_completed": False,
        }
        diagnostics = [
            Diagnostic(
                severity=Severity.INFO,
                code="NATIVE_RENDER_EVIDENCE_CREATED",
                message=(
                    "LibreOffice produced native-compatible layout evidence. "
                    "The page still requires visual or regression review before approval."
                ),
                recoverable=True,
                suggested_actions=[
                    {
                        "action": "inspect_render",
                        "format": format,
                        **(
                            {"page_number": page_number}
                            if page_number is not None
                            else {}
                        ),
                    }
                ],
            ),
            *font_diagnostics,
        ]
        return RenderResult.create(
            format=format,
            media_type=media_type,
            provider=LIBREOFFICE_PROVIDER,
            provider_version=provider_version,
            fidelity="native",
            verification_status="unverified",
            content=content,
            cache_material=json.dumps(
                cache_payload,
                ensure_ascii=False,
                sort_keys=True,
                separators=(",", ":"),
            ).encode("utf-8"),
            diagnostics=diagnostics,
            metadata=metadata,
        )


__all__ = [
    "LIBREOFFICE_PROVIDER",
    "libreoffice_render_capabilities",
    "render_docx_libreoffice",
]
