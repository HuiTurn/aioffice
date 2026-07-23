from __future__ import annotations

import binascii
import struct
import unittest
import zlib
from pathlib import Path
from unittest.mock import patch

from aioffice.core.errors import RenderingError
from aioffice.documents import DocumentBuilder
from aioffice.rendering.libreoffice import _CommandOutput


def _png(width: int, height: int) -> bytes:
    def chunk(kind: bytes, payload: bytes) -> bytes:
        checksum = binascii.crc32(kind + payload) & 0xFFFFFFFF
        return (
            struct.pack(">I", len(payload))
            + kind
            + payload
            + struct.pack(">I", checksum)
        )

    row = b"\x00" + (b"\xff\xff\xff" * width)
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(row * height))
        + chunk(b"IEND", b"")
    )


def _fake_resolve(name: str) -> str:
    return f"/fake/{name}"


def _fake_command(
    command: list[str] | tuple[str, ...],
    *,
    timeout_seconds: float,
    cwd: Path | None = None,
) -> _CommandOutput:
    del timeout_seconds, cwd
    tool = Path(command[0]).name
    if tool == "soffice":
        if "--version" in command:
            return _CommandOutput("LibreOffice 25.2.1\n", "")
        self_profile = next(
            argument
            for argument in command
            if argument.startswith("-env:UserInstallation=file:")
        )
        if not self_profile or "--norestore" not in command:
            raise AssertionError("LibreOffice render was not isolated.")
        outdir = Path(command[command.index("--outdir") + 1])
        (outdir / "document.pdf").write_bytes(b"%PDF-1.7\nfake native render\n")
        return _CommandOutput("convert document.docx as document.pdf\n", "")
    if tool == "pdfinfo":
        if "-v" in command:
            return _CommandOutput("", "pdfinfo version 25.03.0\n")
        return _CommandOutput("Title: fake\nPages:          3\n", "")
    if tool == "pdftoppm":
        if "-v" in command:
            return _CommandOutput("", "pdftoppm version 25.03.0\n")
        Path(f"{command[-1]}.png").write_bytes(_png(12, 16))
        return _CommandOutput("", "")
    if tool == "fc-list":
        return _CommandOutput(
            "/fonts/A.ttf\tA\tRegular\n/fonts/B.ttf\tB\tBold\n",
            "",
        )
    raise AssertionError(f"Unexpected fake command: {command}")


class NativeRenderingTests(unittest.TestCase):
    def setUp(self) -> None:
        self.document = (
            DocumentBuilder(title="Native render")
            .heading("Evidence", id="title")
            .paragraph("Body", id="body")
            .build()
        )

    def test_pdf_and_page_render_have_structured_native_evidence(self) -> None:
        with (
            patch(
                "aioffice.rendering.libreoffice._resolve_tool",
                side_effect=_fake_resolve,
            ),
            patch(
                "aioffice.rendering.libreoffice._run_command",
                side_effect=_fake_command,
            ),
        ):
            capabilities = self.document.capabilities()
            native = capabilities["render"]["providers"][1]
            self.assertTrue(native["available"])
            self.assertEqual(native["formats"], ["pdf", "png"])
            self.assertTrue(
                capabilities["render"]["native_visual_verification_available"]
            )

            pdf = self.document.render(
                format="pdf",
                provider="libreoffice",
                options={"font_environment_hash": "controlled-fonts"},
            )
            self.assertEqual(pdf.format, "pdf")
            self.assertEqual(pdf.fidelity, "native")
            self.assertEqual(pdf.verification_status, "unverified")
            self.assertEqual(pdf.metadata["page_count"], 3)
            self.assertEqual(
                pdf.metadata["font_environment_source"],
                "caller",
            )
            self.assertEqual(
                pdf.diagnostics[0].code,
                "NATIVE_RENDER_EVIDENCE_CREATED",
            )
            summary = pdf.summary()
            self.assertNotIn("content", summary)
            self.assertEqual(summary["content_size"], len(pdf.content))

            page = self.document.render(
                format="png",
                provider="libreoffice",
                options={"page_number": 2, "dpi": 180},
            )
            self.assertEqual(page.metadata["page_number"], 2)
            self.assertEqual(page.metadata["page_count"], 3)
            self.assertEqual(page.metadata["dpi"], 180)
            self.assertEqual(page.metadata["pixel_size"], [12, 16])
            self.assertEqual(page.metadata["font_count"], 2)
            self.assertNotEqual(page.cache_key, pdf.cache_key)

    def test_page_bounds_and_provider_requirements_are_explicit(self) -> None:
        with (
            patch(
                "aioffice.rendering.libreoffice._resolve_tool",
                side_effect=_fake_resolve,
            ),
            patch(
                "aioffice.rendering.libreoffice._run_command",
                side_effect=_fake_command,
            ),
        ):
            with self.assertRaisesRegex(RenderingError, "has 3 pages"):
                self.document.render(
                    format="png",
                    provider="libreoffice",
                    options={"page_number": 4},
                )
            with self.assertRaisesRegex(RenderingError, "valid only"):
                self.document.render(
                    format="pdf",
                    provider="libreoffice",
                    options={"page_number": 1},
                )

        def missing_soffice(name: str) -> str | None:
            return None if name == "soffice" else f"/fake/{name}"

        with patch(
            "aioffice.rendering.libreoffice._resolve_tool",
            side_effect=missing_soffice,
        ):
            with self.assertRaisesRegex(RenderingError, "soffice"):
                self.document.render(
                    format="pdf",
                    provider="libreoffice",
                )


if __name__ == "__main__":
    unittest.main()
