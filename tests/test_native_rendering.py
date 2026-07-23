from __future__ import annotations

import binascii
import struct
import tempfile
import unittest
import zlib
from pathlib import Path
from importlib.util import find_spec
from unittest.mock import patch

from aioffice.core.errors import RenderingError
from aioffice.documents import DocumentBuilder
from aioffice.rendering import analyze_raster_page
from aioffice.rendering.libreoffice import _CommandOutput


def _png(
    width: int,
    height: int,
    *,
    black_pixels: set[tuple[int, int]] | None = None,
) -> bytes:
    def chunk(kind: bytes, payload: bytes) -> bytes:
        checksum = binascii.crc32(kind + payload) & 0xFFFFFFFF
        return (
            struct.pack(">I", len(payload))
            + kind
            + payload
            + struct.pack(">I", checksum)
        )

    active_black_pixels = black_pixels or set()
    rows = []
    for y in range(height):
        row = bytearray(b"\x00")
        for x in range(width):
            row.extend(
                b"\x00\x00\x00"
                if (x, y) in active_black_pixels
                else b"\xff\xff\xff"
            )
        rows.append(bytes(row))
    return (
        b"\x89PNG\r\n\x1a\n"
        + chunk(b"IHDR", struct.pack(">IIBBBBB", width, height, 8, 2, 0, 0, 0))
        + chunk(b"IDAT", zlib.compress(b"".join(rows)))
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
        if "-singlefile" in command:
            Path(f"{command[-1]}.png").write_bytes(_png(12, 16))
        else:
            first = int(command[command.index("-f") + 1])
            last = int(command[command.index("-l") + 1])
            for page_number in range(first, last + 1):
                Path(f"{command[-1]}-{page_number}.png").write_bytes(
                    _png(12, 16)
                )
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

    def test_paginated_render_reuses_one_pdf_and_writes_safely(self) -> None:
        commands: list[tuple[str, ...]] = []

        def recording_command(
            command: list[str] | tuple[str, ...],
            *,
            timeout_seconds: float,
            cwd: Path | None = None,
        ) -> _CommandOutput:
            commands.append(tuple(command))
            return _fake_command(
                command,
                timeout_seconds=timeout_seconds,
                cwd=cwd,
            )

        with (
            patch(
                "aioffice.rendering.libreoffice._resolve_tool",
                side_effect=_fake_resolve,
            ),
            patch(
                "aioffice.rendering.libreoffice._run_command",
                side_effect=recording_command,
            ),
        ):
            result = self.document.render_pages(
                page_numbers=[3, 1],
                options={"dpi": 180},
                max_pages=2,
            )
        conversions = [
            command
            for command in commands
            if Path(command[0]).name == "soffice"
            and "--convert-to" in command
        ]
        self.assertEqual(len(conversions), 1)
        self.assertEqual(result.page_count, 3)
        self.assertEqual(
            [page.page_number for page in result.pages],
            [1, 3],
        )
        self.assertEqual(
            result.metadata["selected_pages"],
            [1, 3],
        )
        self.assertNotIn("content", result.summary()["pdf"])
        self.assertNotIn("content", result.summary()["pages"][0])

        with tempfile.TemporaryDirectory() as directory:
            written = result.write(directory, stem="review")
            self.assertEqual(
                Path(written["pdf"]).name,
                "review.pdf",
            )
            page_paths = written["pages"]
            assert isinstance(page_paths, list)
            self.assertEqual(
                [path.name for path in page_paths],
                ["review-page-0001.png", "review-page-0003.png"],
            )
            with self.assertRaisesRegex(RenderingError, "Refusing"):
                result.write(directory, stem="review")
            result.write(directory, stem="review", overwrite=True)

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
            with self.assertRaisesRegex(ValueError, "duplicates"):
                self.document.render_pages(page_numbers=[1, 1])
            with self.assertRaisesRegex(RenderingError, "max_pages"):
                self.document.render_pages(max_pages=2)

    @unittest.skipUnless(find_spec("PIL") is not None, "Pillow is optional")
    def test_page_analysis_reports_blank_and_edge_contact(self) -> None:
        blank = analyze_raster_page(_png(20, 30), page_number=2)
        self.assertTrue(blank.appears_blank)
        self.assertEqual(blank.diagnostics[0].code, "PAGE_APPEARS_BLANK")
        self.assertEqual(blank.background_color, "#FFFFFF")

        edge = analyze_raster_page(
            _png(
                20,
                30,
                black_pixels={(0, 10), (1, 10), (2, 10)},
            ),
            page_number=3,
        )
        self.assertFalse(edge.appears_blank)
        self.assertIn("left", edge.edge_contact)
        self.assertEqual(
            edge.diagnostics[0].code,
            "PAGE_CONTENT_NEAR_EDGE",
        )


if __name__ == "__main__":
    unittest.main()
