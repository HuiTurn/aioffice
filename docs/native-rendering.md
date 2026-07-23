# Native rendering contract

AiOffice uses native-compatible page evidence as a separate verification layer.
The JSON Spec expresses intent, the DOCX package remains the preservation authority,
and a named office renderer supplies pagination and page pixels.

## Provider contract

The built-in provider is named `libreoffice`. It requires these system tools on
`PATH`:

- `soffice` from LibreOffice for DOCX-to-PDF conversion;
- `pdfinfo` from Poppler for page-count and PDF validity evidence;
- `pdftoppm` from Poppler for PNG page evidence;
- optional `fc-list` from fontconfig for a reproducible font inventory hash.

PDF rendering requires the first two tools. PNG additionally requires `pdftoppm`.
`Document.capabilities()` reports discovery without launching the programs.

AiOffice invokes LibreOffice in headless mode with `writer_pdf_Export` and a fresh,
temporary `UserInstallation` URI. This keeps renderer state isolated from a user's
interactive LibreOffice profile and avoids reusing an already-running process.
Every external command has a bounded timeout. Temporary DOCX, PDF, PNG, and profile
files are removed when the render call returns or fails.

LibreOffice documents its headless startup and writable user-profile requirement in
[Starting LibreOffice with parameters](https://help.libreoffice.org/latest/en-US/text/shared/guide/start_parameters.html).
Its documented PDF CLI uses `--convert-to pdf:writer_pdf_Export`; see
[PDF export command-line parameters](https://help.libreoffice.org/latest/en-US/text/shared/guide/pdf_params.html).
Poppler's `pdftoppm` supports first/last-page selection, one-file output, DPI, and PNG
generation as used by this provider; see the
[pdftoppm manual](https://manpages.debian.org/trixie/poppler-utils/pdftoppm.1.en.html).

## Python API

```python
import aioffice

document = aioffice.open("report.docx")

pdf = document.render(format="pdf", provider="libreoffice")
pdf.write("evidence/report.pdf")

for page_number in range(1, pdf.metadata["page_count"] + 1):
    page = document.render(
        format="png",
        provider="libreoffice",
        options={
            "page_number": page_number,
            "dpi": 144,
            "timeout_seconds": 60,
        },
    )
    page.write(f"evidence/page-{page_number}.png")
```

PNG page numbers are one-based. Omitting `page_number` selects page 1. DPI is
strictly bounded from 72 to 600. `page_number` is rejected for PDF output instead of
being ignored.

## CLI

```bash
aioffice capabilities report.docx
aioffice render report.docx --format pdf -o evidence/report.pdf
aioffice render report.docx --format png --page 2 --dpi 144 \
  -o evidence/page-2.png
```

`--provider` is inferred as `semantic-html` for HTML and `libreoffice` for PDF/PNG.
The command writes the requested artifact and prints a JSON render summary containing
the output path but not the potentially large binary content.

## Evidence and cache identity

Every result includes:

- content size and SHA-256;
- source DOCX SHA-256 and, for page renders, the intermediate PDF SHA-256;
- LibreOffice and Poppler versions;
- page count and selected one-based page;
- raster DPI and pixel dimensions;
- platform and architecture;
- font-environment fingerprint and source;
- a cache key derived from the layout-affecting inputs.

When fontconfig is unavailable, rendering can still succeed but includes
`FONT_ENVIRONMENT_UNVERIFIED`. A caller that provisions a controlled font image can
pass its own `font_environment_hash`.

PDF files may contain renderer-generated metadata, so byte hashes can change even
when page pixels do not. Use PNG evidence and `compare_raster_images()` for visual
regression.

## Verification boundary

`fidelity="native"` means the declared office engine performed layout. It does not
mean all Word engines are pixel-identical. `verification_status="unverified"` means
the artifact has not yet passed human visual review or an accepted regression
threshold.

The intended AI loop is:

1. inspect structure and diagnostics;
2. apply a bounded semantic/native patch;
3. render PDF to discover pagination;
4. render each affected page to PNG;
5. inspect pixels or compare with an approved baseline;
6. revise until the page-level criteria pass.

Renderer success alone must never be used as aesthetic approval.
