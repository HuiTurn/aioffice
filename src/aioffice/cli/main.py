"""Command-line entry point."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Sequence

from aioffice._version import __version__
from aioffice.core.errors import AiOfficeError
from aioffice.documents import DocumentBuilder, open_artifact
from aioffice.operations import TextMatch, TextRange
from aioffice.spec.models import (
    AiOfficeDocumentSpec,
    AssetRef,
    BorderLine,
    DocumentDefaults,
    DocumentField,
    DocumentSection,
    DocumentSettings,
    HeaderFooterBindings,
    HeaderFooterPart,
    ImageBlock,
    ImageInsert,
    ImageUpdate,
    NamedStyle,
    PageSize,
    ParagraphBorders,
    ParagraphStyle,
    SectionLayout,
    TableCell,
    TableCellBorders,
    TableCellFormat,
    TableColumn,
    TableBorders,
    TableLayout,
    TableWidth,
)
from aioffice.workspace import Workspace


def _json_dump(value: Any) -> None:
    print(json.dumps(value, ensure_ascii=False, indent=2))


def _load_patch(path: Path) -> dict[str, Any]:
    payload = json.loads(path.read_text(encoding="utf-8"))
    if isinstance(payload, list):
        return {"operations": payload}
    if not isinstance(payload, dict):
        raise ValueError("Patch JSON must be an operation list or an envelope object.")
    return payload


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="aioffice",
        description="Create and validate office documents from AiOffice Spec.",
    )
    parser.add_argument("--version", action="version", version=f"%(prog)s {__version__}")
    subparsers = parser.add_subparsers(dest="command", required=True)

    build = subparsers.add_parser("build", help="Build a target artifact from JSON or Markdown.")
    build.add_argument("input", type=Path)
    build.add_argument("-o", "--output", type=Path)

    export = subparsers.add_parser("export", help="Export JSON or Markdown to another format.")
    export.add_argument("input", type=Path)
    export.add_argument("--to", required=True, type=Path, dest="output")

    inspect = subparsers.add_parser("inspect", help="Inspect document structure.")
    inspect.add_argument("input", type=Path)
    inspect.add_argument(
        "--response-format",
        choices=("summary", "compact", "expanded"),
        default="compact",
    )

    capabilities = subparsers.add_parser(
        "capabilities", help="Report operations and fidelity available for a document."
    )
    capabilities.add_argument("input", type=Path)

    extract_image = subparsers.add_parser(
        "extract-image",
        help="Extract one verified native image by its projected image ID.",
    )
    extract_image.add_argument("input", type=Path)
    extract_image.add_argument("image_id")
    extract_image.add_argument(
        "-o",
        "--output",
        required=True,
        type=Path,
    )
    extract_image.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace the output file if it already exists.",
    )

    replace_image = subparsers.add_parser(
        "replace-image",
        help="Replace one projected native image through a bounded local binary input.",
    )
    replace_image.add_argument("input", type=Path)
    replace_image.add_argument("image_id")
    replace_image.add_argument("replacement", type=Path)
    replace_image.add_argument(
        "-o",
        "--output",
        type=Path,
    )
    replace_image.add_argument(
        "--media-type",
        help="Optional declared media type; it must match the detected image signature.",
    )
    replace_image.add_argument("--dry-run", action="store_true")
    replace_image.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace the output DOCX if it already exists.",
    )

    insert_image = subparsers.add_parser(
        "insert-image-after",
        help="Insert one native inline image after a mapped top-level node.",
    )
    insert_image.add_argument("input", type=Path)
    insert_image.add_argument("target")
    insert_image.add_argument("replacement", type=Path)
    insert_image.add_argument("--width", required=True, type=float)
    insert_image.add_argument(
        "--width-unit",
        required=True,
        choices=("pt", "in", "cm", "mm", "px"),
    )
    insert_image.add_argument("--height", required=True, type=float)
    insert_image.add_argument(
        "--height-unit",
        required=True,
        choices=("pt", "in", "cm", "mm", "px"),
    )
    insert_image.add_argument("--alt-text", required=True)
    insert_image.add_argument("--title")
    insert_image.add_argument("--name")
    insert_image.add_argument("--image-id")
    insert_image.add_argument("--media-type")
    insert_image.add_argument(
        "--align",
        choices=("left", "center", "right", "justify"),
    )
    insert_image.add_argument("-o", "--output", type=Path)
    insert_image.add_argument("--dry-run", action="store_true")
    insert_image.add_argument("--overwrite", action="store_true")

    render = subparsers.add_parser(
        "render",
        help="Render semantic preview or native-compatible page evidence.",
    )
    render.add_argument("input", type=Path)
    render.add_argument(
        "--provider",
        choices=("semantic-html", "libreoffice"),
        help="Defaults to semantic-html for HTML and libreoffice for PDF/PNG.",
    )
    render.add_argument(
        "--format",
        choices=("html", "pdf", "png"),
        default="html",
    )
    render.add_argument("-o", "--output", required=True, type=Path)
    render.add_argument(
        "--page",
        type=int,
        help="One-based page number for PNG output; defaults to page 1.",
    )
    render.add_argument(
        "--dpi",
        type=int,
        default=144,
        help="PNG raster resolution from 72 to 600 DPI.",
    )
    render.add_argument(
        "--timeout",
        type=float,
        default=60.0,
        help="Per-command native renderer timeout in seconds.",
    )
    render.add_argument(
        "--font-environment-hash",
        help="Caller-supplied font environment fingerprint for reproducible evidence.",
    )

    render_pages = subparsers.add_parser(
        "render-pages",
        help="Render one native PDF and a bounded set of consistent PNG pages.",
    )
    render_pages.add_argument("input", type=Path)
    render_pages.add_argument(
        "-o",
        "--output-directory",
        required=True,
        type=Path,
    )
    render_pages.add_argument(
        "--pages",
        help="One-based comma/range selection such as 1,3-5; defaults to all pages.",
    )
    render_pages.add_argument(
        "--dpi",
        type=int,
        default=144,
        help="PNG raster resolution from 72 to 600 DPI.",
    )
    render_pages.add_argument(
        "--timeout",
        type=float,
        default=60.0,
        help="Per-command native renderer timeout in seconds.",
    )
    render_pages.add_argument(
        "--max-pages",
        type=int,
        default=100,
        help="Maximum number of PNG pages emitted in one call (1–500).",
    )
    render_pages.add_argument(
        "--analyze",
        action="store_true",
        help="Measure blank pages, ink bounds, whitespace, and edge contact with Pillow.",
    )
    render_pages.add_argument(
        "--font-environment-hash",
        help="Caller-supplied font environment fingerprint for reproducible evidence.",
    )
    render_pages.add_argument(
        "--stem",
        help="Output filename stem; defaults to the input filename stem.",
    )
    render_pages.add_argument(
        "--overwrite",
        action="store_true",
        help="Replace matching evidence files; default behavior refuses overwrite.",
    )

    validate = subparsers.add_parser("validate", help="Validate a document.")
    validate.add_argument("input", type=Path)
    validate.add_argument("--json", action="store_true", dest="as_json")

    apply = subparsers.add_parser("apply", help="Apply an atomic JSON patch.")
    apply.add_argument("input", type=Path)
    apply.add_argument("patch", type=Path)
    apply.add_argument("--dry-run", action="store_true")
    apply.add_argument("-o", "--output", type=Path)

    schema = subparsers.add_parser("schema", help="Print a strict AiOffice JSON Schema.")
    schema.add_argument("-o", "--output", type=Path)
    schema.add_argument(
        "--kind",
        choices=(
            "document",
            "asset-ref",
            "border-line",
            "document-defaults",
            "document-section",
            "document-settings",
            "document-field",
            "header-footer-bindings",
            "header-footer-part",
            "image-block",
            "image-insert",
            "image-update",
            "named-style",
            "page-size",
            "paragraph-borders",
            "paragraph-style",
            "section-layout",
            "table-cell",
            "table-cell-borders",
            "table-cell-format",
            "table-column",
            "table-layout",
            "table-borders",
            "table-width",
            "text-range",
            "text-match",
        ),
        default="document",
        help="Select the strict model whose JSON Schema is printed.",
    )

    init = subparsers.add_parser("init", help="Initialize a small AiOffice project.")
    init.add_argument("directory", nargs="?", type=Path, default=Path("."))

    workspace = subparsers.add_parser(
        "workspace",
        help="Manage persistent artifacts and revisions in .aioffice.",
    )
    workspace_commands = workspace.add_subparsers(
        dest="workspace_command",
        required=True,
    )
    workspace_init = workspace_commands.add_parser("init", help="Initialize a workspace.")
    workspace_init.add_argument("root", nargs="?", type=Path, default=Path("."))

    workspace_import = workspace_commands.add_parser(
        "import",
        help="Copy a DOCX into the workspace as a tracked artifact.",
    )
    workspace_import.add_argument("input", type=Path)
    workspace_import.add_argument("--root", type=Path, default=Path("."))

    workspace_list = workspace_commands.add_parser(
        "list",
        help="List tracked workspace artifacts.",
    )
    workspace_list.add_argument("--root", type=Path, default=Path("."))

    workspace_capabilities = workspace_commands.add_parser(
        "capabilities",
        help="Report persistent workspace operations and guarantees.",
    )
    workspace_capabilities.add_argument("artifact_id", nargs="?")
    workspace_capabilities.add_argument("--root", type=Path, default=Path("."))

    workspace_inspect = workspace_commands.add_parser(
        "inspect",
        help="Inspect a tracked artifact revision.",
    )
    workspace_inspect.add_argument("artifact_id")
    workspace_inspect.add_argument("--root", type=Path, default=Path("."))
    workspace_inspect.add_argument("--revision", type=int)
    workspace_inspect.add_argument(
        "--response-format",
        choices=("summary", "compact", "expanded"),
        default="compact",
    )

    workspace_apply = workspace_commands.add_parser(
        "apply",
        help="Apply and persist an atomic patch as a new revision.",
    )
    workspace_apply.add_argument("artifact_id")
    workspace_apply.add_argument("patch", type=Path)
    workspace_apply.add_argument("--root", type=Path, default=Path("."))
    workspace_apply.add_argument("--dry-run", action="store_true")

    workspace_replace_image = workspace_commands.add_parser(
        "replace-image",
        help="Replace one image and persist the result as a new workspace revision.",
    )
    workspace_replace_image.add_argument("artifact_id")
    workspace_replace_image.add_argument("image_id")
    workspace_replace_image.add_argument("replacement", type=Path)
    workspace_replace_image.add_argument("--root", type=Path, default=Path("."))
    workspace_replace_image.add_argument("--media-type")
    workspace_replace_image.add_argument("--dry-run", action="store_true")
    workspace_replace_image.add_argument("--base-revision", type=int)

    workspace_insert_image = workspace_commands.add_parser(
        "insert-image-after",
        help="Insert one image and persist a new workspace revision.",
    )
    workspace_insert_image.add_argument("artifact_id")
    workspace_insert_image.add_argument("target")
    workspace_insert_image.add_argument("replacement", type=Path)
    workspace_insert_image.add_argument("--root", type=Path, default=Path("."))
    workspace_insert_image.add_argument("--width", required=True, type=float)
    workspace_insert_image.add_argument(
        "--width-unit",
        required=True,
        choices=("pt", "in", "cm", "mm", "px"),
    )
    workspace_insert_image.add_argument("--height", required=True, type=float)
    workspace_insert_image.add_argument(
        "--height-unit",
        required=True,
        choices=("pt", "in", "cm", "mm", "px"),
    )
    workspace_insert_image.add_argument("--alt-text", required=True)
    workspace_insert_image.add_argument("--title")
    workspace_insert_image.add_argument("--name")
    workspace_insert_image.add_argument("--image-id")
    workspace_insert_image.add_argument("--media-type")
    workspace_insert_image.add_argument(
        "--align",
        choices=("left", "center", "right", "justify"),
    )
    workspace_insert_image.add_argument("--dry-run", action="store_true")
    workspace_insert_image.add_argument("--base-revision", type=int)

    workspace_reconcile = workspace_commands.add_parser(
        "reconcile",
        help="Preview or commit an externally edited DOCX as a new revision.",
    )
    workspace_reconcile.add_argument("artifact_id")
    workspace_reconcile.add_argument("input", type=Path)
    workspace_reconcile.add_argument("--root", type=Path, default=Path("."))
    workspace_reconcile.add_argument("--commit", action="store_true")

    workspace_export = workspace_commands.add_parser(
        "export",
        help="Export a tracked revision without overwriting by default.",
    )
    workspace_export.add_argument("artifact_id")
    workspace_export.add_argument("output", type=Path)
    workspace_export.add_argument("--root", type=Path, default=Path("."))
    workspace_export.add_argument("--revision", type=int)
    workspace_export.add_argument("--overwrite", action="store_true")

    return parser


def _default_build_output(source: Path) -> Path:
    return source.with_suffix(".docx")


def _parse_page_numbers(
    value: str | None,
    *,
    max_pages: int,
) -> list[int] | None:
    if value is None:
        return None
    if isinstance(max_pages, bool) or not 1 <= max_pages <= 500:
        raise ValueError("--max-pages must be between 1 and 500.")
    selected: set[int] = set()
    for raw_item in value.split(","):
        item = raw_item.strip()
        if not item:
            raise ValueError("--pages contains an empty selection.")
        if "-" in item:
            parts = item.split("-")
            if len(parts) != 2 or not all(part.isdigit() for part in parts):
                raise ValueError(
                    "--pages ranges must use positive integers such as 3-5."
                )
            first, last = (int(part) for part in parts)
            if first < 1 or last < first:
                raise ValueError(
                    "--pages ranges must be positive and ascending."
                )
            if last - first + 1 > max_pages:
                raise ValueError(
                    "--pages selection exceeds the configured --max-pages limit."
                )
            range_pages = set(range(first, last + 1))
            if selected.intersection(range_pages):
                raise ValueError(
                    "--pages cannot contain duplicate or overlapping selections."
                )
            selected.update(range_pages)
        else:
            if not item.isdigit() or int(item) < 1:
                raise ValueError(
                    "--pages entries must be positive one-based integers."
                )
            page_number = int(item)
            if page_number in selected:
                raise ValueError("--pages cannot contain duplicate selections.")
            selected.add(page_number)
        if len(selected) > max_pages:
            raise ValueError(
                "--pages selection exceeds the configured --max-pages limit."
            )
    return sorted(selected)


def _run(args: argparse.Namespace) -> int:
    if args.command == "workspace":
        workspace_command = args.workspace_command
        if workspace_command == "init":
            workspace = Workspace.init(args.root)
            _json_dump(
                {
                    "workspace_id": workspace.id,
                    "root": str(workspace.root),
                }
            )
            return 0

        workspace = Workspace.open(args.root)
        if workspace_command == "import":
            document = workspace.import_document(args.input)
            _json_dump(
                {
                    "workspace_id": workspace.id,
                    "artifact": document.inspect(response_format="summary"),
                }
            )
            return 0
        if workspace_command == "list":
            _json_dump(
                {
                    "workspace_id": workspace.id,
                    "artifacts": workspace.list_artifacts(),
                }
            )
            return 0
        if workspace_command == "capabilities":
            _json_dump(workspace.capabilities(args.artifact_id))
            return 0
        if workspace_command == "inspect":
            document = workspace.open_document(
                args.artifact_id,
                revision=args.revision,
            )
            _json_dump(document.inspect(response_format=args.response_format))
            return 0
        if workspace_command == "apply":
            patch = _load_patch(args.patch)
            operations = patch.get("operations")
            if not isinstance(operations, list):
                raise ValueError("Patch envelope requires an operations array.")
            result = workspace.apply(
                args.artifact_id,
                operations,
                dry_run=args.dry_run,
                base_revision=patch.get("base_revision"),
                idempotency_key=patch.get("idempotency_key"),
            )
            _json_dump(result.model_dump())
            return 0 if result.success else 1
        if workspace_command == "replace-image":
            result = workspace.replace_image(
                args.artifact_id,
                args.image_id,
                args.replacement,
                media_type=args.media_type,
                dry_run=args.dry_run,
                base_revision=args.base_revision,
            )
            _json_dump(result.model_dump())
            return 0 if result.success else 1
        if workspace_command == "insert-image-after":
            result = workspace.insert_image_after(
                args.artifact_id,
                args.target,
                args.replacement,
                width={
                    "value": args.width,
                    "unit": args.width_unit,
                },
                height={
                    "value": args.height,
                    "unit": args.height_unit,
                },
                alt_text=args.alt_text,
                media_type=args.media_type,
                image_id=args.image_id,
                name=args.name,
                title=args.title,
                paragraph_style=(
                    {"alignment": args.align}
                    if args.align is not None
                    else None
                ),
                dry_run=args.dry_run,
                base_revision=args.base_revision,
            )
            _json_dump(result.model_dump())
            return 0 if result.success else 1
        if workspace_command == "reconcile":
            document = workspace.reconcile_document(
                args.artifact_id,
                args.input,
                commit=args.commit,
            )
            diagnostics = [
                diagnostic.model_dump(mode="json") for diagnostic in document.import_diagnostics
            ]
            _json_dump(
                {
                    "committed": args.commit,
                    "artifact": document.inspect(response_format="compact"),
                    "diagnostics": diagnostics,
                }
            )
            return (
                1
                if any(diagnostic["code"] == "IDENTITY_AMBIGUOUS" for diagnostic in diagnostics)
                else 0
            )
        if workspace_command == "export":
            path = workspace.export_document(
                args.artifact_id,
                args.output,
                revision=args.revision,
                overwrite=args.overwrite,
            )
            print(path)
            return 0
        raise AssertionError(f"Unhandled workspace command: {workspace_command}")

    if args.command in {"build", "export"}:
        document = open_artifact(args.input)
        output = args.output if args.output is not None else _default_build_output(args.input)
        path = document.export(output)
        print(path)
        return 0

    if args.command == "inspect":
        document = open_artifact(args.input)
        _json_dump(document.inspect(response_format=args.response_format))
        return 0

    if args.command == "capabilities":
        document = open_artifact(args.input)
        _json_dump(document.capabilities())
        return 0

    if args.command == "extract-image":
        document = open_artifact(args.input)
        image = document.read_image(args.image_id)
        output = image.write(args.output, overwrite=args.overwrite)
        _json_dump(
            {
                "image_id": image.image_id,
                "asset_id": image.asset_id,
                "media_type": image.media_type,
                "filename": image.filename,
                "sha256": image.sha256,
                "size_bytes": image.size_bytes,
                "output": str(output),
            }
        )
        return 0

    if args.command == "replace-image":
        if (
            not args.dry_run
            and args.output is not None
            and args.output.exists()
            and not args.overwrite
        ):
            raise FileExistsError(
                f"Refusing to overwrite existing DOCX {args.output}."
            )
        document = open_artifact(args.input)
        result = document.replace_image(
            args.image_id,
            args.replacement,
            media_type=args.media_type,
            dry_run=args.dry_run,
        )
        report = result.model_dump()
        if result.success and not args.dry_run:
            if args.output is None:
                raise ValueError(
                    "Committed image replacement requires --output; "
                    "the input DOCX is never overwritten."
                )
            assert result.document is not None
            result.document.export(args.output)
            report["output"] = str(args.output)
        _json_dump(report)
        return 0 if result.success else 1

    if args.command == "insert-image-after":
        if (
            not args.dry_run
            and args.output is not None
            and args.output.exists()
            and not args.overwrite
        ):
            raise FileExistsError(
                f"Refusing to overwrite existing DOCX {args.output}."
            )
        document = open_artifact(args.input)
        result = document.insert_image_after(
            args.target,
            args.replacement,
            width={
                "value": args.width,
                "unit": args.width_unit,
            },
            height={
                "value": args.height,
                "unit": args.height_unit,
            },
            alt_text=args.alt_text,
            media_type=args.media_type,
            image_id=args.image_id,
            name=args.name,
            title=args.title,
            paragraph_style=(
                {"alignment": args.align}
                if args.align is not None
                else None
            ),
            dry_run=args.dry_run,
        )
        report = result.model_dump()
        if result.success and not args.dry_run:
            if args.output is None:
                raise ValueError(
                    "Committed image insertion requires --output; "
                    "the input DOCX is never overwritten."
                )
            assert result.document is not None
            result.document.export(args.output)
            report["output"] = str(args.output)
        _json_dump(report)
        return 0 if result.success else 1

    if args.command == "render":
        if args.page is not None and args.format != "png":
            raise ValueError("--page is valid only with --format png.")
        document = open_artifact(args.input)
        provider = args.provider or (
            "semantic-html" if args.format == "html" else "libreoffice"
        )
        result = document.render(
            format=args.format,
            provider=provider,
            options={
                "dpi": args.dpi,
                "page_number": args.page,
                "timeout_seconds": args.timeout,
                "font_environment_hash": args.font_environment_hash,
            },
        )
        output = result.write(args.output)
        summary = result.summary()
        summary["output"] = str(output)
        _json_dump(summary)
        return 0

    if args.command == "render-pages":
        page_numbers = _parse_page_numbers(
            args.pages,
            max_pages=args.max_pages,
        )
        document = open_artifact(args.input)
        result = document.render_pages(
            page_numbers=page_numbers,
            options={
                "dpi": args.dpi,
                "timeout_seconds": args.timeout,
                "font_environment_hash": args.font_environment_hash,
            },
            analyze=args.analyze,
            max_pages=args.max_pages,
        )
        written = result.write(
            args.output_directory,
            stem=args.stem or args.input.stem,
            overwrite=args.overwrite,
        )
        page_paths = written["pages"]
        assert isinstance(page_paths, list)
        summary = result.summary()
        summary["outputs"] = {
            "pdf": str(written["pdf"]),
            "pages": [str(path) for path in page_paths],
        }
        _json_dump(summary)
        return 0

    if args.command == "validate":
        document = open_artifact(args.input)
        result = document.validate()
        if args.as_json:
            _json_dump(
                {
                    "valid": result.valid,
                    "diagnostics": [
                        diagnostic.model_dump(mode="json") for diagnostic in result.diagnostics
                    ],
                }
            )
        elif result.diagnostics:
            for diagnostic in result.diagnostics:
                location = f" ({diagnostic.path})" if diagnostic.path else ""
                print(
                    f"{diagnostic.severity.value.upper()} "
                    f"{diagnostic.code}{location}: {diagnostic.message}"
                )
            print("VALID" if result.valid else "INVALID")
        else:
            print("VALID")
        return 0 if result.valid else 1

    if args.command == "apply":
        document = open_artifact(args.input)
        patch = _load_patch(args.patch)
        operations = patch.get("operations")
        if not isinstance(operations, list):
            raise ValueError("Patch envelope requires an operations array.")
        result = document.apply(
            operations,
            dry_run=args.dry_run,
            base_revision=patch.get("base_revision"),
            idempotency_key=patch.get("idempotency_key"),
        )
        _json_dump(result.model_dump())
        if not result.success:
            return 1
        if not args.dry_run:
            if args.output is None:
                raise ValueError(
                    "Committed patches require --output; input files are never overwritten."
                )
            assert result.document is not None
            result.document.export(args.output)
            print(args.output, file=sys.stderr)
        return 0

    if args.command == "schema":
        schema_models = {
            "document": AiOfficeDocumentSpec,
            "asset-ref": AssetRef,
            "border-line": BorderLine,
            "document-defaults": DocumentDefaults,
            "document-section": DocumentSection,
            "document-settings": DocumentSettings,
            "document-field": DocumentField,
            "header-footer-bindings": HeaderFooterBindings,
            "header-footer-part": HeaderFooterPart,
            "image-block": ImageBlock,
            "image-insert": ImageInsert,
            "image-update": ImageUpdate,
            "named-style": NamedStyle,
            "page-size": PageSize,
            "paragraph-borders": ParagraphBorders,
            "paragraph-style": ParagraphStyle,
            "section-layout": SectionLayout,
            "table-cell": TableCell,
            "table-cell-borders": TableCellBorders,
            "table-cell-format": TableCellFormat,
            "table-column": TableColumn,
            "table-layout": TableLayout,
            "table-borders": TableBorders,
            "table-width": TableWidth,
            "text-range": TextRange,
            "text-match": TextMatch,
        }
        value = (
            json.dumps(
                schema_models[args.kind].model_json_schema(by_alias=True),
                ensure_ascii=False,
                indent=2,
            )
            + "\n"
        )
        if args.output is None:
            print(value, end="")
        else:
            args.output.parent.mkdir(parents=True, exist_ok=True)
            args.output.write_text(value, encoding="utf-8")
            print(args.output)
        return 0

    if args.command == "init":
        directory: Path = args.directory
        directory.mkdir(parents=True, exist_ok=True)
        config = directory / "aioffice.toml"
        report = directory / "report.json"
        existing = [path for path in (config, report) if path.exists()]
        if existing:
            names = ", ".join(str(path) for path in existing)
            raise ValueError(f"Refusing to overwrite existing files: {names}")
        config.write_text(
            '[project]\nname = "my-aioffice-project"\n\n'
            '[build]\nsource = "report.json"\noutput = "output/report.docx"\n',
            encoding="utf-8",
        )
        document = (
            DocumentBuilder(title="My AiOffice Document")
            .heading("My AiOffice Document", id="document_title")
            .paragraph("Edit this declarative source and run aioffice build report.json.")
            .build()
        )
        report.write_text(document.to_json(), encoding="utf-8")
        (directory / "output").mkdir(exist_ok=True)
        print(directory.resolve())
        return 0

    raise AssertionError(f"Unhandled command: {args.command}")


def main(argv: Sequence[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)
    try:
        return _run(args)
    except (AiOfficeError, OSError, ValueError, json.JSONDecodeError) as error:
        print(f"aioffice: error: {error}", file=sys.stderr)
        return 2


if __name__ == "__main__":
    raise SystemExit(main())
