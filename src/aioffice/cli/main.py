"""Command-line entry point."""

from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path
from typing import Any, Sequence

from aioffice._version import __version__
from aioffice.core.errors import AiOfficeError
from aioffice.documents import Document, DocumentBuilder, open_artifact
from aioffice.spec.models import AiOfficeDocumentSpec


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

    validate = subparsers.add_parser("validate", help="Validate a document.")
    validate.add_argument("input", type=Path)
    validate.add_argument("--json", action="store_true", dest="as_json")

    apply = subparsers.add_parser("apply", help="Apply an atomic JSON patch.")
    apply.add_argument("input", type=Path)
    apply.add_argument("patch", type=Path)
    apply.add_argument("--dry-run", action="store_true")
    apply.add_argument("-o", "--output", type=Path)

    schema = subparsers.add_parser("schema", help="Print the Document Spec JSON Schema.")
    schema.add_argument("-o", "--output", type=Path)

    init = subparsers.add_parser("init", help="Initialize a small AiOffice project.")
    init.add_argument("directory", nargs="?", type=Path, default=Path("."))

    return parser


def _default_build_output(source: Path) -> Path:
    return source.with_suffix(".docx")


def _run(args: argparse.Namespace) -> int:
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
                raise ValueError("Committed patches require --output; input files are never overwritten.")
            assert result.document is not None
            result.document.export(args.output)
            print(args.output, file=sys.stderr)
        return 0

    if args.command == "schema":
        value = json.dumps(
            AiOfficeDocumentSpec.model_json_schema(by_alias=True),
            ensure_ascii=False,
            indent=2,
        ) + "\n"
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
