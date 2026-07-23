# AiOffice

AiOffice is an AI-native, declarative document engine. It lets an agent describe the
document it wants, validates that intent as a strict spec, and compiles it into office
formats without exposing low-level Word object APIs.

This first `0.1.0` release is an intentionally small, usable vertical slice of the
larger AiOffice architecture:

- strict AiOffice Document Spec 1.0 draft models;
- stable semantic node IDs;
- a Python API and convenience builder;
- JSON and Markdown input;
- JSON, Markdown, semantic HTML, and DOCX output;
- machine-readable validation diagnostics;
- atomic, revision-checked document patches;
- a CLI shared with the Python core.

Workbook, presentation, PDF, DOCX import, persistent revision history, rendering, and
MCP are planned, but are not claimed by this release.

## Install

```bash
pip install aioffice
```

AiOffice requires Python 3.11 or newer.

## Python quick start

```python
from aioffice.documents import DocumentBuilder

doc = (
    DocumentBuilder(title="Project Report", theme="business-clean")
    .heading("Project Report", id="report_title")
    .paragraph("The first delivery milestone is complete.", id="status")
    .bullet_list(["Validated spec", "Generated DOCX", "Published HTML preview"])
    .build()
)

validation = doc.validate()
assert validation.valid

doc.export("report.json")
doc.export("report.md")
doc.export("report.html")
doc.export("report.docx")
```

You can also create a document directly from the strict spec:

```python
from aioffice.documents import Document

doc = Document.from_spec({
    "metadata": {"title": "Project Report"},
    "theme": {"ref": "business-clean"},
    "content": [
        {"type": "heading", "level": 1, "text": "Project Report"},
        {"type": "paragraph", "text": "The first milestone is complete."},
    ],
})
```

## Atomic patch

Patches never mutate the source `Document`. A successful result contains the next
logical revision:

```python
result = doc.apply(
    [
        {
            "op": "text.replace",
            "target": "#status",
            "search": "complete",
            "replacement": "approved",
        }
    ],
    base_revision=doc.revision,
    dry_run=True,
)

assert result.success
preview = result.document
```

V0.1 supports `text.replace`, `node.append`, `node.insert_after`, `node.remove`, and
`node.update`. Selectors are stable node IDs in this release.

## CLI

```bash
aioffice inspect examples/report.json
aioffice validate examples/report.json
aioffice build examples/report.json --output report.docx
aioffice export examples/report.json --to report.html
aioffice schema --output document.schema.json
```

Patch files may be an operation array or an envelope:

```json
{
  "base_revision": 1,
  "idempotency_key": "agent-task-001",
  "operations": [
    {
      "op": "text.replace",
      "target": "#status",
      "search": "第一阶段",
      "replacement": "第二阶段"
    }
  ]
}
```

Preview or commit it without overwriting the input:

```bash
aioffice apply examples/report.json patch.json --dry-run
aioffice apply examples/report.json patch.json --output updated.json
```

## Development and release

```bash
python -m unittest discover -s tests -v
python -m build
python -m twine check dist/*
```

Production releases use PyPI Trusted Publishing. The tag must match the package
version in `src/aioffice/_version.py`; pushing it starts
`.github/workflows/publish.yml`:

```bash
git tag v0.1.0
git push origin v0.1.0
```

No long-lived PyPI API token is stored in GitHub.

The current spec is a draft. Compatibility will be maintained within the `0.1.x`
series where practical, but the public model can still evolve before 1.0.
