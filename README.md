# AiOffice

AiOffice is an AI-native, declarative document engine. It lets an agent describe the
document it wants, validates that intent as a strict spec, and compiles it into office
formats without exposing low-level Word object APIs.

The `0.1.0` release is an intentionally small, usable vertical slice of the larger
AiOffice architecture:

- strict AiOffice Document Spec 1.0 draft models;
- stable semantic node IDs;
- a Python API and convenience builder;
- JSON and Markdown input;
- JSON, Markdown, semantic HTML, and DOCX output;
- machine-readable validation diagnostics;
- atomic, revision-checked document patches;
- a CLI shared with the Python core.

The development branch is now `0.2.0.dev3`. It adds lossless DOCX opening, semantic
projection over a native package, persistent native identities, local revision
workspaces, copy-on-write native parts, strict paragraph/text formatting, semantic
diffs, render contracts, and fidelity reports. Workbook, presentation, PDF, native
visual rendering, and MCP remain planned.

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

Open an existing DOCX without rebuilding its unknown or unsupported parts:

```python
import aioffice

doc = aioffice.open("existing.docx", roundtrip="preserve_unknown")
assert doc.origin == "native"

result = doc.apply([
    {
        "op": "text.replace",
        "target": "#para_000001",
        "search": "Draft",
        "replacement": "Approved",
    }
], dry_run=True)

assert result.success
print(result.fidelity)
result.document.export("updated.docx")
```

Exporting an imported DOCX without changes returns the exact original package bytes.
When a supported edit is applied, AiOffice rewrites only the affected native part and
preserves untouched part payloads.

AiOffice-generated DOCX files embed a versioned identity manifest. Artifact IDs,
semantic node IDs, native anchors, and revisions therefore survive export and reopen.
Third-party documents can keep the same guarantees through a local workspace:

```python
from aioffice import Workspace

workspace = Workspace.init("project")
doc = workspace.import_document("existing.docx")

result = workspace.apply(
    doc.id,
    [{
        "op": "text.replace",
        "target": f"#{doc.to_spec()['content'][0]['id']}",
        "search": "Draft",
        "replacement": "Approved",
    }],
    base_revision=doc.revision,
    idempotency_key="approve-first-paragraph",
)

assert result.success
revision_one = workspace.checkout(doc.id, revision=1)
revision_two = workspace.open_document(doc.id)
```

Use `workspace.reconcile_document(...)` to preview an externally edited DOCX. A
commit is refused when native identity is ambiguous. The detailed invariants are in
[the native round-trip architecture](docs/native-roundtrip.md).

Native DOCX lowering in this development version supports `text.replace`,
`paragraph.format`, `text.format`, and `node.remove`. Ask the artifact before
planning an edit:

```python
capabilities = doc.capabilities()
assert "text.replace" in capabilities["operations"]
```

Formatting values always include units. This prevents an agent from confusing
points, pixels, inches, and native OOXML twips:

```python
result = doc.apply([
    {
        "op": "paragraph.format",
        "target": "#para_000001",
        "set": {
            "alignment": "justify",
            "spacing_after": {"value": 8, "unit": "pt"},
            "line_spacing": {"rule": "multiple", "value": 1.25},
        },
    },
    {
        "op": "text.format",
        "target": "#para_000001",
        "match": {
            "text": "重要结论",
            "occurrence": 1,
        },
        "set": {
            "font_size": {"value": 10.5, "unit": "pt"},
            "color": "#1F4E78",
        },
    },
])

assert result.success
print(result.diff.summary)
```

`text.format` can target the whole node, an exact text occurrence, or a half-open
Unicode code-point range such as
`{"range": {"start": 4, "end": 10, "unit": "unicode_codepoint"}}`. Imported
mixed Word runs and hyperlinks are projected as rich `TextSpan` content, so an
agent can inspect and edit local formatting without losing link targets.

`doc.render()` currently returns a semantic HTML preview whose contract explicitly
reports `fidelity="approximate"` and `verification_status="preview_only"`. It must
not be treated as proof of Word pagination. See
[style, diff, and rendering contracts](docs/style-rendering.md).

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

Semantic documents support `text.replace`, `paragraph.format`, `text.format`,
`node.append`, `node.insert_after`, `node.remove`, and `node.update`. Imported DOCX
documents expose the smaller native-safe subset reported by `capabilities()`.
Selectors are stable node IDs in this release.

## CLI

```bash
aioffice inspect examples/report.json
aioffice capabilities existing.docx
aioffice validate examples/report.json
aioffice build examples/report.json --output report.docx
aioffice export examples/report.json --to report.html
aioffice schema --output document.schema.json
aioffice schema --kind text-range --output text-range.schema.json

aioffice workspace init project
aioffice workspace import existing.docx --root project
aioffice workspace list --root project
aioffice workspace capabilities ARTIFACT_ID --root project
aioffice workspace inspect ARTIFACT_ID --root project
aioffice workspace apply ARTIFACT_ID patch.json --root project
aioffice workspace reconcile ARTIFACT_ID edited.docx --root project
aioffice workspace reconcile ARTIFACT_ID edited.docx --root project --commit
aioffice workspace export ARTIFACT_ID updated.docx --root project
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
python -m pip install -e ".[dev]"
python -m unittest discover -s tests -v
ruff check src tests
pyright src
python -m build
python -m twine check dist/*
```

Production releases use PyPI Trusted Publishing. The tag must match the package
version in `src/aioffice/_version.py`; pushing it starts
`.github/workflows/publish.yml`:

```bash
git tag vX.Y.Z
git push origin vX.Y.Z
```

No long-lived PyPI API token is stored in GitHub.

The current spec is a draft. Compatibility will be maintained within the `0.1.x`
series where practical, but the public model can still evolve before 1.0.
