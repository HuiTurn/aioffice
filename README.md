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

The development branch is now `0.2.0.dev8`. It adds lossless DOCX opening, semantic
projection over a native package, persistent native identities, local revision
workspaces, copy-on-write native parts, exact text-range formatting, AI-addressable
named styles, document defaults, ordered page/section models, reusable header/footer
parts, structured dynamic fields, explicit table geometry, semantic diffs, render
contracts, and fidelity reports. Workbook, presentation, PDF, native visual
rendering, and MCP remain planned.

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
`paragraph.format`, `text.format`, `node.remove`, `style.define`, `style.apply`,
`style.format`, `section.format`, `field.update`, `table.format`, and
`table.column.format`. Ask the artifact before planning an edit:

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

Named styles are stable, AI-addressable layout rules rather than copied formatting.
The resolver applies document defaults, the complete `based_on` chain, node direct
formatting, and finally span formatting:

```python
doc = (
    DocumentBuilder()
    .define_style({
        "id": "Executive",
        "name": "Executive",
        "semantic_role": "custom",
        "based_on": "Normal",
        "paragraph_style": {
            "spacing_after": {"value": 14, "unit": "pt"},
            "keep_together": True,
        },
        "text_style": {
            "font_size": {"value": 13, "unit": "pt"},
            "color": "#7A1F5B",
            "bold": True,
        },
    })
    .paragraph("Board decision", id="decision", style_ref="Executive")
    .build()
)

result = doc.apply([
    {
        "op": "style.format",
        "target": "@Executive",
        "paragraph": {
            "set": {"spacing_after": {"value": 18, "unit": "pt"}}
        },
        "text": {
            "set": {"color": "#1F4E78"},
            "clear": ["bold"],
        },
    }
])
```

Imported `w:style` definitions, `w:docDefaults`, inheritance links, quick-style
metadata, and paragraph `w:pStyle` references are projected into the Spec. Native
style patches update only supported properties in `word/styles.xml`; unknown style
XML and every untouched package part remain byte-for-byte preserved.

Sections are ordered, AI-addressable page regions. The first section starts at the
document root; each later section is anchored at its first content node. Page size,
orientation, margins, gutter, header/footer distance, columns, vertical alignment,
first-page behavior, and Word section-start type all use strict values:

```python
doc = DocumentBuilder(
    sections=[
        {
            "id": "cover_section",
            "start_at": None,
            "layout": {
                "page_size": {"preset": "letter"},
                "margin_top": {"value": 1, "unit": "in"},
                "margin_right": {"value": 1, "unit": "in"},
                "margin_bottom": {"value": 1, "unit": "in"},
                "margin_left": {"value": 1, "unit": "in"},
            },
        },
        {
            "id": "analysis_section",
            "start_at": "analysis",
            "layout": {
                "start_type": "next_page",
                "page_size": {
                    "preset": "a4",
                    "orientation": "landscape",
                },
                "columns": {
                    "count": 2,
                    "spacing": {"value": 24, "unit": "pt"},
                    "separator": True,
                },
                "page_number_start": 1,
                "page_number_format": "lower_roman",
            },
        },
    ]
).paragraph("Cover", id="cover").paragraph("Analysis", id="analysis").build()

result = doc.apply([
    {
        "op": "section.format",
        "target": "#analysis_section",
        "set": {"margin_left": {"value": 18, "unit": "mm"}},
        "clear": ["footer_distance"],
    }
])
```

Generated multi-section DOCX uses Word's native section placement rules. Imported
paragraph-level and final body-level `w:sectPr` elements are projected separately,
and `section.format` changes only the selected native section properties. Unknown
section XML remains native and untouched. See
[the page and section contract](docs/section-layout.md).

Headers and footers use reusable parts rather than copied strings. Each section may
explicitly bind `default`, `first`, and `even` header/footer variants; a missing
binding means “inherit the same slot from the previous section”:

```python
doc = DocumentBuilder(
    settings={"even_and_odd_headers": True},
    header_footers=[
        {
            "id": "report_header",
            "kind": "header",
            "content": [
                {
                    "id": "report_header_text",
                    "type": "paragraph",
                    "text": "Confidential",
                }
            ],
        },
        {
            "id": "report_footer",
            "kind": "footer",
            "content": [
                {
                    "id": "report_footer_text",
                    "type": "paragraph",
                    "content": [
                        {"text": "Page "},
                        {
                            "id": "current_page",
                            "type": "field",
                            "kind": "page_number",
                            "cached_result": "1",
                        },
                        {"text": " of "},
                        {
                            "id": "total_pages",
                            "type": "field",
                            "kind": "page_count",
                            "cached_result": "1",
                        },
                    ],
                }
            ],
        },
    ],
    sections=[
        {
            "id": "main_section",
            "header_footer": {
                "header_default": "report_header",
                "footer_default": "report_footer",
            },
        }
    ],
).paragraph("Report body", id="body").build()
```

The paragraph IDs inside a header/footer are regular edit selectors. `text.replace`,
`text.format`, and `paragraph.format` lower directly into the referenced
`headerN.xml` or `footerN.xml` part. PAGE, NUMPAGES, SECTION, and SECTIONPAGES are
structured fields with their own stable IDs. Their displayed result is explicitly a
non-authoritative cache:

```python
result = doc.apply([
    {
        "op": "field.update",
        "target": "#current_page",
        "set": {"number_format": "upper_roman"},
    }
])
```

Generated fields are marked dirty and `update_fields_on_open` is enabled unless
explicitly disabled. Unknown field instructions remain structured but read-only;
drawings, objects, tables, and malformed field structures remain opaque. See
[the dynamic field contract](docs/dynamic-fields.md) and
[the header/footer contract](docs/header-footer.md).

Document tables keep semantic column keys and stable column/row IDs while exposing
layout geometry in explicit units:

```python
table_doc = DocumentBuilder().table(
    id="metrics",
    columns=[
        {
            "id": "metric_column",
            "key": "metric",
            "title": "Metric",
            "width": {"value": 120, "unit": "pt"},
        },
        {
            "id": "value_column",
            "key": "value",
            "title": "Value",
            "data_type": "number",
            "width": {"value": 180, "unit": "pt"},
        },
    ],
    rows=[
        {
            "id": "revenue_row",
            "values": {"metric": "Revenue", "value": 42},
            "allow_break_across_pages": False,
        }
    ],
    layout={
        "preferred_width": {"mode": "percent", "value": 90},
        "alignment": "center",
        "algorithm": "fixed",
        "repeat_header": True,
        "cell_margin_left": {"value": 6, "unit": "pt"},
        "cell_margin_right": {"value": 6, "unit": "pt"},
    },
).build()

result = table_doc.apply(
    [
        {
            "op": "table.column.format",
            "target": "#metrics",
            "column": "#value_column",
            "set": {"width": {"value": 200, "unit": "pt"}},
        }
    ]
)
```

Imported regular DOCX grids support selective table and column formatting. Merged,
shifted, or otherwise irregular grids remain readable but reject column-width
mutation atomically; cell content is currently projected as plain text. See
[the table layout contract](docs/table-layout.md).

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
`node.append`, `node.insert_after`, `node.remove`, `node.update`, `style.define`,
`style.apply`, `style.format`, `section.format`, `field.update`, `table.format`, and
`table.column.format`. Imported DOCX documents expose the native-safe subset reported
by `capabilities()`. Selectors use stable content, section, header/footer block,
field, table, column, or row identities in this release.

## CLI

```bash
aioffice inspect examples/report.json
aioffice capabilities existing.docx
aioffice validate examples/report.json
aioffice build examples/report.json --output report.docx
aioffice export examples/report.json --to report.html
aioffice schema --output document.schema.json
aioffice schema --kind named-style --output named-style.schema.json
aioffice schema --kind document-defaults --output document-defaults.schema.json
aioffice schema --kind page-size --output page-size.schema.json
aioffice schema --kind section-layout --output section-layout.schema.json
aioffice schema --kind document-section --output document-section.schema.json
aioffice schema --kind document-settings --output document-settings.schema.json
aioffice schema --kind document-field --output document-field.schema.json
aioffice schema --kind table-width --output table-width.schema.json
aioffice schema --kind table-layout --output table-layout.schema.json
aioffice schema --kind table-column --output table-column.schema.json
aioffice schema --kind header-footer-bindings --output header-footer-bindings.schema.json
aioffice schema --kind header-footer-part --output header-footer-part.schema.json
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
