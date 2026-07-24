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

The development branch is now `0.2.0.dev24`. It adds lossless DOCX opening, semantic
projection over a native package, persistent native identities, local revision
workspaces, copy-on-write native parts, exact text-range formatting, AI-addressable
named styles, document defaults, ordered page/section models, reusable header/footer
parts, structured dynamic fields, explicit table geometry, logical merged cells,
rich table-cell paragraphs, explicit table/cell border control, paragraph
background/border surfaces, conservative native image projection, verified asset
extraction, selective native image metadata and geometry updates, occurrence-scoped
copy-on-write image replacement, addressable native inline image insertion, direct
image-paragraph layout formatting, semantic diffs, isolated LibreOffice/Poppler
native rendering, root append plus bidirectional stable-ID native
paragraph/heading insertion and block reordering, consistent multi-page evidence,
page occupancy diagnostics, visual-regression contracts, and fidelity reports.
Workbook, presentation, PDF editing, and MCP remain planned.

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

Image bytes deliberately stay out of the JSON Spec. A simple body paragraph
containing exactly one embedded inline DrawingML picture is projected as an
AI-addressable `image` block with physical extent, alternative text, media type,
filename, byte count, and SHA-256 asset identity:

```python
image = next(
    node for node in doc.inspect()["nodes"]
    if node["type"] == "image"
)

verified = doc.read_image(image["id"])
assert verified.sha256 == image["asset"]["sha256"]
verified.write("extracted/" + verified.filename)
```

Supported projected images can be resized or given accessible metadata without
rewriting their binary part or relationship:

```python
result = doc.apply([
    {
        "op": "image.update",
        "target": f"#{image['id']}",
        "set": {
            "width": {"value": 3, "unit": "in"},
            "alt_text": "Quarterly revenue by region",
            "title": "Revenue chart",
        },
    }
])
assert result.success
result.document.export("updated.docx")
```

Setting one dimension preserves the current aspect ratio; setting both applies the
exact requested extent. `alt_text` and `title` may be removed with
`"clear": ["alt_text", "title"]`. The native patch updates both DrawingML extent
records while preserving image bytes and package relationships. It requires the
attached native DOCX, so a detached JSON snapshot cannot perform this operation.

The projected image ID also addresses its native host paragraph. Reuse
`paragraph.format` to control layout around an existing picture without touching its
DrawingML or bytes:

```python
result = doc.apply([
    {
        "op": "paragraph.format",
        "target": f"#{image['id']}",
        "set": {
            "alignment": "center",
            "spacing_before": {"value": 10, "unit": "pt"},
            "spacing_after": {"value": 12, "unit": "pt"},
            "keep_together": True,
        },
    }
])
assert result.success
```

The same strict `ParagraphStyle` fields and `set`/`clear` semantics used by text
paragraphs apply to the image paragraph, including indentation, page-flow controls,
solid background, and supported physical borders. The operation appears on each
projected image's `supported_operations` list.

Image binaries also use an explicit out-of-band write path:

```python
result = doc.replace_image(
    image["id"],
    "assets/revenue-chart.png",
    media_type="image/png",
)
assert result.success
result.document.export("replaced.docx")
```

AiOffice signature-checks and bounds the raster input, creates a content-addressed
native image part and a new relationship for only that occurrence, and preserves its
stable image ID, displayed extent, alternative text, and title. Other occurrences
that shared the old image remain unchanged. Raw JSON Patch cannot carry the binary.

New inline pictures use the same bounded asset channel and require explicit layout:

```python
result = doc.insert_image_after(
    "#status",
    "assets/expert-workflow.png",
    width={"value": 3, "unit": "in"},
    height={"value": 1.5, "unit": "in"},
    alt_text="Expert workflow with three approval stages",
    image_id="expert_workflow",
    paragraph_style={"alignment": "center"},
)
assert result.success
result.document.export("inserted.docx")
```

The target must be a mapped top-level body node. AiOffice inserts after its last
native element, which keeps multi-paragraph lists addressable as one semantic node.
The new paragraph, DrawingML geometry, relationship, asset and identity manifest are
created atomically. Explicit width, height and nonblank alternative text avoid
model-side DPI guessing and inaccessible output.

The equivalent CLI is:

```bash
aioffice extract-image existing.docx IMAGE_ID -o extracted.png
aioffice replace-image existing.docx IMAGE_ID replacement.png -o replaced.docx
aioffice insert-image-after existing.docx TARGET replacement.png \
  --width 3 --width-unit in --height 1.5 --height-unit in \
  --alt-text "Expert workflow" --align center -o inserted.docx
```

Persistent workspaces expose the same operations through
`Workspace.replace_image(...)`, `Workspace.insert_image_after(...)`, and matching
CLI commands, recording verified asset and insertion metadata but never base64 in the
revision log.

The read path re-resolves the trusted native paragraph and OPC relationship, then
verifies the asset record, media type, size, and content hash before returning bytes.
Mixed text/picture paragraphs, floating anchors, linked images, multiple pictures,
crops, transforms, effects, tables, headers/footers, VML, OLE, and embedded objects
remain explicit opaque native content. They are preserved losslessly and rendered
through the native provider rather than flattened into a misleading image model. See
[the native image and asset contract](docs/native-images.md).

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
`style.format`, `section.format`, `field.update`, `image.insert_after`,
`image.replace`, `image.update`, `table.format`, `table.column.format`, and
`table.cell.format`. Ask the artifact before planning an edit:

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
            "background_color": "#EAF2F8",
            "borders": {
                "left": {
                    "style": "single",
                    "width": {"value": 3, "unit": "pt"},
                    "color": "#1F4E78",
                    "space": {"value": 8, "unit": "pt"},
                },
            },
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

`paragraph_style.background_color` creates a solid paragraph-wide surface.
`paragraph_style.borders` controls top/right/bottom/left edges with the same strict
border line model used by tables. Border edges inherit independently through the
named-style chain: a direct bottom edge can override a style while its other edges
continue to inherit. Clearing removes direct XML; `style: "none"` explicitly
suppresses an inherited edge. Pattern/theme shading and Word's `between`/`bar`
borders remain native-only and losslessly preserved. See
[the paragraph surface contract](docs/paragraph-surfaces.md).

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

Document tables keep semantic column keys and stable column/row/cell IDs while
exposing layout geometry in explicit units. The `values` form remains a compact
input shorthand and is normalized to cells:

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
        "borders": {
            "top": {
                "style": "single",
                "width": {"value": 1.5, "unit": "pt"},
                "color": "#1F4E78",
            },
            "right": {
                "style": "single",
                "width": {"value": 1.5, "unit": "pt"},
                "color": "#1F4E78",
            },
            "bottom": {
                "style": "single",
                "width": {"value": 1.5, "unit": "pt"},
                "color": "#1F4E78",
            },
            "left": {
                "style": "single",
                "width": {"value": 1.5, "unit": "pt"},
                "color": "#1F4E78",
            },
            "inside_horizontal": {
                "style": "single",
                "width": {"value": 0.5, "unit": "pt"},
                "color": "#D9E2F3",
            },
            "inside_vertical": {"style": "none"},
        },
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

Logical cells can span rows or columns, contain multiple rich paragraphs, and carry
cell-local formatting:

```python
from aioffice import Document

rich_table = Document.from_spec({
    "content": [{
        "id": "summary_table",
        "type": "table",
        "columns": [
            {"key": "summary", "title": "Summary"},
            {"key": "detail", "title": "Detail"},
        ],
        "rows": [{
            "id": "summary_row",
            "cells": [{
                "id": "summary_cell",
                "column_key": "summary",
                "column_span": 2,
                "content": [{
                    "id": "summary_text",
                    "type": "paragraph",
                    "content": [
                        {"type": "text", "text": "Approved", "marks": ["strong"]},
                        {"type": "text", "text": " for release"},
                    ],
                }],
                "format": {
                    "vertical_alignment": "center",
                    "background_color": "#EAF2F8",
                    "margin_left": {"value": 8, "unit": "pt"},
                },
            }],
        }],
    }],
})

result = rich_table.apply([{
    "op": "table.cell.format",
    "target": "#summary_table",
    "cell": "#summary_cell",
    "set": {
        "background_color": "#FFF2CC",
        "borders": {
            "bottom": {
                "style": "double",
                "width": {"value": 2, "unit": "pt"},
                "color": "#C00000",
            },
        },
    },
}])
```

Imported DOCX grids are analyzed as logical cells before `gridSpan` and `vMerge`
are exposed. Regular grids support selective column widths; merged grids reject
column-width mutation but still support formatting a mapped anchor cell. Supported
cell paragraphs use the normal text and paragraph operations. Cells containing
drawings, nested tables, dynamic fields, or malformed content fall back to a
read-only text projection while their native XML remains intact. See
[the table layout contract](docs/table-layout.md) and
[the table cell contract](docs/table-cells.md).

Border edges use explicit styles, widths, colors, and optional spacing. Clearing the
`borders` property removes known direct border XML so table styles can apply again;
`{"style": "none"}` writes an explicit no-border edge. A direct cell edge wins over
the conflicting table perimeter or internal-grid edge.

`doc.render()` defaults to a semantic HTML preview whose contract explicitly reports
`fidelity="approximate"` and `verification_status="preview_only"`. A local
LibreOffice and Poppler installation enables native-compatible PDF and page PNG
evidence:

```python
pdf = doc.render(format="pdf", provider="libreoffice")
pdf.write("report-render.pdf")
assert pdf.metadata["page_count"] >= 1

page = doc.render(
    format="png",
    provider="libreoffice",
    options={"page_number": 1, "dpi": 144},
)
page.write("report-page-1.png")
```

For an entire document, render the PDF only once and derive a bounded, consistent
page set from it:

```python
evidence = doc.render_pages(
    options={"dpi": 144},
    analyze=True,
    max_pages=100,
)
paths = evidence.write("evidence", stem="report")
assert len(evidence.pages) == evidence.page_count
```

Page analysis reports the background, ink ratio, content bounding box, four-side
whitespace, apparent blank pages, and visible content near a page edge. It requires
`pip install "aioffice[render]"`; rendering pages without analysis does not require
Pillow.

Each job uses an isolated LibreOffice user profile and reports engine versions,
source/output hashes, page count, font-environment hash, page dimensions, and
diagnostics. Native evidence still reports `verification_status="unverified"`:
successful rendering proves that inspectable pages exist, not that an aesthetic
review has passed. See [the native rendering contract](docs/native-rendering.md)
and [style, diff, and rendering contracts](docs/style-rendering.md).

The equivalent CLI workflow is:

```bash
aioffice render report.docx --format pdf -o report-render.pdf
aioffice render report.docx --format png --page 1 --dpi 144 -o report-page-1.png
aioffice render-pages report.docx --analyze --output-directory evidence
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

Imported DOCX documents can receive a new paragraph or heading without rebuilding
their existing content:

```python
result = doc.apply([
    {
        "op": "node.insert_after",
        "target": "#executive_summary",
        "content": {
            "id": "recommendation",
            "type": "paragraph",
            "content": [
                {"type": "text", "text": "Recommendation: ", "marks": ["strong"]},
                {"type": "text", "text": "approve the proposed plan."},
            ],
            "paragraph_style": {
                "spacing_before": {"value": 8, "unit": "pt"},
                "spacing_after": {"value": 8, "unit": "pt"},
            },
        },
    }
])
assert result.success
```

Only the new `w:p` is compiled. Existing XML, DrawingML, relationships, and
unsupported native features remain untouched. A caller-selected new ID can be
targeted again later in the same Patch; an omitted ID is returned in change
evidence. Rich text, direct formatting, internal/external hyperlinks, and normalized
document fields are supported. Use `node.insert_before` for symmetric placement,
including insertion at the beginning of the document. Inserting before a later
section's first node safely rebinds that section's `start_at`. Use `node.append`
with target `$` when the AI should add content to the last section without first
discovering the final content ID; native lowering inserts it before the terminal
body `w:sectPr`. See
[native text insertion](docs/native-text-insertion.md).

Existing top-level content can be reordered without reconstructing it or addressing
an array index:

```python
result = doc.apply([
    {
        "op": "node.move_before",
        "target": "#risk_table",
        "before": "#executive_summary",
    }
])
assert result.success
```

For imported DOCX, AiOffice moves the target's complete mapped XML range. A
multi-paragraph list remains one contiguous unit, DrawingML and unknown XML stay in
their original elements, and every native reference is reindexed. `node.move_after`
and `node.move_before` cover every relative position without array indexes. The
conservative dev24 boundary permits moves only within one semantic section, refuses
moving a section start anchor, and rebinds `section.start_at` when prepending within
a later section. Native elements carrying `w:sectPr` remain immovable. See
[the structural editing contract](docs/structural-editing.md).

`node.remove` uses the same native-authority boundary. It removes the complete mapped
XML range, refuses native section carriers, attaches an identity manifest on the
first structural edit to a third-party DOCX, and preserves now-unreferenced
relationships or parts rather than guessing that they are safe to delete.

Semantic documents support `text.replace`, `paragraph.format`, `text.format`,
`node.append`, `node.insert_after`, `node.insert_before`, `node.move_after`,
`node.move_before`, `node.remove`, `node.update`, `style.define`, `style.apply`, `style.format`,
`section.format`, `field.update`,
`table.format`, `table.column.format`, and `table.cell.format`. Imported DOCX
documents support incremental before/after insertion for paragraphs and headings
plus root append, and additionally expose safe native image operations reported by
`capabilities()`.
Selectors use stable content, section, header/footer block, field, image, table,
column, row, cell, or rich cell-paragraph identities in this release.

## CLI

```bash
aioffice inspect examples/report.json
aioffice capabilities existing.docx
aioffice validate examples/report.json
aioffice build examples/report.json --output report.docx
aioffice export examples/report.json --to report.html
aioffice schema --output document.schema.json
aioffice schema --kind named-style --output named-style.schema.json
aioffice schema --kind paragraph-style --output paragraph-style.schema.json
aioffice schema --kind paragraph-borders --output paragraph-borders.schema.json
aioffice schema --kind document-defaults --output document-defaults.schema.json
aioffice schema --kind page-size --output page-size.schema.json
aioffice schema --kind section-layout --output section-layout.schema.json
aioffice schema --kind document-section --output document-section.schema.json
aioffice schema --kind document-settings --output document-settings.schema.json
aioffice schema --kind document-field --output document-field.schema.json
aioffice schema --kind table-width --output table-width.schema.json
aioffice schema --kind table-layout --output table-layout.schema.json
aioffice schema --kind table-column --output table-column.schema.json
aioffice schema --kind table-cell --output table-cell.schema.json
aioffice schema --kind table-cell-format --output table-cell-format.schema.json
aioffice schema --kind border-line --output border-line.schema.json
aioffice schema --kind table-borders --output table-borders.schema.json
aioffice schema --kind table-cell-borders --output table-cell-borders.schema.json
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
