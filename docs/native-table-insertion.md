# Incremental native table insertion

AiOffice `0.2.0.dev26` can insert a complete semantic table into an imported DOCX
without rebuilding any existing body block. The JSON Spec describes the new table;
the attached OPC package remains authoritative for everything already in the file.

## Operations

A table uses the ordinary stable-ID structural operations:

```json
[
  {
    "op": "node.insert_after",
    "target": "#executive_summary",
    "content": {
      "id": "metrics_table",
      "type": "table",
      "columns": [
        {
          "id": "metric_column",
          "key": "metric",
          "title": "Metric",
          "width": {"value": 120, "unit": "pt"}
        },
        {
          "id": "value_column",
          "key": "value",
          "title": "Value",
          "width": {"value": 96, "unit": "pt"}
        }
      ],
      "rows": [
        {
          "id": "growth_row",
          "cells": [
            {
              "id": "growth_label",
              "column_key": "metric",
              "value": "Growth"
            },
            {
              "id": "growth_value",
              "column_key": "value",
              "value": "18%"
            }
          ]
        }
      ],
      "layout": {
        "style_ref": "TableGrid",
        "algorithm": "fixed",
        "repeat_header": true
      }
    }
  },
  {
    "op": "table.cell.format",
    "target": "#metrics_table",
    "cell": "#growth_value",
    "set": {
      "vertical_alignment": "center",
      "background_color": "#E2F0D9"
    }
  }
]
```

`node.insert_before` provides symmetric relative placement. `node.append` with
target `$` places the new table in the final semantic section, immediately before an
optional terminal body-level `w:sectPr`. Inserting before the first node of a later
section rebinds that section's `start_at` to the table with explicit change evidence.

## Supported table surface

The inserted table uses the same strict models as semantic DOCX generation:

- ordered columns with stable IDs, semantic keys, data types, titles, and physical
  widths;
- ordered rows with stable IDs, exact or minimum height, and page-break behavior;
- scalar cells or ordered rich paragraphs;
- horizontal `column_span` and vertical `row_span` after rectangular-grid
  validation;
- table preferred width, alignment, fixed/autofit algorithm, indent, cell spacing,
  default cell margins, borders, and repeated header behavior;
- cell vertical alignment, no-wrap, fit-text, fill, direct borders, and independent
  margins;
- paragraph and text formatting inside rich cells;
- external hyperlinks and internal `#bookmark` hyperlinks inside rich cells.

Dynamic fields, drawings, nested tables, and lists inside cells are not accepted by
this insertion contract. They remain future native surfaces rather than being
approximated from display text.

## Native compilation and identity

AiOffice compiles one new `w:tbl` and inserts that element at the proven body
location. Existing `w:p`, `w:tbl`, DrawingML, section properties, relationships,
and unknown XML are not reconstructed.

The native identity map covers:

| Semantic object | Native reference |
| --- | --- |
| table | `w:tbl` |
| column | `w:tblGrid/w:gridCol` |
| row | `w:tr` |
| cell | `w:tc` |
| rich cell paragraph | `w:tc/w:p` |

IDs omitted by the caller are generated once during semantic normalization and then
reused by native lowering. Rich cell paragraphs receive collision-safe
`w14:paraId` values. The identity manifest is refreshed after all operations, so
the generated or caller-selected IDs survive export, Workspace persistence, and
standalone reopen.

## Same-Patch addressability

Native lowering tracks the inserted XML objects directly rather than relying on
indices that earlier operations may have shifted. Later operations in the same
atomic Patch can therefore:

- apply `table.format`, `table.column.format`, or `table.cell.format`;
- apply `text.replace`, `paragraph.format`, `text.format`, or `style.apply` to a
  rich cell paragraph with a known ID;
- use the table as the anchor for another insertion;
- move or remove the complete table through its stable root ID.

`table.column.format` retains its conservative regular-grid rule. It refuses a
merged or shifted table because a semantic column no longer has a provable
one-to-one native width mapping. Table-wide and cell-local formatting remain
available for valid merged tables.

## Styles and hyperlinks

An explicit `layout.style_ref` must identify exactly one native table style. If it
is omitted, AiOffice requires the native `TableGrid` style. A `style_ref` on any rich
cell paragraph must likewise exist in the native paragraph-style catalog. AiOffice
does not silently copy or synthesize a missing style because doing so could change
theme and inheritance behavior.

External links receive fresh collision-safe relationship IDs in
`word/_rels/document.xml.rels`. Internal links are lowered to `w:anchor` and do not
create package relationships. Existing relationships remain unchanged.

## Atomic safety boundary

Before committing the native package, AiOffice proves:

1. the target/anchor location is a complete mapped top-level body range;
2. the semantic table forms a complete, non-overlapping logical grid;
3. the required native table and paragraph styles exist exactly once;
4. no table, column, row, cell, or rich paragraph claims a pre-existing
   `source_ref`;
5. every external or internal hyperlink has unambiguous native evidence;
6. every semantic component resolves to the expected new native element;
7. all generated XML is safe and parseable;
8. section ownership and terminal `w:sectPr` placement remain consistent.

Any failed proof aborts the whole Patch. The source `Document`, native package, and
revision remain unchanged.

## Fidelity and verification

The operation rewrites `word/document.xml`, the identity manifest, and
`word/_rels/document.xml.rels` only when an external hyperlink requires a new
relationship. A first structural edit to a third-party package may also attach the
identity manifest relationship and content type. Other package parts are preserved
byte-for-byte.

The JSON Spec is intent and audit evidence, not a Word layout engine. Use native
LibreOffice rendering and page-image inspection for pagination, typography, table
width, row splitting, and overall visual quality before treating an expert-facing
document as approved.

## CLI and Workspace

The generic Patch transport is sufficient:

```bash
aioffice apply report.docx insert-table.json --output report-updated.docx
aioffice workspace apply ARTIFACT_ID insert-table.json --root project
```

Check `Document.capabilities()["structural_editing"]["insertable_native_blocks"]`
before planning. A detached JSON projection that declares native authority omits all
native structural insertion operations until its original DOCX package is attached.
