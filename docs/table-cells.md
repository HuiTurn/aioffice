# Table cell contract

AiOffice models the logical table seen by an editor, while retaining the physical
WordprocessingML cell sequence as native authority. This distinction matters because
a merged Word table does not contain a rectangular matrix of independent `w:tc`
elements.

## Canonical semantic form

Every data row contains `cells`. A cell has:

- a stable ID and semantic `column_key` anchor;
- `column_span` and `row_span`;
- either one scalar `value` or ordered rich `Paragraph` content;
- cell-local formatting;
- an optional native reference and fidelity metadata.

The older row `values` object remains accepted as an input shorthand. Validation
normalizes each key/value pair to a one-column scalar cell, and serialized Spec output
uses the canonical cell form.

Rich content uses the same paragraph, text-span, hyperlink, direct formatting, and
named-style models as body content. Dynamic fields inside semantic table-cell
paragraphs are not yet accepted.

## Logical grid proof

For imported DOCX, AiOffice reads `w:tblGrid` and walks physical cells in row order.
It recognizes modern horizontal spans through
[`w:gridSpan`](https://learn.microsoft.com/lb-lu/dotnet/api/documentformat.openxml.wordprocessing.gridspan?view=openxml-2.8.0)
and vertical restart/continuation chains through
[`w:vMerge`](https://learn.microsoft.com/es-es/dotnet/api/documentformat.openxml.wordprocessing.verticalmerge?view=openxml-3.0.1).

The projection is marked `logical_grid=true` only when every data row covers the
declared grid exactly and every vertical continuation has a matching anchor. Legacy
`w:hMerge`, shifted rows (`gridBefore`/`gridAfter`), invalid spans, orphan
continuations, and out-of-grid cells fail the proof.

When proof succeeds, only anchor cells appear in the semantic row; continuation
cells are represented by the anchor's `row_span`. When proof fails, AiOffice emits a
one-cell-per-column heuristic view for inspection and keeps `grid_diagnostics`.
Column-width mutation remains unavailable because logical ownership is uncertain.

## Rich-content boundary

A cell is projected as editable rich paragraphs when its direct block content is a
supported paragraph sequence and contains no dynamic fields, drawing, embedded
object, or nested table. Each paragraph receives a stable native reference and can
use:

- `text.replace`;
- `text.format`;
- `paragraph.format`;
- `style.apply`.

Simple generated scalar cells deliberately have no paragraph identity; they reopen as
scalar values. Unsupported or complex cells expose
`content_projection="plain_text_read_only"` and `content_editable=false`. The display
text is useful for planning but is not an editable reconstruction of the cell.

## Cell formatting

`TableCellFormat` supports:

| Semantic property | WordprocessingML |
| --- | --- |
| vertical alignment | `w:vAlign` |
| no wrap | `w:noWrap` |
| fit text | `w:tcFitText` |
| background color | `w:shd/@w:fill` |
| top/right/bottom/left border | `w:tcBorders` |
| top/right/bottom/left margin | `w:tcMar` |

`table.cell.format` addresses a cell by stable ID. It sets or clears only selected
known properties in that cell's `w:tcPr`; unknown attributes, unknown children,
merged-grid markers, widths, paragraphs, drawings, and all unrelated package parts
remain in place.

Direct cell edges use the same `BorderLine` style, width, color, and spacing
constraints as table borders. Under WordprocessingML's border conflict rules, a
direct cell edge overrides the conflicting table perimeter or internal-grid edge;
see Microsoft's
[TableCellBorders contract](https://learn.microsoft.com/es-es/dotnet/api/documentformat.openxml.wordprocessing.tablecellborders?view=openxml-3.0.1).
Setting an edge to `{"style": "none"}` explicitly disables it. Clearing `borders`
removes the supported direct edge attributes and allows the table or named style to
apply again. Supplying a new `borders` object replaces the supported direct edge set,
so omitted sides are cleared rather than silently retained.

This operation is different from `table.column.format`: cell formatting remains safe
for a mapped merged anchor, while a merged grid cannot safely expose a one-to-one
column-width operation.

## Generation

Semantic horizontal spans generate one anchor `w:tc` with `w:gridSpan`. A vertical
span generates `w:vMerge restart` on the anchor and matching continuation cells with
`w:vMerge continue` in subsequent physical rows. Continuations repeat the horizontal
span and contain the required empty paragraph.

Validation rejects overlapping cells, cells extending beyond the final column or
row, duplicate identities, and uncovered logical grid positions before generation.

## Preview limitations

Semantic HTML uses `colspan`, `rowspan`, stable cell/paragraph IDs, cell padding,
vertical alignment, no-wrap, fill color, and edge-aware border overrides. Markdown
necessarily flattens merged cells and rich paragraphs to display text. Neither
preview proves native pagination, line wrapping, font substitution, border conflict
resolution, or fit-text behavior; use the `libreoffice` provider to create
inspectable native-compatible PDF/PNG evidence.
