# Table layout contract

AiOffice separates an AI-friendly semantic table model from the native DOCX package.
The semantic model is compact enough for planning and tool calls; the original OOXML
remains the source of truth for lossless persistence.

## Semantic model

A table contains:

- stable table, column, data-row, logical-cell, and rich cell-paragraph IDs;
- semantic column keys, titles, and data-type hints;
- explicit column widths;
- preferred table width (`auto`, percent, or exact);
- alignment, layout algorithm, indent, and cell spacing;
- independent top/right/bottom/left cell margins;
- explicit perimeter and internal horizontal/vertical borders;
- repeated-header behavior;
- per-row page-break permission, height, and height rule;
- logical row/column spans and cell-local presentation.

All physical lengths use explicit `pt`, `in`, `cm`, `mm`, or `px` units.
Percent widths use ordinary percentages in the Spec rather than Word's internal
fiftieths-of-a-percent representation.

## DOCX mapping

The native adapter projects and generates the corresponding WordprocessingML
properties:

| Semantic property | WordprocessingML |
| --- | --- |
| preferred width | `w:tblW` |
| alignment | `w:jc` |
| layout algorithm | `w:tblLayout` |
| indent | `w:tblInd` |
| cell spacing | `w:tblCellSpacing` |
| cell margins | `w:tblCellMar` |
| perimeter and internal borders | `w:tblBorders` |
| repeat header | `w:tblHeader` on the first row |
| prevent row splitting | `w:cantSplit` |
| row height/rule | `w:trHeight` |
| column width | `w:tblGrid/w:gridCol` and regular-grid `w:tcW` |

These structures follow Microsoft's
[WordprocessingML table model](https://learn.microsoft.com/en-us/office/open-xml/word/working-with-wordprocessingml-tables),
including the distinction between
[table width](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.tablewidth?view=openxml-3.0.1),
[table properties](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.tableproperties?view=openxml-3.0.1),
[table borders](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.tableborders?view=openxml-2.20.0),
[grid spans](https://learn.microsoft.com/lb-lu/dotnet/api/documentformat.openxml.wordprocessing.gridspan?view=openxml-2.8.0),
and [vertical merges](https://learn.microsoft.com/es-es/dotnet/api/documentformat.openxml.wordprocessing.verticalmerge?view=openxml-3.0.1).

## Border contract

`TableBorders` addresses `top`, `right`, `bottom`, `left`,
`inside_horizontal`, and `inside_vertical`. Each edge is either absent or a strict
`BorderLine`:

- `style`: `none`, `single`, `double`, `dotted`, `dashed`, or `thick`;
- `width`: required for visible styles and bounded from 0.25pt through 12pt;
- `color`: `#RRGGBB` or `auto`;
- `space`: optional and bounded from 0pt through 31pt.

An absent semantic edge means AiOffice did not project a supported direct edge; it
does not claim that Word will display no border because a table style can still
contribute one. Setting `{"style": "none"}` writes `w:val="none"` and explicitly
suppresses that edge. Clearing the whole `borders` property removes all known direct
table-edge attributes, allowing style inheritance to resume. A supplied `borders`
object replaces the supported direct edge set, so omitted sides are cleared.

For compatibility, generated DOCX uses physical `left`/`right` edge names. Projection
also understands the logical `start`/`end` aliases. Unsupported native border styles,
theme-only colors, and unknown extension data stay authoritative in OOXML and are not
invented in the semantic projection.

## Selective native Patch

`table.format` sets or clears selected table-wide properties, including the complete
supported direct border set. It preserves unrelated attributes, children, cell
content, styles, and package parts. Updating a known border edge removes conflicting
known theme attributes but preserves unknown attributes and child elements on that
edge.

`table.column.format` addresses a column by stable ID or semantic key and changes
only its width. Before lowering, AiOffice proves that:

1. a `w:tblGrid` exists;
2. every row has exactly one physical cell per grid column;
3. no row uses `gridBefore`, `gridAfter`, `gridSpan`, `hMerge`, or `vMerge`.

If the proof fails, the entire Patch returns `NATIVE_PATCH_FAILED`; no partial
revision is emitted. This is deliberately conservative because a logical column in
an irregular Word table may correspond to different physical cells in different
rows.

`table.cell.format` addresses one stable anchor cell and changes only selected
`w:tcPr` properties. It does not depend on a one-to-one column mapping, so it remains
safe for a mapped merged cell. See [the table cell contract](table-cells.md).

## Projection boundary

The first physical row is currently treated as the semantic header. Remaining rows
become data rows. Proven `gridSpan`/`vMerge` structures become logical cells with
column and row spans. Supported paragraph sequences become editable rich content.
Nested tables, drawings, objects, fields, and malformed content fall back to
read-only display text; they remain intact during table, column, and cell formatting.

Tables in header/footer parts remain opaque. Irregular body tables expose their
table-wide layout, `logical_grid=false`, grid diagnostics, and a heuristic inspection
view, but do not expose a writable column geometry contract.

## Preview and validation

Semantic HTML renders widths, spans, alignment, layout algorithm, table/cell spacing
and margins, perimeter/internal/direct cell borders, fill, rich paragraph content,
row height, and page-break hints for planning. Internal borders stop at the logical
grid edge and direct cell borders take precedence in the preview. It is approximate
and is not evidence of Word pagination.

Validation warns when explicit column widths or an exact preferred width exceed the
active section's printable width. Fixed-layout tables also warn when a column width
is missing. Final pagination, repeated headers, and exact line wrapping still
require the `libreoffice` native provider or another declared office engine.
