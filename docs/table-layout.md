# Table layout contract

AiOffice separates an AI-friendly semantic table model from the native DOCX package.
The semantic model is compact enough for planning and tool calls; the original OOXML
remains the source of truth for lossless persistence.

## Semantic model

A table contains:

- stable table, column, and data-row IDs;
- semantic column keys, titles, and data-type hints;
- explicit column widths;
- preferred table width (`auto`, percent, or exact);
- alignment, layout algorithm, indent, and cell spacing;
- independent top/right/bottom/left cell margins;
- repeated-header behavior;
- per-row page-break permission, height, and height rule.

All physical lengths use explicit `pt`, `in`, `cm`, `mm`, `px`, or `emu` units.
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
| repeat header | `w:tblHeader` on the first row |
| prevent row splitting | `w:cantSplit` |
| row height/rule | `w:trHeight` |
| column width | `w:tblGrid/w:gridCol` and regular-grid `w:tcW` |

These structures follow Microsoft's
[WordprocessingML table model](https://learn.microsoft.com/en-us/office/open-xml/word/working-with-wordprocessingml-tables),
including the distinction between
[table width](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.tablewidth?view=openxml-3.0.1),
[grid spans](https://learn.microsoft.com/lb-lu/dotnet/api/documentformat.openxml.wordprocessing.gridspan?view=openxml-2.8.0),
and [vertical merges](https://learn.microsoft.com/es-es/dotnet/api/documentformat.openxml.wordprocessing.verticalmerge?view=openxml-3.0.1).

## Selective native Patch

`table.format` sets or clears selected table-wide properties. It preserves unrelated
attributes, children, cell content, styles, and package parts.

`table.column.format` addresses a column by stable ID or semantic key and changes
only its width. Before lowering, AiOffice proves that:

1. a `w:tblGrid` exists;
2. every row has exactly one physical cell per grid column;
3. no row uses `gridBefore`, `gridAfter`, `gridSpan`, `hMerge`, or `vMerge`.

If the proof fails, the entire Patch returns `NATIVE_PATCH_FAILED`; no partial
revision is emitted. This is deliberately conservative because a logical column in
an irregular Word table may correspond to different physical cells in different
rows.

## Projection boundary

The first physical row is currently treated as the semantic header. Remaining rows
become data rows. Cell content is projected as plain display text, so rich cell
paragraphs, nested tables, drawings, content controls, and cell-level native
formatting are not yet editable through the semantic table API. They remain intact
in the native package during table-wide and regular-grid column formatting.

Tables in header/footer parts remain opaque. Irregular body tables expose their
table-wide layout and `regular_grid=false`, but do not expose a writable column
geometry contract.

## Preview and validation

Semantic HTML renders widths, alignment, layout algorithm, cell spacing/margins,
row height, and page-break hints for planning. It is approximate and is not evidence
of Word pagination.

Validation warns when explicit column widths or an exact preferred width exceed the
active section's printable width. Fixed-layout tables also warn when a column width
is missing. Final pagination, repeated headers, and exact line wrapping still
require rendering by a native office engine.
