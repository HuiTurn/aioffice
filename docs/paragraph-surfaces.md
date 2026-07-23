# Paragraph surface contract

Paragraph surfaces provide a native, reusable foundation for title bars, executive
callouts, warnings, quotations, and emphasized regions without pretending that a
floating text box is an ordinary paragraph.

## Semantic model

`ParagraphStyle` adds:

- `background_color`: one solid `#RRGGBB` paragraph background;
- `borders.top`, `right`, `bottom`, and `left`: optional `BorderLine` values.

Each visible border has an explicit style, width, color, and optional gap between the
text and edge:

```json
{
  "background_color": "#EAF2F8",
  "borders": {
    "left": {
      "style": "single",
      "width": {"value": 3, "unit": "pt"},
      "color": "#1F4E78",
      "space": {"value": 8, "unit": "pt"}
    },
    "bottom": {
      "style": "double",
      "width": {"value": 1.5, "unit": "pt"},
      "color": "#5B9BD5"
    }
  }
}
```

Supported border styles are `none`, `single`, `double`, `dotted`, `dashed`, and
`thick`. Visible widths are bounded from 0.25pt through 12pt; optional space is
bounded from 0pt through 31pt. Colors are normalized to uppercase sRGB.

The model is available on body paragraphs, headings, editable rich table-cell
paragraphs, headers/footers, document defaults, and named paragraph styles.

## Style inheritance

Effective formatting follows the normal AiOffice chain:

```text
theme defaults
  → document defaults
  → named style based_on chain
  → direct node paragraph_style
```

Border edges merge independently. If a named style supplies four edges and a direct
paragraph supplies only `bottom`, the effective result keeps the inherited top,
right, and left edges and replaces the bottom edge.

There are two distinct removal intents:

- clearing the `borders` field removes supported direct border XML so the previous
  style level can apply;
- setting an edge to `{"style": "none"}` writes an explicit no-border edge and
  suppresses the inherited edge.

Setting a `borders` object replaces the supported direct edge set at that level, so
omitted direct sides are cleared. Clearing `background_color` similarly removes
direct paragraph shading and resumes style inheritance.

## WordprocessingML mapping

The native mapping is:

| Semantic property | WordprocessingML |
| --- | --- |
| solid background | `w:pPr/w:shd` with clear pattern and sRGB `w:fill` |
| four border edges | `w:pPr/w:pBdr/w:top|right|bottom|left` |
| border width | `w:sz` in eighths of a point |
| text-to-border gap | `w:space` in points |

Microsoft documents paragraph properties, including border and shading, in
[Working with paragraphs](https://learn.microsoft.com/en-us/office/open-xml/word/working-with-paragraphs).
The Open XML SDK contracts identify
[`w:pBdr`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.paragraphborders?view=openxml-3.0.1)
and
[`w:shd`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.shading?view=openxml-3.0.1)
as paragraph properties.

Generated DOCX writes canonical physical edge names. Projection accepts an edge only
when its style, size, color, and spacing are all representable by the strict model.
An imported document with no edits is always exported byte-for-byte unchanged.

## Selective native Patch

`paragraph.format` mutates one mapped `w:pPr`. `style.format` mutates one paragraph
style in `word/styles.xml`. Both preserve paragraph content, unrelated properties,
unknown attributes and child elements, and every unaffected package part.

Updating a supported edge removes conflicting known theme attributes on that edge,
while extension attributes and children remain. Clearing a surface removes all known
attributes but retains unknown extension data so future producers can recover it.

## Deliberate preservation boundary

AiOffice does not currently project:

- paragraph `between` or `bar` borders;
- theme-based border colors and tint/shade transforms;
- patterned shading;
- theme-based shading, tint, or shade transforms.

Word combines borders across adjacent paragraphs with identical border settings, and
theme/pattern shading depends on more context than one resolved sRGB value. Exposing
those features as ordinary four-side borders or a guessed color would be lossy.
Their OOXML therefore remains native-only and untouched by unrelated edits.

This follows the Open XML shading model, whose pattern values include clear, solid,
stripes, percentages, and crosses; see Microsoft's
[ShadingPatternValues](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.shadingpatternvalues?view=openxml-3.0.1).

## Preview and verification

Semantic HTML maps solid background, four edge styles, and border spacing to CSS.
It resolves named styles for body, table-cell, and header/footer paragraphs. The
preview is useful for AI planning, but browser box layout does not prove Word's
adjacent-paragraph border grouping, pagination, font substitution, or line wrapping.
Use the LibreOffice PDF/PNG provider and inspect the returned page images before
approving layout or aesthetics.
