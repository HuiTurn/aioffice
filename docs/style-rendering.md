# Style, diff, and rendering contracts

AiOffice keeps design intent strict enough for an agent to reason about while the
native DOCX package remains authoritative for imported documents.

## Explicit style values

`paragraph_style` and `text_style` describe direct formatting. Every length is an
object with a numeric `value` and one of `pt`, `in`, `cm`, `mm`, or `px`:

```json
{
  "type": "paragraph",
  "id": "executive_summary",
  "text": "The program is on schedule.",
  "paragraph_style": {
    "alignment": "justify",
    "spacing_before": {"value": 6, "unit": "pt"},
    "spacing_after": {"value": 8, "unit": "pt"},
    "line_spacing": {"rule": "multiple", "value": 1.2},
    "keep_together": true
  },
  "text_style": {
    "font_family": "Aptos",
    "font_family_east_asia": "Microsoft YaHei",
    "font_size": {"value": 10.5, "unit": "pt"},
    "color": "#222222"
  }
}
```

Unitless lengths, unknown properties, invalid colors, negative paragraph spacing,
and conflicting first-line/hanging indents fail validation. `set` changes selected
properties; `clear` removes direct formatting so the native named style or document
default can take effect again.

## Named styles and inheritance

`defaults` and `styles` are first-class Spec fields. A paragraph or heading uses
`style_ref` to address a reusable definition by stable style ID. The effective
formatting order is deterministic:

```text
theme defaults
  → document defaults
  → named style based_on chain
  → node paragraph_style / text_style
  → TextSpan style
```

A `NamedStyle` declares a human name, semantic role, optional heading level,
`based_on`, `next_style`, paragraph properties, text properties, and UI visibility
hints. Missing references, inheritance cycles, and a heading style applied to an
incompatible semantic node produce stable validation diagnostics.

Three atomic Patch operations are available:

- `style.define` creates a complete style definition;
- `style.apply` applies or clears a node style reference and keeps heading
  role/level consistent;
- `style.format` sets or clears selected paragraph/text properties on a style
  without replacing the rest of its definition.

For generated DOCX, the theme and local catalog compile into `word/styles.xml`.
For imported DOCX, AiOffice projects paragraph styles and `w:docDefaults` while the
native style part remains authoritative. A native `style.format` patch changes only
the selected supported `w:pPr` / `w:rPr` properties. Unknown attributes, extension
elements, table/character styles, and untouched package parts remain preserved.

`text.format` supports three explicit scopes:

- no selector: all text runs and the paragraph mark;
- `match`: the requested one-based exact occurrence;
- `range`: a half-open `[start, end)` interval measured in Unicode code points.

Range operations split only the boundary runs, preserve their existing `w:rPr`,
hyperlink container, attributes, and untouched text, then update the selected
clones. A partial boundary run containing fields, drawings, or another unsupported
inline child causes an atomic `NATIVE_PATCH_FAILED` result rather than a lossy
approximation.

## Native lowering

On a generated document, the same models compile to `w:pPr` and `w:rPr`. On an
imported DOCX, AiOffice projects supported direct properties into the semantic
model. A format patch updates only the requested native properties:

- paragraph alignment, spacing, line spacing, indentation, pagination controls;
- paragraph outline level used by heading/navigation semantics;
- font families, size, foreground/background color;
- bold, italic, underline, strike, small/all caps;
- character spacing and baseline.

Unknown OOXML children and untouched package parts are retained. Mixed run
formatting and hyperlinks are projected as rich `TextSpan` content.
`text_style` on the projected paragraph contains only direct properties common to
every text-bearing run; each span contains the residual direct properties and link
target. Cross-span text replacement inherits the first replaced run's formatting,
matching the native DOCX lowering behavior.

## Semantic diff

Every successful patch returns `PatchResult.diff`. Paths use persistent IDs rather
than array indexes, for example:

```text
content.#executive_summary.paragraph_style.alignment
```

`Document.diff(other)` can also compare arbitrary revisions. Native source
references and revision bookkeeping are excluded by default; pass
`include_native=True` when diagnosing identity changes.

## Rendering evidence

Render results declare:

- provider and provider version;
- output format and content hash;
- cache key inputs;
- `fidelity`: `approximate` or `native`;
- verification status and diagnostics.

The built-in `semantic-html` provider is an inspectable preview only. It helps an AI
review hierarchy, content, and declared styling, but it is not authoritative for
Word pagination, font substitution, line breaking, or floating-object placement.

`compare_raster_images()` provides deterministic PNG regression metrics for current
and future native render providers. It refuses dimension changes instead of resizing
them away, and reports normalized mean error plus changed-pixel ratio. Install the
optional dependency with:

```bash
pip install "aioffice[render]"
```

The next rendering milestone is a native Word/LibreOffice provider that returns page
images plus font-environment metadata. Until that provider exists, AiOffice requires
native visual verification after layout-affecting DOCX patches.
