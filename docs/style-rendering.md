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

`text.format` currently applies to all native runs mapped to one heading or
paragraph. Range-level run splitting is a later native capability and is not
silently approximated.

## Native lowering

On a generated document, the same models compile to `w:pPr` and `w:rPr`. On an
imported DOCX, AiOffice projects supported direct properties into the semantic
model. A format patch updates only the requested native properties:

- paragraph alignment, spacing, line spacing, indentation, pagination controls;
- font families, size, foreground/background color;
- bold, italic, underline, strike, small/all caps;
- character spacing and baseline.

Unknown OOXML children and untouched package parts are retained. Mixed run
formatting remains in the native package; `text_style` on the projected paragraph
contains only direct properties common to every text-bearing run.

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
