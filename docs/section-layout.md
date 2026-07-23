# Page and section contract

AiOffice models pages through ordered document sections instead of one global page
setup. This matches Word documents that change paper, margins, columns, or header
behavior partway through the body while keeping the exchange model easy for an AI to
inspect and patch.

## Semantic model

`AiOfficeDocumentSpec.sections` always contains at least one `DocumentSection`.

- The first section has `start_at: null`.
- Every non-empty later section names its first content node in `start_at`.
- Section IDs and content IDs share one selector namespace.
- Sections must be ordered by their content anchors.
- Imported native documents may expose an unanchored later section only when that
  section contains no projected content; validation reports a warning.

`SectionLayout` supports:

- `start_type`: `continuous`, `next_page`, `even_page`, `odd_page`, or
  `next_column`;
- standard paper presets or an exact custom width and height;
- portrait or landscape orientation;
- top, right, bottom, and left margins plus binding gutter;
- header and footer distances;
- equal-width or explicit unequal-width columns and an optional separator;
- top, center, justified, or bottom vertical alignment;
- different-first-page header/footer behavior.

All physical values carry an explicit `pt`, `in`, `cm`, `mm`, or `px` unit. The
validator rejects impossible page geometry, unordered/missing anchors, duplicate
section IDs, malformed unequal columns, and columns wider than the printable area.

## DOCX mapping

For each section except the final one, Word stores `w:sectPr` in the paragraph
properties at that section's end. The final `w:sectPr` is a direct child of
`w:body`. AiOffice-generated files use a dedicated empty boundary paragraph for a
section boundary; the importer recognizes that paragraph as structure and does not
invent an empty semantic paragraph.

Imported sections retain a `NativeRef` to the exact `w:sectPr`. The native package,
not the JSON projection, remains authoritative for unsupported section properties.
A no-op export therefore returns the original DOCX bytes.

`section.format` accepts only `set` and `clear` fields from `SectionLayout`. Native
lowering changes the corresponding known OOXML values in one mapped `w:sectPr`.
Unknown attributes, unknown children, relationship references, and all untouched
package parts are preserved. Patch validation and lowering are atomic.

Header/footer bindings are intentionally separate from `SectionLayout`: they point
to reusable package parts and have inheritance semantics. See
[the header/footer contract](header-footer.md).

## Preview boundary

Semantic HTML uses page dimensions, margins, section breaks, and CSS multi-column
hints to make page intent inspectable. It remains an approximate preview:
font metrics, pagination, fields, floating objects, headers/footers, and Word's
layout engine can change the native result. Only a future native render provider and
visual comparison may be treated as pagination evidence.
