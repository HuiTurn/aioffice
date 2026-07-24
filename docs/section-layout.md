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
- different-first-page header/footer behavior;
- page-number restart and decimal, Roman, or alphabetic numbering formats.

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

## Inserting a section boundary

`section.insert_before` splits the section containing an existing top-level content
node:

```json
{
  "op": "section.insert_before",
  "target": "#wide_appendix",
  "section": {
    "id": "wide_appendix_section",
    "layout": {
      "page_size": {
        "preset": "a4",
        "orientation": "landscape"
      },
      "margin_left": {"value": 18, "unit": "mm"},
      "margin_right": {"value": 18, "unit": "mm"}
    }
  }
}
```

The target becomes the new section's `start_at`. The new section inherits all other
known layout properties and the containing section's semantic header/footer
bindings. When `start_type` is omitted, it becomes `next_page`. A null
`start_type` is rejected because a missing native type on a later Word section also
means `next_page`; accepting null would make the JSON result disagree after reopen.

For imported DOCX, AiOffice:

1. proves the target's complete native range is top-level and contiguous;
2. proves there is semantic content before it in the same section;
3. copies the containing section's exact `w:sectPr` into one new hidden boundary
   paragraph before the target;
4. keeps the original `w:sectPr` as the new section's ending boundary;
5. patches only `start_type` and caller-selected layout fields on that original
   boundary;
6. remaps both section identities and every shifted native reference.

Existing content elements are never reconstructed. A multi-paragraph list is a
valid target and remains intact after the new boundary. Unknown section attributes,
children, header/footer references, page-border settings, and other supported or
opaque native properties are copied for the preceding section and retained on the
following section.

The operation refuses a target that already starts its containing section because
that would create an empty section. It also refuses tracked `w:sectPrChange`
properties, direct header/footer rebinding in the insertion payload, detached native
projections, stale or noncontiguous ranges, and any result that cannot reproduce the
semantic section order. The transaction is atomic.

The new section can receive `section.format`, or have content inserted before its
start with ordered `start_at` rebinding, later in the same Patch. See
[the dedicated native insertion contract](native-section-insertion.md).

`page_number_start` and `page_number_format` map to selected attributes of
`w:pgNumType`. They define section pagination; PAGE/NUMPAGES fields remain separate
inline content. See [the dynamic field contract](dynamic-fields.md).

Header/footer bindings are intentionally separate from `SectionLayout`: they point
to reusable package parts and have inheritance semantics. See
[the header/footer contract](header-footer.md).

## Preview boundary

Semantic HTML uses page dimensions, margins, section breaks, and CSS multi-column
hints to make page intent inspectable. It remains an approximate preview:
font metrics, pagination, fields, floating objects, headers/footers, and Word's
layout engine can change the native result. Use the `libreoffice` provider for
provider-specific PDF/PNG pagination evidence and page-level visual comparison.
