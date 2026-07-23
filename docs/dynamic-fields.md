# Dynamic field contract

Dynamic fields are computed document content such as the current page number or the
total page count. AiOffice models them as inline objects instead of flattening their
last rendered value into text.

## Supported field kinds

| Spec kind | DOCX instruction | Typical use |
| --- | --- | --- |
| `page_number` | `PAGE` | Current page number |
| `page_count` | `NUMPAGES` | Total document page count |
| `section_number` | `SECTION` | Current section number |
| `section_page_count` | `SECTIONPAGES` | Page count for the current section |

Each `DocumentField` has a stable ID, a semantic kind, an optional number format,
and an optional `cached_result`. The cached result is only the value last stored by
the native editor. It is useful for previews, but it is never authoritative because
pagination depends on fonts, layout, printer metrics, and the rendering engine.

Supported number formats are decimal, upper/lower Roman, and upper/lower alphabetic.
Section page-number restart and default numbering format live on
`SectionLayout.page_number_start` and `SectionLayout.page_number_format`.

## Example

```python
from aioffice import DocumentField, Paragraph

footer = Paragraph(
    content=[
        {"text": "Page "},
        DocumentField(kind="page_number", number_format="decimal"),
        {"text": " of "},
        DocumentField(kind="page_count", number_format="decimal"),
    ]
)
```

AiOffice-generated DOCX files emit conventional complex fields with separate
instruction and cached-result runs. The document settings request that a native
editor updates dirty fields when the file is opened. Setting
`document.settings.update_fields_on_open` to `False` is allowed, but validation
warns when dynamic fields are present.

## Native import and editing

AiOffice recognizes both `w:fldSimple` and well-bounded complex DOCX fields. PAGE,
NUMPAGES, SECTION, and SECTIONPAGES become editable semantic fields. An isolated
unknown field becomes `kind="native"` and retains its instruction, source reference,
and cached display value, but is read-only.

Malformed or ambiguous complex-field containment is projected as an opaque paragraph.
This is intentional: guessing field boundaries could duplicate or delete native XML.

A selective update uses the field ID:

```python
result = document.apply(
    [{
        "op": "field.update",
        "target": "#page-field",
        "set": {"number_format": "upper_roman"},
    }]
)
```

For an imported DOCX, native lowering changes only that field's instruction payload
and dirty flag. It does not overwrite the cached result, rebuild the paragraph, or
touch another package part. Field identities are part-scoped, so the same mechanism
works in the body, headers, and footers.

Text replacement and character-range formatting refuse paragraphs containing
fields because flattening computed content into plain text is destructive. Whole
paragraph formatting and whole-inline-content text formatting remain available.

## Preview and verification

Semantic HTML and Markdown use the cached result when one exists and otherwise show
a field placeholder. These are approximate previews, not pagination evidence.
Expert layout verification should use the `libreoffice` provider to render the DOCX
and inspect the resulting pages. AiOffice therefore reports field presence and the
native update policy separately from the preview text. LibreOffice evidence remains
provider-specific; fields marked dirty may still require the target office
application to refresh cached values.
