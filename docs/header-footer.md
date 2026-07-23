# Header and footer contract

AiOffice separates reusable header/footer content from section bindings. This keeps
the model faithful to DOCX relationships and prevents an AI from accidentally
duplicating or unlinking content while changing page setup.

## Normalized model

`AiOfficeDocumentSpec.header_footers` contains unique `HeaderFooterPart` objects.
Each part is either a header or footer and has stable, selectable block IDs.

`DocumentSection.header_footer` may bind six slots:

- `header_default`, `header_first`, and `header_even`;
- `footer_default`, `footer_first`, and `footer_even`.

A missing slot is not blank content. It means the same slot is inherited from the
previous section; in the first section it resolves to Word's blank/default behavior.
An explicitly blank region is represented by an explicit part with no semantic
blocks. Multiple sections may intentionally reference the same part.

First-page bindings become active when that section's
`different_first_page` is true. Even-page bindings become active when the
document-wide `settings.even_and_odd_headers` switch is true. Validation warns when
a binding exists but its switch is inactive.

## Native mapping

Each explicit section binding becomes a `w:headerReference` or
`w:footerReference` with `w:type` equal to `default`, `first`, or `even`. The
relationship resolves from `word/document.xml` to a header/footer part. Generated
parts receive the correct OPC relationship and content type, and document-wide
even/odd behavior is written to `word/settings.xml`.

Imported parts and their ordinary paragraphs receive persistent native identities
scoped by part URI. This is important because paragraph IDs and identical XML can
legitimately occur in different parts.

`text.replace`, `paragraph.format`, and `text.format` may target an ordinary
header/footer paragraph ID. Native lowering mutates only that part and the AiOffice
identity manifest. Unknown XML in the same part remains in place.

## Conservative projection boundary

Fields such as PAGE and NUMPAGES, drawings, embedded objects, tables, and unknown
header/footer elements are projected as non-editable opaque blocks. Their native XML
remains authoritative. AiOffice does not turn a field's cached display result into
plain editable text, because doing so would destroy its computation semantics.

Semantic generation currently supports ordinary paragraph blocks, including rich
text and hyperlinks. Structured fields, tables, images, and binding creation/removal
on an already imported native package are planned behind explicit capabilities.

## Preview boundary

Semantic HTML shows the effective inherited default or first-page region for each
section. It is an inspectable approximation, not proof of Word pagination, field
evaluation, or physical header/footer placement.
