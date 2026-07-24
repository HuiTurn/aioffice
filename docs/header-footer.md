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

Safe internal header/footer relationships are projected even when no current section
references the part. This keeps a region reusable after a binding change instead of
hiding preserved but temporarily unbound native content from the AI.

`text.replace`, `paragraph.format`, and `text.format` may target an ordinary
header/footer paragraph ID. Native lowering mutates only that part and the AiOffice
identity manifest. Unknown XML in the same part remains in place.

Paragraph background, border, paragraph, and text styles use the same inheritance
and Patch contract as body paragraphs. Semantic HTML resolves those styles for
header/footer previews; native PDF/PNG remains the pagination and visual authority.

## Section binding edits

`section.header_footer.bind` changes which existing reusable parts a section
explicitly references:

```json
{
  "op": "section.header_footer.bind",
  "target": "#appendix_section",
  "set": {
    "header_default": "appendix_header",
    "footer_default": "report_footer"
  },
  "clear": ["header_first"]
}
```

`set` maps one or more of the six binding slots to stable header/footer part IDs.
The selected part must exist in `header_footers`, have the correct kind, and
for imported DOCX resolve through exactly one internal relationship from
`word/document.xml`. A leading `#` on a part ID is accepted and normalized away.

`clear` removes that section's explicit native reference. It does not create a blank
region: Word inherits the same slot from the preceding section, or uses blank/default
behavior for the first section. Bind an explicit empty part when the intended result
is an intentionally blank region.

The operation can target a section created earlier in the same Patch. Only selected
direct `w:headerReference` or `w:footerReference` children change; the reusable part
XML, its relationships, unselected section references, and unknown section
properties remain untouched. A standalone identity manifest is attached when needed
so section and part IDs remain stable after reopen.

The native transaction refuses a missing or type-incompatible part, external or
duplicate relationships, duplicate references for one slot, stale section
boundaries, and clearing an invalid native reference that was not safely projected.
Creating or cloning a reusable part is a separate explicit operation. Deleting
parts remains a future operation. See
[the full native binding contract](native-header-footer-binding.md).

## Reusable part creation

`header_footer.create` adds one independent reusable header or footer:

```json
{
  "op": "header_footer.create",
  "part": {
    "id": "appendix_header",
    "kind": "header",
    "content": [
      {
        "id": "appendix_header_line",
        "type": "paragraph",
        "text": "Appendix"
      }
    ]
  }
}
```

It may be followed by `section.header_footer.bind` in the same atomic Patch.
AiOffice allocates the native part URI, document relationship, content-type
override, optional hyperlink relationships, and native identities. The caller
cannot provide those package-owned values. Existing parts and unrelated package
content remain untouched. See
[the full native creation contract](native-header-footer-creation.md).

## Reusable part cloning

`header_footer.clone` forks an existing reusable header/footer as an independent
part while retaining native content that the semantic projection cannot recreate:

```json
{
  "op": "header_footer.clone",
  "target": "#report_header",
  "part": {
    "id": "appendix_header",
    "metadata": {
      "role": "appendix"
    }
  }
}
```

AiOffice generates deterministic new semantic IDs, copies the native story and its
part-local relationship XML, shares relationship targets such as media, and rebases
native paragraph and DrawingML IDs. The clone may be bound in the same Patch. Edit
its content in a later Patch so creation-time graph evidence remains independently
verifiable. See
[the full native cloning contract](native-header-footer-cloning.md).

## Conservative projection boundary

PAGE, NUMPAGES, SECTION, and SECTIONPAGES fields are projected as structured inline
objects. Their cached display result is kept separate from the semantic instruction
and never treated as authoritative text. Unknown field instructions are visible as
read-only native fields. A paragraph containing exactly one conservative embedded
inline picture or offset-positioned, square-wrapped floating picture is projected as
a stable `ImageBlock`: its bytes can be verified or extracted, and its accessibility
metadata, extent, bounded rectangular source crop, paragraph layout, or binary can
be changed through the same native image APIs as a body picture. Floating anchor
layout is explicit native-authoritative evidence; for the conservative projected
subset, `image.anchor.update` can selectively change its offset positions,
square-wrap group, relative height, and boolean flags.

Image replacement is occurrence-scoped copy-on-write and allocates the new
relationship in the containing header/footer part. This makes the expert workflow
safe: clone and bind a shared header in one Patch, then replace only the cloned logo
in a subsequent transaction. The source story, source relationship, original media,
and unrelated sections remain unchanged.

The same story-local rule applies to `image.update`: changing or clearing a cloned
logo's `crop` mutates only that cloned header/footer XML. The source story,
relationship, and media payload remain byte-exact.

`image.anchor.update` is likewise story-local. On a projected floating logo it
changes only the selected anchor fields in that header/footer part; the main
document, part-local relationships, media payload, crop, extent, and source story
remain unchanged.

Malformed field containment, complex drawings, embedded objects, tables, and unknown
header/footer elements remain non-editable opaque blocks.

Semantic creation currently supports ordinary paragraph blocks, including rich
text, hyperlinks, and normalized dynamic fields. Native cloning can preserve
supported drawings and images; supported inline and conservative floating images
become addressable native occurrences after import. Direct image
insertion/deletion within a reusable story, tables, deleting parts, and direct
editing of opaque region content remain planned behind explicit capabilities.
See [the dynamic field contract](dynamic-fields.md).

## Preview boundary

Semantic HTML shows the effective inherited default or first-page region for each
section. It is an inspectable approximation, not proof of Word pagination, field
evaluation, or physical header/footer placement.
