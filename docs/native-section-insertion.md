# Incremental native section insertion

AiOffice `0.2.0.dev28` can split an existing DOCX section before a stable top-level
content node. The JSON Spec expresses the new section's identity, anchor, and layout
intent; the attached `/word/document.xml` remains authoritative for unsupported
section properties and exact native placement.

## Operation

```json
{
  "op": "section.insert_before",
  "target": "#wide_appendix",
  "section": {
    "id": "wide_appendix_section",
    "metadata": {
      "purpose": "wide appendix"
    },
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

The target must resolve to one existing top-level semantic content block. Its
complete native range may be one paragraph, a multi-paragraph list, a table, or
another mapped body block. The target becomes the new section's `start_at`.

The optional section payload accepts `id`, `type: "section"`, `metadata`, and
`layout`. When the ID is omitted, the semantic transaction generates one and
returns it in `changes[].created_sections`.

## Inheritance and defaults

The containing section is split into two non-empty semantic regions:

- the existing section keeps all content before the target;
- the new section owns the target and following content up to the old boundary.

The new section begins with a copy of the containing section's complete known
`SectionLayout`, then applies caller-supplied fields. An omitted `start_type`
becomes `next_page`. Other fields can explicitly use null to remove an inherited
known property; `start_type: null` is rejected because Word interprets a missing
type on a later section as `next_page`.

Header/footer IDs are inherited from the containing semantic section. Direct
rebinding is excluded from this operation because it requires a separate proof over
document relationships, reusable header/footer parts, first/even-page settings, and
inheritance across neighboring sections.

## Native lowering

Word stores the properties for a non-final section at that section's end, inside
`w:p/w:pPr/w:sectPr`. The final section uses a body-level `w:sectPr`. To split a
section without mutating the previous visible block, AiOffice:

1. resolves every XML element in the target's stable native range;
2. proves the range is present, top-level, ordered, and contiguous;
3. resolves the containing section's exact ending `w:sectPr`;
4. proves the previous and ending section boundaries surround the target;
5. creates one empty, hidden `w:p` immediately before the target with a
   collision-safe `w14:paraId`;
6. deep-copies the old `w:sectPr` into that paragraph for the preceding section;
7. keeps the old boundary object in place as the new section's ending boundary;
8. patches only the default/selected layout fields on that existing boundary;
9. remaps both section IDs and every shifted native identity;
10. serializes the changed document part and identity manifest atomically.

The visible target and all existing content elements remain the same XML objects.
If the target is a list, its complete `w:p` group remains contiguous after the new
carrier. If the old boundary is paragraph-carried, it stays paragraph-carried for
the new section; if it is the terminal body boundary, it stays terminal.

## Unknown native properties

The copied boundary retains unknown attributes and children, header/footer
references, footnote/endnote settings, page borders, line numbering, document-grid
settings, printer settings, and other native data that the semantic projection does
not claim to model. The new section's original boundary retains the same payload,
except for explicitly patched supported layout fields.

AiOffice deliberately refuses `w:sectPrChange`. Duplicating tracked section
property history could misrepresent Word revision semantics, so it remains
read-only until revision markup has its own native model and validation contract.

## Same-Patch behavior

Native lowering tracks live section boundary objects and ordered section indices.
After insertion, the new section can immediately receive:

- `section.format`;
- `node.insert_before` at its first node, with `start_at` rebound to the created
  predecessor;
- `node.insert_after` and other ordinary edits that remain inside proven section
  boundaries.

Multiple section insertions can run in one Patch. Each later split uses the current
semantic section order and live native boundaries rather than stale element
indices.

## Fail-closed boundary

The operation fails atomically when:

- the target already begins its containing section and would create an empty
  section;
- the target is missing, detached, non-top-level, stale, or noncontiguous;
- the containing or preceding native boundary cannot be proven;
- the requested section has unknown fields or invalid page/column geometry;
- `start_type` is null;
- the payload attempts direct header/footer rebinding or supplies a forged native
  source reference;
- the source boundary contains tracked `w:sectPrChange`;
- semantic change evidence, section order, identity refresh, or reprojection does
  not match the native result.

The source `Document` and source package bytes remain unchanged on every failure.
A detached JSON projection that declares native authority omits and rejects the
operation.

## CLI, Workspace, and verification

The operation uses the ordinary JSON Patch channel:

```bash
aioffice apply report.docx split-section.json --output report-updated.docx
aioffice workspace apply ARTIFACT_ID split-section.json --root project
```

Inspect `Document.capabilities()["formatting"]["section_contract"]` before
planning. It reports the target, inheritance, default start type, native patch
scope, supported same-Patch operations, and explicit unsupported cases.

JSON is the model-facing intent and evidence protocol, not a lossless replacement
for `w:sectPr`. Native standalone reopen proves structural identity; rendered PDF
and page PNGs remain the authority for orientation, paper geometry, pagination,
columns, headers, footers, and visual balance.
