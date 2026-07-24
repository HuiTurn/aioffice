# Lossless native header/footer cloning

AiOffice `0.2.0.dev31` can fork a supported imported header or footer into an
independent reusable part. This is the first copy-on-write primitive for shared
Word regions: JSON identifies the source and the new semantic identity, while the
attached DOCX package remains authoritative for the complete story XML,
relationships, unknown formatting, drawings, and media.

## Operation

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

`target` resolves one existing reusable part. `part.id` is a new stable ID. The
optional `part.type` may only be `header_footer`, and `metadata` may add semantic
labels. Kind and content come from the source; callers cannot replace them or claim
native references and revision fields.

The semantic clone receives deterministic new IDs for every block and dynamic
field. Its metadata records `cloned_from`. Reapplying the operation to the same
base document produces the same semantic ID map.

In `0.2.0.dev32`, a conservative simple inline DrawingML picture receives its own
deterministic cloned image ID while continuing to share the source asset. The new ID
is immediately readable and inspectable. Use `image.update`, `paragraph.format`, or
the out-of-band `replace_image()` API in a subsequent transaction. Binary replacement
then forks only that occurrence through the cloned part's local relationship graph.

In `0.2.0.dev33`, `image.update` can also set or clear a bounded rectangular source
crop on the clone. The crop changes only the cloned story's `a:srcRect`; its
relationship and shared source media remain unchanged, and the source story stays
byte-exact.

In `0.2.0.dev34`, the same clone workflow projects the conservative floating
offset-and-square-wrap subset. The cloned occurrence receives its own image ID while
retaining identical anchor reference frames, offsets, wrap distances, relative
height, and compatibility flags. Its `wp14:anchorId` and `wp14:editId` identities
are collision-safely rebased along with the existing DrawingML IDs. Subsequent crop,
resize, metadata, paragraph, or binary edits re-prove that layout rather than
reconstructing it.

In `0.2.0.dev35`, `image.anchor.update` may then selectively move or rewrap the
cloned floating occurrence. The patch remains local to the cloned story and keeps
its rebased `wp14` identities, relationship graph, media, crop, extent, and source
story unchanged.

In `0.2.0.dev37`, the projected and clone-safe floating subset also accepts
`wp:align` positioning. A later `image.anchor.update` can switch either complete
position group between alignment and physical offset while preserving the cloned
story's rebased native identities and local relationship graph.

In `0.2.0.dev40`, clone-safe floating pictures may use square, no-wrap,
top-and-bottom, tight, or through wrapping. Tight/through polygons retain their
ordered raw native points and optional `edited` flag. A later
`image.anchor.update` can replace the complete supported wrap group story-locally
while retaining optional parent distances, supported wrap-local geometry, rebased
identities, media, crop, and source story.

In `0.2.0.dev41`, clone-safe horizontal and vertical positions may also use Office
2010 percentage offsets. The clone retains the exact signed Int32 native percentage
values, and a later story-local `image.anchor.update` may switch among physical
offset, alignment, and percentage modes without changing the source story.

In `0.2.0.dev42`, a clone-safe floating picture may additionally carry independent
Office 2010 relative width and height rules. The clone retains those rules together
with its absolute `wp:extent` fallback. A story-local update can replace or clear
the complete relative-size group without changing the source story.

In `0.2.0.dev43`, clone-safe inline and floating pictures may carry DrawingML
rotation plus horizontal and vertical mirror state. Cloning retains the exact
native `a:xfrm` attributes, while the semantic projection exposes canonical
clockwise degrees and booleans. A story-local `image.update`, anchor update, or
copy-on-write binary replacement re-proves that the transform remains isolated to
the intended occurrence.

In `0.2.0.dev44`, clone-safe pictures may also carry one supported direct-RGB
DrawingML outline with an explicit native width and preset dash. The clone retains
the complete `a:ln` subtree and shares the same media target; later outline,
transform, anchor, paragraph, or binary updates re-prove that only the selected
story occurrence changed.

In `0.2.0.dev45`, clone-safe pictures may carry one direct fixed-opacity effect.
The clone preserves the exact `a:blip/a:alphaModFix` value and shares the media
target; later opacity, outline, transform, anchor, paragraph, or binary updates
remain isolated to the selected story occurrence.

In `0.2.0.dev46`, clone-safe pictures may also carry one strict direct-RGB outer
shadow. The clone preserves the complete `a:effectLst/a:outerShdw` subtree and
shares the media target; later shadow, opacity, outline, transform, anchor,
paragraph, or binary updates remain isolated to the selected story occurrence.

In `0.2.0.dev47`, a strict inline `mc:AlternateContent` DrawingML/VML picture can be
projected, read, resized, and conditionally replaced inside an existing header or
footer. Cloning a part that contains the VML compatibility branch is still refused:
the current clone proof intentionally does not rebase legacy VML shape identities or
duplicate producer-specific fallback assets.

## Clone and bind atomically

The new part may be assigned to a section in the same Patch:

```json
[
  {
    "op": "header_footer.clone",
    "target": "#report_header",
    "part": {
      "id": "appendix_header"
    }
  },
  {
    "op": "section.header_footer.bind",
    "target": "#appendix_section",
    "set": {
      "header_default": "#appendix_header"
    }
  }
]
```

This changes the selected section to an independent story without modifying the
source region. To keep clone creation separately provable, edit the new clone in a
subsequent Patch. Editing the source later in the same Patch is safe and does not
change the clone.

## Native graph behavior

For each clone, AiOffice:

1. proves the source part and its single internal document relationship;
2. allocates a collision-free `/word/headerN.xml` or `/word/footerN.xml`;
3. deep-copies the complete supported native story;
4. assigns new `w14:paraId` values to every paragraph;
5. assigns collision-free IDs to `wp:docPr`, `a:cNvPr`, `pic:cNvPr`, and present
   `wp14:anchorId` / `wp14:editId` identities;
6. copies the source part-local `.rels` payload exactly;
7. keeps each local relationship target unchanged, so images and other package
   resources are shared rather than duplicated;
8. creates one new document relationship and one content-type override;
9. persists the new semantic-to-native identities in the manifest.

The source story, source local relationships, shared media, styles, body content,
and unrelated package parts remain byte-for-byte untouched. The cloned story is
structurally identical except for the identities that must be rebased.

## Conservative refusal boundary

Cloning fails atomically when the source contains identities that are not yet safe
to duplicate across stories, including:

- bookmarks, comments, permission ranges, and tracked move/change ranges;
- structured document tags, `customXml`, `altChunk`, and subdocuments;
- embedded Word objects, legacy `w:pict`, or VML;
- malformed, duplicate, or incomplete part-local relationships;
- a missing relationship target or ambiguous document relationship.

Ordinary rich text, normalized or preserved fields, external hyperlinks, and
DrawingML images are supported. Relationship targets are still governed by the
native package security policy.

## Proof and verification

Native lowering records the clone's creation-time story and relationship
signatures. At transaction end it proves that the new part, content type, document
relationship, local relationship graph, section bindings, and persistent identities
all match the semantic result. Any mismatch aborts without changing the source
document.

Structural equality does not establish visual quality. Reopen and render
layout-sensitive results through the native LibreOffice/Poppler path, then inspect
the affected pages for selection of first/even/default variants, header/footer
distance, image placement, clipping, overlap, and field display.

The JSON Spec is the AI-facing intent and evidence protocol. The DOCX package and
native render remain the authorities for lossless storage and appearance.
