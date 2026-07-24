# Incremental native header/footer binding

AiOffice `0.2.0.dev30` can change the explicit header/footer references of an
existing or newly inserted DOCX section without rebuilding any reusable region
part. JSON carries the model-facing binding intent; the attached package remains
authoritative for relationship IDs, section-property ordering, unknown XML, and
the exact region content.

## Operation

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

The target must resolve to one semantic section. `set` accepts any of:

- `header_default`, `header_first`, and `header_even`;
- `footer_default`, `footer_first`, and `footer_even`.

Every value is a stable ID from the document's `header_footers` collection. A
leading `#` is accepted for selector consistency and removed during normalization.
The part's `kind` must match the selected header or footer slot.

`clear` removes an explicit reference from that section. This means native Word
inheritance, not blank content. Bind an existing empty part when the desired result
is explicitly blank.

At least one requested slot must actually change. Unknown operation fields,
unknown slots, overlapping `set`/`clear` entries, duplicate clear entries, null or
empty IDs, missing parts, and kind mismatches are rejected before mutation.

## Native relationship proof

For each selected part, native lowering proves:

1. its semantic source reference identifies a header or footer package part;
2. the package part exists and has the expected `w:hdr` or `w:ftr` root;
3. `word/document.xml` owns exactly one internal relationship of the expected
   header/footer type to that part;
4. an existing projected section reference uses that exact relationship ID;
5. the live section boundary still belongs to the requested stable section.

External relationships and multiple relationships to the same projected part are
ambiguous and fail atomically. AiOffice does not select one by array position or
relationship-ID spelling.

## Minimal XML patch

An existing selected `w:headerReference` or `w:footerReference` keeps its element
and unknown attributes; only `r:id` changes. A missing reference is inserted among
the leading section reference children in OOXML schema order:

- headers before footer references;
- footer references after all header references;
- every reference before the remaining `w:sectPr` properties.

Clearing removes only the one proven direct child. Other variants, unknown section
attributes and children, layout properties, tracked content outside the selected
reference, reusable part XML, content types, and relationship parts stay unchanged.

If one slot contains duplicate direct references, AiOffice refuses the operation.
If `clear` targets a malformed native reference that was not safely represented in
the semantic projection, it also refuses rather than deleting opaque state. An
explicit `set` may replace one such non-ambiguous reference because the caller
provides a fully proven destination.

## Same-Patch state

Native lowering tracks the live binding map for every section. This supports:

- multiple binding changes in one atomic Patch;
- binding a section created earlier by `section.insert_before`;
- splitting a section after an earlier binding change and inheriting the current
  explicit binding map;
- mixing binding edits with `section.format` and ordinary content edits.

At the end of lowering, the complete live native binding map must equal the final
semantic section model. A mismatch aborts the transaction.

## Identity and package preservation

Changing a reference can alter the fingerprint of `w:sectPr`, so AiOffice ensures
that a standalone identity manifest is present and refreshes section and region
source references. This preserves stable IDs after export and reopen.

All safe internal header/footer relationships remain projected, including parts that
become temporarily unbound after this operation. The AI can therefore bind back to
the old part in a later transaction. Repeated fallback-ID collisions for third-party
parts without native paragraph anchors are resolved through an unbounded,
deterministic suffix sequence.

For a package that already contains the AiOffice manifest, the expected changed
parts are normally:

```text
/word/document.xml
/customXml/aioffice-manifest.xml
```

For a third-party package without a manifest, the root relationship and content-type
control parts are additionally changed once to attach it. Header/footer parts and
`/word/_rels/document.xml.rels` are not rewritten by this operation.

## Deliberate boundary

This operation reuses projected parts. It does not:

- create a part implicitly or delete a header/footer part;
- clone a shared part for copy-on-write editing;
- add or remove a document relationship;
- bind an unprojected orphan package part;
- repair multiple or external relationships;
- enable first-page or even/odd layout switches automatically.

Use `section.format` to set `different_first_page`. Even-page variants also require
the document-wide `settings.even_and_odd_headers` switch, which remains separately
modeled. Capabilities expose the supported slots, set/clear semantics, native patch
scope, preservation guarantees, and unsupported part lifecycle operations.

Use [`header_footer.create`](native-header-footer-creation.md) to create a reusable
semantic part explicitly, optionally immediately before this binding operation in
the same atomic Patch.

## Visual verification

Binding correctness is first proven structurally by exact relationship targets,
section order, semantic reprojection, and standalone reopen. Native PDF/PNG remains
the authority for:

- first/even/default page selection;
- header and footer distance;
- interaction with section breaks and page-number restarts;
- inherited versus explicitly blank regions;
- clipping, overlap, and visual alignment.

JSON is therefore the AI-facing intent and evidence protocol, while the native DOCX
package and rendered pages remain the fidelity authorities.
