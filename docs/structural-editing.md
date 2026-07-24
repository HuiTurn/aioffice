# Stable-ID structural editing

AiOffice structural edits address semantic nodes, not array positions. The dev25
contract exposes insertion, a deliberately narrow pair of relative moves, and
removal:

```json
{
  "op": "node.move_after",
  "target": "#risk_table",
  "after": "#executive_summary"
}
```

```json
{
  "op": "node.move_before",
  "target": "#executive_summary",
  "before": "#report_title"
}
```

The target and relative anchor must resolve to different top-level content IDs. On
success, the target becomes the immediate semantic predecessor or successor of the
anchor and receives the new document revision. Together the operations can reach the
first and last position without exposing an array index. The source `Document`
remains immutable; dry run, diff, optimistic revision checks, idempotency, CLI
application, and Workspace persistence use the ordinary Patch transaction.

Deletion uses the same stable selector and transaction boundary:

```json
{
  "op": "node.remove",
  "target": "#obsolete_appendix"
}
```

For an imported DOCX, the complete mapped XML range is removed. A multi-paragraph
list is deleted as one group; a projected image removes its complete host paragraph.
AiOffice deliberately keeps relationships and package parts that become unreferenced
because another unknown native consumer may still depend on them.

## Native insertion versus movement

`node.append`, `node.insert_after`, and `node.insert_before` create a new semantic
node. For an imported DOCX, dev25 accepts new `paragraph`, `heading`, and
`page_break` blocks and compiles exactly one new `w:p`. Root append targets `$` and
works even when the document has no semantic content; it places the new block
before the optional final body-level `w:sectPr`. The relative operations preserve
the complete anchor. Rich text, direct paragraph/text formatting, internal and
external hyperlinks, and normalized document fields are supported for text blocks;
a page break is compiled as an isolated `w:r/w:br` control. See
[native text insertion](native-text-insertion.md).

Reconstructing an *existing* imported DOCX block from its JSON projection would
still lose unsupported native detail. The move operations therefore relocate the
existing native XML objects:

1. resolve the target and anchor through their trusted stable identities;
2. resolve every native element belonging to each semantic node;
3. prove that both ranges are present, top-level, disjoint, and contiguous;
4. remove the complete target range from `w:body`;
5. insert those same element objects before or after the complete anchor range;
6. recompute all affected `NativeRef` indices and fingerprints;
7. serialize the changed document part and refreshed identity manifest atomically.

A list backed by two or more `w:p` elements therefore moves as one unit. A table's
`w:tbl`, an image's complete host `w:p`, and unknown children inside a moved element
are retained rather than regenerated. Relationship IDs and related package parts do
not change merely because their referring element moved.

Sequential structural operations in one Patch use live XML element objects.
AiOffice does not look up a later operation using an element index that an earlier
insert or move has already invalidated. A newly inserted caller-selected or
stable ID can immediately anchor another insertion or receive text, formatting,
move, and removal operations. When the ID is omitted, the generated value is
returned in `changes[].created_nodes` for a later Patch.

## Section safety boundary

Word section semantics depend on the position of `w:sectPr`. A move that casually
crosses one of those boundaries can silently change page size, margins, columns,
headers, footers, numbering, or vertical alignment. Dev25 therefore requires:

- target and anchor belong to the same semantic section;
- the target is not the `start_at` node of a later section;
- neither native range contains a `w:sectPr`;
- both ranges belong directly to `/word/document.xml`'s `w:body`.

When `node.move_before` prepends a node within a later section, the native section
carrier remains in place and the semantic section's `start_at` is rebound to the
moved node. The change evidence records the section ID and old/new anchors. A
text-bearing paragraph that itself carries `w:sectPr` is refused. Cross-section
movement will require an explicit future operation that updates section ownership
and proves header/footer semantics together; dev25 does not approximate it.

Insertion is placed before or after the anchor's complete contiguous range. An
after-insertion anchor that carries `w:sectPr` is refused because placement after
that paragraph would change the native section. Before-insertion is safe because the
boundary remains after the unchanged anchor. If the anchor is the first node of a
later semantic section, its created predecessor becomes the new `section.start_at`;
native lowering proves that both remain after the preceding section boundary.

Root append always belongs to the last semantic section. AiOffice inserts before a
unique terminal body `w:sectPr`, never after it. If direct body-level section
properties are duplicated or nonterminal, the Patch fails atomically.

## Identity and third-party packages

Structural edits change native indices even when the XML payloads are otherwise
identical. AiOffice refreshes identity records for all mapped content, fields,
sections, tables, cells, and header/footer blocks.

If a third-party DOCX has no embedded AiOffice identity manifest, the first
successful insertion, move, or removal attaches one root relationship, the manifest
part, and a valid content type before export. Stable IDs then survive standalone
reopen rather than depending on the new ordinal positions. Workspace identity
evidence continues to be written alongside each revision.

## Diagnostics

The operations fail atomically with actionable diagnostics for:

- missing, ambiguous, identical, or already-adjacent target/anchor IDs;
- unknown operation fields;
- cross-section requests;
- movement of a section start anchor;
- detached or non-top-level native ranges;
- overlapping, missing, or non-contiguous native ranges;
- target or anchor elements carrying a native section boundary;
- removal targets carrying a native section boundary;
- insertion of unsupported block types, native-only fields, forged source
  references, missing paragraph styles, or unsafe XML text;
- any result that fails semantic validation or native identity refresh.

The machine-readable `structural_editing` capability reports root append, relative
insertion and both move operations, the remove operation, supported inserted blocks
and inline content, placement rules, selector type, native scope, multi-element
behavior, section policy, section-prepend behavior, identity behavior, conservative
orphan policy, source immutability, and dry-run support.

## CLI and Workspace

Write the operation as a JSON array or Patch envelope:

```bash
aioffice apply report.docx move.json --output reordered.docx
aioffice workspace apply ARTIFACT_ID move.json --root project
```

Workspace patch records store only the stable IDs and structured change evidence:

```json
{
  "operation": "node.move_after",
  "moved_nodes": ["risk_table"],
  "from_after": "appendix_heading",
  "after": "executive_summary",
  "section_index": 0
}
```

A section-prepend move additionally reports:

```json
{
  "operation": "node.move_before",
  "moved_nodes": ["recommendations"],
  "from_after": "appendix",
  "before": "section_heading",
  "section_index": 2,
  "section_start_updated": {
    "section_id": "recommendations_section",
    "from": "section_heading",
    "to": "recommendations"
  }
}
```

`from_after` is the previous stable predecessor, or `null` when the target was first.
It is audit evidence, not a selector to be cached for later edits.

For a document whose JSON extension declares native authority, keep the native DOCX
package attached while applying any structural operation. AiOffice refuses
insertion, movement, or removal against a detached native projection because JSON
alone cannot prove the complete XML range or native placement. Its capability
response reports structural editing as unavailable and omits all five operations
from the executable operation list. Documents created as semantic AiOffice specs do
not have this restriction.
