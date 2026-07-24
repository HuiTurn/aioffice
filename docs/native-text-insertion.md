# Incremental native paragraph and heading insertion

AiOffice `0.2.0.dev24` can insert a new paragraph or heading into an imported DOCX
without rebuilding the document from its JSON projection. JSON remains the
AI-facing intent and evidence layer; the attached OPC package remains the native
authority.

## Operations

```json
{
  "op": "node.insert_after",
  "target": "#summary",
  "content": {
    "id": "decision",
    "type": "heading",
    "level": 2,
    "text": "Decision"
  }
}
```

Use the symmetric operation to place content before an anchor:

```json
{
  "op": "node.insert_before",
  "target": "#summary",
  "content": {
    "id": "context",
    "type": "paragraph",
    "text": "Context for the summary."
  }
}
```

The target must be a mapped top-level body node. AiOffice resolves its complete
native range and inserts one freshly compiled `w:p` after its last element or before
its first element. This makes a multi-paragraph list a valid anchor without placing
the new paragraph inside the list. `node.insert_before` also reaches the beginning
of the document without a synthetic root or array index.

Use the document root when the content belongs at the end of the final section:

```json
{
  "op": "node.append",
  "target": "$",
  "content": {
    "id": "appendix",
    "type": "heading",
    "level": 2,
    "text": "Appendix"
  }
}
```

Unlike `node.insert_after`, root append does not require an existing anchor and
therefore works for an empty document. In native DOCX, AiOffice inserts the new
`w:p` immediately before the optional final body-level `w:sectPr`. The original
section properties stay terminal and unchanged, so the block belongs to the final
semantic section.

The `id` is optional. When omitted, the semantic transaction generates one and
passes that exact identity to native lowering. Supplying IDs is recommended for
multi-operation agent plans because later operations can refer to them directly.

## Supported content

The native subset accepts:

- `paragraph` and `heading` blocks;
- plain text or ordered rich `TextSpan` / normalized `DocumentField` content;
- strong, emphasis, underline, strike, code, subscript, superscript, highlight, and
  link marks;
- direct `ParagraphStyle`, block `TextStyle`, and per-span `TextStyle`;
- internal `#bookmark` hyperlinks and external relationship-backed hyperlinks;
- PAGE, NUMPAGES, SECTION, and SECTIONPAGES fields, including normalized number
  formats and cached display results;
- an existing named paragraph style through `style_ref`; a heading without
  `style_ref` uses the document's `Heading1` through `Heading6` style.

Native-only field instructions are read-only and cannot be inserted from their
display projection. Lists, tables, images, page breaks, and opaque blocks require
dedicated native operations rather than generic text insertion.

## Batch object tracking

Native lowering consumes the semantic change record for each operation and binds a
new ID to the actual inserted XML object. The following is one atomic Patch:

```json
[
  {
    "op": "node.insert_after",
    "target": "#summary",
    "content": {
      "id": "recommendation",
      "type": "paragraph",
      "text": "Approve the plan."
    }
  },
  {
    "op": "paragraph.format",
    "target": "#recommendation",
    "set": {"alignment": "center"}
  },
  {
    "op": "text.replace",
    "target": "#recommendation",
    "search": "Approve",
    "replacement": "Approve and fund"
  }
]
```

The same live object can anchor another insertion or be moved or removed later in
the batch. Native indices are computed only after all operations finish.

## Section starts

For a later document section, `start_at` names its first semantic content node. When
`node.insert_before` targets that node, the created node becomes the section's new
`start_at`. Native lowering proves that the new paragraph is placed after the
preceding section's existing `w:sectPr` and before the old first node.

Repeated prepends in one Patch update this state in order. Each change record reports
the section ID and the old and new anchors:

```json
{
  "section_start_updated": {
    "section_id": "analysis_section",
    "from": "analysis_heading",
    "to": "analysis_context"
  }
}
```

Insertion before a text-bearing paragraph that itself carries `w:sectPr` is safe:
the unchanged boundary still follows that paragraph. Insertion *after* the same
paragraph remains refused because it would silently change the inserted block's
section ownership.

## Fidelity and safety

Before relative insertion AiOffice proves that:

1. the anchor belongs directly to `/word/document.xml`'s `w:body`;
2. every mapped anchor element is present and contiguous;
3. after-insertion anchors do not carry a `w:sectPr`; before-insertion preserves
   any boundary inside the unchanged anchor;
4. the new block has no forged `source_ref`;
5. the required named style exists in the native style catalog;
6. generated text and attributes form safe, valid XML;
7. the generated field count matches its semantic field identities.

For root append, AiOffice separately proves that a direct body-level `w:sectPr`, if
present, is unique and terminal. A malformed or ambiguous body layout is refused
atomically instead of guessing an insertion point.

External hyperlinks receive collision-safe relationship IDs. New paragraphs receive
collision-safe deterministic `w14:paraId` values. All shifted content, section,
field, table, cell, and header/footer references are refreshed. If the input is a
third-party DOCX without an AiOffice identity manifest, the first successful
insertion attaches one so the new and shifted IDs survive standalone reopen.

Any failed proof aborts the entire native transaction. The source `Document` and
source package bytes remain unchanged.

## CLI and Workspace

Generic Patch paths require no special binary channel:

```bash
aioffice apply report.docx insert.json --output report-updated.docx
aioffice workspace apply ARTIFACT_ID insert.json --root project
```

Inspect `Document.capabilities()["structural_editing"]` before planning. A detached
JSON projection whose extension declares native authority omits root append and
both relative insertion operations, rejecting them until the original DOCX package
is attached again.
