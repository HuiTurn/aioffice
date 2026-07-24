# Incremental native paragraph and heading insertion

AiOffice `0.2.0.dev22` can insert a new paragraph or heading into an imported DOCX
without rebuilding the document from its JSON projection. JSON remains the
AI-facing intent and evidence layer; the attached OPC package remains the native
authority.

## Operation

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

The target must be a mapped top-level body node. AiOffice resolves its complete
native range and inserts one freshly compiled `w:p` after the range's last element.
This makes a multi-paragraph list a valid anchor without placing the new paragraph
inside the list.

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

## Fidelity and safety

Before mutation AiOffice proves that:

1. the anchor belongs directly to `/word/document.xml`'s `w:body`;
2. every mapped anchor element is present and contiguous;
3. the anchor does not carry a `w:sectPr` section boundary;
4. the new block has no forged `source_ref`;
5. the required named style exists in the native style catalog;
6. generated text and attributes form safe, valid XML;
7. the generated field count matches its semantic field identities.

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
JSON projection whose extension declares native authority omits `node.insert_after`
and rejects it until the original DOCX package is attached again.
