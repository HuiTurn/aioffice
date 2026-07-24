# Incremental native header/footer creation

AiOffice `0.2.0.dev30` can create a reusable header or footer inside an imported
DOCX without rebuilding the document package. JSON describes the semantic region;
the attached native package remains authoritative for package-part names,
relationship IDs, content types, and all preserved content.

## Operation

```json
{
  "op": "header_footer.create",
  "part": {
    "id": "appendix_header",
    "kind": "header",
    "metadata": {
      "role": "running_header"
    },
    "content": [
      {
        "id": "appendix_header_line",
        "type": "paragraph",
        "paragraph_style": {
          "alignment": "right"
        },
        "content": [
          {
            "type": "text",
            "text": "Appendix · ",
            "marks": ["strong"]
          },
          {
            "id": "appendix_page",
            "type": "field",
            "kind": "page_number"
          }
        ]
      }
    ]
  }
}
```

The part has a stable AI-facing ID and a required `kind` of `header` or `footer`.
Its content may contain zero or more ordinary paragraphs. Paragraphs support the
same direct paragraph/text styles, rich spans, normalized PAGE/NUMPAGES/SECTION/
SECTIONPAGES fields, and external hyperlinks as generated DOCX content.

An empty content list creates an explicitly blank but reusable native region.
Tables, images, drawings, embedded objects, opaque XML, and caller-supplied
`source_ref` values are rejected. Native identities and revision fields are owned
by AiOffice and assigned during the transaction.

## Create and bind atomically

Creation does not silently change page behavior. Bind the new part explicitly,
including in the same Patch:

```json
[
  {
    "op": "header_footer.create",
    "part": {
      "id": "appendix_footer",
      "kind": "footer",
      "content": [
        {
          "id": "appendix_footer_line",
          "type": "paragraph",
          "text": "Confidential"
        }
      ]
    }
  },
  {
    "op": "section.header_footer.bind",
    "target": "#appendix_section",
    "set": {
      "footer_default": "#appendix_footer"
    }
  }
]
```

Semantic validation sees operations in order. Native lowering also registers the
new part as live state before processing the binding, so the second operation must
resolve through the exact relationship created by the first. Any failure aborts
the complete Patch.

## Native package transaction

For each created part, AiOffice:

1. scans existing package part URIs case-insensitively and allocates the first free
   `/word/headerN.xml` or `/word/footerN.xml`;
2. compiles the semantic paragraphs to a `w:hdr` or `w:ftr` root and assigns
   collision-free native paragraph anchors;
3. allocates a unique `rIdAiOfficeHeaderN` or `rIdAiOfficeFooterN` in
   `/word/_rels/document.xml.rels`;
4. adds exactly one internal document relationship of the required type;
5. adds exactly one matching override to `/[Content_Types].xml`;
6. creates a part-local relationship file only when external hyperlinks require it;
7. refreshes persistent identities in the AiOffice manifest.

Before changing the copy-on-write package, lowering rejects malformed relationship
roots, missing or ambiguous relationship IDs, ambiguous content-type overrides,
part-name collisions, and an invalid relationship content-type default. Package
size, XML size, part-count, compression, external-relationship, and macro policies
remain enforced by the native package security layer.

At the end of the Patch, AiOffice proves that:

- the final semantic reusable-part collection equals the live native collection;
- the created part exists with the correct root and content type;
- exactly one internal document relationship with the allocated ID and type reaches
  that part;
- exactly one matching content-type override exists;
- every created paragraph and supported field has a part-scoped persistent identity;
- the complete final section-binding map equals the semantic result.

## Preservation boundary

For a package that already has an AiOffice identity manifest, creating and binding a
plain region normally changes:

```text
/[Content_Types].xml
/word/_rels/document.xml.rels
/word/document.xml
/customXml/aioffice-manifest.xml
```

and adds:

```text
/word/headerN.xml or /word/footerN.xml
```

A region containing external hyperlinks also adds its own `.rels` part. Existing
body content, styles, media, unrelated relationships, old header/footer parts, and
unknown XML remain byte-for-byte untouched. A third-party package without a
manifest additionally receives the one-time manifest relationship and content-type
entries.

Deleting a part, cloning a shared part, and copy-on-write editing of complex native
region content remain separate lifecycle operations. `header_footer.create` never
guesses which section should use the new part and never enables first-page or
even-page switches automatically.

## Verification

Structural proof is necessary but does not establish visual quality. For
layout-sensitive output, reopen the result, render it through the native
LibreOffice/Poppler path, and inspect the affected pages for header/footer distance,
alignment, clipping, overlap, field display, and section inheritance.

The JSON Spec is the AI-facing intent and evidence protocol; the native DOCX package
and rendered pages remain the fidelity authorities.
