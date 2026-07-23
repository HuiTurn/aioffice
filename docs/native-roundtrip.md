# Native round-trip architecture

AiOffice uses two coordinated representations because one model cannot provide both
AI-friendly semantics and lossless Office preservation.

## Sources of truth

| Origin | Semantic Spec | Native package |
| --- | --- | --- |
| Created by AiOffice | Authoritative intent | Compiled representation |
| Imported DOCX | Editable projection | Authoritative native content |

The AiOffice Spec is the model exchange protocol. It contains stable IDs, semantic
content, declarative operations, diagnostics, and revision metadata. It does not try
to serialize every OOXML element.

An imported native package remains an immutable base with copy-on-write part
overrides. A no-op export returns the exact original bytes. A supported edit rewrites
only affected parts; opaque and unsupported parts remain in the package.

## Persistent identity

AiOffice-generated DOCX files contain:

```text
/customXml/aioffice-manifest.xml
```

The manifest stores the artifact ID, revision, Spec version, content, section,
header/footer part, region-block, and dynamic-field IDs, native references,
structural paths, native object IDs, and fingerprints. Paragraph anchors are emitted
as `w14:paraId` values and declared through Markup Compatibility. Inline field
references additionally carry their ordinal within the paragraph, so a field does
not compete with its containing paragraph for identity. Section identities point to
the exact paragraph-carried or body-level `w:sectPr` without pretending the semantic
Spec contains all of its XML.

Third-party documents use the same identity model in a `.aioffice/` workspace
sidecar. Rebinding after an external edit follows this order:

1. exact package hash and native path;
2. native object ID;
3. unique native fingerprint;
4. structural path confirmed by a neighboring fingerprint.

If no unique match can be proven, AiOffice emits `IDENTITY_AMBIGUOUS`, assigns a new
ID to the unbound projection, and refuses a workspace reconcile commit. It never
silently gives an uncertain object an old semantic ID.

## Workspace revisions

```text
.aioffice/
├── workspace.json
└── artifacts/
    └── <artifact-id>/
        ├── manifest.json
        ├── manifests/<revision>.json
        ├── revisions/<revision>.docx
        ├── snapshots/<revision>.json
        └── patches/<revision>.json
```

Native revisions, semantic snapshots, identity manifests, and patch records are
written atomically. The workspace index is replaced last, so incomplete writes do
not become visible revisions. Commits use `base_revision` optimistic concurrency.
Idempotency keys replay the original result and are rejected if reused for different
operations.

The original imported file is never overwritten. Workspace export also refuses to
overwrite unless the caller explicitly opts in.

## Current native lowering boundary

The current DOCX native layer lowers `text.replace`, `paragraph.format`,
`text.format`, `node.remove`, `style.define`, `style.apply`, `style.format`,
`section.format`, and `field.update`.
Text replacement can cross Word run boundaries while retaining run properties and
unknown XML. Paragraph and text formatting mutate only selected supported
`w:pPr` / `w:rPr` properties and preserve unrelated or unknown children. Character
ranges can cross multiple runs and hyperlinks; boundary runs are split only around
selected `w:t` content. If a partial boundary run contains unsupported inline
children, the complete Patch is refused rather than duplicating or dropping them.
List nodes may reference multiple native paragraphs, and removing a list removes
that complete native range atomically.

Paragraph `w:pStyle` references, supported paragraph style definitions,
`basedOn`/`next` links, `w:docDefaults`, and heading outline semantics are projected
into strict Spec models. Imported documents use the empty `native-docx` theme so
AiOffice never overlays `business-clean` defaults on an existing template. Style
edits mutate only `word/styles.xml` and, when a node reference changes, its single
`w:pStyle`. Unknown style XML remains in place.

Paragraph-carried section properties and the final body section are projected as
ordered semantic sections. Native section edits update only explicitly selected
`w:type`, `w:pgSz`, `w:pgMar`, `w:cols`, `w:vAlign`, `w:titlePg`, or `w:pgNumType`
values.
Unrecognized attributes, children, header/footer relationships, and other section
settings remain untouched.

Header/footer parts are separately mapped by part URI. Ordinary paragraph edits are
lowered to the referenced `headerN.xml` or `footerN.xml`; `document.xml` and unrelated
region parts remain byte-identical. Shared parts stay shared, while an absent section
binding remains an inheritance instruction. PAGE, NUMPAGES, SECTION, and
SECTIONPAGES fields are projected as stable inline `DocumentField` objects with
non-authoritative cached results. `field.update` rewrites only the selected field
instruction and dirty flag in its native part. Unknown isolated fields are
structured but read-only; malformed field containment, drawings, objects, tables,
and unknown elements remain opaque and cannot be selected for destructive text
edits.

Other operations are rejected before a new native revision is committed. Future
iterations will add richer field families, tables, drawings, and further
layout-aware operations behind the same capability and fidelity contracts.
