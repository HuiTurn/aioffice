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

The manifest stores the artifact ID, revision, Spec version, node IDs, native
references, structural paths, native object IDs, and fingerprints. Paragraph anchors
are emitted as `w14:paraId` values and declared through Markup Compatibility.

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

The current DOCX native layer lowers `text.replace` and `node.remove`. Text
replacement can cross Word run boundaries while retaining run properties and
unknown XML. List nodes may reference multiple native paragraphs, and removing a
list removes that complete native range atomically.

Other operations are rejected before a new native revision is committed. Future
iterations will add style, paragraph, table, section, header/footer, drawing, and
layout-aware operations behind the same capability and fidelity contracts.
