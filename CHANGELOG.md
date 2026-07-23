# Changelog

## 0.2.0 (unreleased)

- Added embedded DOCX identity manifests and `w14:paraId` anchors.
- Added persistent `.aioffice/` workspaces with native revisions, semantic snapshots,
  identity manifests, patch logs, checkout, and safe export.
- Added idempotent workspace commits and external DOCX reconciliation.
- Added ambiguity-safe identity matching using native IDs, fingerprints, paths, and
  neighboring fingerprints.
- Added semantic projection for native lists and page breaks.
- Added a safe, copy-on-write native OPC package graph.
- Added DOCX semantic projection with native source references.
- Added capability discovery and refreshed native identity maps across sequential edits.
- Added exact no-op DOCX round trips.
- Added minimal native lowering for text replacement and node removal.
- Added structured fidelity policies and reports.
- Added package limits and defenses against traversal, ZIP bombs, unsafe XML, and macros.
- Marked the AiOffice Document Spec as `0.2-draft.1`.

## 0.1.0

- Added the strict AiOffice Document Spec 1.0 draft.
- Added stable artifact and semantic node IDs.
- Added document creation through a declarative API and builder.
- Added JSON and Markdown input.
- Added JSON, Markdown, semantic HTML, and DOCX output.
- Added machine-readable document validation.
- Added atomic, revision-checked patch operations.
- Added the `aioffice` CLI.
