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
- Added native lowering for text replacement, paragraph/text direct formatting, and
  node removal while preserving unknown OOXML.
- Added explicit-unit paragraph and text style models shared by semantic HTML,
  generated DOCX, and imported DOCX projections.
- Added exact Unicode code-point range and occurrence selectors for text formatting.
- Added minimal DOCX run splitting across mixed formatting and hyperlinks, with
  atomic refusal when a partial selection crosses unsupported inline content.
- Added rich run-style and hyperlink projection for imported paragraphs and headings.
- Added cross-span semantic text replacement and rich heading content.
- Added document defaults and AI-addressable named paragraph styles with explicit
  semantic roles, inheritance, heading levels, and stable references.
- Added `style.define`, `style.apply`, and selective `style.format` operations for
  semantic and native DOCX documents.
- Added native `w:docDefaults` / `w:style` projection and minimal `styles.xml`
  mutation that preserves unknown style XML.
- Added inherited named-style rendering in semantic HTML and compact style discovery
  through inspect/capabilities.
- Added stable semantic document diffs and per-property format Patch results.
- Added render contracts that distinguish approximate previews from native evidence.
- Added optional PNG visual-regression metrics for future native render providers.
- Added strict, ordered page and section models with standard/custom paper sizes,
  orientation, margins, gutter, header/footer distances, equal or unequal columns,
  section start types, vertical alignment, and first-page behavior.
- Added correct DOCX projection for paragraph-carried and final body `w:sectPr`
  elements, including structural carrier suppression and persistent section IDs.
- Added deterministic multi-section DOCX generation and section-aware semantic HTML
  page hints.
- Added selective `section.format` native lowering that mutates one mapped
  `w:sectPr`, preserves unknown section XML, and refreshes section fingerprints.
- Added section anchor, ordering, page-geometry, and column-overflow diagnostics.
- Added normalized, reusable header/footer parts and per-section bindings for
  default, first, and even page variants.
- Added document-wide `even_and_odd_headers` settings with native
  `word/settings.xml` generation and projection.
- Added header/footer relationship, content-type, inheritance, persistent identity,
  and semantic HTML preview support.
- Extended `text.replace`, `paragraph.format`, and `text.format` native lowering to
  ordinary header/footer paragraphs while touching only the target part.
- Added conservative opaque projection for native header/footer drawings, objects,
  tables, malformed field containment, and unknown elements so unsupported content
  is never reconstructed from display text.
- Added stable inline `DocumentField` objects for PAGE, NUMPAGES, SECTION, and
  SECTIONPAGES, with cached results explicitly separated from instructions.
- Added complex-field generation, `w:fldSimple`/complex-field projection,
  `update_fields_on_open`, semantic HTML hints, and part-scoped field identities.
- Added `field.update` native lowering that changes one field instruction, marks its
  result dirty, and preserves cached results and unknown surrounding XML.
- Added section page-number restart and decimal, Roman, or alphabetic formats through
  selective `w:pgNumType` projection and Patch.
- Changed unknown native fields from lossy display-text projection to structured,
  read-only preservation; malformed field containment remains opaque.
- Serialized OPC content-type and relationship control parts with default package
  namespaces for LibreOffice interoperability.
- Added structured fidelity policies and reports.
- Added package limits and defenses against traversal, ZIP bombs, unsafe XML, and macros.
- Advanced the AiOffice Document Spec to `0.2-draft.7`.

## 0.1.0

- Added the strict AiOffice Document Spec 1.0 draft.
- Added stable artifact and semantic node IDs.
- Added document creation through a declarative API and builder.
- Added JSON and Markdown input.
- Added JSON, Markdown, semantic HTML, and DOCX output.
- Added machine-readable document validation.
- Added atomic, revision-checked patch operations.
- Added the `aioffice` CLI.
