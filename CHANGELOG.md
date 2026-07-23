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
- Added optional PNG visual-regression metrics for native render providers.
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
- Added explicit table width, alignment, layout algorithm, indent, cell spacing,
  cell margins, repeated-header, row height, and row pagination contracts.
- Added stable table-column and table-row identities, preserving semantic column
  keys and data types through generated DOCX reopen and external reconciliation.
- Added `table.format` and `table.column.format` native lowering with selective
  `w:tblPr`, `w:tblGrid`, and one-to-one `w:tcW` mutation.
- Added conservative regular-grid detection: table-wide formatting remains safe for
  irregular tables, while column-width edits refuse merged or shifted grids.
- Added table geometry to compact inspection, capabilities, semantic HTML, and
  printable-width diagnostics.
- Added normalized logical `TableCell` objects with stable IDs, semantic column
  anchors, rich paragraph content, and backward migration from compact row values.
- Added horizontal `gridSpan` and vertical `vMerge` generation/projection with
  rectangular-grid proof, overlap diagnostics, and atomic refusal for invalid spans.
- Added stable native identities for table cells and editable cell paragraphs.
- Added cell-local vertical alignment, no-wrap, fit-text, fill color, and independent
  margins through strict `TableCellFormat` models.
- Added `table.cell.format` native lowering that mutates one mapped `w:tcPr` while
  preserving cell content and unknown OOXML.
- Extended text, paragraph, and style operations to supported rich paragraphs inside
  body-table cells.
- Added conservative read-only fallback for cells containing drawings, nested tables,
  dynamic fields, malformed content, or an unprovable logical grid.
- Added a native-compatible LibreOffice render provider for DOCX-to-PDF evidence.
- Added Poppler page rasterization with explicit one-based page selection and
  bounded 72–600 DPI PNG output.
- Added isolated per-render LibreOffice user profiles, external-command timeouts,
  output validation, page-count discovery, and explicit missing-tool failures.
- Added render evidence metadata for source/PDF/content hashes, engine versions,
  page count, page number, pixel dimensions, platform, and font-environment hashes.
- Added `aioffice render` with structured JSON summaries and dynamic render-provider
  capability discovery.
- Kept successful native evidence `unverified` until visual or regression review,
  so renderer success cannot be mistaken for aesthetic approval.
- Added `Document.render_pages()` to convert DOCX to PDF once and derive a bounded,
  internally consistent set of selected or complete page PNGs.
- Added contiguous page-range batching, deterministic page ordering, duplicate and
  out-of-range rejection, and a configurable 1–500 page resource ceiling.
- Added `PageVisualAnalysis` for background, ink ratio, content bounding box,
  four-side whitespace, apparent blank-page, and near-edge content diagnostics.
- Added `PaginatedRenderResult` / `RenderedPage` contracts with binary-free summaries,
  per-page hashes, safe output names, staged writes, rollback-on-error exclusive
  creation, and explicit overwrite behavior.
- Added `aioffice render-pages` with bounded comma/range selection, optional analysis,
  structured outputs, and overwrite refusal by default.
- Added strict table and cell border models with explicit edge style, width, color,
  and optional spacing constraints.
- Added DOCX generation and projection for table perimeter/internal borders and
  direct four-side cell borders.
- Added selective `table.format` and `table.cell.format` border lowering that
  preserves unknown edge XML and distinguishes clearing direct formatting from an
  explicit `none` edge.
- Added border conflict semantics and edge-aware semantic HTML preview without
  leaking fallback grid lines onto explicitly styled tables.
- Exposed standalone border JSON Schemas and machine-readable border capabilities
  for AI planning.
- Added paragraph-wide solid sRGB backgrounds and strict top/right/bottom/left
  border surfaces to `ParagraphStyle`.
- Added shared loss-aware WordprocessingML border codecs used by table, cell, and
  paragraph lowering.
- Added DOCX generation and conservative projection for `w:pPr/w:shd` and
  `w:pPr/w:pBdr` across body, heading, rich table-cell, header/footer, document
  defaults, and named-style paragraphs.
- Added selective `paragraph.format` and `style.format` surface Patch support with
  per-edge style inheritance, explicit `none`, direct-format clearing, unknown XML
  preservation, and part-scoped fidelity reports.
- Kept pattern/theme shading and paragraph `between`/`bar` borders native-only and
  losslessly preserved instead of exposing misleading semantic values.
- Added resolved paragraph/text styling to header/footer HTML previews and exposed
  paragraph surface schemas and capability metadata.
- Added conservative `image` block projection for one embedded, inline DrawingML
  picture in an otherwise empty body paragraph, including explicit physical extent,
  native alternative text, title/name, and stable paragraph identity.
- Added content-addressed `AssetRef` records with media type, filename, byte count,
  and full SHA-256 while deliberately keeping image bytes out of the JSON Spec.
- Added verified `Document.read_image()`, `image_bytes()`, and `extract_image()` APIs
  plus `aioffice extract-image`; every read re-resolves the trusted native paragraph
  and OPC relationship and rechecks the binary hash and size.
- Added selective native `image.update` for alternative text, title, and displayed
  width/height, including aspect-ratio preservation when only one dimension is set.
- Added coordinated `wp:extent` and DrawingML transform extent mutation while
  preserving image bytes, relationships, asset identity, and unrelated OPC parts.
- Added strict native-package, selector, positive-EMU, clearability, XML-text, and
  pre/post image-shape validation with atomic failure outside the proven subset.
- Exposed the image update schema and machine-readable supported-operation,
  update-field, resize-mode, and native-geometry capabilities for AI planning.
- Added `Document.replace_image()` and `aioffice replace-image` as explicit
  out-of-band binary write channels; image bytes still never enter the JSON Spec.
- Added `Workspace.replace_image()` and `aioffice workspace replace-image` with
  revision persistence and binary-free patch logs.
- Added bounded signature and declared-media-type validation for PNG, JPEG, GIF, BMP,
  and TIFF replacement inputs with deterministic full-SHA-256 asset identities.
- Added occurrence-scoped copy-on-write image replacement using a content-addressed
  media part, content-type override, independent relationship, and selective
  `a:blip/@r:embed` mutation.
- Preserved shared-image occurrences, original image parts, displayed extent,
  accessibility metadata, stable occurrence identity, and unknown OPC consumers.
- Added relationship-graph refresh and prospective package part/uncompressed-size
  enforcement to copy-on-write native part mutation.
- Added atomic rejection for detached packages, raw JSON binary replacement, invalid
  signatures, MIME mismatches, oversized inputs, part collisions, and failed
  post-mutation native proofs.
- Added `Document.insert_image_after()` and `aioffice insert-image-after` for
  explicit-size, accessible inline images at a mapped top-level body position.
- Added stable caller-selected image IDs, collision-free `w14:paraId`/drawing IDs,
  coordinated DrawingML extents, and direct paragraph-style insertion.
- Added last-element anchoring for semantic nodes backed by multiple native elements,
  including lists, while preserving following content and section structure.
- Added automatic embedded identity-manifest attachment for structural edits to
  third-party DOCX files so inserted IDs survive standalone export and reopen.
- Added `Workspace.insert_image_after()` and
  `aioffice workspace insert-image-after` with binary-free revision logs.
- Added strict insertion schemas, explicit geometry/accessibility capability
  metadata, CLI overwrite safety, and atomic refusal for invalid or detached inputs.
- Extended `paragraph.format` to projected native image IDs so an AI can adjust the
  picture paragraph's alignment, spacing, indentation, page-flow controls, solid
  background, and supported physical borders through the existing strict style
  contract.
- Preserved the complete DrawingML tree, image relationship, binary asset, stable
  image ID, and unrelated OPC parts while selectively patching only requested
  `w:pPr` properties.
- Exposed image layout fields and `paragraph.format` in per-image and document
  capabilities, including the same paths through JSON Patch, CLI, and Workspace.
- Added set/clear round-trip coverage, native-minimality proof, invalid-style atomic
  refusal, real patch persistence, and image-binary stability checks.
- Added stable-ID `node.move_after` for semantic documents and imported DOCX files.
- Added exact top-level native range relocation: a multi-paragraph list moves as one
  contiguous group while tables, images, DrawingML, unknown XML, relationships, and
  binary parts remain unreconstructed.
- Added sequential structural Patch correctness by resolving original XML element
  objects instead of stale element indices, followed by complete native-reference
  and embedded-manifest reindexing.
- Added strict same-section movement, section-start-anchor protection, native
  `w:sectPr` carrier refusal, overlap/contiguity proofs, no-op diagnostics, and atomic
  rejection for unsupported structural requests.
- Added automatic identity-manifest attachment for the first successful third-party
  move, preserving caller-visible IDs across standalone export and reopen.
- Exposed machine-readable structural-editing capabilities and added CLI, Workspace,
  list-range, image, third-party, multi-operation, section-boundary, package
  minimality, and reopen tests.
- Changed mixed text/drawing paragraphs and complex image cases to explicit opaque
  projections so the semantic layer can no longer hide a picture inside an ordinary
  text node.
- Added accessible, dimensioned semantic HTML image placeholders, stable Markdown
  asset references, image/asset JSON Schemas, and machine-readable capability
  boundaries while keeping native rendering as visual authority.
- Added explicit semantic compiler refusal for native-only image and opaque body
  blocks instead of silently omitting them.
- Serialized OPC content-type and relationship control parts with default package
  namespaces for LibreOffice interoperability.
- Added structured fidelity policies and reports.
- Added package limits and defenses against traversal, ZIP bombs, unsafe XML, and macros.
- Advanced the AiOffice Document Spec to `0.2-draft.19`.

## 0.1.0

- Added the strict AiOffice Document Spec 1.0 draft.
- Added stable artifact and semantic node IDs.
- Added document creation through a declarative API and builder.
- Added JSON and Markdown input.
- Added JSON, Markdown, semantic HTML, and DOCX output.
- Added machine-readable document validation.
- Added atomic, revision-checked patch operations.
- Added the `aioffice` CLI.
