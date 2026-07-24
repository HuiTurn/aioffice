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
- Added the symmetric stable-ID `node.move_before` operation so AI callers can reach
  the first or last position without computing array indices.
- Preserved exact native target and anchor ranges for before/after placement across
  sequential mixed-direction moves, multi-paragraph lists, tables, and images.
- Added safe section-prepend semantics: the native section carrier stays fixed while
  `section.start_at` is rebound to the moved node with explicit change evidence.
- Kept detached native projections fail-closed and made capabilities omit both
  structural operations when their XML authority package is unavailable.
- Upgraded `node.remove` to a proven native structural edit: third-party packages
  now receive an identity manifest on first removal so surviving stable IDs persist
  after their native indices change.
- Added atomic refusal for removal ranges carrying `w:sectPr`, preventing deletion
  from silently collapsing Word section layout, header/footer, or numbering state.
- Made detached native-authority projections omit and reject `node.remove` instead
  of rebuilding a lossy DOCX from JSON after deletion.
- Defined a conservative removal orphan policy that preserves unreferenced
  relationships and package parts until future graph-wide garbage collection can
  prove they have no unknown consumers.
- Corrected semantic Diff ordering evidence so pure additions or removals no longer
  masquerade as moves; `moved` now requires a relative-order change among identities
  present in both revisions.
- Added incremental native `node.insert_after` lowering for paragraphs and headings:
  only the new `w:p` is compiled while every existing native XML element remains
  unreconstructed.
- Added rich inserted text spans, direct paragraph/text formatting, named heading
  styles, internal/external hyperlinks, and normalized dynamic fields with fresh,
  collision-safe native identities and relationship IDs.
- Made nodes created earlier in one Patch immediately addressable by later insertion,
  text, formatting, move, and removal operations through live XML object tracking
  rather than stale indices.
- Added fail-closed top-level, contiguity, section-boundary, style-existence,
  source-reference, safe-XML, and supported-block proofs for native text insertion.
- Attached persistent identity manifests on the first text insertion into third-party
  DOCX packages, refreshed shifted content/field/section/table references, and exposed
  the operation through capabilities, CLI apply, and Workspace persistence.
- Added native `node.append` lowering for imported DOCX so AI callers can add a
  paragraph or heading through the document-root selector without first discovering
  the final content ID.
- Appended content is inserted before the optional terminal body-level `w:sectPr`,
  keeping it in the final semantic section while preserving the original section
  properties byte-for-byte; empty document bodies are supported.
- Added ordered batch tracking for appended nodes, terminal-section fail-closed
  validation, strict unknown-field validation, CLI/Workspace capability coverage,
  and exact preservation/reopen tests.
- Removed `node.append` from detached native-authority projections and reject it
  before semantic mutation when the authoritative DOCX package is unavailable.
- Added native `page_break` lowering through `node.append`, `node.insert_after`, and
  `node.insert_before`, compiling exactly one stable-ID `w:p/w:r/w:br` with
  `w:type="page"`.
- Page breaks now participate in live batch targeting, section-start rebinding,
  identity refresh, CLI apply, Workspace persistence, and standalone reopen without
  reconstructing any existing native element.
- Added exact XML tests for all three placement modes and rendered a real third-party
  DOCX to prove that the explicit break changes pagination while existing image,
  table, relationship, and section payloads remain unchanged.
- Added incremental native table lowering through `node.append`,
  `node.insert_after`, and `node.insert_before`, compiling only the new `w:tbl`
  while preserving every existing body element.
- Reused the semantic table compiler for regular and merged cells, rich cell
  paragraphs, table/row/cell geometry and formatting, borders, and repeated headers;
  inserted external links receive collision-safe native relationships while internal
  links remain relationship-free bookmarks.
- Assigned persistent native references to inserted tables, columns, rows, cells,
  and rich cell paragraphs, including IDs generated during semantic normalization.
  Those objects survive standalone reopen and are immediately addressable later in
  the same Patch.
- Enabled same-Patch table, column, cell, paragraph, and text formatting for inserted
  tables, plus use of the new table as an insertion, move, or removal anchor.
- Exposed machine-readable table-insertion capabilities covering native component
  kinds, supported and unsupported cell content, style policy, hyperlink lowering,
  same-Patch operations, generated-ID behavior, and the regular-grid width boundary.
- Added table-aware section-start rebinding, terminal-section root append,
  CLI/Workspace persistence, generated-ID reconciliation, and exact XML/relationship
  preservation tests for all three placement modes.
- Rendered and visually inspected a real two-section third-party DOCX after inserting
  a fixed-width rich table: page count and Letter geometry remained stable, the
  untouched first-page PNG stayed byte-identical, and existing image, table,
  relationship, section, and body-element evidence remained exact.
- Added atomic refusal for recursively forged native references, missing table or
  cell-paragraph styles, invalid table grids, unsafe XML, and column-width edits on
  merged grids.
- Added native `bullet_list` and `ordered_list` lowering through `node.append`,
  `node.insert_after`, and `node.insert_before`, compiling one contiguous `w:p`
  range with one stable root identity and one paragraph per plain-text item.
- Reused one list compiler for generated and imported DOCX paths while assigning
  collision-safe `w14:paraId` anchors to every newly inserted item.
- Added one independent single-level `abstractNum` and `num` definition per inserted
  list, giving ordered lists deterministic restart-at-one behavior and preventing
  accidental continuation of adjacent native lists.
- Kept every numbering part in strict OOXML child order with all abstract
  definitions before concrete numbering instances, and used Word-compatible
  Symbol-font bullet glyphs for consistent visual weight.
- Added safe creation of a missing numbering part, canonical document relationship,
  and OPC content-type override while preserving every existing numbering child and
  unrelated package part.
- Made inserted lists immediately usable as before/after anchors, contiguous move
  targets, and removal targets in the same Patch, including later-section
  `start_at` rebinding and terminal-section root append.
- Exposed machine-readable list-insertion capabilities and added CLI, Workspace,
  standalone-reopen, all-position, numbering-isolation, exact-preservation, and
  transaction-atomicity coverage.
- Added atomic refusal for forged native references, unsafe item text, malformed or
  duplicate numbering IDs, and conflicting numbering relationships or content
  types.
- Rendered and visually inspected a real two-section third-party DOCX with two
  independent numbered lists and one bullet list: numbering restarted correctly,
  hanging-indent wrapping and bullet weight were clean, page count and Letter
  geometry remained stable, and the untouched first-page PNG stayed byte-identical.
- Added semantic and native `section.insert_before` so an AI can split the section
  containing any non-leading top-level content node without addressing a body
  position or reconstructing existing content.
- Made the target the new section's stable `start_at`, inherited all unselected
  layout values and semantic header/footer bindings, and defaulted an omitted
  section-start type to `next_page`.
- Lowered imported section insertion through one collision-safe hidden `w:p`
  boundary carrying an exact copy of the old `w:sectPr`, while retaining the
  original boundary for the new section and patching only requested layout fields.
- Preserved unknown section XML and relationship references on both sides of the
  split, including multi-paragraph list targets, while refreshing all content,
  section, table, field, and header/footer identities.
- Added live same-Patch section state so newly created sections can receive
  `section.format`, and content inserted before their first node rebinds `start_at`
  against the newly created native boundary.
- Added CLI and Workspace persistence plus machine-readable insertion strategy,
  inheritance, default-break, native-scope, same-Patch, and unsupported-case
  capabilities.
- Added atomic refusal for empty-section requests, detached native projections,
  direct header/footer rebinding, tracked section properties, stale/noncontiguous
  targets, forged semantic evidence, and unsafe boundary placement.
- Verified a real two-section, third-party DOCX by splitting its second section
  before an existing table, prepending a heading and paragraph in the same Patch,
  and rendering two Letter portrait pages followed by one Letter landscape page:
  all pre-existing content XML and the copied old `w:sectPr` stayed byte-exact,
  the untouched first-page PNG stayed byte-identical, and the wide page rendered
  without clipping or overlap.
- Added semantic and native `section.header_footer.bind` for explicitly reusing an
  existing header/footer part in any default, first, or even section slot, or
  clearing a slot back to Word's native inheritance behavior.
- Lowered binding changes to only the selected direct `w:headerReference` or
  `w:footerReference`, preserving reusable part XML, document relationships,
  unselected references, unknown section properties, and exact OOXML child order.
- Added strict proof of native part kind, root element, internal relationship type
  and uniqueness, existing projected references, live section identity, sequential
  binding state, and final semantic/native binding-map equality.
- Made binding work on sections created earlier in the same Patch and made later
  section splits inherit the current explicit binding state.
- Ensured standalone identity attachment for third-party binding changes and added
  CLI, Workspace, capabilities, generated/native round-trip, exact-part
  preservation, same-Patch, kind-mismatch, duplicate-reference, detached-projection,
  and transaction-atomicity coverage.
- Projected safe internal header/footer relationships even when their part is
  temporarily unbound, keeping preserved old regions AI-addressable for later reuse,
  and fixed deterministic native-ID allocation beyond a second identical fallback
  collision.
- Verified a manifest-free, independently generated three-section DOCX by replacing
  the appendix header with the body header and clearing its footer to inherit: all
  six region parts, `document.xml.rels`, existing body XML, and unknown section XML
  remained exact; standalone reopen retained even the unbound old parts; three
  Letter pages rendered without clipping, while the untouched first two page PNGs
  stayed byte-identical and only the appendix header/footer changed visually.
- Added symmetric native `node.insert_before` so AI callers can prepend content to
  the document or place a new paragraph/heading before any complete mapped range,
  including multi-paragraph list anchors.
- Added ordered section-start state tracking across mixed insert and move operations.
  Inserting before a later section's first semantic node now rebinds `section.start_at`
  to the created node while proving it remains after the preceding native `w:sectPr`.
- Allowed safe insertion before a text-bearing paragraph that carries `w:sectPr`
  because the boundary remains after the unchanged anchor; insertion after that same
  anchor remains fail-closed.
- Added semantic/native cross-checks for section rebind evidence, sequential section
  prepends, document-head insertion, exact preservation of existing multi-element
  ranges, detached projections, CLI application, and Workspace persistence.
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
- Advanced the AiOffice Document Spec to `0.2-draft.29`.

## 0.1.0

- Added the strict AiOffice Document Spec 1.0 draft.
- Added stable artifact and semantic node IDs.
- Added document creation through a declarative API and builder.
- Added JSON and Markdown input.
- Added JSON, Markdown, semantic HTML, and DOCX output.
- Added machine-readable document validation.
- Added atomic, revision-checked patch operations.
- Added the `aioffice` CLI.
