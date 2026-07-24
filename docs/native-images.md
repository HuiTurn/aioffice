# Native image and asset contract

AiOffice separates image semantics from image storage:

- the JSON Spec describes an image occurrence and a content-addressed asset;
- image bytes remain in the native OPC package rather than the JSON projection;
- native OOXML remains the authority for exact layout and rendering;
- bytes are returned only through a verified, explicit read path.

This avoids two opposite failures: base64 payloads consuming an AI model's context,
and a simplified JSON model pretending it can losslessly reconstruct every Word
drawing.

## Projected semantic shape

The supported vertical slice is one embedded DrawingML picture in an otherwise
empty body, header, or footer paragraph. Placement may be inline or one conservative
floating offset anchor with square text wrapping. A body occurrence lives in
`content`; a reusable story occurrence lives in its `header_footers[].content`:

```json
{
  "content": [
    {
      "id": "image_3A17C04E",
      "type": "image",
      "asset_id": "asset_8c5f5d1ccabbe00dcfeab72c4c718c7dfb9f71c6eaa4ec1f36b78356d9f1e19a",
      "placement": "inline",
      "width": {"value": 144, "unit": "pt"},
      "height": {"value": 72, "unit": "pt"},
      "crop": {
        "left": 12.5,
        "top": 5,
        "right": 12.5,
        "bottom": 5
      },
      "name": "Expert diagram",
      "alt_text": "A compact expert workflow diagram",
      "title": "Workflow",
      "capabilities": ["inspect", "extract", "delete", "render"],
      "editable": false
    }
  ],
  "assets": [
    {
      "id": "asset_8c5f5d1ccabbe00dcfeab72c4c718c7dfb9f71c6eaa4ec1f36b78356d9f1e19a",
      "sha256": "8c5f5d1ccabbe00dcfeab72c4c718c7dfb9f71c6eaa4ec1f36b78356d9f1e19a",
      "media_type": "image/png",
      "filename": "image1.png",
      "size_bytes": 21743
    }
  ]
}
```

A floating occurrence changes `placement` and adds explicit native layout evidence:

```json
{
  "placement": "floating",
  "floating": {
    "horizontal": {
      "relative_to": "column",
      "offset": {"value": 36, "unit": "pt"}
    },
    "vertical": {
      "relative_to": "paragraph",
      "offset": {"value": 0.4, "unit": "pt"}
    },
    "wrap": {
      "mode": "square",
      "side": "both_sides",
      "distance_top": {"value": 0, "unit": "pt"},
      "distance_right": {"value": 4, "unit": "pt"},
      "distance_bottom": {"value": 0, "unit": "pt"},
      "distance_left": {"value": 4, "unit": "pt"}
    },
    "relative_height": 1026,
    "behind_text": false,
    "locked": false,
    "layout_in_cell": true,
    "allow_overlap": true
  }
}
```

`asset_id` is derived from the full lowercase SHA-256 digest. Repeated occurrences
of identical bytes can therefore share one asset record while retaining separate,
stable image node IDs and native paragraph references.

The model stores no base64, relationship target, local path, URL, or arbitrary read
location. `metadata.native_part_uri` is inspection evidence only; it is never trusted
as a filesystem or package read request.

`editable: false` means the image binary and full DrawingML object are not represented
as generally editable JSON. On an attached native DOCX, compact inspection separately
advertises the operations proven safe for that story. Body images expose insertion
after the occurrence, replacement, update, paragraph formatting, and removal.
Header/footer images expose replacement, update, and paragraph formatting; direct
story-local insertion and deletion are not claimed. This keeps the lossless boundary
explicit while still exposing the small set of native mutations AiOffice can prove.

`capabilities()["assets"]["projected_story_scopes"]` reports
`["document_body", "header_footer"]`. The header/footer contract independently
reports `simple_native_image`, the three safe operations, and
`occurrence_copy_on_write` replacement so an AI does not have to infer scope from
examples. `aioffice schema --kind image-block` describes the body occurrence;
`aioffice schema --kind header-footer-image-block` describes the story-scoped
occurrence and fixes its capability list to `inspect`, `extract`, and `render`.

## Conservative projection proof

An image becomes an `image` block only when all of these conditions hold:

1. the paragraph contains only paragraph properties and runs;
2. the runs contain only run properties and exactly one `w:drawing`;
3. the drawing contains one supported `wp:inline` or `wp:anchor`;
4. `wp:extent` has positive `cx` and `cy` values;
5. the graphic data contains one DrawingML picture;
6. one `a:blip` uses `r:embed`, with no `r:link`;
7. that relationship is one internal image relationship from the containing body,
   header, or footer part;
8. the target exists and has an `image/*` OPC content type;
9. the picture uses one rectangular stretch fill, optionally preceded by one
   non-negative, non-overconstrained `a:srcRect`;
10. the picture has no rotation, flip, visible outline, non-zero effect extent, or
    recognized visual effect.

For `wp:anchor`, the proof additionally requires `simplePos="0"` with zero simple
coordinates, one horizontal and vertical `wp:posOffset`, recognized `relativeFrom`
values, four non-negative text distances, one `wp:wrapSquare` and recognized wrap
side, bounded `relativeHeight`, and strict boolean values for `behindDoc`, `locked`,
`layoutInCell`, and `allowOverlap`. Child order must match the native anchor schema.

Microsoft's Open XML contracts distinguish
[`wp:inline`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.inline?view=openxml-3.0.1)
from floating placement, define
[`wp:extent`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.extent?view=openxml-3.0.1)
in English Metric Units, and expose name/title/description through
[`wp:docPr`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.docproperties?view=openxml-3.0.1).
The standard insertion pattern connects `a:blip/@r:embed` to an image part through a
package relationship; see Microsoft's
[picture insertion example](https://learn.microsoft.com/zh-cn/office/open-xml/word/how-to-insert-a-picture-into-a-word-processing-document).

AiOffice converts EMUs to points using 12,700 EMUs per point. It does not infer DPI,
resample bytes, decode the bitmap, or guess missing dimensions.

## Floating anchor projection and update

`FloatingImageLayout` is a normalized, AI-readable view of the supported native
anchor. Horizontal and vertical positions deliberately keep both the reference frame
and signed physical offset; an offset without its frame is not meaningful. Square
wrapping keeps its side and all four distances from text. Native flags remain
separate rather than being collapsed into a vague “floating” boolean.

The accepted horizontal frames are `character`, `column`, `inside_margin`,
`left_margin`, `margin`, `outside_margin`, `page`, and `right_margin`. The accepted
vertical frames are `bottom_margin`, `inside_margin`, `line`, `margin`,
`outside_margin`, `page`, `paragraph`, and `top_margin`. Every offset and distance
has an explicit unit. `relative_height` preserves Word's unsigned relative stacking
value; `behind_text` is kept independently because it materially changes the layer
group.

This model follows Wordprocessing Drawing's
[`wp:anchor`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.anchor?view=openxml-3.0.1)
container and its distinct position and wrap children. Existing verified image
operations may extract, resize, crop, change accessibility metadata, format the host
paragraph, or replace the binary; native lowering proves that the complete anchor
layout remains identical afterward.

In dev35, `image.anchor.update` selectively changes that projected layout:

```json
{
  "op": "image.anchor.update",
  "target": "#image_3A17C04E",
  "set": {
    "horizontal": {
      "relative_to": "page",
      "offset": {"value": 72, "unit": "pt"}
    },
    "vertical": {
      "relative_to": "margin",
      "offset": {"value": -18, "unit": "pt"}
    },
    "wrap": {
      "mode": "square",
      "side": "right",
      "distance_top": {"value": 5, "unit": "pt"},
      "distance_right": {"value": 6, "unit": "pt"},
      "distance_bottom": {"value": 7, "unit": "pt"},
      "distance_left": {"value": 8, "unit": "pt"}
    },
    "relative_height": 2048,
    "behind_text": true,
    "locked": false,
    "layout_in_cell": false,
    "allow_overlap": false
  }
}
```

Every top-level field is optional, but `set` must contain at least one non-null
change. `horizontal`, `vertical`, and `wrap` are complete grouped values: callers
must keep the reference frame with its offset and the wrap side with all four
distances. This avoids partial native geometry whose meaning depends on hidden
state. No anchor field is clearable. Signed offsets must fit OOXML Int64 EMUs;
distances and `relative_height` must fit UInt32; booleans are strict JSON booleans.
`aioffice schema --kind floating-image-layout-update` exposes the same contract.

Native lowering changes only the selected `wp:positionH`, `wp:positionV`,
`wp:wrapSquare`, and `wp:anchor` attributes in the already-proven tree. It preserves
the drawing, extents, crop, image relationship and bytes, accessibility metadata,
host paragraph, unknown surrounding XML, and present `wp14:anchorId` /
`wp14:editId`. The operation re-projects the result and fails the complete Patch if
any requested semantic value does not round-trip exactly. It composes with
`image.update` in either order inside one atomic Patch.

Physical lengths are compared by their rounded native EMU value, not by unit
spelling. For example, an accepted `1 in` offset reopens as the canonical native
projection `72 pt` without becoming a semantic or integrity mismatch.

An inline image, detached JSON projection, opaque or stale drawing, unsupported
anchor variant, empty update, partial group, null, unknown field, invalid unit/range,
or non-boolean flag fails closed. The operation does not convert an inline picture
to floating and does not create or reconstruct an anchor.

Alignment-based positioning, active simple positioning, `wrapNone`,
`wrapTopAndBottom`, tight/through polygons, relative-size extensions, missing or
unknown compatibility values, extra children, and mixed text plus drawing paragraphs
stay opaque. This is a deliberate protocol boundary, not an indication that their
native XML is discarded.

## Verified binary access

Use the projected image node ID, never a part path:

```python
image = document.read_image("image_3A17C04E")
payload = document.image_bytes("image_3A17C04E")
document.extract_image(
    "image_3A17C04E",
    "review/image.png",
    overwrite=False,
)
```

`read_image` performs the proof again against the attached native package. It:

1. resolves the image node's trusted `NativeRef`;
2. loads that exact native paragraph;
3. rechecks the conservative DrawingML shape;
4. resolves the embedded relationship from the correct source part;
5. confirms the image target and OPC media type;
6. compares native identity, asset ID, placement, floating layout, displayed extent,
   crop, accessibility metadata, declared size, and SHA-256;
7. hashes the returned bytes again.

Any stale, forged, missing, external, or structurally changed reference fails closed.
An exported JSON snapshot can describe an image, but it cannot return the binary
after the native package is detached.

The CLI follows the same path:

```bash
aioffice extract-image report.docx IMAGE_ID -o image.png
```

It refuses to overwrite by default and reports the image ID, asset ID, media type,
filename, size, SHA-256, and output path as JSON.

## Selective native updates

`image.update` changes only the supported native picture's accessibility metadata,
displayed extent, and/or rectangular source crop:

```python
result = document.apply([
    {
        "op": "image.update",
        "target": "#image_3A17C04E",
        "set": {
            "width": {"value": 3, "unit": "in"},
            "crop": {
                "left": 12.5,
                "top": 5,
                "right": 12.5,
                "bottom": 5
            },
            "alt_text": "Expert workflow with three approval stages",
            "title": "Expert workflow",
        },
    }
])
assert result.success
```

The operation accepts `width`, `height`, `crop`, `alt_text`, and `title` in `set`.
`crop`, `alt_text`, and `title` are clearable:

```json
{
  "op": "image.update",
  "target": "#image_3A17C04E",
  "clear": ["crop", "alt_text", "title"]
}
```

Widths and heights must convert to a positive signed 64-bit EMU value. Setting one
dimension preserves the current aspect ratio; setting both dimensions uses the exact
requested size. AiOffice writes the final EMU values to both
the active `wp:inline` or `wp:anchor` extent and `pic:spPr/a:xfrm/a:ext`. The latter
uses DrawingML's
[`PositiveSize2DType`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.positivesize2dtype?view=openxml-2.20.0);
the former follows the Wordprocessing Drawing
[`Extent`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.extent?view=openxml-3.0.1)
contract. A picture is projected as editable by this operation only when both native
extent records exist once and agree before the patch.

### Rectangular source crop

`crop.left`, `crop.top`, `crop.right`, and `crop.bottom` are percentage points of the
original source image. Each edge is in `[0, 100)`, left plus right must be below
`100`, and top plus bottom must be below `100`, so some source area always remains
visible. At least one edge must be non-zero; use `clear` to remove the crop.

AiOffice rounds every edge to three decimal places, matching DrawingML's integer
thousandths of one percent. For example, `12.3454` becomes the native value `12345`
and reopens as `12.345`. Setting `crop` replaces the complete rectangle; an omitted
edge defaults to zero rather than inheriting its previous value:

```json
{
  "op": "image.update",
  "target": "#image_3A17C04E",
  "set": {
    "crop": {"left": 15, "right": 15}
  }
}
```

The native lowering creates or minimally updates `pic:blipFill/a:srcRect`, before
the existing `a:stretch`, and emits only non-zero `l`, `t`, `r`, or `b` attributes.
Clearing `crop` removes only that element. It never rewrites the raster, changes the
displayed extent, or adjusts the relationship. Binary replacement and header/footer
cloning preserve the occurrence's current crop.

Compact inspection reports the normalized crop and advertises
`crop_unit: "percentage_points"`, `crop_precision: 0.001`, and
`crop_visible_area_required: true`. Semantic HTML exposes the four normalized edges
as `data-aioffice-crop-*` evidence on its placeholder, but does not simulate native
cropping. DOCX rendering remains the visual authority.

An empty or all-zero native `a:srcRect` has no semantic crop and is preserved by
unrelated edits. Negative values, values outside the bounded range, overconstrained
edge sums, duplicate crop elements, unknown attributes, alternate fill structures,
or malformed child order fail the conservative proof and remain opaque.

The native lowering re-proves the conservative image shape before and after mutation.
It does not decode, resample, replace, or recompress the image, and it does not change
the OPC image part or relationship. The asset ID, filename, media type, byte count,
SHA-256, and image occurrence ID therefore remain stable.

The operation fails atomically when the target is not a supported image, the request
is empty or malformed, a dimension is cleared, metadata is blank or invalid XML text,
the geometry is unsafe, or the native package is detached. A detached JSON snapshot
may still be inspected but cannot authorize a native DrawingML mutation.

## Image paragraph layout

A projected image is one picture occurrence hosted by one native paragraph. Its
stable image ID therefore also provides a safe target for `paragraph.format`:

```python
result = document.apply([
    {
        "op": "paragraph.format",
        "target": "#image_3A17C04E",
        "set": {
            "alignment": "center",
            "spacing_before": {"value": 10, "unit": "pt"},
            "spacing_after": {"value": 12, "unit": "pt"},
            "indent_left": {"value": 18, "unit": "pt"},
            "indent_right": {"value": 18, "unit": "pt"},
            "keep_together": True,
        },
    }
])
assert result.success
```

This is the same `ParagraphStyle` and selective `set`/`clear` contract used for an
ordinary paragraph. Supported fields cover alignment, solid sRGB background, four
physical borders, before/after and line spacing, left/right/first-line/hanging
indentation, keep-with-next, keep-together, page-break-before, widow control, and
outline level. Every length keeps an explicit unit.

The native lowering resolves the image's trusted paragraph reference in the body,
header, or footer part and mutates only the requested supported `w:pPr` properties.
It does not alter `w:drawing`,
`wp:inline`, either DrawingML extent, `a:blip`, the image relationship, or the image
part. Clearing a field removes only its supported direct native value so normal Word
style inheritance can apply again. Unknown paragraph-property XML is retained.

For a floating image whose position is relative to its host paragraph, changing that
paragraph's spacing, indentation, or pagination can still move the rendered anchor
indirectly according to Word's layout rules. The anchor XML and projected
`FloatingImageLayout` remain unchanged; use native rendering to verify the resulting
page composition.

`text.format` remains invalid for an image because the conservative image paragraph
has no model-editable text. Complex or opaque drawings do not gain this capability.
Invalid styles, detached packages, stale identities, or non-paragraph native targets
fail the complete Patch atomically.

The operation uses ordinary model-authored JSON, so the existing
`aioffice apply` and `aioffice workspace apply` commands require no special binary
channel. Capability metadata exposes `native_layout_operation`,
`native_layout_fields`, and `native_layout_target` for host-paragraph planning.
Floating planning separately exposes `floating_layout_update_operation`,
`floating_layout_update_fields`, group-replacement semantics, and an empty
clearable-field list.

## Out-of-band binary replacement

Binary replacement deliberately bypasses JSON:

```python
result = document.replace_image(
    "#image_3A17C04E",
    "generated/expert-workflow.png",
    media_type="image/png",
)
assert result.success
result.document.export("replaced.docx")
```

`source` may be a local path, bytes-like value, or a previously verified
`ImageAsset`. The optional declared media type must match the detected signature.
The first safe subset accepts PNG, JPEG, GIF, BMP, and TIFF. SVG, metafiles, icons,
unknown image types, empty payloads, and inputs above the active security policy's
file-size limit fail before native mutation. AiOffice performs bounded signature
validation; it does not claim that this replaces native rendering or a full image
decoder's security and compatibility checks.

The replacement receives a full-SHA-256 asset ID and canonical native filename:

```text
asset_<64 lowercase hex characters>
/word/media/aioffice-<64 lowercase hex characters>.png
```

The native lowering follows occurrence-scoped copy-on-write:

1. re-prove the target's conservative inline-picture shape;
2. add or reuse the exact content-addressed image part;
3. add a matching `[Content_Types].xml` override;
4. add a new internal image relationship with a collision-free ID;
5. change only the target occurrence's `a:blip/@r:embed`;
6. re-prove relationship, content type, hash, size, and projection shape;
7. refresh the semantic asset record and native identity manifest.

This mirrors the Open XML package model in which pictures are backed by
[`ImagePart`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.packaging.imagepart?view=openxml-3.0.1)
objects connected from a containing part. Microsoft's
[picture insertion example](https://learn.microsoft.com/en-us/office/open-xml/word/how-to-insert-a-picture-into-a-word-processing-document)
likewise adds an image part and refers to it through a relationship ID.

Always allocating a relationship for the target occurrence is essential. Two
`a:blip` elements can share the same relationship and image part; retargeting that
relationship would silently replace both pictures. AiOffice instead changes only the
selected blip. The original relationship and image part remain untouched, so other
occurrences and unknown native consumers are preserved.

The relationship is allocated in the occurrence's own story. Replacing a logo in a
cloned header therefore adds a relationship to the cloned header's `.rels`, changes
only its `a:blip`, and leaves the source header relationship and shared source media
unchanged. This is the intended copy-on-write workflow:

1. clone the reusable header/footer and bind it to the selected section;
2. reopen or use the returned document to obtain the cloned image ID;
3. call `replace_image()` for that ID in a subsequent transaction;
4. reopen and render affected pages through the native provider.

The image occurrence ID, placement, complete floating anchor layout, displayed
width/height, source crop, alternative text, title, paragraph formatting, and
surrounding layout remain stable. The asset ID, media type, native filename, byte
count, and SHA-256 change to describe the replacement. No decode, resample,
recompression, automatic resizing, or orphan cleanup occurs.

The CLI uses the same binary channel and refuses output overwrite by default:

```bash
aioffice replace-image \
  existing.docx IMAGE_ID replacement.jpg \
  --media-type image/jpeg \
  -o replaced.docx
```

Tracked documents use the same boundary while committing a new workspace revision:

```python
result = workspace.replace_image(
    document.id,
    image_id,
    "replacement.jpg",
    media_type="image/jpeg",
    base_revision=document.revision,
)
```

```bash
aioffice workspace replace-image \
  ARTIFACT_ID IMAGE_ID replacement.jpg \
  --root project \
  --media-type image/jpeg
```

The workspace patch log stores the verified `AssetRef`, transport label, semantic
change, fidelity report, and diff. It does not duplicate the binary or encode it as
base64; the committed native DOCX revision remains the binary authority.

A raw `image.replace` JSON Patch fails with `BINARY_ASSET_REQUIRED`; accepting a path,
URL, or base64 field inside model-authored JSON would cross the trust boundary and
make patch replay ambiguous. `Document.replace_image()` and the CLI bind the verified
bytes to the metadata operation inside one atomic in-memory transaction.

## Addressable inline image insertion

Insertion uses a dedicated binary API because neither a local path nor binary payload
belongs in model-authored JSON:

```python
result = document.insert_image_after(
    "#analysis",
    "generated/expert-workflow.png",
    width={"value": 3, "unit": "in"},
    height={"value": 1.5, "unit": "in"},
    alt_text="Expert workflow with three approval stages",
    image_id="expert_workflow",
    name="Expert workflow",
    title="Approval workflow",
    paragraph_style={
        "alignment": "center",
        "spacing_before": {"value": 6, "unit": "pt"},
        "spacing_after": {"value": 8, "unit": "pt"},
    },
)
assert result.success
result.document.export("inserted.docx")
```

The contract is deliberately explicit:

- `target` must resolve to one top-level semantic node mapped into
  `/word/document.xml`;
- the picture is inserted after the target's last native body element, so a
  multi-paragraph list remains one valid anchor;
- `width` and `height` are both required and must fit positive signed 64-bit EMUs;
- `alt_text` is required and cannot be blank;
- `image_id` is optional but, when supplied, must be globally unique;
- `paragraph_style` may set direct paragraph alignment, spacing, indentation,
  background, and supported borders around the inline picture paragraph;
- placement is `inline`; floating anchors, wrap modes, crop, rotation and effects are
  not silently inferred.

The native lowering adds or reuses the content-addressed image part, creates a fresh
relationship ID, creates a `w:p/w:r/w:drawing/wp:inline` tree, writes identical outer
and inner extents, generates collision-free `wp:docPr/@id` and `w14:paraId` values,
applies the requested paragraph style, inserts at the proven body position, and then
re-runs the conservative projection proof. The operation either returns a fully
readable `ImageBlock` or leaves the original document unchanged.

If a third-party DOCX has no AiOffice identity manifest, the first successful
structural insertion also attaches:

- `/customXml/aioffice-manifest.xml`;
- one root package relationship with the AiOffice manifest relationship type;
- an `application/xml` content-type override.

That metadata is necessary to restore a caller-selected image ID after standalone
export and reopen. Existing native IDs are retained; Workspace documents continue to
keep their external revision identity evidence as well.

The CLI exposes the same required geometry and accessibility inputs:

```bash
aioffice insert-image-after \
  existing.docx TARGET replacement.png \
  --width 3 --width-unit in \
  --height 1.5 --height-unit in \
  --alt-text "Expert workflow with three approval stages" \
  --image-id expert_workflow \
  --align center \
  -o inserted.docx
```

Tracked insertion creates one new workspace revision without storing binary data in
the patch log:

```python
result = workspace.insert_image_after(
    document.id,
    "#analysis",
    "replacement.png",
    width={"value": 3, "unit": "in"},
    height={"value": 1.5, "unit": "in"},
    alt_text="Expert workflow",
    base_revision=document.revision,
)
```

The equivalent command is `aioffice workspace insert-image-after` with the same
width, height and alternative-text flags plus `--root`.

A raw `image.insert_after` JSON operation fails with `BINARY_ASSET_REQUIRED`. This
keeps asset acquisition, signature validation and package mutation in one trusted
transaction while leaving the operation metadata explainable to an AI.

## Deliberate opaque boundary

These cases remain native and explicit `opaque`/read-only projections:

- text plus a drawing in one paragraph;
- alignment-based, active-simple-position, non-square-wrap, relative-size,
  malformed, or otherwise unsupported floating anchors;
- multiple pictures or alternate representations;
- linked or external images;
- negative, overconstrained, malformed, or otherwise unsupported crop rectangles;
- rotations, flips, outlines, or effects;
- drawings inside tables;
- complex header/footer drawings that do not satisfy the same one-picture proof;
- VML pictures, OLE objects, embedded files, charts, SmartArt, and other graphic
  data types.

The boundary is intentionally based on what the semantic layer can prove, not what it
can approximately display. Unrelated edits preserve every original package part and
unknown XML. Deleting a top-level body image deletes its mapped paragraph only;
header/footer image deletion, orphan cleanup, insertion outside the proven top-level
body subset, and replacement outside the proven inline-or-conservative-floating
subset are not claimed in this release.

## Preview and visual authority

Semantic HTML emits an accessible, dimensioned placeholder with the asset ID, media
type, placement, and normalized crop evidence. Markdown emits an `aioffice-asset:`
reference. Neither exporter embeds binary data or claims to reproduce native
floating placement, wrapping, crop, or picture content.

Use the native LibreOffice PDF/PNG provider to judge the actual picture, cropping,
position, pagination, and surrounding layout. Native rendering remains the visual
authority even for the supported inline and floating subsets.
