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

The first supported vertical slice is one embedded, inline DrawingML picture in an
otherwise empty body paragraph:

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

`asset_id` is derived from the full lowercase SHA-256 digest. Repeated occurrences
of identical bytes can therefore share one asset record while retaining separate,
stable image node IDs and native paragraph references.

The model stores no base64, relationship target, local path, URL, or arbitrary read
location. `metadata.native_part_uri` is inspection evidence only; it is never trusted
as a filesystem or package read request.

`editable: false` means the image binary and full DrawingML object are not represented
as generally editable JSON. On an attached native DOCX, compact inspection separately
advertises
`supported_operations: ["image.insert_after", "image.replace", "image.update",
"node.remove"]`. This keeps the lossless boundary explicit while still exposing the
small set of native mutations that AiOffice can prove safe.

## Conservative projection proof

An image becomes an `image` block only when all of these conditions hold:

1. the paragraph contains only paragraph properties and runs;
2. the runs contain only run properties and exactly one `w:drawing`;
3. the drawing contains one `wp:inline`, not `wp:anchor`;
4. `wp:extent` has positive `cx` and `cy` values;
5. the graphic data contains one DrawingML picture;
6. one `a:blip` uses `r:embed`, with no `r:link`;
7. that relationship is one internal image relationship from the containing part;
8. the target exists and has an `image/*` OPC content type;
9. the picture uses one rectangular stretch fill;
10. the picture has no crop, rotation, flip, visible outline, non-zero effect
    extent, or recognized visual effect.

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
6. compares native identity, asset ID, declared size, and SHA-256;
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

`image.update` changes only the supported inline picture's accessibility metadata
and/or displayed extent:

```python
result = document.apply([
    {
        "op": "image.update",
        "target": "#image_3A17C04E",
        "set": {
            "width": {"value": 3, "unit": "in"},
            "alt_text": "Expert workflow with three approval stages",
            "title": "Expert workflow",
        },
    }
])
assert result.success
```

The operation accepts `width`, `height`, `alt_text`, and `title` in `set`.
`alt_text` and `title` are the only clearable fields:

```json
{
  "op": "image.update",
  "target": "#image_3A17C04E",
  "clear": ["alt_text", "title"]
}
```

Widths and heights must convert to a positive signed 64-bit EMU value. Setting one
dimension preserves the current aspect ratio; setting both dimensions uses the exact
requested size. AiOffice writes the final EMU values to both
`wp:inline/wp:extent` and `pic:spPr/a:xfrm/a:ext`. The latter uses DrawingML's
[`PositiveSize2DType`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.positivesize2dtype?view=openxml-2.20.0);
the former follows the Wordprocessing Drawing
[`Extent`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.extent?view=openxml-3.0.1)
contract. A picture is projected as editable by this operation only when both native
extent records exist once and agree before the patch.

The native lowering re-proves the conservative image shape before and after mutation.
It does not decode, resample, replace, or recompress the image, and it does not change
the OPC image part or relationship. The asset ID, filename, media type, byte count,
SHA-256, and image occurrence ID therefore remain stable.

The operation fails atomically when the target is not a supported image, the request
is empty or malformed, a dimension is cleared, metadata is blank or invalid XML text,
the geometry is unsafe, or the native package is detached. A detached JSON snapshot
may still be inspected but cannot authorize a native DrawingML mutation.

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

The image occurrence ID, displayed width/height, alternative text, title, paragraph
formatting, and surrounding layout remain stable. The asset ID, media type, native
filename, byte count, and SHA-256 change to describe the replacement. No decode,
resample, recompression, automatic resizing, or orphan cleanup occurs.

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
- floating or anchored drawings;
- multiple pictures or alternate representations;
- linked or external images;
- crops, rotations, flips, outlines, or effects;
- drawings inside tables, headers, or footers;
- VML pictures, OLE objects, embedded files, charts, SmartArt, and other graphic
  data types.

The boundary is intentionally based on what the semantic layer can prove, not what it
can approximately display. Unrelated edits preserve every original package part and
unknown XML. Deleting a top-level projected image deletes its mapped paragraph only;
orphan cleanup, insertion outside the proven top-level inline subset, and replacement
outside the proven inline subset are not claimed in this release.

## Preview and visual authority

Semantic HTML emits an accessible, dimensioned placeholder with the asset ID and
media type. Markdown emits an `aioffice-asset:` reference. Neither exporter embeds
binary data or claims to reproduce the picture.

Use the native LibreOffice PDF/PNG provider to judge the actual picture, cropping,
position, pagination, and surrounding layout. Native rendering remains the visual
authority even for the supported inline subset.
