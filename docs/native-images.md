# Native image and asset contract

AiOffice separates image semantics from image storage:

- the JSON Spec describes an image occurrence and a content-addressed asset;
- the original image bytes remain in the native OPC package;
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
orphan cleanup and image mutation are not claimed in this release.

## Preview and visual authority

Semantic HTML emits an accessible, dimensioned placeholder with the asset ID and
media type. Markdown emits an `aioffice-asset:` reference. Neither exporter embeds
binary data or claims to reproduce the picture.

Use the native LibreOffice PDF/PNG provider to judge the actual picture, cropping,
position, pagination, and surrounding layout. Native rendering remains the visual
authority even for the supported inline subset.
