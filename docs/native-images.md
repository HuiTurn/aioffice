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
floating offset/alignment/percentage anchor with square, no-wrap, top-and-bottom, tight, or
through text wrapping. A body occurrence lives in `content`; a reusable story
occurrence lives in its `header_footers[].content`:

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
      "transform": {
        "rotation_degrees_clockwise": 90,
        "flip_horizontal": true,
        "flip_vertical": false
      },
      "outline": {
        "width": {"value": 2, "unit": "pt"},
        "color": "#CC0000",
        "dash": "dash_dot"
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
    "anchor_distances": {
      "top": {"value": 0, "unit": "pt"},
      "right": {"value": 4, "unit": "pt"},
      "bottom": {"value": 0, "unit": "pt"},
      "left": {"value": 4, "unit": "pt"}
    },
    "anchor_effect_extent": {
      "left": {"value": 0, "unit": "pt"},
      "top": {"value": 0, "unit": "pt"},
      "right": {"value": 0, "unit": "pt"},
      "bottom": {"value": 0, "unit": "pt"}
    },
    "wrap": {
      "mode": "square",
      "side": "both_sides"
    },
    "relative_height": 1026,
    "behind_text": false,
    "locked": false,
    "layout_in_cell": true,
    "allow_overlap": true
  }
}
```

No-wrap has no wrap-local fields. Top-and-bottom may carry its own optional top and
bottom distances and its own effect extent, independently of the parent anchor:

```json
{
  "wrap": {
    "mode": "top_and_bottom",
    "distances": {
      "top": {"value": 6, "unit": "pt"},
      "bottom": {"value": 6, "unit": "pt"}
    },
    "effect_extent": {
      "left": {"value": 0, "unit": "pt"},
      "top": {"value": 1, "unit": "pt"},
      "right": {"value": 0, "unit": "pt"},
      "bottom": {"value": 1, "unit": "pt"}
    }
  }
}
```

Tight and through wrapping require a side plus an ordered native polygon. Their
optional wrap-local distances are left/right only. Coordinates deliberately remain
raw signed OOXML integers: producers commonly use a normalized 0–21600 space, so
labeling them as EMUs or another physical `Length` would be incorrect. The last
point is not forced to equal the start because the native renderer infers the
closing edge when it is omitted:

```json
{
  "wrap": {
    "mode": "tight",
    "side": "both_sides",
    "distances": {
      "left": {"value": 3, "unit": "pt"},
      "right": {"value": 4, "unit": "pt"}
    },
    "polygon": {
      "edited": true,
      "start": {"x": 0, "y": 0},
      "line_to": [
        {"x": 0, "y": 21600},
        {"x": 21600, "y": 21600},
        {"x": 21600, "y": 0}
      ]
    }
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
10. `a:xfrm` has zero `a:off`, one `a:ext` matching the outer extent, and only
    valid signed-Int32 rotation plus strict horizontal/vertical flip attributes;
11. `pic:spPr` uses exact supported child order and has at most one direct picture
    line: either a neutral absent/zero-width `a:noFill` line or one positive
    `ST_LineWidth` line with direct `a:srgbClr`, a supported preset dash, and only
    the supported default cap, compound, alignment, and round-join semantics;
12. the picture has no recognized DrawingML visual effect; LibreOffice's optional
    neutral `pic:spPr/@bwMode="auto"` and at most one empty shape `a:noFill`
    normalization are allowed.

For `wp:anchor`, the proof additionally requires `simplePos="0"` with zero simple
coordinates; one horizontal and vertical position, each containing exactly one
recognized `wp:posOffset`, `wp:align`, or axis-correct Office 2010 percentage
child; recognized `relativeFrom` values; zero or more optional non-negative native
anchor distances; an optional strict parent
`wp:effectExtent`; exactly one empty `wp:wrapNone`, or a `wp:wrapTopAndBottom` /
`wp:wrapSquare` carrying only its schema-defined optional distances and at most one
strict child `wp:effectExtent`, or one `wp:wrapTight` / `wp:wrapThrough` with a
recognized side, optional left/right distances, and exactly one strict
`wp:wrapPolygon`. A polygon preserves optional boolean `edited`, exactly one first
`wp:start`, and 2–4096 ordered `wp:lineTo` points whose signed coordinates fit
OOXML `ST_Coordinate`. `relativeHeight` is bounded and `behindDoc`, `locked`,
`layoutInCell`, and `allowOverlap` are strict booleans. Child order must match the
native anchor schema. After `a:graphic`, the anchor may additionally contain one
strict `wp14:sizeRelH`, one strict `wp14:sizeRelV`, or both, in that order. Each
axis contains its axis-correct percentage child, an allowed reference frame, and a
non-negative signed-Int32 native percentage.

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
anchor. Horizontal and vertical positions deliberately keep the reference frame plus
exactly one positioning mode: a signed physical `offset`, semantic `alignment`, or
signed `percentage_offset` in percentage points of the selected frame.
No mode is meaningful without its frame. `wrap.mode` is `square`, `none`,
`top_and_bottom`, `tight`, or `through`; square, tight, and through wrapping have a
`side`. `anchor_distances` preserves
each optional `wp:anchor/@dist*` attribute without inventing missing values.
`anchor_effect_extent` preserves the optional parent element. Square wrap has four
optional local distances; top-and-bottom has optional top and bottom distances.
Both can carry a wrap-child `effect_extent`. These fields are deliberately not
collapsed: for square and top-and-bottom wrapping, a present child effect extent
defines that wrap boundary instead of the parent value. Native flags remain
separate rather than being collapsed into a vague “floating” boolean.
Tight and through have optional left/right distances and require a polygon instead
of a child effect extent. Their point order, optional `edited` attribute, and
explicit or inferred closure are preserved without normalizing producer coordinates.

The accepted horizontal frames are `character`, `column`, `inside_margin`,
`left_margin`, `margin`, `outside_margin`, `page`, and `right_margin`. The accepted
vertical frames are `bottom_margin`, `inside_margin`, `line`, `margin`,
`outside_margin`, `page`, `paragraph`, and `top_margin`. Horizontal alignments are
`left`, `right`, `center`, `inside`, and `outside`; vertical alignments are `top`,
`bottom`, `center`, `inside`, and `outside`. Every offset and distance has an
explicit unit. Percentage offsets use percentage points directly, have `0.001`
precision, and preserve the full signed Int32 `ST_Percentage` native range.
`relative_height` preserves Word's unsigned relative stacking value;
`behind_text` is kept independently because it materially changes the layer group.

This model follows Wordprocessing Drawing's
[`wp:anchor`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.anchor?view=openxml-3.0.1)
container and its distinct position and wrap children. Microsoft defines
[`wp:effectExtent`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.effectextent?view=openxml-3.0.1)
as the additional edge extents used for wrapping objects with drawing effects, while
[`wp:wrapSquare`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.wrapsquare?view=openxml-3.0.1)
and `wp:wrapTopAndBottom` may carry their own schema-defined boundary data.
[`wp:wrapNone`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.wrapnone?view=openxml-3.0.1)
causes no text reflow; `behind_text` then determines whether the picture is behind
or in front of document text.
[`wp:wrapTopAndBottom`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.wraptopbottom?view=openxml-3.0.1)
prevents text beside the picture. The
[`wp:wrapTight`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.wraptight?view=openxml-3.0.1)
and
[`wp:wrapThrough`](https://learn.microsoft.com/zh-tw/dotnet/api/documentformat.openxml.drawing.wordprocessing.wrapthrough?view=openxml-2.8.1)
elements require a
[`wp:wrapPolygon`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.wrappolygon?view=openxml-3.0.1);
its first point is `wp:start` and subsequent
[`wp:lineTo`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.lineto?view=openxml-3.0.1)
points describe the object-relative outline. Tight prevents text inside the
polygon's maximum left/right extents, while through permits text in those interior
regions. The
[`wp:align`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.wordprocessing.horizontalalignment?view=openxml-3.0.1)
element expresses alignment relative to its parent position frame. Existing verified
image operations may extract, resize, crop, change accessibility metadata, format
the host paragraph, or replace the binary; native lowering proves that the complete
anchor layout remains identical afterward.

Office 2010 adds
[`wp14:pctPosHOffset` and `wp14:pctPosVOffset`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.office2010.word.drawing?view=openxml-3.0.1)
as alternative position children. Microsoft defines the underlying percentage unit
as thousandths of a percent, so native `37500` projects as
`"percentage_offset": 37.5`; see the
[`CT_Percentage` interoperability definition](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-docx/690e8fcb-f555-4da7-8155-6848901a34df).
AiOffice writes the Office 2010 namespace into `mc:Ignorable` when it introduces
this mode, but does not invent optional `wp14:anchorId` or `wp14:editId` values.

The same Office 2010 namespace defines optional `wp14:sizeRelH` and
`wp14:sizeRelV` rules after `a:graphic`. AiOffice projects them as:

```json
{
  "relative_size": {
    "width": {"relative_to": "margin", "percentage": 75},
    "height": {"relative_to": "page", "percentage": 40}
  }
}
```

Width and height are independent and at least one axis is required. Width accepts
`inside_margin`, `left_margin`, `margin`, `outside_margin`, `page`, or
`right_margin`; height accepts `bottom_margin`, `inside_margin`, `margin`,
`outside_margin`, `page`, or `top_margin`. Values are human-facing percentage
points with native `0.001` precision, cannot be negative, and must fit signed
Int32 after multiplying by 1,000.

Relative size is a layout rule, not a replacement for native geometry. The image
node's `width` and `height` continue to project the positive absolute `wp:extent`
fallback. `image.update` can change that fallback without changing
`relative_size`; `image.anchor.update` can change the rule without changing the
fallback. This dual representation avoids guessing page geometry, section margins,
or the producer's compatibility behavior.

`image.anchor.update` selectively changes that projected layout:

```json
{
  "op": "image.anchor.update",
  "target": "#image_3A17C04E",
  "set": {
    "horizontal": {
      "relative_to": "margin",
      "alignment": "center"
    },
    "vertical": {
      "relative_to": "margin",
      "percentage_offset": -12.5
    },
    "anchor_distances": {
      "top": {"value": 5, "unit": "pt"},
      "right": {"value": 6, "unit": "pt"},
      "bottom": {"value": 7, "unit": "pt"},
      "left": {"value": 8, "unit": "pt"}
    },
    "anchor_effect_extent": {
      "left": {"value": -0.5, "unit": "pt"},
      "top": {"value": 1, "unit": "pt"},
      "right": {"value": 1.5, "unit": "pt"},
      "bottom": {"value": 2, "unit": "pt"}
    },
    "wrap": {
      "mode": "square",
      "side": "right",
      "distances": {
        "top": {"value": 2, "unit": "pt"},
        "right": {"value": 3, "unit": "pt"},
        "bottom": {"value": 2, "unit": "pt"},
        "left": {"value": 3, "unit": "pt"}
      }
    },
    "relative_size": {
      "width": {
        "relative_to": "right_margin",
        "percentage": 62.345
      }
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
change. `horizontal`, `vertical`, `anchor_distances`, `anchor_effect_extent`,
`wrap`, and `relative_size` are complete grouped values. Position callers keep the
reference frame with exactly one of `offset`, `alignment`, or
`percentage_offset`. `side` is required
and non-null exactly for
`square`, `tight`, and `through`; it must be omitted for `none` and
`top_and_bottom`. No-wrap forbids wrap-local distances, effect extent, and polygon.
Top-and-bottom forbids left/right local distances and polygons. Tight/through
require a 2–4096-segment polygon, allow only left/right local distances, and forbid
a child effect extent. Use `clear: ["anchor_distances"]`,
`clear: ["anchor_effect_extent"]`, or `clear: ["relative_size"]` to remove those
optional native groups. A relative-size set replaces both axes, so a width-only
payload removes an existing height rule. A wrap update omitting its optional
`distances` or `effect_extent` removes them as part of the complete wrap replacement.
Signed offsets must fit OOXML Int64
EMUs; percentage offsets are quantized to `0.001` percentage points and must fit
signed Int32; relative-size percentages use the same precision but must be
non-negative; text distances and `relative_height` must fit UInt32; effect extents
must fit OOXML `ST_Coordinate`; polygon coordinates fit the same signed range but remain raw
integers; booleans are strict JSON booleans.
`aioffice schema --kind floating-image-layout-update` exposes the same contract,
including machine-readable `oneOf` constraints for position and wrap modes.
`aioffice schema --kind floating-image-relative-size` exposes the axis rule, and
the `floating-image-relative-width` and `floating-image-relative-height` schema
kinds expose their distinct frame vocabularies.

For example, removing both optional parent-native groups and the relative-size rule
without changing the wrap child is a clear-only operation:

```json
{
  "op": "image.anchor.update",
  "target": "#image_3A17C04E",
  "clear": ["anchor_distances", "anchor_effect_extent", "relative_size"]
}
```

Native lowering changes only the selected `wp:positionH`, `wp:positionV`,
the one selected wrap child, optional `wp14:sizeRelH` / `wp14:sizeRelV`, and
`wp:anchor` attributes in the already-proven tree.
It preserves the drawing, extents, crop, image relationship and bytes, accessibility
metadata, host paragraph, unknown surrounding XML, and present `wp14:anchorId` /
`wp14:editId`. A wrap update may replace `wp:wrapSquare`, `wp:wrapNone`,
`wp:wrapTopAndBottom`, `wp:wrapTight`, or `wp:wrapThrough` at the same schema
position. Polygon lowering writes the requested start and line sequence exactly and
does not invent a closing point. The operation re-projects the result and fails the
complete Patch if any requested semantic value does not round-trip exactly. It
composes with `image.update` in either order inside one atomic Patch. A complete
position update may also switch the native child among `wp:posOffset`, `wp:align`,
and the axis-specific `wp14` percentage child without reconstructing the surrounding
anchor.

Physical lengths are compared by their rounded native EMU value, not by unit
spelling. For example, an accepted `1 in` offset reopens as the canonical native
projection `72 pt` without becoming a semantic or integrity mismatch.

Dev38 payloads that placed `distance_top`, `distance_right`, `distance_bottom`, and
`distance_left` inside `wrap` are still accepted when validating a complete layout
or layout update. They are migrated to `anchor_distances`, which is where those
values were always lowered natively. Dev39 emits only the corrected layered form;
new callers should not author the legacy keys.

An inline image, detached JSON projection, opaque or stale drawing, unsupported
anchor variant, empty update, partial group, multiple/no positioning modes, null,
unknown field, invalid unit/range/alignment, or non-boolean flag fails closed. The
operation does not convert an inline picture to floating and does not create or
reconstruct an anchor.

Active simple positioning, malformed or out-of-range relative-size,
distance/effect/polygon metadata, missing or
unknown compatibility values, extra position/wrap children, and mixed text plus
drawing paragraphs stay opaque. This is a deliberate protocol boundary, not an
indication that their native XML is discarded.

LibreOffice may add `bwMode="auto"` and an empty `a:noFill` while saving an otherwise
unchanged picture. AiOffice treats only those exact optional forms as neutral and
preserves them through later edits. Other black-and-white modes,
duplicate/non-empty no-fill elements, solid/gradient/pattern fills, and unknown
shape attributes still fail the projection proof.

Office-compatible producers may also normalize layered wrap geometry. In the dev39
interoperability fixture, a LibreOffice open/save retained the parent anchor
distance attributes but quantized some values, removed wrap-local distances and the
wrap-child effect extent, and reset the parent effect extent to zero. AiOffice
preserves the original fields exactly during its own no-op round trip; after an
external producer rewrites them, AiOffice projects that producer's new native state
instead of pretending the old source values survived.

The dev40 tight/through fixture shows a stronger producer rewrite: LibreOffice
removed local left/right distances, changed `edited`, regenerated the point
coordinates, appended an explicit closing point, and also normalized parent anchor
attributes. AiOffice preserves the authored non-closed polygon byte-for-byte on a
no-op and exactly through its own unrelated edits. If another producer saves the
file, the reopened projection truthfully reports that producer's new sequence.

The dev41 percentage-position fixture confirms a different interoperability edge:
AiOffice preserves the original `wp14` percentage children byte-for-byte on an
exact no-op and preserves their signed Int32 values through unrelated edits.
LibreOffice 26.8 currently rewrites those children to absolute `wp:posOffset`
values on save. AiOffice therefore projects that new absolute state after an
external save instead of claiming the percentage rule survived.

The dev42 relative-size fixture preserves `wp14:sizeRelH` and `wp14:sizeRelV`
byte-for-byte on exact no-op and through unrelated AiOffice edits. LibreOffice 26.8
currently lays out the picture from `wp:extent`, then removes both relative-size
rules on save while retaining that absolute fallback. AiOffice therefore reopens
the saved file as an absolute-size picture instead of claiming the removed rules
survived.

The dev43 transform fixture preserves negative or multi-turn `a:xfrm/@rot` values
and explicit flip booleans byte-for-byte on exact no-op and through unrelated
AiOffice edits, while projecting a canonical angle. LibreOffice 26.8 produced a
pixel-identical rendering in the tested round trip, but normalized some equivalent
rotation/flip combinations and quantized or rewrote extent evidence. When that new
geometry no longer has exact matching inner and outer extents, AiOffice deliberately
reopens it as opaque instead of claiming the original editable subset survived.

The dev44 outline fixture preserves its complete supported `a:ln` subtree
byte-for-byte on exact no-op and through unrelated AiOffice edits. LibreOffice 26.8
retained the tested visible red outline, width, and direct RGB color but removed the
preset dash plus explicit default line attributes and round-join element on save,
making the rewritten native line solid. AiOffice reopens that producer-authored
solid outline instead of claiming the original dash rule survived.

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
   crop, picture transform, picture outline, accessibility metadata, declared size,
   and SHA-256;
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
displayed extent, rectangular source crop, picture transform, and/or direct outline:

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
            "transform": {
                "rotation_degrees_clockwise": 90,
                "flip_horizontal": true
            },
            "outline": {
                "width": {"value": 2, "unit": "pt"},
                "color": "#CC0000",
                "dash": "dash_dot"
            },
            "alt_text": "Expert workflow with three approval stages",
            "title": "Expert workflow",
        },
    }
])
assert result.success
```

The operation accepts `width`, `height`, `crop`, `transform`, `outline`, `alt_text`,
and `title` in `set`. `crop`, `transform`, `outline`, `alt_text`, and `title` are
clearable:

```json
{
  "op": "image.update",
  "target": "#image_3A17C04E",
  "clear": ["crop", "transform", "outline", "alt_text", "title"]
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

### Picture rotation and mirroring

`transform.rotation_degrees_clockwise` is the canonical clockwise angle in
`[0, 360)`. AiOffice rounds it to DrawingML's signed `ST_Angle` unit of
1/60000 degree. `flip_horizontal` and `flip_vertical` are strict booleans. At
least one field must produce a visible transform; clear the complete `transform`
group to restore the identity transform.

The native authority is `pic:spPr/a:xfrm`: `rot` stores the angle and `flipH` /
`flipV` store the mirror state. The protocol normalizes a valid negative or
multi-turn native angle modulo one turn for easier AI reasoning, but unrelated
edits retain the original raw attributes. An explicit transform update writes one
canonical representation and never changes `wp:extent` or `a:xfrm/a:ext`.
This follows the Open XML
[`Transform2D`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.transform2d?view=openxml-3.0.1)
contract; native DOCX rendering remains the authority for final composition.

Setting `transform` replaces the complete rotation/mirror group, so omitted flips
become false and omitted rotation becomes zero. Unknown transform attributes,
invalid XML booleans, non-integer or out-of-Int32 rotations, nonzero `a:off`,
mismatched extents, extra children, or malformed child order fail closed as opaque.
Semantic HTML publishes normalized `data-aioffice-rotation-degrees-clockwise` and
flip attributes as evidence but does not pretend to be Word's layout engine.

### Picture outline

`outline.width` is an explicit `Length` that must quantize to DrawingML
`ST_LineWidth` from 1 through 20,116,800 EMUs. AiOffice normalizes it to points at
the exact nearest-EMU value. `outline.color` is one direct six-digit sRGB value,
normalized to uppercase. `outline.dash` is one of:

`solid`, `dot`, `system_dot`, `dash`, `system_dash`, `large_dash`, `dash_dot`,
`system_dash_dot`, `large_dash_dot`, `large_dash_dot_dot`, or
`system_dash_dot_dot`.

Those names map one-to-one onto DrawingML preset-dash tokens. Setting `outline`
replaces the complete supported line group. AiOffice writes explicit flat caps,
single compound lines, centered pen alignment, direct RGB solid fill, one preset
dash, and a round join, making the authored subset independent of inherited line
attributes. Clearing it removes the direct `a:ln` and restores the ordinary
unoutlined picture state.

For import, the same values may be omitted only where Office's supported default
resolution is equivalent. A missing preset dash projects as `solid`; a missing join
projects as the Office round default. A neutral absent/zero-width line containing
only `a:noFill` projects as no semantic outline and remains byte-exact through
unrelated edits.

Theme, system, preset, HSL, or scRGB colors; color transforms; gradient or pattern
fills; custom dash arrays; non-default caps, compound modes, or alignment; bevel or
miter joins; arrowheads; extension lists; unknown attributes; duplicate lines; and
malformed child order all fail closed as opaque. This bounded subset follows
Microsoft's [`LinePropertiesType`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.linepropertiestype?view=openxml-3.0.1),
[`RGBColorModelHex`](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.drawing.rgbcolormodelhex?view=openxml-3.0.1),
and [Office DrawingML default-resolution notes](https://learn.microsoft.com/en-us/openspecs/office_standards/ms-oe376/a9897c2b-0404-4676-aa5c-8f25bc6d66ca).
Semantic HTML exposes width, color, and dash as evidence attributes; DOCX rendering
remains the visual authority.

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
`floating_layout_update_fields`, group-replacement semantics, three clearable
fields, relative-size frames/precision/fallback authority, native distance-layer
authority, and effect-extent precedence.

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

## Addressable inline and floating image insertion

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

Omitting `floating` creates the established inline form. Supplying one strict
`FloatingImageLayout` creates the same conservative offset, alignment, or percentage
anchor with square, no-wrap, top-and-bottom, tight, or through wrapping that import and
`image.anchor.update` already understand:

```python
result = document.insert_image_after(
    "#analysis",
    "generated/expert-workflow.png",
    width={"value": 3, "unit": "in"},
    height={"value": 1.5, "unit": "in"},
    alt_text="Floating expert workflow",
    image_id="floating_workflow",
    floating={
        "horizontal": {
            "relative_to": "column",
            "offset": {"value": 72, "unit": "pt"}
        },
        "vertical": {
            "relative_to": "paragraph",
            "percentage_offset": 12.5
        },
        "anchor_distances": {
            "top": {"value": 2, "unit": "pt"},
            "right": {"value": 6, "unit": "pt"},
            "bottom": {"value": 2, "unit": "pt"},
            "left": {"value": 6, "unit": "pt"}
        },
        "wrap": {
            "mode": "square",
            "side": "both_sides"
        },
        "relative_size": {
            "width": {
                "relative_to": "margin",
                "percentage": 75
            },
            "height": {
                "relative_to": "page",
                "percentage": 40
            }
        },
        "relative_height": 1024,
        "behind_text": False,
        "locked": False,
        "layout_in_cell": True,
        "allow_overlap": True
    },
)
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
  background, and supported borders around the picture's host paragraph;
- optional `transform` sets one complete clockwise rotation and horizontal/vertical
  mirror group using the same strict semantics as `image.update`;
- optional `outline` sets one complete direct-RGB picture line using the same
  width and preset-dash semantics as `image.update`;
- placement defaults to `inline`; a non-null `floating` value selects floating
  placement and must validate as one complete `FloatingImageLayout`;
- every horizontal and vertical position must select exactly one explicit `offset`,
  allowed semantic `alignment`, or signed `percentage_offset`;
- optional `relative_size` must contain width, height, or both; each axis selects an
  allowed frame and a non-negative percentage, while the required top-level
  `width` and `height` remain the absolute fallback extent;
- `wrap.mode` must be `square`, `none`, `top_and_bottom`, `tight`, or `through`;
  `side` is required for square/tight/through; tight/through also require a native
  ordered polygon; parent distances are optional and separate from schema-defined
  wrap-local geometry;
- active simple-position anchors, malformed relative-size rules, polygons, or
  transforms, outlines, crop, effects, and other drawing features are not silently
  inferred.

The native lowering adds or reuses the content-addressed image part, creates a fresh
relationship ID, creates a `w:p/w:r/w:drawing` tree with `wp:inline` or canonical
`wp:anchor`, writes identical outer and inner extents, generates collision-free
`wp:docPr/@id` and `w14:paraId` values, applies the requested paragraph style,
inserts at the proven body position, and then re-runs the conservative projection
proof. The generated anchor uses inactive zero `simplePos`, horizontal and vertical
positions expressed by the caller's selected offset, alignment, or percentage mode,
the caller's selected supported wrap child, optional relative-size rules, and
explicit flags. Percentage-position or relative-size insertion also emits the
namespace compatibility declaration. Optional Office 2010
anchor IDs are not invented. The operation either returns a fully readable and
subsequently editable `ImageBlock` or leaves the original document unchanged.

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
  --transform image-transform.json \
  --outline image-outline.json \
  --floating-layout floating-layout.json \
  -o inserted.docx
```

`image-transform.json`, `image-outline.json`, and `floating-layout.json` are exactly
the objects accepted by `aioffice schema --kind image-transform`,
`aioffice schema --kind image-outline`, and
`aioffice schema --kind floating-image-layout`. Omitting the floating-layout flag
keeps inline placement. The Workspace CLI accepts the same flags and records the
normalized transform, outline, placement, and layout in its binary-free patch log.

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
- active-simple-position, unsupported wrap-specific effects, malformed
  relative-size rules, or otherwise unsupported floating anchors;
- non-default picture black-and-white modes or non-neutral shape fills;
- multiple pictures or alternate representations;
- linked or external images;
- negative, overconstrained, malformed, or otherwise unsupported crop rectangles;
- malformed/unknown picture transforms or nonzero transform offsets;
- unsupported picture outline fills, colors, joins, compound modes, arrowheads,
  custom dashes, extensions, or malformed line structures;
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
type, placement, and normalized crop, transform, and outline evidence. Markdown
emits an `aioffice-asset:` reference. Neither exporter embeds binary data or claims
to reproduce native floating placement, wrapping, crop, transform, outline, or
picture content.

Use the native LibreOffice PDF/PNG provider to judge the actual picture, cropping,
position, pagination, and surrounding layout. Native rendering remains the visual
authority even for the supported inline and floating subsets.
