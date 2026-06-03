---
name: images-to-editable-pptx
description: Rebuild slide screenshots, generated PPT images, or image-only presentation pages into a genuinely editable PowerPoint deck using native text, shapes, lines, tables, and chart primitives. Use when Codex or another agent must create an editable .pptx from slide images and must not embed each reference image as a full-page picture.
---

# Images To Editable PPTX

Reconstruct slide images as native PowerPoint objects. Treat every source image
as a visual reference only, never as the slide itself.

## Hard Requirements

- Do not add the source slide image as a full-page picture or background.
- Do not deliver a slide whose only meaningful object is a picture.
- Recreate all visible text as native editable text.
- Recreate diagrams, cards, rules, arrows, tables, and charts with editable
  PowerPoint primitives.
- Use raster images only for bounded photo or illustration assets that cannot be
  reasonably rebuilt as shapes. Do not use text-bearing screenshots.
- Prefer editability over pixel-perfect imitation when the two conflict.

## Workflow

1. Inspect every slide image and extract its text, hierarchy, geometry, palette,
   reading order, and semantic relationships.
2. Plan a native reconstruction for each visual object.
3. Build the deck using the host's native presentation authoring capability.
   In Codex environments with artifact-tool, follow the installed Presentations
   skill and use editable artifact-tool primitives.
4. For a portable workflow in other tools, create a reconstruction JSON spec and
   run `scripts/build_editable_pptx.py`. Read `references/spec-format.md` when
   using the bundled renderer.
5. Render or open the result, compare it against the references, and iterate.
6. Run `scripts/validate_editable_pptx.py` before delivery. Fix every failure.

## Reconstruction Rules

- Use one text box per independently editable text block.
- Use shapes with text only when the text semantically belongs to the shape.
- Use native lines and connectors for relationships; do not rasterize diagrams.
- Build bar and line charts from editable shapes or native charts.
- Keep titles, body copy, data labels, and footnotes separately editable.
- Approximate complex gradients or decorative effects with simple editable
  styling instead of flattening the slide.
- If a slide contains a complex photograph, use the photograph only as a
  bounded asset and rebuild all overlay content natively.

## Portable Renderer

Initialize a reconstruction spec from a folder of reference images:

```bash
python scripts/init_rebuild_spec.py \
  --images-dir path/to/reference-images \
  --out path/to/rebuild-spec.json
```

Build the editable PPTX:

```bash
python scripts/build_editable_pptx.py \
  --spec path/to/rebuild-spec.json \
  --out path/to/editable-deck.pptx
```

Validate the result:

```bash
python scripts/validate_editable_pptx.py path/to/editable-deck.pptx
```

Use `--allow-full-bleed-images` only when the deck intentionally contains a
verified full-bleed photo asset. Never use it to permit source slide screenshots.

## Delivery Gate

Deliver only when:

- the slide count matches the references;
- all visible text is editable;
- diagrams, tables, and charts remain editable;
- no reference slide screenshot is embedded;
- the validator passes;
- the reconstructed deck is visually coherent at thumbnail and full-slide size.
