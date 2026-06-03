# Reconstruction Spec Format

Use this JSON format with `scripts/build_editable_pptx.py`. Coordinates and sizes
are in inches. The default slide size is 13.333 x 7.5 inches.

## Top Level

```json
{
  "title": "Editable deck",
  "slide_size": {"width": 13.333, "height": 7.5},
  "theme": {
    "font_face": "Microsoft YaHei",
    "text_color": "#0F172A",
    "accent_color": "#0D9488"
  },
  "slides": []
}
```

Each slide accepts `name`, `source_image`, `background`, and `elements`.
`source_image` is provenance only and is never embedded.

```json
{
  "name": "Slide 1",
  "source_image": "reference-images/slide-01.png",
  "background": "#F8FAFC",
  "elements": []
}
```

## Common Element Fields

All elements accept:

- `type`: required element type.
- `name`: optional PowerPoint object name.
- `x`, `y`, `w`, `h`: element bounds in inches.

Colors use six-digit hex strings such as `#0F172A`.

## Text

```json
{
  "type": "text",
  "name": "Title",
  "x": 0.8,
  "y": 0.6,
  "w": 6.0,
  "h": 0.8,
  "text": "A native editable title",
  "font_face": "Microsoft YaHei",
  "font_size": 30,
  "bold": true,
  "italic": false,
  "color": "#0F172A",
  "align": "left",
  "valign": "middle",
  "margin": 0.04
}
```

Use `align` values `left`, `center`, or `right`. Use `valign` values `top`,
`middle`, or `bottom`.

## Shape

```json
{
  "type": "shape",
  "name": "Metric card",
  "shape": "rounded_rectangle",
  "x": 0.8,
  "y": 2.0,
  "w": 3.2,
  "h": 1.5,
  "fill": "#FFFFFF",
  "line": "#CBD5E1",
  "line_width": 1,
  "text": "42%",
  "font_size": 28,
  "bold": true,
  "color": "#0F172A",
  "align": "center",
  "valign": "middle"
}
```

Supported shapes: `rectangle`, `rounded_rectangle`, `ellipse`, `triangle`,
`chevron`, `diamond`, `hexagon`, and `parallelogram`.

## Line

```json
{
  "type": "line",
  "name": "Connector",
  "x1": 2.0,
  "y1": 3.0,
  "x2": 5.0,
  "y2": 3.0,
  "color": "#0D9488",
  "width": 2,
  "dash": "solid",
  "arrow_start": false,
  "arrow_end": true
}
```

Supported dash values: `solid`, `dash`, and `dot`.

## Image Asset

Use images only for bounded photos or illustrations. Never point `path` to a
source slide screenshot.

```json
{
  "type": "image",
  "name": "Product photo",
  "path": "assets/product-photo.png",
  "x": 8.0,
  "y": 0.8,
  "w": 4.5,
  "h": 5.8
}
```

## Table

```json
{
  "type": "table",
  "name": "Comparison table",
  "x": 0.8,
  "y": 2.0,
  "w": 11.8,
  "h": 3.8,
  "rows": [
    ["Metric", "Current", "Target"],
    ["Growth", "18%", "25%"],
    ["Retention", "91%", "95%"]
  ],
  "header_fill": "#0D9488",
  "header_color": "#FFFFFF",
  "body_fill": "#FFFFFF",
  "body_color": "#0F172A",
  "font_size": 14
}
```

## Editable Bar Chart

```json
{
  "type": "bar_chart",
  "name": "Growth chart",
  "x": 0.8,
  "y": 2.0,
  "w": 7.0,
  "h": 4.2,
  "data": [
    {"label": "2024", "value": 48, "color": "#99F6E4"},
    {"label": "2025", "value": 72, "color": "#0D9488"}
  ],
  "value_suffix": "%",
  "label_color": "#475569",
  "value_color": "#0F172A"
}
```

## Editable Line Chart

```json
{
  "type": "line_chart",
  "name": "Trend chart",
  "x": 0.8,
  "y": 2.0,
  "w": 7.0,
  "h": 4.2,
  "data": [
    {"label": "Q1", "value": 12},
    {"label": "Q2", "value": 18},
    {"label": "Q3", "value": 27}
  ],
  "line_color": "#0D9488",
  "marker_color": "#0D9488",
  "value_suffix": "%"
}
```

The chart renderer uses editable shapes, lines, and text labels.
