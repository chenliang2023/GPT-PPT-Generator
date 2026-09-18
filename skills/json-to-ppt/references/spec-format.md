# JSON-to-PPT Spec Format

`skills/json-to-ppt` 的 spec 文档。本文件是
[`skills/images-to-editable-pptx/references/spec-format.md`](../../images-to-editable-pptx/references/spec-format.md)
的**超集**（superset）：旧文档全文**原样**收录在下面的「继承自 images-to-editable-pptx」
小节里，新字段追加在其后的「新增字段」小节。

读者：给生成 spec 的 LLM（GPT / Claude）。每个字段都附完整 JSON 示例，示例可直接拷进
spec 文件使用。坐标与尺寸单位一律为英寸（inch）；颜色一律是六位 hex 字符串（如 `#0F172A`）。

- 本 skill 渲染入口：
  `python skills/json-to-ppt/scripts/build_pptx.py --spec spec.json --out out.pptx [--template base.pptx]`
- 旧 skill 渲染入口（仍然可用）：`python skills/images-to-editable-pptx/scripts/build_editable_pptx.py --spec spec.json --out out.pptx`
- 所有路径字段（`source_image`、`elements[].path`、`slide.template`、`background.image`）
  一律相对 **base_dir** 解析；CLI 里 base_dir 就是 `--spec` 文件所在目录。

## 文档结构

| 小节 | 内容 |
|---|---|
| 继承自 images-to-editable-pptx | 旧 spec 全文，逐字未改 |
| 新增字段 | `slide.template`、`slide.background.image` |
| slide_size 的解析规则 | 传 `--template` 时尺寸以模板为准 |
| 向后兼容 | 旧 JSON 一个字不改即可用新 skill 渲染 |
| 原文同步校验 | 用 `diff` 验证「继承自」小节与原文件逐字一致 |

## 继承自 images-to-editable-pptx

以下内容逐字来自 `skills/images-to-editable-pptx/references/spec-format.md`。
为避免与本文件的标题层级冲突，**原样保留其 `#` 一级标题**——这样可以直接用 `diff` 逐字比对
（比对方法见文末「原文同步校验」）。

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

## 新增字段

以下是 `skills/json-to-ppt` 相对旧 spec 的扩展，共两个字段。

> ⚠️ **本次只做文档化，渲染器尚未读取这两个字段。** 语义写在这里，是为了让生成 spec 的
> LLM 提前有统一口径、不要在 JSON 里自创别的键名。实际实现（读字段 + 走渲染路径）见后续 ticket。

### `slide.template`（字符串，相对 base_dir）

声明「这一张 slide 用哪个 PPTX 模板」。取值是 `.pptx` 模板文件的路径字符串，相对 base_dir
解析（口径与 `elements[].path`、`source_image` 一致）。

与 CLI `--template` 的关系：

- CLI 的 `--template` 是**全局**的——整份 deck 一个模板，生成的 slide 追加到该模板的 slide 之后；
- `slide.template` 是**逐 slide** 的——spec 自己声明每张 slide 的模板来源；
- 二者**并存**。本版**不定义优先级**（渲染器还没读这个字段），优先级留到实现它的 ticket 再拍板。

最小示例（两张 slide 各自声明模板）：

```json
{
  "title": "Deck with per-slide templates",
  "slides": [
    {
      "name": "Opener",
      "template": "templates/base.pptx",
      "background": "#F8FAFC",
      "elements": []
    },
    {
      "name": "Detail",
      "template": "templates/base.pptx",
      "elements": []
    }
  ]
}
```

带元素的完整示例：

```json
{
  "name": "Opener",
  "template": "templates/deck-base.pptx",
  "background": "#F8FAFC",
  "elements": [
    {
      "type": "text",
      "name": "Title",
      "x": 0.8,
      "y": 0.6,
      "w": 6.0,
      "h": 0.8,
      "text": "模板来自 templates/deck-base.pptx",
      "font_face": "Microsoft YaHei",
      "font_size": 30,
      "bold": true,
      "color": "#0F172A",
      "align": "left",
      "valign": "middle"
    }
  ]
}
```

不写 `template` 的 slide：沿用 CLI 传入的模板；没有 `--template` 时则由渲染器从空白起。

### `slide.background.image`（字符串，相对 base_dir）

现有 `slide.background` 是颜色 hex 字符串（如 `"#F8FAFC"`）。新增的 `background.image`
与之**并列**：类型为**路径字符串**，指向一张背景图，相对 base_dir 解析。

只给背景图：

```json
{
  "name": "Cover",
  "background": {"image": "assets/backgrounds/cover-bg.png"},
  "elements": []
}
```

颜色 + 背景图同时给（图片叠在颜色之上；具体叠加/缩放方式由实现它的 ticket 定）：

```json
{
  "name": "Cover",
  "background": {"color": "#0F172A", "image": "assets/backgrounds/cover-bg.png"},
  "elements": [
    {
      "type": "text",
      "x": 0.9,
      "y": 2.6,
      "w": 8.0,
      "h": 1.2,
      "text": "Q3 业务回顾",
      "font_size": 40,
      "bold": true,
      "color": "#FFFFFF",
      "align": "left",
      "valign": "middle"
    }
  ]
}
```

旧写法（纯 hex 字符串）继续有效，一个字都不用改：

```json
{
  "name": "Cover",
  "background": "#0F172A",
  "elements": []
}
```

### 完整示例：旧字段 + 新字段混用

下面这份 spec 同时用到旧字段（`title` / `slide_size` / `theme` / `elements`）和新字段
（`template` / `background.image`），可以直接当一个可运行模板：

```json
{
  "title": "Q3 业务回顾",
  "slide_size": {"width": 13.333, "height": 7.5},
  "theme": {
    "font_face": "Microsoft YaHei",
    "text_color": "#0F172A",
    "accent_color": "#0D9488"
  },
  "slides": [
    {
      "name": "封面",
      "template": "templates/base.pptx",
      "background": {"color": "#0F172A", "image": "assets/backgrounds/cover-bg.png"},
      "elements": [
        {
          "type": "text",
          "name": "Deck title",
          "x": 0.9,
          "y": 2.6,
          "w": 8.0,
          "h": 1.2,
          "text": "Q3 业务回顾",
          "font_size": 40,
          "bold": true,
          "color": "#FFFFFF",
          "align": "left",
          "valign": "middle"
        },
        {
          "type": "line",
          "x1": 0.9,
          "y1": 4.2,
          "x2": 5.4,
          "y2": 4.2,
          "color": "#0D9488",
          "width": 3,
          "dash": "solid",
          "arrow_end": false
        }
      ]
    },
    {
      "name": "指标页",
      "template": "templates/base.pptx",
      "background": "#F8FAFC",
      "elements": [
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
        },
        {
          "type": "table",
          "x": 4.6,
          "y": 2.0,
          "w": 7.8,
          "h": 2.6,
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
      ]
    }
  ]
}
```

## slide_size 的解析规则

`slide_size` 是顶层字段，写法与旧 spec 完全一致：

```json
{"slide_size": {"width": 13.333, "height": 7.5}}
```

省略时渲染器用默认尺寸 **13.333 x 7.5 英寸**（16:9）。

**传了 `--template` 时，尺寸以模板为准**：模板自带的 `slide_width` / `slide_height` 保留，
spec 里显式声明的 `slide_size` 与之不一致时被忽略，并在 stderr 给一条 warning——不会静默
改掉模板尺寸（例如 10 x 7.5 英寸的模板，输出仍是 10 x 7.5 英寸）。

不传 `--template` 时，`slide_size` 按 spec 声明生效（省略则用上面的默认值）。

也就是说：想让 `slide_size` 说了算，就不要传 `--template`；传了模板，就以模板为准。

## 向后兼容

**原 JSON 文件不需要任何修改就能用新 skill 渲染。**

- 本文件是旧文档的超集：旧文档里的
  `title` / `slide_size` / `theme` / `slides` / `source_image` / `background`（hex）/ `elements`
  以及全部八种元素类型（`text`、`shape`、`line`、`image`、`table`、`bar_chart`、`line_chart`，
  外加通用元素字段）语义一字未改。
- `slide.template` 与 `slide.background.image` 是**可选**新字段：不写就是旧行为，写了才生效，
  且当前版本渲染器还不读取它们（见「新增字段」的说明）。
- 旧 CLI `skills/images-to-editable-pptx/scripts/build_editable_pptx.py` 保持可用，
  用哪份 spec 都能跑；同一份 JSON 交给新 CLI（`build_pptx.py`）也能跑。
- 迁移方式：**不用迁移**。想用新能力再按需加字段，加与不加可以混在同一份 deck 的不同 slide 上。

## 原文同步校验

「继承自 images-to-editable-pptx」小节是逐字拷贝，可以被机械验证。用下面的命令抽出该小节
（从 `# Reconstruction Spec Format` 到 `## 新增字段` 之前，去掉分隔空行），与原文件逐字比对：

```bash
awk '/^# Reconstruction Spec Format$/{f=1} /^## 新增字段$/{f=0} f' \
    skills/json-to-ppt/references/spec-format.md | head -n -1 > /tmp/inherited.md
diff /tmp/inherited.md skills/images-to-editable-pptx/references/spec-format.md && echo "逐字一致"
```

`diff` 无输出（并打印 `逐字一致`）即说明旧 spec 全文在该小节里完整且逐字保留。
