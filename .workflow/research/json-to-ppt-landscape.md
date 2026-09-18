# JSON/结构化文本转 PPT 方案调研

> 调研问题：市面上主流的 JSON → PPT 方案有哪些？各有什么特点和优缺点？
>
> 调研日期：2026-09-17

---

## 🎯 问题

JSON/结构化文本转 PPT 的主流方案是什么？我们项目应该采用哪种？

---

## ✅ 结论

### 这个领域的方案分三类

| 类型 | 代表方案 | 适用场景 |
|:--|:--|:--|
| **底层库** | python-pptx、pptxgenjs | 需要完全控制输出质量 |
| **中间格式** | Markdown → reveal.js → Decktape | 快速原型、程序员友好 |
| **商业服务** | Beautiful.ai、Canva API | 非技术用户、批量生产 |

---

### 推荐：python-pptx + 自定义 JSON Schema

理由：

| 维度 | 分析 |
|:--|:--|
| **输入匹配** | 项目输入是 GPT 生成的 JSON（`sample-deck.json`） |
| **元素覆盖** | 图表、表格、图片、形状全部支持 |
| **输出质量** | 生成原生 `.pptx`，用户可后续编辑 |
| **技术栈** | Python 项目，与现有代码无缝集成 |
| **License** | MIT，商业可用 |

---

## 📊 方案对比表

| 方案 | 输入格式 | 文本 | 形状 | 图表 | 表格 | 图片 | 动画 | License |
|:--|:--|:--:|:--:|:--:|:--:|:--:|:--|:--|
| **python-pptx** | JSON（需映射） | ✅ | ✅ | ✅ | ✅ | ✅ | ❌ | MIT |
| **pptxgenjs** | JSON（原生） | ✅ | ✅ | ✅ | ✅ | ✅ | ❌ | MIT |
| **Decktape** | HTML | ✅ | ✅ | ✅ | ✅ | ✅ | ✅ | MIT |
| **Beautiful.ai** | 文字 | ✅ | ✅ | ✅ | ❌ | ✅ | ❌ | 商业 |
| **Canva API** | 文字/图片 | ✅ | ✅ | ✅ | ✅ | ✅ | ❌ | 商业 |
| **Gamma.app** | 文字 | ✅ | ✅ | ❌ | ❌ | ✅ | ❌ | 商业 |

---

## 🐍 python-pptx 详情

### 基本信息

| 项目 | 信息 |
|:--|:--|
| GitHub | scanny/python-pptx |
| License | MIT |
| Python | ≥ 3.8 |

### 支持的元素类型

| 元素 | 支持程度 | 备注 |
|:--|:--|:--|
| 文本框 | ✅ 完整 | 字体、颜色、大小、对齐、行间距 |
| 形状 | ✅ 完整 | 矩形、圆形、线条、箭头等 |
| 图片 | ✅ 完整 | PNG、JPEG、SVG 嵌入 |
| 表格 | ✅ 完整 | 单元格合并、样式、边框 |
| 图表 | ✅ 完整 | 柱状图、饼图、折线图、面积图 |
| 母版/布局 | ✅ 完整 | 读取和修改现有模板 |
| 动画 | ❌ 不支持 | pptx 是静态格式 |

---

## 📦 pptxgenjs 详情

### 基本信息

| 项目 | 信息 |
|:--|:--|
| GitHub | sauchelli/pptxgenjs |
| 语言 | JavaScript / TypeScript |
| Stars | ~3k+ |

### 特点

| ✅ 优点 | ❌ 缺点 |
|:--|:--|
| JSON 驱动的 API（原生支持） | 不支持完整母版/模板 |
| TypeScript 原生支持 | 图表依赖 Chart.js 风格数据 |
| 可在浏览器端运行 | 不如 python-pptx 成熟 |

---

## 🔮 其他方案

| 方案 | 说明 |
|:--|:--|
| **python-pptx-template** | Jinja2 模板引擎，适合固定模板 + 动态数据 |
| **markitdown** | Markdown → PowerPoint，输入是 Markdown 不是 JSON |
| **SlidesAI.io** | Google Slides 插件，面向 Google Slides 用户 |

---

## 🏁 具体实施建议

1. **定义 JSON Schema** — 规范 GPT 生成的 deck spec 格式
2. **实现映射层** — 将 JSON 转换为 python-pptx API 调用
3. **支持模板** — 读取现有 `.pptx` 模板，填充内容
4. **分步验证** — 先生成静态 PPTX，再考虑动态更新

---

## 📚 参考资料

- python-pptx: https://github.com/scanny/python-pptx
- pptxgenjs: https://github.com/sauchelli/pptxgenjs
- Decktape: https://github.com/astefanutti/decktape
- Beautiful.ai: https://www.beautiful.ai/
- Canva API: https://www.canva.com/developers/

---

## 🧭 对我们的影响

1. **确认**：python-pptx 是我们项目的最优选择
2. **下一步**：定义 JSON Schema 规范 deck spec 格式
3. **注意**：动画不在 python-pptx 支持范围内（需要另做调研）
