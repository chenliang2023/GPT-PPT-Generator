# 主流 AI PPT 产品调研

> 调研问题：Gamma、WPS AI PPT、豆包 AI PPT、Kimi PPT 等主流产品的输入、输出、核心思路、技术栈、动画支持、二次编辑能力
>
> 调研日期：2026-09-17

---

## 🎯 问题

1. 主流 AI PPT 产品是「图片型」还是「结构化」输出？
2. 为什么多数产品看起来更像「图片型」？
3. 我们「图片 + JSON」方案的定位、优劣势是什么？

---

## ✅ 结论速览

| 结论 | 一句话 |
|:--|:--|
| 主流产品分两派 | 一派走「图片型渲染」（Gamma、Tome 历史路线），一派走「原生 PPTX」（Kimi、WPS、Doubao、Microsoft Copilot） |
| 看似图片型的根因 | 自由版式设计在网页里用 HTML/CSS/SVG 渲染最简单，导出成可编辑 PPTX 成本陡增 |
| 我们的方案 | 走中间路线：图片是参考，重建为 python-pptx 原生对象，结构化、可编辑 |

---

## 📊 产品横向对比

| 产品 | 输入 | 核心思路 | 输出格式 | 可二次编辑 | 动画 |
|:--|:--|:--|:--|:--|:--|
| **Gamma.app** | 一句话/大纲/粘贴文本/导入文件 | 网页富文本 + AI 生成 | PPTX（受限）+ PDF | ✅ 网页端，导出受限 | ❌ |
| **WPS AI PPT** | 一句话/大纲/文档 | WPS Office + AI 模板填充 | `.pptx`（原生） | ✅ 完全 | ✅ |
| **豆包 AI PPT** | 一句话/文档/网页 | 豆包大模型 + 智能体编排 | `.pptx` | ✅ | ✅ |
| **Kimi PPT** | 一句话/文本/PDF/Word/PPTX/图片 | K3 LLM + 智能布局/经典模板 | `.pptx`（原生） | ✅ 完全 | ✅ |
| **Microsoft Copilot** | Word/大纲/提示词 | PowerPoint 内部调用 OpenAI | 原生 `.pptx` | ✅ 完全 | ✅ |
| **Beautiful.ai** | 一句话/大纲 | Smart Slides 模板驱动 | `.pptx` | ⚠️ 受限 | ❌ |

---

## 🔍 各产品技术栈要点

### Gamma

- 编辑器内核：ProseMirror + Yjs
- PPTX 渲染：自研 [pptx-renderer](https://github.com/gamma-app/pptx-renderer)（TypeScript，Apache-2.0）
- 图表：ECharts
- 绘图：tldraw
- 提示词框架：AIJSX

**自述不支持**：3D 效果、动画、公式、EMF/WMF 矢量、阴影/反射/辉光

### Kimi PPT

- 模型：Kimi K3 LLM
- 两种模式：智能布局（5-10 分钟）、经典模板（3-5 分钟）
- 支持自定义模板上传

### WPS AI PPT

- 客户端：WPS Office（Windows/macOS/Linux/Web/移动端）
- AI 入口：ai.wps.cn 网页版 + 客户端内置
- 第三方插件生态：OfficeAIWork/PptGPT、it235/office-ai-agent

---

## 🤔 为什么多数产品看起来像「图片型」

| 派系 | 代表 | 优势 | 代价 |
|:--|:--|:--|:--|
| **图片型/富网页型** | Gamma（早期）、Tome、美图 AI PPT | 视觉表现力强 | 导出 PPT 困难 |
| **原生 PPTX 型** | Kimi、WPS AI、豆包、Microsoft Copilot | 完全可编辑 | 版式灵活性低 |

**多数人觉得 AI PPT 是「图片型」，原因**：

1. **视觉冲击力大**——网页 AI 排版效果远好于传统 PPTX，宣传片都用网页版截图
2. **导出 PPTX 成本太高**——需要写完整 OOXML 反向工程（Gamma 自研 pptx-renderer）
3. **动画在 PPTX 里很复杂**——OOXML 动画系统基于 SMIL，schema 跨多命名空间，没有库能完整支持
4. **多数用户其实是在线用**——核心交付是「链接 + 在线编辑」

---

## 📐 我们「图片 + JSON」方案的定位

### 我们两段式

```
阶段 1：GPT-PPT-Generator（已有）
  用户输入提示词 → gpt-image-2 出图 → 每页一张 PNG

阶段 2：images-to-editable-pptx（已有 skill）
  PNG 作为视觉参考 → AI 视觉分析 → JSON Spec → python-pptx → 可编辑 PPTX
```

### 与主流产品对比

| 维度 | 主流 AI PPT（Kimi/WPS/Doubao） | Gamma（网页型） | 我们 |
|:--|:--|:--|:--|
| 输入 | 文本/大纲/文档 | 文本/大纲/文件 | **图像（每页一张 PNG）** |
| 中间表示 | 模型内部结构化 | ProseMirror 文档树 | **显式 JSON Spec** |
| 输出 | 原生 `.pptx` | 网页 + 受限 `.pptx` | **原生 `.pptx`** |
| 视觉表现力 | 受 PPTX 约束 | 最高 | 第一阶段 GPT 自由发挥 |
| 可编辑性 | 完全 | 受限 | **完全** |
| 动画 | 模型自动 + 模板默认 | 不支持 | **暂不支持** |
| 离线运行 | ❌ | ❌ | ✅ |

### 优势

- **离线 + 本地化**：不需要订阅 AI PPT 服务
- **视觉风格可定制**：第一阶段 GPT 出图，参考图由用户上传
- **结构化中间产物**：JSON Spec 显式产物，可读、可版本管理
- **完全可编辑**：第二阶段所有元素都是 python-pptx 原生对象
- **可分阶段升级**：可替换更新模型，无需重写接口

### 劣势

- **流程比一站式长**：要等两个阶段（出图 + 重建）
- **视觉还原依赖 AI 精度**：复杂布局漏识别会丢细节
- **动画短板**：输出静态 PPTX，需手动加
- **不支持在线链接分享**：交付 `.pptx` 文件，无 SaaS 形态

### 适用场景

- 对视觉风格有强要求（设计感、艺术化版面），接受两次出稿
- 需要离线/本地化处理，不愿上云
- 拿到 PPT 后要做大幅二次编辑
- 团队工作流需「图像版快速迭代 + 文档版正式交付」

### 不适用场景

- 赶时间、要「一句话出稿、立刻能讲」——用 Kimi/WPS/Doubao 更合适
- 要带动画、过渡的发布会 PPT——目前 AI PPT 在动画上都没做好

---

## 📚 参考资料

### Gamma

- [github.com/gamma-app/pptx-renderer](https://github.com/gamma-app/pptx-renderer)
- [github.com/gamma-app/aijsx](https://github.com/gamma-app/aijsx)
- [github.com/gamma-app/y-prosemirror](https://github.com/gamma-app/y-prosemirror)

### Kimi

- [kimi.com/slides](https://www.kimi.com/slides)
- [kimi.com/help/ppt/ppt-overview](https://www.kimi.com/help/ppt/ppt-overview)
- [kimi.com/help/ppt/ppt-creation-mode](https://www.kimi.com/help/ppt/ppt-creation-mode)

### WPS AI

- [ai.wps.cn](https://ai.wps.cn)
- [github.com/OfficeAIWork/PptGPT](https://github.com/OfficeAIWork/PptGPT)

### 豆包

- [www.doubao.com](https://www.doubao.com)
- [www.volcengine.com/product/doubao](https://www.volcengine.com/product/doubao)

---

## 🔬 查证过程

| 预想 | 让它站不住的材料 | 现在的写法 |
|:--|:--|:--|
| Gamma 输出图片型 PPT | Gamma 自研 pptx-renderer 支持解析/渲染 OOXML，能导出 PPTX | 修正为「Gamma 导出 PPTX 但能力受限」 |
| Kimi PPT 输出图片 | 帮助中心明确说「使用 PowerPoint、WPS 或 Keynote 打开」 | 修正为「原生 PPTX、可编辑」 |
| 豆包 PPT 资料全无 | 主页有「PPT 生成」入口，火山方舟文档公开，但生成管线未公开 | 写成「模型与平台可查，**生成管线技术细节未公开**」 |

---

## 🚧 不确定的部分

- **豆包 PPT 的具体生成管线**——火山方舟文档未单独描述 PPT 场景
- **Beautiful.ai 当前的 AI 能力**——官网连接重置，历史资料指向「Smart Slides 模板驱动」
- **WPS AI PPT 的具体技术栈**——金山办公未公开模型选型
- **Kimi 智能布局模式的内部表示**——未披露是富文本还是直接生成 PPTX

---

## 🧭 对我们的影响

1. **定位清晰**：我们的方案是「图片型视觉 + 结构化输出」的中间路线
2. **承认动画短板**：动画不是 python-pptx 当前能解决的问题，**短期内不要承诺**
3. **下一阶段可选项**：
   - 接入 Kimi/WPS/Doubao API（输出原生 PPTX），扩展「图片 + 重建 + 一站式兜底」
   - 给 images-to-editable-pptx 加网页端预览/编辑界面（ProseMirror + Yjs）
4. **不要把赌注压在「追上 Kimi/WPS」**——它们是大模型公司 + 办公软件公司的联合，我们做工具型 skill 应聚焦「**可本地化、可控中间产物、可编程**」