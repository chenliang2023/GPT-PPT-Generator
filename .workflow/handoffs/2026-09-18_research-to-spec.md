# Handoff · research → spec

> **日期**：2026-09-18
> **来源阶段**：research（5 份调研报告）
> **目标阶段**：spec（1 份设计稿）
> **写入位置**：`.workflow/specs/json-to-ppt-skill.md`

---

## Source

从以下 5 份调研报告中**提取结论**并合并到 spec：

| 调研文件 | 关键结论（被采纳） | 对应 spec 段 |
|:--|:--|:--|
| [gpt-ppt-generator-as-skill.md](../research/gpt-ppt-generator-as-skill.md) | GPT-PPT-Generator 应拆分为**多个 skill**；JSON → PPTX 是独立子模块 | 🎯 问题陈述 / 💡 解决方案 / 📎 附注 |
| [image-ppt-to-editable.md](../research/image-ppt-to-editable.md) | python-pptx 是 PPTX 生成/编辑的稳定底座；可直接复用其代码 | 💡 解决方案 / 🏗️ 实现决策（"复用现有渲染逻辑"） |
| [json-to-ppt-landscape.md](../research/json-to-ppt-landscape.md) | 市面无标准；推荐自建 JSON schema + python-pptx；避免引入重量级依赖 | 💡 解决方案 / 🏗️ 实现决策（schema 自定义 + 渲染引擎） |
| [ai-ppt-animation.md](../research/ai-ppt-animation.md) | python-pptx **没有公共 API** 创建形状动画，OOXML schema 复杂 | 🏗️ 实现决策（⏸️ 不做动画） |
| [ai-ppt-products-landscape.md](../research/ai-ppt-products-landscape.md) | Gamma/Kimi/豆包走"AI 生图片 + 模板"路线，**不可编辑**；本项目差异化方向：JSON → 原生可编辑 PPTX | 🎯 问题陈述 / 🚫 范围外 |

**未采纳的结论**（仅记录，避免后续回归）：

- `gpt-ppt-generator-as-skill.md` 提到的"独立 CLI + Flask 后端 + Web 预览"完整拆分方案 → 第一版先做 CLI 端到端，Web 预览留到后续版本
- `ai-ppt-products-landscape.md` 提到的"Gamma 风格的 LLM 直接生成 PPT"路线 → 与"用户已有 JSON 输入"的定位冲突，**不采纳**

---

## Target

| 输出文件 | 写入位置 | 性质 |
|:--|:--|:--|
| [`.workflow/specs/json-to-ppt-skill.md`](../specs/json-to-ppt-skill.md) | 整体新建 | 第一版唯一 spec 文件 |

下游消费者：`spec-to-tickets` handoff 会引用本文件。

---

## Decisions preserved（spec 锁定的 8 个决策）

grill-me 8 问，决定如下，全部写入 spec 的「Grill-me 决定（定稿）」表格：

| 编号 | 决策点 | 选定 | 含义 |
|:--|:--|:--|:--|
| 范围 | 仅做 JSON → PPTX 渲染 | **a** 只渲染 | 不接受 Markdown、图片等其他输入 |
| Schema | 元素类型扩展方式 | **b** 扩展（向后兼容） | 新 spec 是旧 spec 的超集；旧 JSON 仍能渲染 |
| 输入 | 单 JSON / 单 PPTX | **a** 单 JSON → 单 PPTX | 不做批量、多 deck、merge |
| 测试 seam | 测试边界 | **a** CLI 端到端 | 不测 python-pptx 内部 API |
| **Q1** 暴露方式 | skill 放哪里 | **b** 单独 `skills/json-to-ppt/` | 不挂到 `images-to-editable-pptx` 下 |
| **Q2** 向后兼容 | 与 `images-to-editable-pptx` 关系 | **a** 100% 向后兼容 | 旧 spec、旧 CLI、旧测试一律不动 |
| **Q3** validate 选项 | v1 是否带 `--validate` | **a** v1 不带 | 校验交给下游工具 |
| **Q4** 失败行为 | 渲染异常退出策略 | **b** warning + 继续 | 元素级异常不阻断整个 deck |

**额外工程决策**（不在 grill-me 内，但同样锁定）：

- 🔧 **复用现有渲染层**：直接抽 `build_editable_pptx.py` 作为核心脚本，**不重写** python-pptx 调用层
- 🚪 **入口脚本独立**：新建 `scripts/build_pptx.py` CLI，内部 `from build_editable_pptx import render` 调用
- 🧩 **支持模板输入**：CLI 加 `--template <file.pptx>`，追加生成 slide 到模板尾部
- 🛡️ **全页图片防护保留**：`validate_editable_pptx.py` 的 ≥90% 面积非授权图拒绝逻辑**必须**沿用
- ⏸️ **不做动画**：明确不支持

---

## Open questions（spec 未解决，留给下游）

1. **T003 `--template` 追加模式的实现细节** — 决策时点：T003 开始前必须拍板。
   - 选项 a：动 `build_editable_pptx.py`，给 `build_pptx()` 加 `existing_presentation` 参数
   - 选项 b：不动旧文件，T003 CLI 内用 `Presentation(template).slides.add_clone(...)` 拼装
2. **元素的 spec-format 字段扩展范围** — T004（spec-format 扩展 ticket）落地时定
   - 已知要扩展：`slide.template`、`background.image`
   - 待定：是否引入 `slide.layout`（layout 名称映射）、`element.animations`（即使不做 v1 渲染也要在 schema 里占位？）
3. **fixture 文件粒度** — T005（fixture + SKILL.md）落地时定
   - 草拟 4 个 fixture：纯文字、多元素、表格、含模板路径
   - 待定：是否再增加"含坏元素的 fixture"（用于验证 Q4 warning 行为）

---

## 交叉引用

- 上游：5 份 research，路径见上方 Source 表
- 下游：spec-to-tickets handoff（[2026-09-18_spec-to-tickets.md](2026-09-18_spec-to-tickets.md)）
- 平行：spec 本身写入 `.workflow/specs/json-to-ppt-skill.md`