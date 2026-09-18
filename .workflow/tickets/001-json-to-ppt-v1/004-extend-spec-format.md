<!-- status: in_review task:#13 已实测验证，待 Merge -->

# [004] 扩展 `spec-format.md`（template 字段 + slide_size 描述）

## 📌 Spec 引用
来源：[`.workflow/specs/json-to-ppt-skill.md`](../specs/json-to-ppt-skill.md) 的「📐 扩展 JSON Schema（向后兼容）」

## 🎯 任务
在 `skills/json-to-ppt/references/spec-format.md` 写新 spec 文档。它是 `skills/images-to-editable-pptx/references/spec-format.md` 的**超集**：原文一字不动地放在文档里作为「## 继承自 images-to-editable-pptx」，再追加「## 新增字段」。

现在 `json-to-ppt/references/` 还没有文件。任务完成后：

- `skills/json-to-ppt/references/spec-format.md` 存在
- 原 spec-format 内容**完整**复制在「继承自」小节
- 新增两个字段的文档：`slide.template`（指向 PPTX 模板路径，字符串，相对 base_dir）和 `slide.background.image`（字符串路径）
- 「向后兼容」明示：原 JSON 文件不需要任何修改就能用新 skill 渲染

## ✅ 验收标准
1. `cat skills/json-to-ppt/references/spec-format.md` 文档 ≥ 200 行
2. 文档包含完整原 spec-format 的所有元素类型（text/shape/line/image/table/bar_chart/line_chart）
3. 「新增字段」小节有 `template` 字段示例（指向 base 模板）和 `background.image` 示例
4. 文档明确写出"原 JSON 不需修改即可渲染"
5. 没有「破坏性变更」「removed」字样

## 🧪 测试 seam
- shell：`wc -l skills/json-to-ppt/references/spec-format.md`、`grep -c "^##" skills/json-to-ppt/references/spec-format.md`
- 不需要 pytest

## 🚧 阻塞
- 依赖：[001] scaffold（必须先有 `references/` 目录）

## 🤖 建议 Agent
Claude
理由：写文档 + 保持向后兼容叙事属于 `core-domain` 范畴

## 🌿 分支
`ticket/004-extend-spec-format`

## 🏃 执行约定
- **直接动手，不要先出计划等确认**——本 ticket 已定稿，实现 → 跑验收命令 → 提交，一次做完
- 分支由 CodeG 管理（`task/<id>`）：不要自建分支、不要 push、不要改 `.workflow/`
- 验收命令必须真跑，并把**实际输出**贴回结果；跑不起来就直说，不要声称通过

## 🔍 派发后验证（2026-09-18，服务器侧实测，待 Merge）

`task/13`（分支 `task/13`，提交 `c3679ab`）5 条验收全部实测通过：

| # | 实测结果 |
|:--|:--|
| 1 | 482 行（要求 ≥ 200） |
| 2 | 7 种元素类型全覆盖：text / shape / line / image / table / bar_chart / line_chart |
| 3 | 「新增字段」小节含 `slide.template` 与 `slide.background.image` 的完整 JSON 示例，并写明与 CLI `--template` 的并存关系 |
| 4 | 「向后兼容」小节 + 文档头部表格都明示"旧 JSON 一个字不改即可渲染" |
| 5 | 无「破坏性变更」「removed」字样 |

额外核对了「超集」这条硬要求：旧 `spec-format.md`（195 行）**整体作为连续子串**出现在新文档中（`old in new == True`），逐行比对 **0 行丢失**。

## 🧭 上下文
- 原文：[`skills/images-to-editable-pptx/references/spec-format.md`](../../skills/images-to-editable-pptx/references/spec-format.md)（约 130 行）
- 「超集」实现方式有两条：
  - **A**：直接 copy 原文 + 在末尾追加「新增字段」小节
  - **B**：写一个总览 + 用 `<!-- include -->` 引用旧文件
- 选 A：spec 是给人看的，copy 更直观；不会因旧文件改动而漏同步（doc 是 frozen 在某次 grill-me 的快照）
- `slide.template` 字段语义：spec 这一层声明「这一张 slide 用哪个模板」，CLI `--template` 是全局的，二者并存但优先级在 ticket 里**不用定**——本次只文档化字段，不实现读取
- `background.image`：与现有 `background`（颜色 hex）并列，类型为路径字符串；本次也只文档化
- 这俩字段的**实际实现**留到后续 ticket；本次范围仅文档
- 文档读者：要给 GPT/Claude 生成 spec 的 LLM 看，每个字段都要有完整 JSON 示例