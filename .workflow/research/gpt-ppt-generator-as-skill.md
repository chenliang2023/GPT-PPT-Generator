# GPT-PPT-Generator 转换为 Skill 调研报告

> 调研问题：如何将现有的 GPT-PPT-Generator 项目转换为一个可复用的 skill
>
> 调研日期：2026-09-17

---

## 🎯 问题

GPT-PPT-Generator 能否转换为可复用的 Skill？如何转换？

---

## ✅ 结论

### 能转，但需要分拆

GPT-PPT-Generator **不能整体作为一个 Skill**，原因：

| 问题 | 说明 |
|------|------|
| **Flask Web 应用 ≠ Skill** | Skill 系统是被调用的执行单元，不是独立 Web 服务器 |
| **Stateful 任务管理** | `app.py` 的 `JOBS` 字典管理异步状态，无法跨调用持久化 |
| **已有重复** | `images-to-editable-pptx` 已存在，负责图片→可编辑 PPTX |

---

### 推荐的转换架构

```
GPT-PPT-Generator
├── skill-1: gpt-image-batch-generator   ← 新建
│   ├── SKILL.md                        ← 必选（参照 images-to-editable-pptx/SKILL.md）
│   ├── scripts/generate_batch.py       ← 从 app.py 提取
│   └── references/deck-format.md       ← 推荐
└── skill-2: images-to-editable-pptx   ← 已有 ✅
```

---

### 所需文件

| 优先级 | 文件 | 说明 |
|--------|------|------|
| 🔴 必选 | `SKILL.md` | 格式参照 `skills/images-to-editable-pptx/SKILL.md` |
| 🔴 必选 | `scripts/generate_batch.py` | 从 `app.py` 提取 compose_prompt、request_image 等 |
| 🟡 推荐 | `references/deck-format.md` | deck.json 格式文档 |
| 🟡 推荐 | `scripts/requirements.txt` | 复用 `requirements.txt` |
| 🟢 可选 | `agents/openai.yaml` | Agent 配置 |

---

### 转换步骤

1. **提取** — 从 `app.py` 提取 API 调用 + prompt 组合逻辑为独立 CLI
2. **封装** — 编写 `SKILL.md`（YAML frontmatter + Markdown 正文）
3. **注册** — 在 `skills-lock.json` 添加新条目

---

## 📚 依据

- `app.py` — 包含 Flask Web 应用和 JOBS 状态管理
- `skills/images-to-editable-pptx/SKILL.md` — Skill 格式参照
- `skills-lock.json` — Skill 注册位置

---

## 🔬 查证过程

| 预想 | 让它站不住的材料 | 现在的写法 |
|------|-----------------|-----------|
| 可以整体转换为一个 Skill | Flask app.py 是 Web 服务器架构，与 Skill 调用模式不兼容 | 需要拆分：提取核心逻辑为 CLI + SKILL.md |

---

## 🧭 对我们的影响

1. **新 Skill**：需要创建 `gpt-image-batch-generator` skill
2. **提取工作**：从 `app.py` 提取核心逻辑
3. **格式参照**：以 `images-to-editable-pptx` 为模板
