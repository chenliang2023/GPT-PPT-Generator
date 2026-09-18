# 图片型 PPT 转可编辑 PPT 调研报告

> 调研问题：如何将图片型 PPT（由 AI 生成工具生成的不可编辑的图片幻灯片）转换为可编辑的 PPT 格式
>
> 调研日期：2026-09-17

---

## 🎯 问题

图片型 PPT（AI 生成的不可编辑图片幻灯片）如何转换为可编辑的 PPT 格式？

---

## ✅ 结论

### python-pptx 的定位

**python-pptx 是输出层工具，不是解析层工具。** 它负责将结构化数据渲染为可编辑 PPT，但不能从图片中理解内容。

| 能力 | 是否支持 |
|------|:--------:|
| 从图片中提取文字 | ❌ |
| 分析图片内容 | ❌ |
| 创建可编辑文本框 | ✅ |
| 创建可编辑形状 | ✅ |
| 创建可编辑表格/图表 | ✅ |

---

### 可行方案对比

| 方案 | 核心思路 | 文字可编辑 | 形状可重建 | 自动化程度 | 成本 |
|------|----------|:----------:|:----------:|:----------:|------|
| **A: AI 视觉 + python-pptx** | 用 VLM 识别图片，手动重建 | ✅ | ✅ | ⭐⭐⭐ | AI API |
| **B: OCR + python-pptx** | OCR 提取文字 | ⚠️ 仅文字 | ❌ | ⭐⭐ | 免费 |
| **C: 多模态大模型** | LLM 直接生成 PPTX | ✅ | ✅ | ⭐⭐⭐⭐ | 商业 API |
| **D: 商用工具** | Adobe/Smallpdf 等 | ❌ | ❌ | ⭐⭐⭐⭐⭐ | 💰 |

---

### 本项目方案

本项目的 `images-to-editable-pptx` skill 采用**方案 A（AI 视觉 + python-pptx）**：

```
原始幻灯片图片 → AI 视觉分析 → 提取内容/布局/样式 → 构建 JSON Spec → build_editable_pptx.py → 可编辑 PPTX
```

**交付门槛**：
- 所有可见文字可编辑
- 图表、表格保持可编辑
- 无源幻灯片截图被嵌入
- `validate_editable_pptx.py` 通过

**优势**：质量最高，完全可编辑
**局限**：依赖 AI 视觉模型识别精度

---

## 📚 依据

- [python-pptx Shapes API](https://python-pptx.readthedocs.io/en/latest/api/shapes.html) — `Picture` 对象无 `has_text_frame`，图片内文字不可提取
- [images-to-editable-pptx/SKILL.md](skills/images-to-editable-pptx/SKILL.md) — Delivery Gate 交付条件
- [build_editable_pptx.py L180-L195](skills/images-to-editable-pptx/scripts/build_editable_pptx.py#L180-L195) — 图片防护逻辑

---

## 🔬 查证过程

| 预想 | 让它站不住的材料 | 现在的写法 |
|------|-----------------|-----------|
| python-pptx 可以从图片提取文字 | `Picture.shape` 不暴露图片内文字，无 OCR 能力 | python-pptx 是输出层工具，不是解析层工具 |
| 商用工具可以完成转换 | 大多数在线工具转为 PDF 或图片包裹的 PPT，实际仍不可编辑 | 方案 D 仅适合临时转换，质量差 |

---

## 🚧 范围外

- 不同语言的 OCR 识别率对比（需专门测试环境）
- 商业工具的具体定价和限制

---

## 🧭 对我们的影响

1. **确认**：`images-to-editable-pptx` skill 的技术路线是图片 PPT 转可编辑 PPT 的最优方案
2. **改进方向**：提高 AI 视觉识别的自动化程度（如：骨架自动生成 + 细节 AI 填充）
