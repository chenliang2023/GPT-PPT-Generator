<!-- status: dispatched to:pi via:codeg-todos at:2026-09-18T10:47:05+08:00 task:#11 前置门禁：等 [005] 合并后才真正开工 -->

# [006] 注册到顶层 `skills-lock.json`

## 📌 Spec 引用
来源：[`.workflow/specs/json-to-ppt-skill.md`](../specs/json-to-ppt-skill.md) 的「Q1 暴露方式 b」「💡 解决方案」

## 🎯 任务
把新 skill 注册进仓库根的 `skills-lock.json`（如果该文件存在并维护 skill 列表），让其它工具/agent 能发现 `json-to-ppt`。

现在仓库根有 `skills-lock.json` 但只覆盖既有 skill。任务完成后：

- `skills-lock.json` 包含 `json-to-ppt` 条目（name + path 形式 + entry_point）
- `python -c "import json; d=json.load(open('skills-lock.json')); assert 'json-to-ppt' in d['skills']"` 退出码 0

## ✅ 验收标准
1. `cat skills-lock.json` 显示新增条目，键名按现有约定
2. 上述内联断言退出码 0
3. JSON 文件本身 `python -c "import json; json.load(open('skills-lock.json'))"` 仍合法
4. 没动 `images-to-editable-pptx` 既有条目

## 🧪 测试 seam
- shell + python 内联 JSON 解析

## 🚧 阻塞
- 依赖：[005] 新 SKILL.md 存在（要先有 name/description 才能注册）

## 🤖 建议 Agent
Pi
理由：`config` 类改动

## 🌿 分支
`ticket/006-skills-lock-registration`

## 🏃 执行约定
- **直接动手，不要先出计划等确认**——本 ticket 已定稿，实现 → 跑验收命令 → 提交，一次做完
- 分支由 CodeG 管理（`task/<id>`）：不要自建分支、不要 push、不要改 `.workflow/`
- 验收命令必须真跑，并把**实际输出**贴回结果；跑不起来就直说，不要声称通过

## 🧭 上下文
- `skills-lock.json` 的 schema：先 `cat skills-lock.json` 看现有格式再追加
- entry_point 通常指向 `scripts/build_pptx.py`（CLI 入口）
- 如果 `skills-lock.json` 是 per-skill 而非全局列表：把 `skills/json-to-ppt/skills-lock.json` 单独建一份
- 不要改 `requirements.txt` / `requirements-dev.txt`（如果新 skill 需要额外依赖，由 T005 决定是否加）