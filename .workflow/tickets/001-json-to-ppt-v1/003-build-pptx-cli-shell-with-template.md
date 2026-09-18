<!-- status: todo -->

# [003] 新建 `build_pptx.py` CLI 入口（薄壳 + `--template`）

## 📌 Spec 引用
来源：[`.workflow/specs/json-to-ppt-skill.md`](../specs/json-to-ppt-skill.md) 的「🚪 入口脚本独立」「🧩 支持模板输入」

## 🎯 任务
在 `skills/json-to-ppt/scripts/build_pptx.py` 新建一个 CLI 入口，调旧 `build_editable_pptx.build_pptx()` 渲染。**仅在传 `--template <file.pptx>` 时**：用模板作为起始 Presentation，把生成的 slide 追加到模板的 slide 列表**之后**。

现在没有这个 CLI；新 skill 还没有渲染入口。任务完成后：

- `python skills/json-to-ppt/scripts/build_pptx.py spec.json out.pptx` 正常工作（无模板 = 从空白起）
- `python skills/json-to-ppt/scripts/build_pptx.py spec.json out.pptx --template base.pptx` 把生成的 slide 追加到 base.pptx 已有的 slide 之后
- **不修改** `images-to-editable-pptx` 的任何文件
- 模板的母版/配色/layouts 被保留（因为 `Presentation(template_path)` 复用其 slide_master）

## ✅ 验收标准
1. `python skills/json-to-ppt/scripts/build_pptx.py tests/fixtures/spec-text-only.json /tmp/out.pptx` 成功生成
2. `--template` 用一个含 2 张空 slide 的 fixture：生成的 PPTX 有「模板原有 slide 数 + spec 中 slide 数」张
3. **不传** `--template` 时，PPTX 的 slide 数 = spec 中 slide 数（行为等同 T002 完成后的旧 CLI）
4. `--template` 指向不存在文件时，CLI 退出非零，stderr 含 `template`
5. `python -c "from pptx import Presentation; p=Presentation('/tmp/out.pptx'); print(len(p.slides))"` 能解析生成的 PPTX

## 🧪 测试 seam
- shell：直接跑 CLI 命令并断言 exit code、slide 数
- pytest：可加一个集成测试（不强求）

## 🚧 阻塞
- 依赖：[001] scaffold `skills/json-to-ppt/` 目录（必须先有 `scripts/`）
- 依赖：[002] warning + continue 错误路径（CLI 依赖 `build_pptx()` 的新行为）

## 🤖 建议 Agent
Pi
理由：`api-endpoint` 类（CLI 入口 + 参数解析 + 调用下游）

## 🌿 分支
`ticket/003-build-pptx-cli-shell-with-template`

## 🏃 执行约定
- **直接动手，不要先出计划等确认**——本 ticket 已定稿，实现 → 跑验收命令 → 提交，一次做完
- 分支由 CodeG 管理（`task/<id>`）：不要自建分支、不要 push、不要改 `.workflow/`
- 验收命令必须真跑，并把**实际输出**贴回结果；跑不起来就直说，不要声称通过

## 🧭 上下文
- 复用的库函数：`skills/images-to-editable-pptx/scripts/build_editable_pptx.py:478` 的 `build_pptx(spec, output_path, base_dir, allow_full_bleed_images=False)`
- 模板加载：`from pptx import Presentation; tpl = Presentation(template_path)`
- 模板追加模式的关键：先用 `Presentation(template_path)` 拿到 `slide_layouts`，**直接调** `build_pptx()` 会有问题——它内部 `presentation = Presentation()` 创建新的。最简实现：
  - 无模板：`build_pptx(spec, out, base_dir)`
  - 有模板：手动复制模板 slide 到新 deck，再调旧 `build_pptx` 拿到新 slide 列表（这条路复杂）
  - **更直接**：在 `build_pptx()` 不动的前提下，**自己用模板渲染**——copy `tpl.slides` 的 XML 到 out，模仿 `build_pptx()` 的渲染循环（但这违反"不重写渲染层"原则）
  - 推荐做法：给 `build_pptx()` 加一个 `existing_presentation: Presentation | None = None` 参数——但这又动了旧文件
  - **最简**：本次 ticket 把 `--template` 实现为"先把模板复制成 out，再调 `build_pptx` 把新 slide 追加到 out 末尾"。需要在 `build_editable_pptx.py` 加一个 `append_to_pptx(spec, presentation_path, base_dir)` 函数（轻量），**或**在 T003 里就地实现追加逻辑（直接 `from pptx import Presentation; pres = Presentation(out); pres.slides.add_clone(...)`）。
- ⚠️ 决策点：本次 ticket 必须解决"如何在不动 `build_pptx()` 内部实现的前提下，把生成 slide 追加到模板后面"。三条候选实现：

  | 选项 | 改动范围 | 是否动 `build_editable_pptx.py` | 与 T002 关系 |
  |---|---|---|---|
  | **a** 给 `build_pptx()` 加 `existing_presentation: Presentation \| None = None` 参数；在 CLI 里 `build_pptx(spec, out, base_dir, existing_presentation=Presentation(template))` 一次调用搞定 | CLI + 旧 `build_pptx` 签名 | ✗ **要动** | ❌ **不能并行**——需要先把 T002 的修改扩到同时支持 append |
  | **b** 不动旧文件。CLI 流程：① `Presentation(template).save(out)` 把模板原样复制到 out；② 重新 `Presentation(out)` 打开；③ 模仿 `build_pptx()` 的渲染循环，逐个 `add_slide(layout)` + `add_text(...)`，把 spec slide 追加到末尾 | 仅 CLI（重写渲染循环） | ✓ **不动** | ✅ 可并行——但 T003 自带渲染代码，违反"不重写渲染层"原则 |
  | **c** 不动旧文件。CLI 流程：① 把模板复制成 out；② 重新打开 out；③ **直接调** 旧 `build_pptx()` 但传 `output_path=临时文件`；④ 用 `Presentation(临时文件)` 拿到新 slide，逐个 `out.slides.add_clone(s)` 追加 | 仅 CLI（用 add_clone 复用旧渲染） | ✓ **不动** | ✅ 可并行——`add_clone` 是 python-pptx 原生 API，可行 |

  **✅ 已选定 a**（2026-09-18）：在旧 `build_pptx()` 加 `existing_presentation` 参数。
  - 后果：T003 与 T002 **串行**，T002 的 PR 必须同时包含 `existing_presentation` 支持；INDEX 里 W2/W1 的依赖关系同步收紧。
  - 收益：渲染逻辑零分叉，旧 `--template` 行为（保留母版/配色/layouts）天然成立。
- 旧 CLI `skills/images-to-editable-pptx/scripts/build_editable_pptx.py:551 main()` 的 argparse 模式可参考