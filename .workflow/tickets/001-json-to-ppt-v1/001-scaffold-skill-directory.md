<!-- status: todo -->

# [001] scaffold `skills/json-to-ppt/` directory

## 📌 Spec 引用
来源：[`.workflow/specs/json-to-ppt-skill.md`](../specs/json-to-ppt-skill.md) 的「💡 解决方案」「Q1 暴露方式 b」

## 🎯 任务
让仓库里出现 `skills/json-to-ppt/` 目录及其子目录骨架（`scripts/`、`references/`、`agents/`）。

现在仓库里只有 `skills/images-to-editable-pptx/`，新 skill 没有任何目录。任务完成后：

- `skills/json-to-ppt/SKILL.md` 占位文件存在
- `skills/json-to-ppt/scripts/`、`references/`、`agents/` 三个子目录存在，且各含一个空的 `.gitkeep`
- 仓库里其他位置零修改

> ⚠️ git 不追踪空目录——没有 `.gitkeep` 的话这三个子目录合并后会消失，后续 ticket（T003 的 CLI、T004 的 spec-format、T005 的测试）就失去脚手架。所以 `.gitkeep` 是本 ticket 的必需产物。

## ✅ 验收标准
1. `ls skills/json-to-ppt/` 列出 `SKILL.md  agents  references  scripts` 四个条目
2. `cat skills/json-to-ppt/SKILL.md` 输出 placeholder 内容 `# json-to-ppt (placeholder)`
3. `git status --short` 只显示新增的 `SKILL.md` 与三个 `.gitkeep`，不出现其他文件的改动
4. `python -c "import pathlib; p=pathlib.Path('skills/json-to-ppt'); assert p.is_dir() and (p/'scripts').is_dir() and (p/'references').is_dir() and (p/'agents').is_dir()"` 退出码 0
5. `git ls-files skills/json-to-ppt | wc -l` 输出 `4`（证明目录结构真的进了索引，而不只是工作区里存在）

## 🧪 测试 seam
- shell：`ls`、`cat`、`python -c` 内联断言
- 不需要 pytest

## 🚧 阻塞
- 无（可立即开始）

## 🤖 建议 Agent
Pi
理由：纯脚手架创建（mkdir + 占位文件），属于 `config` / `scaffold` 类

## 🌿 分支
`ticket/001-scaffold-skill-directory`

## 🧭 上下文
- 工作区根：**仓库根目录**（在哪个 agent 上跑就用哪边的根，不要把它写进产物里）
- 参考结构：`skills/images-to-editable-pptx/`（已有 `SKILL.md`、`scripts/`、`references/`、`agents/` 四件套）
- 这个 ticket **只**创建目录结构，不写任何真实文件内容（后续 ticket 会各自填）
- `.gitkeep` 用空文件即可（`touch`），不要往里面写说明文字
- 不要新建 `tests/` 子目录——测试在仓库根的 `tests/` 目录，由 T005 负责