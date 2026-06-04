# Image2 PPT 图片批量生成器

这是一个简单的本地化工具，用 `gpt-image-2` 或兼容中转站并行批量生成 PPT 页面图片。生成完成后，同一个任务会同时提供三种下载方式：

- 图片集合压缩包：`*-images.zip`
- 图片合并纯 PDF：`*-images.pdf`
- 图片版纯 PPTX：`*-image-only.pptx`

项目还包含一个独立 skill，用于把图片参考重建为真正可编辑的 PPTX。

## 启动

```powershell
py -3 -m pip install -r requirements.txt
py -3 app.py
```

然后访问：

```text
http://127.0.0.1:7860
```

Windows 也可以直接运行 `start.bat`。

可选环境变量：

```powershell
$env:OPENAI_API_KEY="..."
$env:OPENAI_BASE_URL="https://api.openai.com/v1"
$env:OPENAI_IMAGE_MODEL="gpt-image-2"
$env:OPENAI_IMAGE_PROTOCOL="auto"
```

页面中填写的 API Key 只保存在当前服务进程内存中，不会写入磁盘。

## 使用

1. 填写 API Key、Base URL、模型、图片尺寸和质量。
2. 选择并发页数，范围为 1-20。建议先从 2-5 开始，避免中转站限流。
3. 选择整体风格，也可以选择“自定义风格”并填写自己的视觉要求。
4. 选择输出语言，或填写自定义语言要求。
5. 为每一页添加卡片，输入提示词，并可上传最多 3 张参考图片。
6. 点击“并行生成图片”，完成后下载 ZIP、PDF 或图片版 PPTX。

生成结果保存在：

```text
outputs/<任务编号>/
```

其中包含 `images/`、`prompts.json`、图片 ZIP、纯 PDF 和图片版 PPTX。

## 中转站配置

在页面的 `API 配置` 中填写：

- `Base URL`：中转站根地址，例如 `http://216.234.142.96:3000` 或 `https://www.dreamfield.top`。通常不要手动加 `/v1/images/generations`。
- `API Key`：中转站提供的 Key。
- `模型`：例如 `gpt-image-2`。
- `API 协议`：建议先选 `自动（推荐）`。

协议含义：

- `自动`：先尝试标准 `/v1/images/generations`，失败后尝试 `/v1/chat/completions`。
- `Images API`：只使用标准图片生成接口。
- `Chat Completions`：使用部分中转站为 `gpt-image-2` 提供的对话出图格式。

包含参考图片的页面会自动使用 Chat Completions 多模态格式：

```json
{
  "content": [
    {"type": "text", "text": "..."},
    {"type": "image_url", "image_url": {"url": "data:image/...;base64,..."}}
  ]
}
```

注意：如果中转站地址是 `http://`，API Key 在网络传输时没有 TLS 加密，建议优先使用 `https://` 中转站。

## 前端文件

前端已经单独放在：

```text
templates/index.html
```

`app.py` 只负责读取这个 HTML 并提供本地 API。后续改界面、样式、表单文案，主要改这个文件即可。

## 可编辑 PPT Skill

独立 skill 位于：

```text
skills/images-to-editable-pptx
```

它会将图片作为视觉参考，用 PowerPoint 原生文字、形状、线条、表格和图表重建真正可编辑的 PPTX，并禁止把整页参考图片直接当作幻灯片内容。
