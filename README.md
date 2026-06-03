# Image2 PPT 图片批量生成器

这是一个简单的单文件本地服务。它使用 `gpt-image-2` 或兼容中转站，并行批量生成每一页 PPT 图片。

## 启动

```powershell
py -3 -m pip install -r requirements.txt
py -3 app.py
```

然后访问 `http://127.0.0.1:7860`。Windows 也可以直接运行 `start.bat`。

可选环境变量：

```powershell
$env:OPENAI_API_KEY="..."
$env:OPENAI_BASE_URL="https://api.openai.com/v1"
$env:OPENAI_IMAGE_MODEL="gpt-image-2"
$env:OPENAI_IMAGE_PROTOCOL="auto"
```

页面中填写的 API Key 只保存在当前服务进程内存中，不会写入磁盘。

## 使用

1. 填写 API Key、Base URL、模型、图片尺寸和并发数。
2. 选择固定 PPT 风格和输出语言。
3. 为每一页添加卡片，输入提示词，并可上传最多 3 张参考图片。
4. 选择并发页数，点击“并行生成图片”，完成后下载 ZIP。

输出语言支持简体中文、繁体中文、英语、日语、韩语、法语、德语、
西班牙语，以及自定义要求，例如“标题英文，正文简体中文”或“中英双语对照”。

API 协议支持：

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

每张页面卡片都是独立生成任务，因此可以按页面并发处理。

生成结果保存在 `outputs/<任务编号>/`，其中包含图片和最终提示词记录。

## 可编辑 PPT skill

独立 skill 位于 `skills/images-to-editable-pptx`。它将图片作为视觉参考，使用 PowerPoint 原生文字、形状、线条、表格和图表重建真正可编辑的 PPT，禁止把整页参考图片作为幻灯片内容。
