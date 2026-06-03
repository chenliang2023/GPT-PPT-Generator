from __future__ import annotations

import base64
import binascii
import json
import os
import re
import threading
import time
import uuid
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any

import requests
from flask import Flask, Response, jsonify, request, send_file, send_from_directory


ROOT = Path(__file__).resolve().parent
OUTPUT_ROOT = ROOT / "outputs"
MAX_SLIDES = 50
MAX_REFERENCE_IMAGES_PER_SLIDE = 3
MAX_REFERENCE_IMAGE_DATA_URL_LENGTH = 12_000_000
OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

app = Flask(__name__)
JOBS: dict[str, dict[str, Any]] = {}
JOBS_LOCK = threading.Lock()

STYLE_PRESETS = {
    "modern-tech": {
        "name": "现代科技",
        "prompt": "Modern technology keynote style, deep navy background, teal and cyan accents, crisp information hierarchy, refined glow, spacious 16:9 composition.",
    },
    "minimal-business": {
        "name": "极简商务",
        "prompt": "Minimal premium business presentation, warm white background, charcoal typography, restrained accent color, generous whitespace, editorial 16:9 layout.",
    },
    "dark-luxury": {
        "name": "深色高级",
        "prompt": "Dark luxury presentation style, near-black background, subtle gold accents, elegant typography, cinematic lighting, premium restrained composition.",
    },
    "colorful-creative": {
        "name": "多彩创意",
        "prompt": "Colorful creative presentation style, bold geometric forms, energetic but controlled palette, playful visual rhythm, clear hierarchy, polished 16:9 layout.",
    },
    "data-report": {
        "name": "数据报告",
        "prompt": "Professional data report presentation, light neutral background, blue-green accents, precise chart-like visual language, structured grid, highly readable 16:9 layout.",
    },
    "chinese-red": {
        "name": "中国红",
        "prompt": "Contemporary Chinese red presentation style, rich red and warm ivory palette, modern editorial composition, subtle cultural visual cues, formal and polished 16:9 layout.",
    },
}

LANGUAGE_PRESETS = {
    "zh-cn": {
        "name": "简体中文",
        "prompt": "Use Simplified Chinese for all visible slide text. Do not use Traditional Chinese, Japanese, or other languages.",
    },
    "zh-tw": {
        "name": "繁体中文",
        "prompt": "Use Traditional Chinese for all visible slide text. Do not use Simplified Chinese, Japanese, or other languages.",
    },
    "en": {
        "name": "English",
        "prompt": "Use English for all visible slide text. Do not use any other language.",
    },
    "ja": {
        "name": "日本語",
        "prompt": "Use Japanese for all visible slide text. Do not use any other language.",
    },
    "ko": {
        "name": "한국어",
        "prompt": "Use Korean for all visible slide text. Do not use any other language.",
    },
    "fr": {
        "name": "Français",
        "prompt": "Use French for all visible slide text. Do not use any other language.",
    },
    "de": {
        "name": "Deutsch",
        "prompt": "Use German for all visible slide text. Do not use any other language.",
    },
    "es": {
        "name": "Español",
        "prompt": "Use Spanish for all visible slide text. Do not use any other language.",
    },
}


def clean_text(value: Any, limit: int = 20_000) -> str:
    return str(value or "").strip()[:limit]


def safe_filename(value: str, fallback: str = "ppt-images") -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "-", value).strip(" .-")
    return value[:80] or fallback


def normalize_endpoint(base_url: str, protocol: str = "images") -> str:
    base_url = base_url.strip().rstrip("/")
    if not base_url:
        base_url = "https://api.openai.com/v1"
    for suffix in ("/images/generations", "/chat/completions"):
        if base_url.endswith(suffix):
            base_url = base_url[: -len(suffix)]
            break
    if not base_url.endswith("/v1"):
        base_url = f"{base_url}/v1"
    suffix = "/chat/completions" if protocol == "chat" else "/images/generations"
    return f"{base_url}{suffix}"


def compose_prompt(
    global_style: str,
    page_prompt: str,
    index: int,
    total: int,
    language_instruction: str = "",
) -> str:
    return f"""
Create slide {index} of {total} as one polished presentation slide image.
Aspect ratio: 16:9 landscape.

Overall presentation style:
{global_style or "Clean modern presentation design, strong visual hierarchy, refined spacing."}

This slide's content and visual direction:
{page_prompt}

Output language requirement:
{language_instruction or "Use the same language as this slide's content prompt."}

Keep the visual language consistent with the full deck. Make the slide readable,
well composed, and presentation-ready.
If the slide contains text, follow the output language requirement exactly.
Do not drift into another language. Render any explicitly provided text verbatim
when possible. Do not invent brands, logos, watermarks, or signatures.
""".strip()


def extract_image_bytes(payload: dict[str, Any], session: requests.Session) -> bytes:
    data = payload.get("data")
    if not isinstance(data, list) or not data or not isinstance(data[0], dict):
        raise RuntimeError("图片 API 返回结果中没有可识别的 data")

    item = data[0]
    encoded = (
        item.get("b64_json")
        or item.get("image_base64")
        or item.get("base64")
        or item.get("image")
    )
    if isinstance(encoded, str) and encoded:
        if encoded.startswith("data:") and "," in encoded:
            encoded = encoded.split(",", 1)[1]
        try:
            return base64.b64decode(encoded)
        except (ValueError, binascii.Error) as exc:
            raise RuntimeError("图片 API 返回的 Base64 数据无效") from exc

    image_url = item.get("url")
    if isinstance(image_url, str) and image_url:
        response = session.get(image_url, timeout=300)
        response.raise_for_status()
        return response.content

    raise RuntimeError("图片 API 结果中没有 b64_json 或 url")


def image_bytes_from_string(
    value: str,
    session: requests.Session,
    allow_plain_base64: bool = False,
) -> bytes | None:
    value = value.strip()
    data_uri = re.search(
        r"data:image/[^;]+;base64,([A-Za-z0-9+/=\r\n]+)",
        value,
        flags=re.IGNORECASE,
    )
    if data_uri:
        return base64.b64decode(data_uri.group(1))

    markdown_url = re.search(r"!?\[[^\]]*\]\((https?://[^)]+)\)", value)
    bare_url = re.search(r"https?://[^\s)]+", value)
    image_url = (
        markdown_url.group(1)
        if markdown_url
        else bare_url.group(0)
        if bare_url
        else value
    )
    if image_url.startswith(("https://", "http://")):
        response = session.get(image_url, timeout=300)
        response.raise_for_status()
        return response.content

    if allow_plain_base64 and value:
        try:
            return base64.b64decode(value, validate=True)
        except (ValueError, binascii.Error):
            return None
    return None


def extract_chat_image_bytes(
    payload: dict[str, Any],
    session: requests.Session,
) -> bytes:
    if isinstance(payload.get("data"), list):
        try:
            return extract_image_bytes(payload, session)
        except RuntimeError:
            pass

    choices = payload.get("choices")
    if not isinstance(choices, list) or not choices or not isinstance(choices[0], dict):
        raise RuntimeError("Chat Completions 返回结果中没有可识别的 choices")
    message = choices[0].get("message")
    if not isinstance(message, dict):
        raise RuntimeError("Chat Completions 返回结果中没有可识别的 message")

    candidates: list[tuple[Any, bool]] = []
    images = message.get("images")
    if isinstance(images, list):
        candidates.extend((item, True) for item in images)
    content = message.get("content")
    if isinstance(content, list):
        candidates.extend((item, True) for item in content)
    elif isinstance(content, str):
        candidates.append((content, False))

    for candidate, allow_plain_base64 in candidates:
        if isinstance(candidate, str):
            result = image_bytes_from_string(candidate, session, allow_plain_base64)
            if result:
                return result
            continue
        if not isinstance(candidate, dict):
            continue

        for key in ("b64_json", "image_base64", "base64", "data"):
            value = candidate.get(key)
            if isinstance(value, str):
                result = image_bytes_from_string(value, session, True)
                if result:
                    return result

        for key in ("url", "image_url"):
            value = candidate.get(key)
            if isinstance(value, dict):
                value = value.get("url")
            if isinstance(value, str):
                result = image_bytes_from_string(value, session, False)
                if result:
                    return result

    raise RuntimeError("Chat Completions 响应中没有找到图片 Base64 或 URL")


def translate_api_error(last_response: requests.Response) -> RuntimeError:
    detail = last_response.text[:1500]
    if "Image generation is not enabled for this group" in detail:
        return RuntimeError(
            "中转站已识别请求，但这把 API Key 所属分组未开启图片生成权限"
        )
    if "upstream image connection failed" in detail:
        return RuntimeError(
            "中转站图片上游连接失败，已自动重试 3 次，请稍后重试或更换图片渠道"
        )
    return RuntimeError(f"图片 API 请求失败 ({last_response.status_code}): {detail}")


def request_with_retries(
    endpoint: str,
    headers: dict[str, str],
    payloads: list[dict[str, Any]],
    extractor,
) -> bytes:
    transient_statuses = {429, 500, 502, 503, 504}

    with requests.Session() as session:
        last_response: requests.Response | None = None
        last_exception: requests.RequestException | None = None
        for attempt in range(3):
            should_retry = False
            for payload in payloads:
                try:
                    response = session.post(
                        endpoint, headers=headers, json=payload, timeout=600
                    )
                except requests.RequestException as exc:
                    last_exception = exc
                    should_retry = True
                    break
                last_response = response
                if response.ok:
                    return extractor(response.json(), session)
                if response.status_code in transient_statuses:
                    should_retry = True
                    break
                if response.status_code not in {400, 404, 422}:
                    break
            if should_retry and attempt < 2:
                time.sleep(2**attempt)
                continue
            if not should_retry or attempt == 2:
                break

        if last_response is None and last_exception is not None:
            raise RuntimeError(f"图片 API 连接失败: {last_exception}") from last_exception
        assert last_response is not None
        raise translate_api_error(last_response)


def request_images_api(
    base_url: str,
    api_key: str,
    model: str,
    prompt: str,
    size: str,
    quality: str,
) -> bytes:
    endpoint = normalize_endpoint(base_url, "images")
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    base_payload = {
        "model": model,
        "prompt": prompt,
        "n": 1,
        "size": size,
        "quality": quality,
    }
    payloads = [
        {**base_payload, "output_format": "png"},
        {**base_payload, "response_format": "b64_json"},
        base_payload,
    ]
    return request_with_retries(endpoint, headers, payloads, extract_image_bytes)


def request_chat_api(
    base_url: str,
    api_key: str,
    model: str,
    prompt: str,
    size: str,
    quality: str,
    reference_images: list[str] | None = None,
) -> bytes:
    endpoint = normalize_endpoint(base_url, "chat")
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    chat_prompt = (
        f"{prompt}\n\nRequested output size: {size}. Rendering quality: {quality}."
    )
    if reference_images:
        chat_prompt += (
            "\nUse the attached images as visual or content references for this new "
            "slide. Do not copy any unintended watermark, logo, or signature."
        )
    content: str | list[dict[str, Any]] = chat_prompt
    if reference_images:
        content = [{"type": "text", "text": chat_prompt}]
        content.extend(
            {"type": "image_url", "image_url": {"url": image}}
            for image in reference_images
        )
    payloads = [
        {
            "model": model,
            "messages": [{"role": "user", "content": content}],
        },
        {
            "model": model,
            "messages": [{"role": "user", "content": content}],
            "size": size,
            "quality": quality,
        },
    ]
    return request_with_retries(endpoint, headers, payloads, extract_chat_image_bytes)


def request_image(
    base_url: str,
    api_key: str,
    model: str,
    prompt: str,
    size: str,
    quality: str,
    protocol: str = "auto",
    reference_images: list[str] | None = None,
) -> bytes:
    reference_images = reference_images or []
    if reference_images:
        if protocol == "images":
            raise RuntimeError("参考图片输入需要使用 Chat Completions 协议")
        return request_chat_api(
            base_url,
            api_key,
            model,
            prompt,
            size,
            quality,
            reference_images=reference_images,
        )
    if protocol == "images":
        return request_images_api(base_url, api_key, model, prompt, size, quality)
    if protocol == "chat":
        return request_chat_api(base_url, api_key, model, prompt, size, quality)

    errors: list[str] = []
    for label, requester in (
        ("Images API", request_images_api),
        ("Chat Completions", request_chat_api),
    ):
        try:
            return requester(base_url, api_key, model, prompt, size, quality)
        except RuntimeError as exc:
            errors.append(f"{label}: {exc}")
    raise RuntimeError("自动协议尝试失败；" + "；".join(errors))


def public_job(job: dict[str, Any]) -> dict[str, Any]:
    return {
        "id": job["id"],
        "status": job["status"],
        "message": job["message"],
        "total": job["total"],
        "completed": job["completed"],
        "failed": job["failed"],
        "slides": job["slides"],
        "errors": job["errors"],
        "download_url": job.get("download_url"),
    }


def update_job(job_id: str, **changes: Any) -> None:
    with JOBS_LOCK:
        if job_id in JOBS:
            JOBS[job_id].update(changes)


def create_zip(job_dir: Path, deck_name: str) -> Path:
    zip_path = job_dir / f"{safe_filename(deck_name)}-images.zip"
    with zipfile.ZipFile(zip_path, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for image_path in sorted((job_dir / "images").glob("slide-*.png")):
            archive.write(image_path, arcname=image_path.name)
        archive.write(job_dir / "prompts.json", arcname="prompts.json")
    return zip_path


def run_generation_job(job_id: str, deck: dict[str, Any], api: dict[str, Any]) -> None:
    job_dir = OUTPUT_ROOT / job_id
    images_dir = job_dir / "images"
    images_dir.mkdir(parents=True, exist_ok=True)

    prompt_records = [
        {
            "index": index,
            "page_prompt": slide["prompt"],
            "reference_image_count": len(slide["reference_images"]),
            "final_prompt": compose_prompt(
                deck["global_style"],
                slide["prompt"],
                index,
                len(deck["slides"]),
                deck["language_instruction"],
            ),
            "reference_images": slide["reference_images"],
        }
        for index, slide in enumerate(deck["slides"], start=1)
    ]
    (job_dir / "prompts.json").write_text(
        json.dumps(
            {
                "deck_name": deck["deck_name"],
                "style_preset": deck["style_preset"],
                "style_name": deck["style_name"],
                "global_style": deck["global_style"],
                "language": deck["language"],
                "language_name": deck["language_name"],
                "language_instruction": deck["language_instruction"],
                "api_protocol": api.get("protocol", "auto"),
                "slides": [
                    {
                        key: value
                        for key, value in record.items()
                        if key != "reference_images"
                    }
                    for record in prompt_records
                ],
            },
            ensure_ascii=False,
            indent=2,
        ),
        encoding="utf-8",
    )

    def generate_one(record: dict[str, Any]) -> Path:
        index = record["index"]
        with JOBS_LOCK:
            JOBS[job_id]["slides"][index - 1]["status"] = "generating"
        image_bytes = request_image(
            base_url=api["base_url"],
            api_key=api["api_key"],
            model=api["model"],
            prompt=record["final_prompt"],
            size=api["size"],
            quality=api["quality"],
            protocol=api.get("protocol", "auto"),
            reference_images=record["reference_images"],
        )
        image_path = images_dir / f"slide-{index:02d}.png"
        image_path.write_bytes(image_bytes)
        with JOBS_LOCK:
            slide = JOBS[job_id]["slides"][index - 1]
            slide["status"] = "completed"
            slide["image_url"] = f"/api/jobs/{job_id}/images/{image_path.name}"
            JOBS[job_id]["completed"] += 1
        return image_path

    update_job(job_id, status="generating", message="正在并行调用 image2 生成图片")
    errors: list[str] = []
    with ThreadPoolExecutor(max_workers=api["concurrency"]) as pool:
        futures = {pool.submit(generate_one, record): record for record in prompt_records}
        for future in as_completed(futures):
            record = futures[future]
            try:
                future.result()
            except Exception as exc:  # noqa: BLE001 - surface API errors in the UI
                index = record["index"]
                error = f"第 {index} 页生成失败: {exc}"
                errors.append(error)
                with JOBS_LOCK:
                    JOBS[job_id]["failed"] += 1
                    JOBS[job_id]["slides"][index - 1]["status"] = "failed"
                    JOBS[job_id]["slides"][index - 1]["error"] = str(exc)

    if errors:
        update_job(
            job_id,
            status="failed",
            message="部分图片生成失败，请检查 API 配置后重试",
            errors=errors,
        )
        return

    update_job(job_id, status="packing", message="图片已生成，正在打包 ZIP")
    create_zip(job_dir, deck["deck_name"])
    update_job(
        job_id,
        status="completed",
        message="全部图片已生成，可以下载 ZIP",
        download_url=f"/api/jobs/{job_id}/download",
    )


def validate_payload(payload: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    slides_raw = payload.get("slides")
    if slides_raw is None:
        page_prompts = payload.get("page_prompts")
        if not isinstance(page_prompts, list):
            raise ValueError("每页提示词格式无效")
        slides_raw = [{"prompt": item, "reference_images": []} for item in page_prompts]
    if not isinstance(slides_raw, list):
        raise ValueError("页面卡片格式无效")

    slides: list[dict[str, Any]] = []
    for index, raw in enumerate(slides_raw, start=1):
        if not isinstance(raw, dict):
            raise ValueError(f"第 {index} 页卡片格式无效")
        prompt = clean_text(raw.get("prompt"))
        references_raw = raw.get("reference_images") or []
        if not isinstance(references_raw, list):
            raise ValueError(f"第 {index} 页参考图片格式无效")
        if len(references_raw) > MAX_REFERENCE_IMAGES_PER_SLIDE:
            raise ValueError(
                f"第 {index} 页最多上传 {MAX_REFERENCE_IMAGES_PER_SLIDE} 张参考图片"
            )
        references: list[str] = []
        for image_index, value in enumerate(references_raw, start=1):
            image = clean_text(value, MAX_REFERENCE_IMAGE_DATA_URL_LENGTH + 1)
            if len(image) > MAX_REFERENCE_IMAGE_DATA_URL_LENGTH:
                raise ValueError(f"第 {index} 页第 {image_index} 张参考图片过大")
            if not re.match(r"^data:image/(png|jpeg|jpg|webp);base64,", image, re.I):
                raise ValueError(f"第 {index} 页第 {image_index} 张参考图片格式无效")
            references.append(image)
        if not prompt and not references:
            continue
        slides.append({"prompt": prompt, "reference_images": references})

    if not slides:
        raise ValueError("请至少输入一页提示词")
    if len(slides) > MAX_SLIDES:
        raise ValueError(f"一次最多生成 {MAX_SLIDES} 页")

    api_key = clean_text(payload.get("api_key")) or os.getenv(
        "OPENAI_API_KEY", ""
    ).strip()
    if not api_key:
        raise ValueError("请输入 API Key，或设置 OPENAI_API_KEY 环境变量")

    concurrency = max(1, min(int(payload.get("concurrency") or 2), 6))
    protocol = clean_text(payload.get("protocol"), 20) or "auto"
    if protocol not in {"auto", "images", "chat"}:
        raise ValueError("无效的 API 协议")
    style_preset = clean_text(payload.get("style_preset"), 50) or "modern-tech"
    if style_preset not in STYLE_PRESETS:
        raise ValueError("无效的风格预设")
    style = STYLE_PRESETS[style_preset]
    language = clean_text(payload.get("language"), 50) or "zh-cn"
    custom_language_requirement = clean_text(
        payload.get("custom_language_requirement"), 2000
    )
    if language == "custom":
        if not custom_language_requirement:
            raise ValueError("选择自定义语言要求时，请填写具体要求")
        language_name = "自定义要求"
        language_instruction = custom_language_requirement
    elif language in LANGUAGE_PRESETS:
        language_name = LANGUAGE_PRESETS[language]["name"]
        language_instruction = LANGUAGE_PRESETS[language]["prompt"]
    else:
        raise ValueError("无效的输出语言")
    deck = {
        "deck_name": clean_text(payload.get("deck_name"), 200) or "ppt-images",
        "style_preset": style_preset,
        "style_name": style["name"],
        "global_style": style["prompt"],
        "language": language,
        "language_name": language_name,
        "language_instruction": language_instruction,
        "slides": slides,
    }
    api = {
        "api_key": api_key,
        "base_url": clean_text(payload.get("base_url"), 1000)
        or os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1"),
        "model": clean_text(payload.get("model"), 200)
        or os.getenv("OPENAI_IMAGE_MODEL", "gpt-image-2"),
        "size": clean_text(payload.get("size"), 50) or "2048x1152",
        "quality": clean_text(payload.get("quality"), 30) or "high",
        "concurrency": concurrency,
        "protocol": protocol,
    }
    return deck, api


@app.get("/")
def index() -> Response:
    return Response(INDEX_HTML, mimetype="text/html")


@app.get("/health")
def health() -> Response:
    return jsonify({"ok": True})


@app.get("/api/config")
def config() -> Response:
    return jsonify(
        {
            "base_url": os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1"),
            "model": os.getenv("OPENAI_IMAGE_MODEL", "gpt-image-2"),
            "protocol": os.getenv("OPENAI_IMAGE_PROTOCOL", "auto"),
            "api_key_configured": bool(os.getenv("OPENAI_API_KEY", "").strip()),
        }
    )


@app.post("/api/generate")
def generate() -> Response:
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify({"error": "请求内容必须是 JSON"}), 400
    try:
        deck, api = validate_payload(payload)
    except (TypeError, ValueError) as exc:
        return jsonify({"error": str(exc)}), 400

    job_id = f"{time.strftime('%Y%m%d-%H%M%S')}-{uuid.uuid4().hex[:8]}"
    job = {
        "id": job_id,
        "status": "queued",
        "message": "任务已创建",
        "total": len(deck["slides"]),
        "completed": 0,
        "failed": 0,
        "errors": [],
        "slides": [
            {"index": index, "status": "queued"}
            for index in range(1, len(deck["slides"]) + 1)
        ],
    }
    with JOBS_LOCK:
        JOBS[job_id] = job

    threading.Thread(
        target=run_generation_job,
        args=(job_id, deck, api),
        daemon=True,
    ).start()
    return jsonify({"job_id": job_id}), 202


@app.get("/api/jobs/<job_id>")
def job_status(job_id: str) -> Response:
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if job is None:
            return jsonify({"error": "任务不存在"}), 404
        return jsonify(public_job(job))


@app.get("/api/jobs/<job_id>/images/<filename>")
def job_image(job_id: str, filename: str):
    with JOBS_LOCK:
        if job_id not in JOBS:
            return jsonify({"error": "任务不存在"}), 404
    if not re.fullmatch(r"slide-\d{2}\.png", filename):
        return jsonify({"error": "图片文件名无效"}), 400
    return send_from_directory(OUTPUT_ROOT / job_id / "images", filename)


@app.get("/api/jobs/<job_id>/download")
def job_download(job_id: str):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if job is None:
            return jsonify({"error": "任务不存在"}), 404
        if job["status"] != "completed":
            return jsonify({"error": "图片尚未生成完成"}), 409
    zip_files = list((OUTPUT_ROOT / job_id).glob("*.zip"))
    if not zip_files:
        return jsonify({"error": "找不到 ZIP 文件"}), 404
    return send_file(zip_files[0], as_attachment=True, download_name=zip_files[0].name)


INDEX_HTML = r"""<!doctype html>
<html lang="zh-CN">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Image2 PPT 图片批量生成器</title>
  <style>
    :root {
      --bg: #f0fdfa; --surface: #ffffff; --soft: #f7fffd; --text: #134e4a;
      --muted: #475569; --line: #b9e7df; --primary: #0d9488;
      --primary-dark: #0f766e; --accent: #f97316; --danger: #b91c1c;
    }
    * { box-sizing: border-box; }
    body {
      margin: 0; background: var(--bg); color: var(--text);
      font-family: "Microsoft YaHei", "Segoe UI", Arial, sans-serif; line-height: 1.5;
    }
    button, input, textarea, select { font: inherit; }
    button { cursor: pointer; }
    .shell { width: min(1060px, calc(100% - 32px)); margin: 0 auto; padding: 42px 0 64px; }
    .hero {
      display: grid; grid-template-columns: 1fr auto; gap: 24px; align-items: end;
      padding-bottom: 24px; border-bottom: 1px solid var(--line);
    }
    .eyebrow { margin: 0 0 8px; color: var(--primary-dark); font-size: 13px; font-weight: 700; letter-spacing: .12em; text-transform: uppercase; }
    h1 { margin: 0; font-size: clamp(30px, 5vw, 50px); line-height: 1.08; }
    .hero p { max-width: 720px; margin: 14px 0 0; color: var(--muted); }
    .badge { align-self: start; padding: 8px 12px; border: 1px solid var(--line); border-radius: 999px; background: var(--surface); color: var(--primary-dark); font-size: 13px; font-weight: 700; white-space: nowrap; }
    .section { margin-top: 26px; padding: 22px; border: 1px solid var(--line); border-radius: 14px; background: var(--surface); }
    .section-head { display: flex; justify-content: space-between; gap: 16px; align-items: flex-start; margin-bottom: 18px; }
    h2 { margin: 0; font-size: 20px; }
    .hint { margin: 5px 0 0; color: var(--muted); font-size: 13px; }
    .grid { display: grid; gap: 14px; }
    .grid-2 { grid-template-columns: repeat(2, minmax(0, 1fr)); }
    .grid-3 { grid-template-columns: repeat(3, minmax(0, 1fr)); }
    label { display: grid; gap: 7px; color: var(--text); font-size: 13px; font-weight: 700; }
    input, textarea, select {
      width: 100%; border: 1px solid #a7d8d0; border-radius: 10px; background: #fff;
      color: #0f172a; padding: 10px 12px; outline: none;
      transition: border-color 160ms ease, background 160ms ease;
    }
    textarea { min-height: 120px; resize: vertical; }
    input:focus, textarea:focus, select:focus { border-color: var(--primary); background: var(--soft); }
    .btn { border: 1px solid transparent; border-radius: 10px; padding: 10px 15px; font-weight: 700; transition: background 160ms ease, border-color 160ms ease; }
    .btn-primary { background: var(--accent); color: #fff; }
    .btn-primary:hover { background: #ea580c; }
    .btn-primary:disabled { cursor: not-allowed; opacity: .6; }
    .actions { display: flex; justify-content: space-between; align-items: center; gap: 14px; margin-top: 18px; }
    .note { margin-top: 14px; padding: 12px 14px; border-left: 3px solid var(--primary); background: var(--soft); color: var(--primary-dark); font-size: 13px; }
    .slide-cards { display: grid; gap: 14px; }
    .slide-card { padding: 16px; border: 1px solid var(--line); border-left: 4px solid var(--primary); border-radius: 12px; background: var(--soft); }
    .slide-card-head { display: flex; justify-content: space-between; align-items: center; gap: 12px; margin-bottom: 12px; }
    .slide-card h3 { margin: 0; font-size: 15px; }
    .btn-secondary { background: #fff; border-color: var(--line); color: var(--primary-dark); }
    .btn-secondary:hover { border-color: var(--primary); background: var(--soft); }
    .btn-danger { padding: 6px 10px; background: transparent; color: var(--danger); }
    .btn-danger:hover { background: #fef2f2; }
    .file-input { padding: 8px; background: #fff; }
    .reference-previews { display: flex; flex-wrap: wrap; gap: 8px; margin-top: 8px; }
    .reference-thumb { position: relative; width: 116px; overflow: hidden; border: 1px solid var(--line); border-radius: 8px; background: #fff; }
    .reference-thumb img { display: block; width: 100%; aspect-ratio: 16/9; object-fit: cover; }
    .reference-thumb button { position: absolute; top: 4px; right: 4px; width: 22px; height: 22px; padding: 0; border: 0; border-radius: 50%; background: rgba(15, 23, 42, .78); color: #fff; line-height: 1; }
    .status { display: none; }
    .status.active { display: block; }
    .progress-track { height: 10px; overflow: hidden; border-radius: 999px; background: #ccfbf1; }
    .progress-bar { width: 0; height: 100%; background: var(--primary); transition: width 250ms ease; }
    .status-line { display: flex; justify-content: space-between; gap: 12px; margin: 12px 0 8px; color: var(--muted); font-size: 13px; }
    .preview-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px; margin-top: 18px; }
    .preview { margin: 0; overflow: hidden; border: 1px solid var(--line); border-radius: 10px; background: #fff; }
    .preview img { display: block; width: 100%; aspect-ratio: 16/9; object-fit: cover; }
    .preview figcaption { padding: 8px 10px; color: var(--muted); font-size: 12px; }
    .error-list { margin: 12px 0 0; color: var(--danger); font-size: 13px; white-space: pre-wrap; }
    .download { display: none; text-decoration: none; }
    .download.visible { display: inline-flex; }
    footer { margin-top: 24px; color: var(--muted); font-size: 12px; text-align: center; }
    @media (max-width: 760px) {
      .shell { width: min(100% - 20px, 1060px); padding-top: 24px; }
      .hero, .grid-2, .grid-3 { grid-template-columns: 1fr; }
      .section { padding: 16px; }
      .section-head, .actions { align-items: stretch; flex-direction: column; }
      .preview-grid { grid-template-columns: 1fr; }
    }
    @media (prefers-reduced-motion: reduce) { *, *::before, *::after { transition: none !important; } }
  </style>
</head>
<body>
  <main class="shell">
    <header class="hero">
      <div>
        <p class="eyebrow">Local batch image utility</p>
        <h1>Image2 PPT 图片批量生成器</h1>
        <p>输入 PPT 整体风格和每页提示词，并行调用 GPT image2 或兼容中转站，生成可下载的幻灯片图片包。</p>
      </div>
      <div class="badge">API Key 不写入磁盘</div>
    </header>

    <section class="section">
      <div class="section-head">
        <div><h2>1. API 配置</h2><p class="hint">支持 OpenAI 官方地址、标准 Images API，以及使用 Chat Completions 出图的兼容中转站。</p></div>
      </div>
      <div class="grid grid-3">
        <label>Base URL<input id="baseUrl" type="url" value="https://api.openai.com/v1" placeholder="https://api.openai.com/v1"></label>
        <label>API Key<input id="apiKey" type="password" autocomplete="off" placeholder="留空时使用 OPENAI_API_KEY"></label>
        <label>模型<input id="model" value="gpt-image-2" placeholder="gpt-image-2"></label>
        <label>API 协议
          <select id="protocol">
            <option value="auto" selected>自动（推荐）</option>
            <option value="images">Images API</option>
            <option value="chat">Chat Completions</option>
          </select>
        </label>
        <label>图片尺寸
          <select id="size">
            <option value="2048x1152">2048x1152（推荐 16:9）</option>
            <option value="3840x2160">3840x2160（4K 16:9）</option>
            <option value="1536x1024">1536x1024（兼容性较高）</option>
            <option value="1024x1024">1024x1024（快速测试）</option>
          </select>
        </label>
        <label>质量
          <select id="quality"><option value="high">high</option><option value="medium">medium</option><option value="low">low</option><option value="auto">auto</option></select>
        </label>
        <label>并发页数
          <select id="concurrency"><option value="1">1</option><option value="2" selected>2</option><option value="3">3</option><option value="4">4</option><option value="5">5</option><option value="6">6</option></select>
        </label>
      </div>
    </section>

    <section class="section">
      <div class="section-head">
        <div><h2>2. 风格与页面卡片</h2><p class="hint">选择固定风格，每张卡片可以输入文本并上传最多 3 张参考图片。</p></div>
        <button class="btn btn-secondary" id="addSlide" type="button">添加一页</button>
      </div>
      <div class="grid grid-2">
        <label>图片包名称<input id="deckName" value="image2-ppt" placeholder="用于 ZIP 文件名"></label>
        <label>固定风格
          <select id="stylePreset">
            <option value="modern-tech">现代科技</option>
            <option value="minimal-business">极简商务</option>
            <option value="dark-luxury">深色高级</option>
            <option value="colorful-creative">多彩创意</option>
            <option value="data-report">数据报告</option>
            <option value="chinese-red">中国红</option>
          </select>
        </label>
        <label>输出语言
          <select id="language">
            <option value="zh-cn">简体中文</option>
            <option value="zh-tw">繁体中文</option>
            <option value="en">English</option>
            <option value="ja">日本語</option>
            <option value="ko">한국어</option>
            <option value="fr">Français</option>
            <option value="de">Deutsch</option>
            <option value="es">Español</option>
            <option value="custom">自定义要求</option>
          </select>
        </label>
        <label id="customLanguageLabel" style="display:none">自定义语言要求
          <input id="customLanguageRequirement" placeholder="例如：标题英文，正文简体中文；或中英双语对照。">
        </label>
      </div>
      <p class="note">输出语言会统一注入每一页提示词。有参考图片的页面会自动使用 Chat Completions 多模态格式；所有页面仍会按“并发页数”并行生成。</p>
      <div class="slide-cards" id="slideCards"></div>
      <div class="actions">
        <span class="hint" id="pageCount">0 页</span>
        <button class="btn btn-primary" id="generate" type="button">并行生成图片</button>
      </div>
    </section>

    <section class="section status" id="status">
      <div class="section-head">
        <div><h2>生成进度</h2><p class="hint" id="statusMessage">准备中</p></div>
        <a class="btn btn-primary download" id="download" href="#">下载 ZIP</a>
      </div>
      <div class="status-line"><span id="progressText">0 / 0</span><span id="statusState">queued</span></div>
      <div class="progress-track"><div class="progress-bar" id="progressBar"></div></div>
      <div class="error-list" id="errors"></div>
      <div class="preview-grid" id="previews"></div>
    </section>

    <footer>生成结果保存在本项目的 outputs 目录中。</footer>
  </main>
  <script>
    const slideCardsEl = document.getElementById("slideCards");
    const statusEl = document.getElementById("status");
    const generateBtn = document.getElementById("generate");
    let pollTimer = null;

    function updatePageCount() {
      document.getElementById("pageCount").textContent = `${slideCardsEl.querySelectorAll(".slide-card").length} 页`;
    }

    function renderReferencePreviews(card) {
      const previews = card.querySelector(".reference-previews");
      previews.innerHTML = "";
      (card.referenceImages || []).forEach((source, index) => {
        const wrapper = document.createElement("div");
        wrapper.className = "reference-thumb";
        const image = document.createElement("img");
        image.src = source;
        image.alt = `参考图 ${index + 1}`;
        const remove = document.createElement("button");
        remove.type = "button";
        remove.textContent = "×";
        remove.setAttribute("aria-label", `删除参考图 ${index + 1}`);
        remove.addEventListener("click", () => {
          card.referenceImages.splice(index, 1);
          renderReferencePreviews(card);
        });
        wrapper.append(image, remove);
        previews.appendChild(wrapper);
      });
    }

    function addSlide(values = {}) {
      const card = document.createElement("article");
      card.className = "slide-card";
      card.referenceImages = values.reference_images || [];
      card.innerHTML = `
        <div class="slide-card-head">
          <h3></h3>
          <button class="btn btn-danger remove-slide" type="button">删除</button>
        </div>
        <label>本页提示词
          <textarea class="slide-prompt" placeholder="描述本页主题、需要出现的文字、构图和视觉重点。"></textarea>
        </label>
        <label style="margin-top:12px">参考图片（可选，最多 3 张）
          <input class="file-input" type="file" accept="image/png,image/jpeg,image/webp" multiple>
        </label>
        <div class="reference-previews"></div>
      `;
      card.querySelector(".slide-prompt").value = values.prompt || "";
      card.querySelector(".remove-slide").addEventListener("click", () => {
        card.remove();
        renumberSlides();
      });
      card.querySelector(".file-input").addEventListener("change", async event => {
        const files = [...event.target.files];
        const remaining = Math.max(0, 3 - card.referenceImages.length);
        for (const file of files.slice(0, remaining)) {
          card.referenceImages.push(await readFileAsDataUrl(file));
        }
        event.target.value = "";
        renderReferencePreviews(card);
      });
      slideCardsEl.appendChild(card);
      renderReferencePreviews(card);
      renumberSlides();
    }

    function readFileAsDataUrl(file) {
      return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = () => resolve(reader.result);
        reader.onerror = () => reject(new Error(`无法读取图片：${file.name}`));
        reader.readAsDataURL(file);
      });
    }

    function renumberSlides() {
      [...slideCardsEl.querySelectorAll(".slide-card")].forEach((card, index) => {
        card.querySelector("h3").textContent = `第 ${index + 1} 页`;
      });
      updatePageCount();
    }

    function collectSlides() {
      return [...slideCardsEl.querySelectorAll(".slide-card")].map(card => ({
        prompt: card.querySelector(".slide-prompt").value.trim(),
        reference_images: card.referenceImages || []
      }));
    }

    function updateLanguageRequirementVisibility() {
      const isCustom = document.getElementById("language").value === "custom";
      document.getElementById("customLanguageLabel").style.display = isCustom ? "grid" : "none";
    }
    function renderJob(job) {
      statusEl.classList.add("active");
      document.getElementById("statusMessage").textContent = job.message;
      document.getElementById("statusState").textContent = job.status;
      document.getElementById("progressText").textContent = `${job.completed} / ${job.total}`;
      document.getElementById("progressBar").style.width = `${job.total ? Math.round(job.completed / job.total * 100) : 0}%`;
      document.getElementById("errors").textContent = (job.errors || []).join("\n");
      const previews = document.getElementById("previews");
      previews.innerHTML = "";
      (job.slides || []).filter(slide => slide.image_url).forEach(slide => {
        const figure = document.createElement("figure"); figure.className = "preview";
        const image = document.createElement("img"); image.src = slide.image_url; image.alt = `第 ${slide.index} 页`;
        const caption = document.createElement("figcaption"); caption.textContent = `第 ${slide.index} 页`;
        figure.append(image, caption); previews.appendChild(figure);
      });
      const download = document.getElementById("download");
      if (job.download_url) { download.href = job.download_url; download.classList.add("visible"); }
      else { download.classList.remove("visible"); }
    }
    async function pollJob(jobId) {
      const response = await fetch(`/api/jobs/${jobId}`);
      const job = await response.json();
      renderJob(job);
      if (job.status === "completed" || job.status === "failed") {
        clearInterval(pollTimer); pollTimer = null; generateBtn.disabled = false; generateBtn.textContent = "并行生成图片";
      }
    }
    async function startGeneration() {
      generateBtn.disabled = true; generateBtn.textContent = "正在创建任务...";
      document.getElementById("errors").textContent = ""; document.getElementById("download").classList.remove("visible");
      try {
        const response = await fetch("/api/generate", {
          method: "POST", headers: {"Content-Type": "application/json"},
          body: JSON.stringify({
            base_url: document.getElementById("baseUrl").value.trim(),
            api_key: document.getElementById("apiKey").value.trim(),
            model: document.getElementById("model").value.trim(),
            protocol: document.getElementById("protocol").value,
            size: document.getElementById("size").value,
            quality: document.getElementById("quality").value,
            concurrency: Number(document.getElementById("concurrency").value),
            deck_name: document.getElementById("deckName").value.trim(),
            style_preset: document.getElementById("stylePreset").value,
            language: document.getElementById("language").value,
            custom_language_requirement: document.getElementById("customLanguageRequirement").value.trim(),
            slides: collectSlides()
          })
        });
        const result = await response.json();
        if (!response.ok) throw new Error(result.error || "创建任务失败");
        await pollJob(result.job_id);
        pollTimer = setInterval(() => pollJob(result.job_id), 1500);
      } catch (error) {
        statusEl.classList.add("active"); document.getElementById("errors").textContent = error.message;
        generateBtn.disabled = false; generateBtn.textContent = "并行生成图片";
      }
    }
    async function loadConfig() {
      try {
        const response = await fetch("/api/config"); const config = await response.json();
        document.getElementById("baseUrl").value = config.base_url; document.getElementById("model").value = config.model;
        if (["auto", "images", "chat"].includes(config.protocol)) document.getElementById("protocol").value = config.protocol;
        if (config.api_key_configured) document.getElementById("apiKey").placeholder = "已检测到 OPENAI_API_KEY，可留空";
      } catch (_) {}
    }
    document.getElementById("addSlide").addEventListener("click", () => addSlide());
    document.getElementById("language").addEventListener("change", updateLanguageRequirementVisibility);
    generateBtn.addEventListener("click", startGeneration);
    addSlide({ prompt: "封面：生成式 AI 如何重塑内容生产。左侧使用醒目标题区域，右侧使用抽象智能网络与内容创作图形，保持大留白。" });
    addSlide({ prompt: "展示传统线性内容生产流程向实时智能共创循环转变，使用清晰的流程关系和现代信息图构图。" });
    updateLanguageRequirementVisibility();
    loadConfig();
  </script>
</body>
</html>
"""


if __name__ == "__main__":
    port = int(os.getenv("PORT", "7860"))
    print(f"Image2 PPT 图片批量生成器已启动: http://127.0.0.1:{port}")
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)
