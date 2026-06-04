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
from PIL import Image
from pptx import Presentation
from pptx.util import Inches


ROOT = Path(__file__).resolve().parent
OUTPUT_ROOT = ROOT / "outputs"
MAX_SLIDES = 50
MAX_CONCURRENCY = 20
MAX_REFERENCE_IMAGES_PER_SLIDE = 3
MAX_REFERENCE_IMAGE_DATA_URL_LENGTH = 12_000_000
DEFAULT_SLIDE_WIDTH_INCHES = 16
DEFAULT_SLIDE_HEIGHT_INCHES = 9
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
        "downloads": job.get("downloads", {}),
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


def slide_image_paths(job_dir: Path) -> list[Path]:
    return sorted((job_dir / "images").glob("slide-*.png"))


def create_image_pdf(job_dir: Path, deck_name: str) -> Path:
    image_paths = slide_image_paths(job_dir)
    if not image_paths:
        raise RuntimeError("No slide images found for PDF export")

    pdf_path = job_dir / f"{safe_filename(deck_name)}-images.pdf"
    opened_images: list[Image.Image] = []
    try:
        for image_path in image_paths:
            with Image.open(image_path) as image:
                opened_images.append(image.convert("RGB"))
        first, rest = opened_images[0], opened_images[1:]
        first.save(pdf_path, "PDF", save_all=True, append_images=rest, resolution=150.0)
    finally:
        for image in opened_images:
            image.close()
    return pdf_path


def create_image_pptx(job_dir: Path, deck_name: str) -> Path:
    image_paths = slide_image_paths(job_dir)
    if not image_paths:
        raise RuntimeError("No slide images found for PPTX export")

    pptx_path = job_dir / f"{safe_filename(deck_name)}-image-only.pptx"
    presentation = Presentation()
    presentation.slide_width = Inches(DEFAULT_SLIDE_WIDTH_INCHES)
    presentation.slide_height = Inches(DEFAULT_SLIDE_HEIGHT_INCHES)
    blank_layout = presentation.slide_layouts[6]

    for image_path in image_paths:
        slide = presentation.slides.add_slide(blank_layout)
        slide.shapes.add_picture(
            str(image_path),
            0,
            0,
            width=presentation.slide_width,
            height=presentation.slide_height,
        )

    presentation.save(pptx_path)
    return pptx_path


def create_exports(job_dir: Path, deck_name: str) -> dict[str, str]:
    zip_path = create_zip(job_dir, deck_name)
    pdf_path = create_image_pdf(job_dir, deck_name)
    pptx_path = create_image_pptx(job_dir, deck_name)
    return {
        "zip": zip_path.name,
        "pdf": pdf_path.name,
        "pptx": pptx_path.name,
    }


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

    update_job(job_id, status="packing", message="图片已生成，正在导出 ZIP / PDF / PPTX")
    exports = create_exports(job_dir, deck["deck_name"])
    update_job(
        job_id,
        status="completed",
        message="全部图片已生成，可以下载 ZIP / PDF / PPTX",
        download_url=f"/api/jobs/{job_id}/download",
        downloads={
            key: f"/api/jobs/{job_id}/download/{key}" for key in exports
        },
        export_files=exports,
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

    concurrency = max(1, min(int(payload.get("concurrency") or 2), MAX_CONCURRENCY))
    protocol = clean_text(payload.get("protocol"), 20) or "auto"
    if protocol not in {"auto", "images", "chat"}:
        raise ValueError("无效的 API 协议")
    style_preset = clean_text(payload.get("style_preset"), 50) or "modern-tech"
    custom_style = clean_text(payload.get("custom_style"), 4000)
    if style_preset == "custom":
        if not custom_style:
            raise ValueError("选择自定义风格时，请填写具体风格要求")
        style = {"name": "自定义风格", "prompt": custom_style}
    elif style_preset in STYLE_PRESETS:
        style = STYLE_PRESETS[style_preset]
    else:
        raise ValueError("无效的风格预设")
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
    return Response(INDEX_HTML_PATH.read_text(encoding="utf-8"), mimetype="text/html")


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
    return job_download_kind(job_id, "zip")


@app.get("/api/jobs/<job_id>/download/<kind>")
def job_download_kind(job_id: str, kind: str):
    if kind not in {"zip", "pdf", "pptx"}:
        return jsonify({"error": "下载类型无效"}), 400
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if job is None:
            return jsonify({"error": "任务不存在"}), 404
        if job["status"] != "completed":
            return jsonify({"error": "图片尚未生成完成"}), 409
        filename = job.get("export_files", {}).get(kind)
    if not filename:
        pattern = {"zip": "*.zip", "pdf": "*.pdf", "pptx": "*.pptx"}[kind]
        files = list((OUTPUT_ROOT / job_id).glob(pattern))
        if files:
            filename = files[0].name
    if not filename:
        return jsonify({"error": "找不到导出文件"}), 404
    path = OUTPUT_ROOT / job_id / filename
    if not path.exists():
        return jsonify({"error": "导出文件不存在"}), 404
    return send_file(path, as_attachment=True, download_name=path.name)


INDEX_HTML_PATH = ROOT / "templates" / "index.html"


if __name__ == "__main__":
    port = int(os.getenv("PORT", "7860"))
    print(f"Image2 PPT 图片批量生成器已启动: http://127.0.0.1:{port}")
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)
