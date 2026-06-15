from __future__ import annotations

import base64
import binascii
import io
import json
import os
import re
import threading
import time
import uuid
import warnings
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import dataclass
from pathlib import Path
from typing import Any

import requests
from flask import Flask, Response, jsonify, request, send_file, send_from_directory
from PIL import Image
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.util import Inches

# 禁用SSL警告
import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

try:  # Optional: richer PDF reference handling when installed.
    import fitz  # PyMuPDF
except ImportError:  # pragma: no cover - optional dependency
    fitz = None

try:  # Optional fallback for PDF text extraction.
    from pypdf import PdfReader
except ImportError:  # pragma: no cover - optional dependency
    PdfReader = None


ROOT = Path(__file__).resolve().parent
OUTPUT_ROOT = ROOT / "outputs"
STYLE_LIBRARY_PATH = ROOT / "style-library.json"
MAX_SLIDES = 50
MAX_CONCURRENCY = 20
DEFAULT_IMAGE_CONCURRENCY = 1
IMAGE_PAGE_RETRY_ATTEMPTS = 3
MAX_REFERENCE_IMAGES_PER_SLIDE = 3
MAX_GLOBAL_REFERENCE_IMAGES = 6
MAX_TOTAL_REFERENCE_IMAGES_PER_SLIDE = 8
MAX_REFERENCE_IMAGE_DATA_URL_LENGTH = 12_000_000
MAX_REFERENCE_UPLOADS = 12
MAX_REFERENCE_FILE_BYTES = 25_000_000
MAX_REFERENCE_CONTEXT_LENGTH = 12_000
MAX_CUSTOM_STYLES = 100
MAX_PDF_TEXT_PAGES = 80
DEFAULT_SLIDE_WIDTH_INCHES = 16
DEFAULT_SLIDE_HEIGHT_INCHES = 9
OUTPUT_ROOT.mkdir(parents=True, exist_ok=True)

# Model provider constants
PROVIDER_OPENAI = "openai"
PROVIDER_ANTHROPIC = "anthropic"
SUPPORTED_PROVIDERS = [PROVIDER_OPENAI, PROVIDER_ANTHROPIC]

# Provider preset configurations
PROVIDER_PRESETS = {
    PROVIDER_OPENAI: {
        "name": "OpenAI",
        "default_base_url": "https://api.openai.com/v1",
        "default_models": ["gpt-5.5", "gpt-5.1", "gpt-5", "gpt-4.1", "gpt-4o", "gpt-4o-mini"],
        "api_key_env": "OPENAI_TEXT_API_KEY",
        "base_url_env": "OPENAI_TEXT_BASE_URL",
        "model_env": "OPENAI_TEXT_MODEL",
    },
    PROVIDER_ANTHROPIC: {
        "name": "Anthropic",
        "default_base_url": "https://api.anthropic.com/v1",
        "default_models": ["claude-3-5-sonnet-20241022", "claude-3-5-haiku-20241022", "claude-3-opus-20240229", "claude-3-sonnet-20240229"],
        "api_key_env": "ANTHROPIC_API_KEY",
        "base_url_env": "ANTHROPIC_BASE_URL",
        "model_env": "ANALYSIS_MODEL_NAME",
    },
}

# Known models with file upload support
MODELS_WITH_FILE_SUPPORT = {
    # Anthropic models (Claude 3.5+)
    "claude-3-5-sonnet-20241022": True,
    "claude-3-5-sonnet-20240620": True,
    "claude-3-5-haiku-20241022": True,
    "claude-3-opus-20240229": False,
    "claude-3-sonnet-20240229": False,
    "claude-3-haiku-20240307": False,
    # OpenAI models (GPT-4 Vision and newer)
    "gpt-4o": True,
    "gpt-4o-mini": True,
    "gpt-4-turbo": True,
    "gpt-4-vision-preview": True,
    "gpt-4.1": True,
    "gpt-5": True,
    "gpt-5.1": True,
    "gpt-5.5": True,
    "gpt-5.4": True,  # 用户的模型
    # DeepSeek models
    "deepseek-chat": True,
    "deepseek-reasoner": True,
}


@dataclass
class ReferenceFile:
    """参考文件数据结构"""

    name: str
    file_type: str  # "pdf", "pptx", "image", "text"

    # 策略A：原始文件（优先，用于支持文件上传的模型）
    raw_data: bytes | None = None
    mime_type: str | None = None

    # 策略B：提取的内容（降级，用于不支持文件上传的模型）
    text_content: str | None = None
    extracted_images: list[str] | None = None

    # 元数据
    file_size: int = 0
    summary: str = ""


app = Flask(__name__)
JOBS: dict[str, dict[str, Any]] = {}
JOBS_LOCK = threading.Lock()

STYLE_PRESETS = {
    "paper-grain-editorial": {
        "name": "纸感蓝白信息秩序",
        "prompt": (
            "请把画面处理成一种克制、清洁、带有纸面触感的高级视觉：背景不是空白，而是一层有呼吸感的浅色材质，像细腻纸纤维、柔雾、压印或轻微颗粒共同形成的安静底色。"
            "让主体内容从这种底色里浮出来，不追求热闹，而追求清晰、轻盈、可触摸的层次。"
            "画面可以服务于封面、PPT、信息图、报告页、产品页、海报、榜单或数据可视化，但不要被某一种物件限制；无论主题是产品、人物、空间、食物、知识、商业计划还是年度总结，都要把主体的结构秩序、材质边缘、阴影深度和信息节奏作为真正的视觉核心。"
            "文字要成为画面结构的一部分。使用超大尺度的主标题、关键词、汉字、数字或符号作为背景骨架，让它们占据画面的大面积，但通过颗粒化、网点扩散、轻微模糊、半透明遮挡或边缘溶解变成视觉压力，而不是直接喧宾夺主。"
            "前景信息保持精细、干净、可读，像票据、卡片、标签、注释、编号、短句或小型版面系统一样，有明确的行距、层级和留白。"
            "允许竖排英文、窄字距注释、编号、小字说明和手写强调穿插其中，让文本同时承担阅读、装饰和导视三种功能。"
            "标题可以巨大，正文必须克制，强调语要少而准，形成“远看是图形，近看有信息”的双层阅读。"
            "色彩系统以低饱和浅底作为空气和页面温度，以一个深而稳定的主色承担结构、标题、阴影、数据重心或品牌记忆；这个主色可随具体内容转换为更理性、更清洁、更温热、更锋利或更柔软的方向，但必须保持原图那种大面积冷静、少量高识别色、明暗秩序明确的关系。"
            "点睛色只占很小面积，用来承接情绪转折、手写标记、重点词、编号、曲线或微小符号；它应随主题改变语义，例如商业内容更果断、学术内容更冷静、节庆内容更温暖、医疗内容更洁净，儿童内容更轻快，但不要把整张画面染成彩色。"
            "深色负责结构和重量，浅色负责呼吸，暖色或异色负责瞬间注意力，阴影、透明线框、纸边高光和颗粒质感负责纵深。"
            "版式要有明显的前后景关系：后方是巨大而被柔化的字形或信息块，像一面由文字构成的墙；中景可以有透明椭圆、细线、半透边界或轻微光晕，制造被框选的注意力场；前景则放置最重要的主体或信息模块、让它以轻微倾斜、堆叠、弯曲、错位或悬浮的方式形成真实空间感。"
            "不要做普通居中模板，画面重心可以偏右、偏下或被大字推挤出来，左侧或边缘保留一条安静的纵向信息带。"
            "留白必须有控制，密集处要像被压缩的纸页或数据层，稀疏处要像空气；所有元素之间保持精密距离，既有商业海报的完成度，也有编辑设计的冷静。"
            "请在下面这个具体任务中，让画面自然长成它需要的形式：根据我提供的主题、文字、数据、品牌、产品或素材，生成一张具有蓝白颗粒文字背景、轻盈前景层次、少量内容响应式点睛色、细腻纸感与高级信息秩序的视觉作品。"
        ),
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


def style_slug(value: str) -> str:
    slug = re.sub(r"[^a-z0-9]+", "-", value.lower()).strip("-")
    return slug[:40] or "style"


def load_custom_styles() -> list[dict[str, str]]:
    if not STYLE_LIBRARY_PATH.exists():
        return []
    try:
        raw = json.loads(STYLE_LIBRARY_PATH.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return []
    items = raw.get("styles") if isinstance(raw, dict) else raw
    if not isinstance(items, list):
        return []

    styles: list[dict[str, str]] = []
    seen: set[str] = set()
    for item in items:
        if not isinstance(item, dict):
            continue
        style_id = clean_text(item.get("id"), 80)
        name = clean_text(item.get("name"), 80)
        prompt = clean_text(item.get("prompt"), 4000)
        if not style_id or not name or not prompt:
            continue
        if style_id in STYLE_PRESETS or style_id in seen:
            continue
        seen.add(style_id)
        styles.append({"id": style_id, "name": name, "prompt": prompt})
    return styles[:MAX_CUSTOM_STYLES]


def save_custom_styles(styles: list[dict[str, str]]) -> None:
    payload = {"styles": styles[:MAX_CUSTOM_STYLES]}
    STYLE_LIBRARY_PATH.write_text(
        json.dumps(payload, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )


def public_styles() -> list[dict[str, Any]]:
    builtins = [
        {
            "id": style_id,
            "name": style["name"],
            "prompt": style["prompt"],
            "builtin": True,
        }
        for style_id, style in STYLE_PRESETS.items()
    ]
    custom = [
        {
            "id": style["id"],
            "name": style["name"],
            "prompt": style["prompt"],
            "builtin": False,
        }
        for style in load_custom_styles()
    ]
    return builtins + custom


def resolve_style(payload: dict[str, Any]) -> dict[str, str]:
    style_preset = clean_text(payload.get("style_preset"), 80) or "paper-grain-editorial"
    custom_style = clean_text(payload.get("custom_style"), 4000)
    custom_style_name = clean_text(payload.get("custom_style_name"), 80)

    if style_preset == "custom":
        if not custom_style:
            raise ValueError("选择自定义风格时，请填写具体风格要求")
        return {
            "id": "custom",
            "name": custom_style_name or "自定义风格",
            "prompt": custom_style,
        }
    if style_preset in STYLE_PRESETS:
        style = STYLE_PRESETS[style_preset]
        return {"id": style_preset, "name": style["name"], "prompt": style["prompt"]}

    for style in load_custom_styles():
        if style["id"] == style_preset:
            return {"id": style["id"], "name": style["name"], "prompt": style["prompt"]}
    raise ValueError("无效的风格预设")


def resolve_language(payload: dict[str, Any]) -> tuple[str, str, str]:
    language = clean_text(payload.get("language"), 50) or "zh-cn"
    custom_language_requirement = clean_text(
        payload.get("custom_language_requirement"), 2000
    )
    if language == "custom":
        if not custom_language_requirement:
            raise ValueError("选择自定义语言要求时，请填写具体要求")
        return language, "自定义要求", custom_language_requirement
    if language in LANGUAGE_PRESETS:
        preset = LANGUAGE_PRESETS[language]
        return language, preset["name"], preset["prompt"]
    raise ValueError("无效的输出语言")


def validate_image_data_url(value: Any, label: str) -> str:
    image = clean_text(value, MAX_REFERENCE_IMAGE_DATA_URL_LENGTH + 1)
    if len(image) > MAX_REFERENCE_IMAGE_DATA_URL_LENGTH:
        raise ValueError(f"{label}过大")
    if not re.match(r"^data:image/(png|jpeg|jpg|webp);base64,", image, re.I):
        raise ValueError(f"{label}格式无效")
    return image


def validate_reference_images(
    values: Any,
    label: str,
    max_count: int,
) -> list[str]:
    values = values or []
    if not isinstance(values, list):
        raise ValueError(f"{label}格式无效")
    if len(values) > max_count:
        raise ValueError(f"{label}最多 {max_count} 张")
    return [
        validate_image_data_url(value, f"{label}第 {index} 张")
        for index, value in enumerate(values, start=1)
    ]


def safe_filename(value: str, fallback: str = "ppt-images") -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "-", value).strip(" .-")
    return value[:80] or fallback


def normalize_api_base(base_url: str) -> str:
    base_url = base_url.strip().rstrip("/")
    if not base_url:
        base_url = "https://api.openai.com/v1"
    for suffix in ("/images/generations", "/chat/completions", "/models"):
        if base_url.endswith(suffix):
            base_url = base_url[: -len(suffix)]
            break
    if not base_url.endswith("/v1"):
        base_url = f"{base_url}/v1"
    return base_url


def normalize_endpoint(base_url: str, protocol: str = "images") -> str:
    base_url = normalize_api_base(base_url)
    suffix = "/chat/completions" if protocol == "chat" else "/images/generations"
    return f"{base_url}{suffix}"


def compose_prompt(
    global_style: str,
    page_prompt: str,
    index: int,
    total: int,
    language_instruction: str = "",
    global_prompt_addendum: str = "",
    reference_context: str = "",
) -> str:
    prompt = f"""
Create slide {index} of {total} as one polished presentation slide image.
Aspect ratio: 16:9 landscape.

Overall presentation style:
{global_style or "Clean modern presentation design, strong visual hierarchy, refined spacing."}

This slide's content and visual direction:
{page_prompt}

Output language requirement:
{language_instruction or "Use the same language as this slide's content prompt."}
""".strip()

    if global_prompt_addendum:
        prompt += f"""

Deck-level supplementary instructions:
{global_prompt_addendum}
""".rstrip()

    if reference_context:
        prompt += f"""

Uploaded reference context:
{reference_context}
""".rstrip()

    prompt += "\n\n" + """
Keep the visual language consistent with the full deck. Make the slide readable,
well composed, and presentation-ready.
If the slide contains text, follow the output language requirement exactly.
Do not drift into another language. Render any explicitly provided text verbatim
when possible. Do not invent brands, logos, watermarks, or signatures.
""".strip()
    return prompt


def iter_image_candidates(value: Any):
    if isinstance(value, list):
        for item in value:
            yield from iter_image_candidates(item)
        return
    if not isinstance(value, dict):
        return

    yield value
    for key in (
        "data",
        "output",
        "choices",
        "message",
        "content",
        "images",
        "image",
        "image_url",
        "source",
        "result",
    ):
        nested = value.get(key)
        if isinstance(nested, (dict, list)):
            yield from iter_image_candidates(nested)


def image_bytes_from_candidate(
    item: Any,
    session: requests.Session,
    allow_plain_base64: bool = True,
) -> bytes | None:
    if isinstance(item, str):
        return image_bytes_from_string(item, session, allow_plain_base64)
    if not isinstance(item, dict):
        return None

    encoded = (
        item.get("b64_json")
        or item.get("image_base64")
        or item.get("base64")
        or item.get("image")
        or item.get("result")
        or item.get("data")
    )
    if isinstance(encoded, str) and encoded:
        result = image_bytes_from_string(encoded, session, allow_plain_base64)
        if result:
            return result

    image_url = item.get("url") or item.get("file_url")
    if not image_url:
        image_url = item.get("image_url")
    if isinstance(image_url, dict):
        image_url = image_url.get("url")
    if isinstance(image_url, str) and image_url:
        return image_bytes_from_string(image_url, session, False)

    text_content = item.get("text") or item.get("content")
    if isinstance(text_content, str):
        return image_bytes_from_string(text_content, session, False)

    return None


def extract_image_bytes(payload: dict[str, Any], session: requests.Session) -> bytes:
    for candidate in iter_image_candidates(payload):
        result = image_bytes_from_candidate(candidate, session, True)
        if result:
            return result

    raise RuntimeError("图片 API 结果中没有找到图片 Base64 或 URL")


def image_data_url_to_png_bytes(value: str) -> bytes:
    match = re.match(r"^data:image/[^;]+;base64,(.+)$", value, re.I | re.S)
    if not match:
        raise ValueError("图片地址不是有效的 data URL")
    try:
        raw = base64.b64decode(match.group(1), validate=False)
    except (ValueError, binascii.Error) as exc:
        raise ValueError("图片 data URL 中的 Base64 无效") from exc
    try:
        with Image.open(io.BytesIO(raw)) as image:
            if image.mode not in {"RGB", "RGBA"}:
                image = image.convert("RGBA")
            output = io.BytesIO()
            image.save(output, format="PNG")
            return output.getvalue()
    except Exception as exc:  # noqa: BLE001 - surface invalid image details
        raise ValueError(f"图片数据无法识别: {exc}") from exc


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


def collect_text_parts(value: Any) -> list[str]:
    if isinstance(value, str):
        return [value]
    if isinstance(value, list):
        parts: list[str] = []
        for item in value:
            parts.extend(collect_text_parts(item))
        return parts
    if not isinstance(value, dict):
        return []

    parts: list[str] = []
    for key in ("text", "output_text", "content"):
        nested = value.get(key)
        if isinstance(nested, dict) and isinstance(nested.get("value"), str):
            parts.append(nested["value"])
            continue
        parts.extend(collect_text_parts(nested))
    if not parts and isinstance(value.get("value"), str):
        parts.append(value["value"])
    return parts


def clean_joined_text(value: Any) -> str:
    return "\n".join(
        part.strip()
        for part in collect_text_parts(value)
        if part and part.strip()
    ).strip()


def extract_chat_text(payload: dict[str, Any], _session: requests.Session) -> str:
    candidates: list[Any] = []

    choices = payload.get("choices")
    if isinstance(choices, list):
        for choice in choices:
            if not isinstance(choice, dict):
                continue
            message = choice.get("message")
            if isinstance(message, dict):
                candidates.extend(
                    [
                        message.get("content"),
                        message.get("text"),
                        message.get("output_text"),
                    ]
                )
            delta = choice.get("delta")
            if isinstance(delta, dict):
                candidates.append(delta.get("content"))
            candidates.extend(
                [
                    choice.get("content"),
                    choice.get("text"),
                    choice.get("output_text"),
                ]
            )

    output = payload.get("output")
    if isinstance(output, list):
        for item in output:
            if isinstance(item, dict):
                candidates.extend(
                    [item.get("content"), item.get("text"), item.get("output_text")]
                )
            else:
                candidates.append(item)

    data = payload.get("data")
    if isinstance(data, dict):
        candidates.extend(
            [
                data.get("content"),
                data.get("text"),
                data.get("output_text"),
                data.get("result"),
                data.get("response"),
            ]
        )
    elif isinstance(data, list):
        candidates.append(data)

    candidates.extend(
        [
            payload.get("content"),
            payload.get("text"),
            payload.get("output_text"),
            payload.get("result"),
            payload.get("response"),
        ]
    )

    for candidate in candidates:
        text = clean_joined_text(candidate)
        if text:
            return text

    raise RuntimeError("Chat Completions 响应中没有找到文本内容")


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


def translate_text_api_error(last_response: requests.Response) -> RuntimeError:
    detail = last_response.text[:1500]
    return RuntimeError(f"文本模型请求失败 ({last_response.status_code}): {detail}")


def parse_json_from_partial_text(text: str) -> Any | None:
    text = text.lstrip("\ufeff")
    stripped = text.strip()
    if not stripped:
        return None

    if stripped.startswith("data:"):
        data_lines: list[str] = []
        for line in stripped.splitlines():
            line = line.strip()
            if not line.startswith("data:"):
                continue
            value = line[5:].strip()
            if value and value != "[DONE]":
                data_lines.append(value)
        if data_lines:
            try:
                return json.loads("\n".join(data_lines))
            except json.JSONDecodeError:
                pass

    decoder = json.JSONDecoder()
    for start, char in enumerate(text):
        if char not in "[{":
            continue
        try:
            data, end = decoder.raw_decode(text[start:])
        except json.JSONDecodeError:
            continue
        trailing = text[start + end :].strip()
        if not trailing:
            return data
    return None


def read_response_json_early(response: requests.Response) -> tuple[Any, bytes]:
    chunks: list[bytes] = []
    total = 0

    if not hasattr(response, "iter_content"):
        fallback_content = getattr(response, "content", None)
        payload = response.json()
        raw = fallback_content if isinstance(fallback_content, bytes) else b""
        return payload, raw

    for chunk in response.iter_content(chunk_size=65536):
        if not chunk:
            continue
        chunks.append(chunk)
        total += len(chunk)
        raw = b"".join(chunks)
        text = raw.decode(response.encoding or "utf-8", errors="replace")
        parsed = parse_json_from_partial_text(text)
        if parsed is not None:
            print(f"[DEBUG]   ✓ Parsed complete JSON before connection closed ({total:,} bytes)")
            response.close()
            return parsed, raw

    raw = b"".join(chunks)
    text = raw.decode(response.encoding or "utf-8", errors="replace")
    parsed = parse_json_from_partial_text(text)
    if parsed is not None:
        return parsed, raw
    raise RuntimeError(f"API 响应不是有效 JSON: {text[:1500]}")


def read_response_bytes(response: requests.Response) -> bytes:
    if not hasattr(response, "iter_content"):
        content = getattr(response, "content", b"")
        return content if isinstance(content, bytes) else bytes(content)

    chunks: list[bytes] = []
    expected = response.headers.get("content-length")
    expected_length = int(expected) if expected and expected.isdigit() else None
    total = 0
    for chunk in response.iter_content(chunk_size=65536):
        if not chunk:
            continue
        chunks.append(chunk)
        total += len(chunk)
        if expected_length is not None and total >= expected_length:
            response.close()
            break
    return b"".join(chunks)


def detect_provider_from_base_url(base_url: str) -> str:
    """Detect the model provider from the base URL."""
    normalized = base_url.lower().strip().rstrip("/")
    if "anthropic" in normalized:
        return PROVIDER_ANTHROPIC
    return PROVIDER_OPENAI


def get_provider_config(provider: str) -> dict[str, Any] | None:
    """Get provider configuration preset."""
    return PROVIDER_PRESETS.get(provider, PROVIDER_PRESETS[PROVIDER_OPENAI])


def check_model_capabilities(model_name: str, provider: str = None) -> dict[str, Any]:
    """Check model capabilities including file upload support."""
    if provider is None:
        provider = PROVIDER_OPENAI

    # Check known models
    model_lower = model_name.lower()
    for known_model, has_files in MODELS_WITH_FILE_SUPPORT.items():
        if known_model.lower() in model_lower or model_lower in known_model.lower():
            return {
                "supports_files": has_files,
                "supports_images": True,  # Most modern models support images
                "max_file_size": 50_000_000 if provider == PROVIDER_ANTHROPIC else 25_000_000,
                "known_model": True,
            }

    # Default assumptions based on provider
    if provider == PROVIDER_ANTHROPIC:
        # Assume newer Claude models have file support
        return {
            "supports_files": "sonnet" in model_lower or "haiku" in model_lower,
            "supports_images": True,
            "max_file_size": 50_000_000,
            "known_model": False,
        }
    else:  # OpenAI
        return {
            "supports_files": "gpt-4" in model_lower or "gpt-5" in model_lower,
            "supports_images": "gpt-4" in model_lower or "gpt-5" in model_lower,
            "max_file_size": 25_000_000,
            "known_model": False,
        }


def extract_model_ids(payload: dict[str, Any]) -> list[str]:
    candidates = payload.get("data")
    if not isinstance(candidates, list):
        candidates = payload.get("models")
    if not isinstance(candidates, list):
        nested = payload.get("data")
        if isinstance(nested, dict):
            candidates = nested.get("models")
    if not isinstance(candidates, list):
        return []

    models: list[str] = []
    seen: set[str] = set()
    for item in candidates:
        model_id = ""
        if isinstance(item, str):
            model_id = item
        elif isinstance(item, dict):
            raw_id = item.get("id") or item.get("name") or item.get("model")
            if isinstance(raw_id, str):
                model_id = raw_id
        model_id = clean_text(model_id, 200)
        if model_id and model_id not in seen:
            seen.add(model_id)
            models.append(model_id)
    return models


def request_model_list(base_url: str, api_key: str, provider: str = PROVIDER_OPENAI) -> list[str]:
    """Request list of available models from the API."""
    provider = provider or detect_provider_from_base_url(base_url)

    if provider == PROVIDER_ANTHROPIC:
        # Anthropic doesn't have a public models endpoint, return preset models
        return get_provider_config(PROVIDER_ANTHROPIC)["default_models"]

    endpoint = f"{normalize_api_base(base_url)}/models"
    headers = {"Authorization": f"Bearer {api_key}"}
    try:
        with requests.Session() as session:
            response = session.get(endpoint, headers=headers, timeout=60, verify=False)
    except requests.RequestException as exc:
        raise RuntimeError(f"模型列表连接失败: {exc}") from exc
    if not response.ok:
        detail = response.text[:1500]
        raise RuntimeError(f"模型列表请求失败 ({response.status_code}): {detail}")
    try:
        payload = response.json()
    except ValueError as exc:
        raise RuntimeError("模型列表响应不是有效 JSON") from exc
    models = extract_model_ids(payload)
    if not models:
        raise RuntimeError("模型列表响应中没有可识别的模型 ID")
    return models


def resolve_api_settings(payload: dict[str, Any]) -> tuple[str, str, str, str]:
    """Resolve API settings for image or prompt analysis.
    Returns (target, base_url, api_key, provider)
    """
    target = clean_text(payload.get("target"), 20) or "image"
    default_base_url = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")

    # Get provider from payload or detect from base_url
    provider = clean_text(payload.get("provider"), 20) or ""
    if provider and provider not in SUPPORTED_PROVIDERS:
        raise ValueError(f"不支持的提供商: {provider}")

    if target == "prompt":
        # For prompt analysis, check provider first
        if not provider:
            provider = os.getenv("ANALYSIS_MODEL_PROVIDER", "").strip() or PROVIDER_OPENAI
        if provider == PROVIDER_ANTHROPIC:
            api_key = (
                clean_text(payload.get("prompt_api_key"))
                or clean_text(payload.get("api_key"))
                or os.getenv("ANTHROPIC_API_KEY", "").strip()
                or os.getenv("OPENAI_TEXT_API_KEY", "").strip()
                or os.getenv("OPENAI_API_KEY", "").strip()
            )
            base_url = (
                clean_text(payload.get("prompt_base_url"), 1000)
                or clean_text(payload.get("base_url"), 1000)
                or os.getenv("ANTHROPIC_BASE_URL", "").strip()
                or os.getenv("OPENAI_TEXT_BASE_URL", "").strip()
                or get_provider_config(PROVIDER_ANTHROPIC)["default_base_url"]
            )
        else:  # OpenAI
            api_key = (
                clean_text(payload.get("prompt_api_key"))
                or clean_text(payload.get("api_key"))
                or os.getenv("OPENAI_TEXT_API_KEY", "").strip()
                or os.getenv("OPENAI_API_KEY", "").strip()
            )
            base_url = (
                clean_text(payload.get("prompt_base_url"), 1000)
                or clean_text(payload.get("base_url"), 1000)
                or os.getenv("OPENAI_TEXT_BASE_URL", "").strip()
                or default_base_url
            )
        label = "分析模型"
    elif target == "image":
        api_key = (
            clean_text(payload.get("image_api_key"))
            or clean_text(payload.get("api_key"))
            or os.getenv("OPENAI_IMAGE_API_KEY", "").strip()
            or os.getenv("OPENAI_API_KEY", "").strip()
        )
        base_url = (
            clean_text(payload.get("image_base_url"), 1000)
            or clean_text(payload.get("base_url"), 1000)
            or os.getenv("OPENAI_IMAGE_BASE_URL", "").strip()
            or default_base_url
        )
        provider = detect_provider_from_base_url(base_url)
        label = "Image 生成"
    else:
        raise ValueError("无效的 API 测试目标")
    if not api_key:
        raise ValueError(f"请输入 {label} API Key")
    return target, base_url, api_key, provider


def request_with_retries(
    endpoint: str,
    headers: dict[str, str],
    payloads: list[dict[str, Any]],
    extractor,
    error_translator=translate_api_error,
    read_timeout: int = 600,
    attempts: int = 3,
) -> Any:
    transient_statuses = {429, 500, 502, 503, 504}
    request_headers = {
        **headers,
        "Accept": headers.get("Accept", "application/json"),
        "Connection": "close",
    }

    print(f"[DEBUG] request_with_retries: Starting request to {endpoint}")
    print(f"[DEBUG] Payloads count: {len(payloads)}")

    last_response: requests.Response | None = None
    last_exception: requests.RequestException | None = None
    for attempt in range(attempts):
        print(f"[DEBUG] Attempt {attempt + 1}/{attempts}")
        should_retry = False
        for payload_idx, payload in enumerate(payloads):
            with requests.Session() as session:
                # 计算payload大小
                import json
                payload_json = json.dumps(payload)
                payload_size = len(payload_json)
                print(f"[DEBUG]   Trying payload {payload_idx + 1}/{len(payloads)}... (size: {payload_size:,} bytes = {payload_size/1024/1024:.2f} MB)")

                if payload_size > 10 * 1024 * 1024:  # 大于10MB
                    print(f"[WARNING] Payload is very large! May cause timeout or rejection.")

                try:
                    print(f"[DEBUG]   Sending request... (this may take a while)")
                    import time as time_module
                    start_time = time_module.time()
                    response_bytes = b""

                    response = session.post(
                        endpoint,
                        headers=request_headers,
                        json=payload,
                        timeout=(30, read_timeout),
                        verify=False,  # 跳过SSL证书验证
                        stream=True,
                    )

                    elapsed = time_module.time() - start_time
                    content_length = response.headers.get("content-length", "unknown")
                    print(f"[DEBUG]   Response headers received after {elapsed:.2f}s: status={response.status_code}, content-length={content_length}")

                except requests.Timeout as exc:
                    print(f"[ERROR]   ✗ Request TIMEOUT while waiting for API response!")
                    print(f"[ERROR]   This usually means:")
                    print(f"[ERROR]   - Payload too large ({payload_size/1024/1024:.2f} MB)")
                    print(f"[ERROR]   - API server processing too slow or did not flush a response within {read_timeout}s")
                    print(f"[ERROR]   Suggestion: Reduce PDF pages or image quality")
                    last_exception = exc
                    should_retry = True
                    break
                except requests.RequestException as exc:
                    print(f"[ERROR]   ✗ Request exception: {exc}")
                    last_exception = exc
                    should_retry = True
                    break
                last_response = response
                if response.ok:
                    print(f"[DEBUG]   ✓ Success! Extracting data...")
                    try:
                        content_type = response.headers.get("content-type", "")
                        if content_type.lower().startswith("image/"):
                            response_bytes = read_response_bytes(response)
                            print(f"[DEBUG]   ✓ Received direct image response: {len(response_bytes):,} bytes")
                            return response_bytes
                        try:
                            response_payload, response_bytes = read_response_json_early(response)
                        except RuntimeError as exc:
                            print(f"[ERROR]   ✗ {exc}")
                            raise
                        result = extractor(response_payload, session)
                        print(f"[DEBUG]   ✓ Data extracted successfully")
                        return result
                    except Exception as e:
                        print(f"[ERROR]   ✗ Extraction failed: {e}")
                        preview = ""
                        if "response_bytes" in locals():
                            preview = response_bytes[:1500].decode(
                                response.encoding or "utf-8",
                                errors="replace",
                            )
                        print(f"[ERROR]   Response preview: {preview}")
                        raise
                if response.status_code in transient_statuses:
                    print(f"[DEBUG]   Transient error {response.status_code}, will retry")
                    should_retry = True
                    break
                if response.status_code not in {400, 404, 422}:
                    print(f"[DEBUG]   Non-retryable error {response.status_code}")
                    break
                print(f"[DEBUG]   Client error {response.status_code}, trying next payload")
        if should_retry and attempt < attempts - 1:
            wait_time = min(20, 2**attempt)
            print(f"[DEBUG] Waiting {wait_time}s before retry...")
            time.sleep(wait_time)
            continue
        if not should_retry or attempt == attempts - 1:
            break

    if last_response is None and last_exception is not None:
        print(f"[ERROR] ✗ All attempts failed with exception")
        raise RuntimeError(f"API 连接失败: {last_exception}") from last_exception
    assert last_response is not None
    print(f"[ERROR] ✗ All attempts failed, raising error")
    raise error_translator(last_response)


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
    return request_with_retries(
        endpoint,
        headers,
        payloads,
        extract_image_bytes,
        attempts=5,
    )


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
    return request_with_retries(
        endpoint,
        headers,
        payloads,
        extract_chat_image_bytes,
        attempts=5,
    )


def request_chat_text_api(
    base_url: str,
    api_key: str,
    model: str,
    messages: list[dict[str, Any]],
) -> str:
    endpoint = normalize_endpoint(base_url, "chat")
    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }
    base_payload = {
        "model": model,
        "messages": messages,
        "temperature": 0.35,
    }

    # 调试：打印消息结构（不打印完整数据）
    print(f"[DEBUG] Sending to API: {endpoint}")
    print(f"[DEBUG] Model: {model}")
    print(f"[DEBUG] Message structure:")
    for i, msg in enumerate(messages):
        content = msg.get("content")
        if isinstance(content, list):
            print(f"  Message[{i}] ({msg.get('role')}): {len(content)} parts")
            for j, part in enumerate(content):
                part_type = part.get("type")
                if part_type == "text":
                    print(f"    [{j}] type=text, len={len(part.get('text', ''))}")
                elif part_type == "file":
                    file_info = part.get("file", {})
                    print(f"    [{j}] type=file, filename={file_info.get('filename')}, mime={file_info.get('mime_type')}, data_len={len(file_info.get('data', ''))}")
                elif part_type == "image_url":
                    print(f"    [{j}] type=image_url")
        else:
            print(f"  Message[{i}] ({msg.get('role')}): text, len={len(content) if content else 0}")

    payloads = [
        {**base_payload, "response_format": {"type": "json_object"}},
        base_payload,
    ]
    return request_with_retries(
        endpoint,
        headers,
        payloads,
        extract_chat_text,
        error_translator=translate_text_api_error,
        read_timeout=240,
    )


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


def multimodal_user_content(
    text: str,
    reference_images: list[str] | None = None,
) -> str | list[dict[str, Any]]:
    reference_images = reference_images or []
    if not reference_images:
        return text
    content: list[dict[str, Any]] = [{"type": "text", "text": text}]
    content.extend(
        {"type": "image_url", "image_url": {"url": image}}
        for image in reference_images
    )
    return content


def parse_json_object_from_text(text: str) -> dict[str, Any]:
    text = text.strip()
    fenced = re.search(r"```(?:json)?\s*(.*?)```", text, flags=re.I | re.S)
    if fenced:
        text = fenced.group(1).strip()
    try:
        data = json.loads(text)
    except json.JSONDecodeError:
        decoder = json.JSONDecoder()
        starts = [index for index, char in enumerate(text) if char in "[{"]
        data = None
        for start in starts:
            try:
                data, _end = decoder.raw_decode(text[start:])
                break
            except json.JSONDecodeError:
                continue
        if data is None:
            raise ValueError("文本模型没有返回可解析的 JSON") from None
    if isinstance(data, list):
        return {"slides": data}
    if not isinstance(data, dict):
        raise ValueError("文本模型返回的 JSON 顶层必须是对象或数组")
    return data


def normalize_designed_deck(
    data: dict[str, Any],
    fallback_deck_name: str,
) -> dict[str, Any]:
    slides_raw = data.get("slides")
    if not isinstance(slides_raw, list):
        raise ValueError("文本模型返回的 JSON 缺少 slides 数组")
    slides: list[dict[str, str]] = []
    for index, item in enumerate(slides_raw, start=1):
        if isinstance(item, str):
            title = f"第 {index} 页"
            prompt = item
        elif isinstance(item, dict):
            title = clean_text(item.get("title") or f"第 {index} 页", 120)
            prompt = clean_text(
                item.get("prompt")
                or item.get("visual_prompt")
                or item.get("content")
                or item.get("description"),
                6000,
            )
            if not prompt:
                parts = [
                    clean_text(item.get(key), 1200)
                    for key in ("objective", "layout", "visual_direction", "speaker_notes")
                ]
                prompt = "\n".join(part for part in parts if part)
        else:
            continue
        prompt = clean_text(prompt, 6000)
        if prompt:
            slides.append({"title": title, "prompt": prompt})
    if not slides:
        raise ValueError("文本模型没有返回有效的页面提示词")
    return {
        "deck_name": clean_text(data.get("deck_name"), 200)
        or fallback_deck_name
        or "ai-ppt",
        "slides": slides[:MAX_SLIDES],
        "summary": clean_text(data.get("summary"), 2000),
    }


def build_prompt_design_messages(
    brief: str,
    deck_name: str,
    slide_count: int,
    style: dict[str, str],
    language_instruction: str,
    global_prompt_addendum: str,
    reference_context: str,
    reference_images: list[str],
    reference_files: list[dict] | None = None,  # 新增：原始文件
) -> list[dict[str, Any]]:
    user_prompt = f"""
User deck instruction:
{brief}

Target deck name:
{deck_name or "Let the model infer a concise deck name."}

Target slide count:
{slide_count}

Presentation style name:
{style["name"]}

Presentation style prompt:
{style["prompt"]}

Output language requirement:
{language_instruction}

Additional deck-level instructions:
{global_prompt_addendum or "None."}

Uploaded file/reference context:
{reference_context or "None."}

Task:
Design one high-quality image-generation prompt for every slide in this deck.
Each prompt must be directly usable by an image generation model to create a
single 16:9 presentation slide image. Include the intended visible text,
layout, composition, visual hierarchy, chart/table/data requirements when
useful, and the visual direction. Keep the prompts coherent as one deck and
respect the style/reference material.

Return strict JSON only, with this shape:
{{
  "deck_name": "short deck name",
  "summary": "one sentence deck strategy",
  "slides": [
    {{"title": "slide title", "prompt": "complete slide image prompt"}}
  ]
}}
""".strip()
    system_prompt = (
        "You are a senior presentation strategist and art director. "
        "You turn a user's rough instruction into a complete slide-by-slide "
        "prompt plan for generating polished presentation slide images. "
        "Return valid JSON only."
    )

    # 构建多模态消息内容
    reference_files = reference_files or []
    user_content = build_multimodal_content(
        text=user_prompt,
        reference_files=reference_files,  # 优先使用原始文件
        reference_images=reference_images,  # 降级使用提取的图片
    )

    return [
        {"role": "system", "content": system_prompt},
        {"role": "user", "content": user_content},
    ]


def build_multimodal_content(
    text: str,
    reference_files: list[dict] | None = None,
    reference_images: list[str] | None = None,
) -> str | list[dict[str, Any]]:
    """
    构建支持文件上传的多模态消息内容

    PDF 会被提取为文字，避免把大量页面图片塞进中转站请求体。
    """
    reference_files = reference_files or []
    reference_images = reference_images or []

    # 调试日志
    print(f"\n{'='*60}")
    print(f"[DEBUG] build_multimodal_content called:")
    print(f"  - text length: {len(text)} chars")
    print(f"  - reference_files count: {len(reference_files)}")
    print(f"  - reference_images count: {len(reference_images)}")

    if reference_files:
        for i, f in enumerate(reference_files):
            print(f"  - file[{i}]: name={f.get('name')}, mime={f.get('mime_type')}, data_len={len(f.get('data', ''))} bytes")
    print(f"{'='*60}\n")

    # 如果没有文件和图片，返回纯文本
    if not reference_files and not reference_images:
        print(f"[DEBUG] ➜ No files or images, returning plain text\n")
        return text

    # 构建多模态内容
    content: list[dict[str, Any]] = [{"type": "text", "text": text}]
    added_file_text = False

    # 策略A：如果有原始文件，优先提取文本
    if reference_files:
        print(f"[DEBUG] ➜ Processing reference_files...")
        for idx, file in enumerate(reference_files):
            if file.get("data") and file.get("mime_type"):
                mime_type = file["mime_type"]
                file_name = file.get("name", "document")

                # PDF文件：提取文字，不再转图片
                if mime_type == "application/pdf":
                    try:
                        print(f"\n[DEBUG] [{idx}] Extracting PDF text: {file_name}")
                        print(f"[DEBUG]     Original size: {len(file['data'])} bytes (base64)")

                        # 将base64解码为bytes
                        file_bytes = base64.b64decode(file["data"])
                        print(f"[DEBUG]     Decoded size: {len(file_bytes)} bytes (binary)")

                        summary, pdf_text, warnings, _page_count = extract_pdf_text_reference(
                            file_name,
                            file_bytes,
                        )
                        if pdf_text:
                            content[0]["text"] += (
                                "\n\nUploaded PDF extracted text:\n"
                                f"{text_preview(pdf_text, MAX_REFERENCE_CONTEXT_LENGTH)}"
                            )
                            added_file_text = True
                            print(f"[DEBUG]     ✓ Extracted PDF text: {len(pdf_text)} chars")
                        else:
                            print(f"[WARNING]     No PDF text extracted: {summary}")
                        for warning in warnings:
                            print(f"[WARNING]     {warning}")
                    except Exception as exc:
                        # 转换失败，跳过这个文件
                        print(f"[ERROR] ✗ PDF text extraction failed for {file_name}: {exc}")
                        import traceback
                        traceback.print_exc()
                        continue
                else:
                    print(f"[DEBUG] [{idx}] Unsupported file type: {mime_type}, skipping")

        if added_file_text:
            print(f"\n[DEBUG] ➜ Returning text with extracted file content")
            return content[0]["text"]
        if len(content) > 1:  # 有成功添加的图片内容
            print(f"\n[DEBUG] ➜ Total content items: {len(content)} (1 text + {len(content)-1} images)")
            return content
        else:
            print(f"[DEBUG] ➜ No content added from files, falling back\n")

    # 策略B：降级使用提取的图片
    if reference_images:
        print(f"[DEBUG] ➜ Using reference_images (fallback): {len(reference_images)} images")
        for image_url in reference_images:
            content.append({
                "type": "image_url",
                "image_url": {"url": image_url}
            })
        print(f"[DEBUG] ➜ Total content items: {len(content)}\n")
        return content

    print(f"[DEBUG] ➜ Returning plain text\n")
    return text


def convert_pdf_to_images(pdf_bytes: bytes, max_pages: int | None = None) -> list[str]:
    """
    将PDF转换为图片列表（base64编码的data URLs）

    Args:
        pdf_bytes: PDF文件的二进制数据
        max_pages: 最多转换几页，None表示转换全部

    Returns:
        图片的data URL列表
    """
    import time
    start_time = time.time()

    try:
        import fitz  # PyMuPDF
    except ImportError:
        raise RuntimeError("需要安装 PyMuPDF: pip install pymupdf")

    images = []
    try:
        print(f"[DEBUG]     Opening PDF document...")
        doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        total_pages = len(doc)
        page_count = total_pages if max_pages is None else min(total_pages, max_pages)

        print(f"[DEBUG]     ✓ PDF opened successfully")
        print(f"[DEBUG]     Total pages: {total_pages}")
        print(f"[DEBUG]     Converting: {'all pages' if max_pages is None else f'first {page_count} pages'}")
        print(f"[DEBUG]     Matrix scale: 2x (high quality)")
        print(f"[DEBUG]")

        for page_num in range(page_count):
            page_start = time.time()

            page = doc[page_num]
            # 渲染为图片（2倍缩放以提高质量）
            pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
            img_bytes = pix.tobytes("png")

            # 转为data URL
            img_base64 = base64.b64encode(img_bytes).decode()
            data_url = f"data:image/png;base64,{img_base64}"
            images.append(data_url)

            page_time = time.time() - page_start
            print(f"[DEBUG]     Page {page_num + 1}/{page_count}: {len(img_bytes):,} bytes, {page_time:.2f}s")

        doc.close()

        total_time = time.time() - start_time
        total_size = sum(len(img.split(',')[1]) for img in images)  # base64 size
        avg_time = total_time / len(images) if images else 0

        print(f"[DEBUG]")
        print(f"[DEBUG]     ✓ Conversion complete!")
        print(f"[DEBUG]     Total images: {len(images)}")
        print(f"[DEBUG]     Total size: {total_size:,} bytes (base64)")
        print(f"[DEBUG]     Total time: {total_time:.2f}s")
        print(f"[DEBUG]     Average time per page: {avg_time:.2f}s")

    except Exception as exc:
        print(f"[ERROR]     ✗ PDF conversion failed: {exc}")
        import traceback
        traceback.print_exc()
        raise RuntimeError(f"PDF转图片失败: {exc}")

    return images


def image_data_url_from_bytes(data: bytes, fallback_mime: str = "image/png") -> tuple[str, str]:
    with Image.open(io.BytesIO(data)) as image:
        image.thumbnail((1600, 1000))
        has_alpha = image.mode in {"RGBA", "LA"} or (
            image.mode == "P" and "transparency" in image.info
        )
        buffer = io.BytesIO()
        if has_alpha:
            image.save(buffer, format="PNG", optimize=True)
            mime = "image/png"
        else:
            image.convert("RGB").save(buffer, format="JPEG", quality=86, optimize=True)
            mime = "image/jpeg"
        encoded = base64.b64encode(buffer.getvalue()).decode("ascii")
        return f"data:{mime};base64,{encoded}", f"{image.width}x{image.height}"


def text_preview(value: str, limit: int = 2200) -> str:
    value = re.sub(r"\s+", " ", value).strip()
    return value[:limit]


def extract_pdf_text_reference(
    name: str,
    data: bytes,
    max_pages: int = MAX_PDF_TEXT_PAGES,
) -> tuple[str, str, list[str], int]:
    warnings: list[str] = []
    text_parts: list[str] = []
    page_count = 0

    if fitz is not None:
        try:
            doc = fitz.open(stream=data, filetype="pdf")
            page_count = len(doc)
            pages_to_read = min(page_count, max_pages)
            for page_index in range(pages_to_read):
                text = doc[page_index].get_text("text").strip()
                if text:
                    text_parts.append(f"Page {page_index + 1}:\n{text}")
            doc.close()
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"{name}: PDF 文本提取失败：{exc}")
    elif PdfReader is not None:
        try:
            reader = PdfReader(io.BytesIO(data))
            page_count = len(reader.pages)
            pages_to_read = min(page_count, max_pages)
            for page_index, page in enumerate(reader.pages[:pages_to_read]):
                text = (page.extract_text() or "").strip()
                if text:
                    text_parts.append(f"Page {page_index + 1}:\n{text}")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"{name}: PDF 文本提取失败：{exc}")
    else:
        warnings.append(f"{name}: 未安装 PDF 解析依赖，无法提取文字")

    if not text_parts:
        warnings.append(f"{name}: 没有提取到可用文字，可能是扫描版 PDF")

    header = f"{name}: PDF reference"
    if page_count:
        header += f", {page_count} pages"
    if page_count > max_pages:
        header += f", first {max_pages} pages extracted"
        warnings.append(f"{name}: PDF 共 {page_count} 页，仅提取前 {max_pages} 页文字")

    text = "\n\n".join(text_parts)
    summary = header
    if text:
        summary += ". Extracted text: " + text_preview(text, 3000)
    return summary, text, warnings, page_count


def process_pdf_reference(
    name: str,
    data: bytes,
    image_slots: int,
) -> tuple[str, list[str], list[str]]:
    images: list[str] = []
    warnings: list[str] = []
    text_parts: list[str] = []
    page_count = 0

    if fitz is not None:
        try:
            doc = fitz.open(stream=data, filetype="pdf")
            page_count = len(doc)
            for page_index in range(min(page_count, 6)):
                text = doc[page_index].get_text("text").strip()
                if text:
                    text_parts.append(f"Page {page_index + 1}: {text_preview(text, 900)}")
            for page_index in range(min(page_count, image_slots, 3)):
                page = doc[page_index]
                pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
                data_url, _size = image_data_url_from_bytes(pix.tobytes("png"))
                images.append(data_url)
            doc.close()
        except Exception as exc:  # noqa: BLE001 - return a user-visible warning
            warnings.append(f"{name}: PDF 页面预览提取失败：{exc}")
    elif PdfReader is not None:
        try:
            reader = PdfReader(io.BytesIO(data))
            page_count = len(reader.pages)
            for page_index, page in enumerate(reader.pages[:6]):
                text = (page.extract_text() or "").strip()
                if text:
                    text_parts.append(f"Page {page_index + 1}: {text_preview(text, 900)}")
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"{name}: PDF 文本提取失败：{exc}")
    else:
        warnings.append(f"{name}: 未安装 PDF 解析依赖，无法提取参考内容")

    summary = f"{name}: PDF reference"
    if page_count:
        summary += f", {page_count} pages"
    if text_parts:
        summary += ". Extracted text: " + " | ".join(text_parts)
    if images:
        summary += f". Rendered {len(images)} page preview image(s) for visual style reference."
    return summary, images, warnings


def process_pptx_reference(
    name: str,
    data: bytes,
    image_slots: int,
) -> tuple[str, list[str], list[str]]:
    images: list[str] = []
    warnings: list[str] = []
    slide_summaries: list[str] = []
    try:
        presentation = Presentation(io.BytesIO(data))
    except Exception as exc:  # noqa: BLE001
        return f"{name}: PPTX reference could not be parsed.", [], [f"{name}: PPTX 解析失败：{exc}"]

    for slide_index, slide in enumerate(presentation.slides, start=1):
        texts: list[str] = []
        shape_count = 0
        picture_count = 0
        for shape in slide.shapes:
            shape_count += 1
            if getattr(shape, "has_text_frame", False) and shape.text:
                texts.append(text_preview(shape.text, 300))
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                picture_count += 1
                if len(images) < image_slots:
                    try:
                        data_url, _size = image_data_url_from_bytes(shape.image.blob)
                        images.append(data_url)
                    except Exception:  # noqa: BLE001
                        pass
        if slide_index <= 12:
            text = " / ".join(text for text in texts if text)
            slide_summaries.append(
                f"Slide {slide_index}: {shape_count} shapes, {picture_count} pictures"
                + (f", text: {text_preview(text, 700)}" if text else "")
            )
    summary = (
        f"{name}: PPTX reference, {len(presentation.slides)} slides. "
        + " | ".join(slide_summaries)
    )
    if images:
        summary += f". Extracted {len(images)} embedded image(s) for visual reference."
    return summary, images, warnings


def process_reference_file(
    name: str,
    data: bytes,
    content_type: str,
    image_slots: int,
) -> tuple[str, list[str], list[str]]:
    suffix = Path(name).suffix.lower()
    warnings: list[str] = []
    if content_type.startswith("image/") or suffix in {".png", ".jpg", ".jpeg", ".webp"}:
        data_url, size = image_data_url_from_bytes(data)
        return f"{name}: image reference, {size}.", [data_url], warnings
    if suffix == ".pdf" or content_type == "application/pdf":
        return process_pdf_reference(name, data, image_slots)
    if suffix == ".pptx":
        return process_pptx_reference(name, data, image_slots)
    if suffix == ".ppt":
        warnings.append(f"{name}: 旧版 .ppt 暂不支持直接解析，请另存为 .pptx 或 PDF")
        return f"{name}: legacy PowerPoint reference uploaded but not parsed.", [], warnings
    if suffix in {".txt", ".md"} or content_type.startswith("text/"):
        try:
            text = data.decode("utf-8")
        except UnicodeDecodeError:
            text = data.decode("gb18030", errors="ignore")
        return f"{name}: text reference. {text_preview(text, 3000)}", [], warnings
    warnings.append(f"{name}: 不支持的参考文件类型，已忽略内容")
    return f"{name}: unsupported reference file type.", [], warnings


def process_reference_file_v2(
    name: str,
    data: bytes,
    content_type: str,
    model_supports_files: bool,
) -> ReferenceFile:
    """
    处理参考文件（支持原生文件上传）

    Args:
        name: 文件名
        data: 文件二进制数据
        content_type: MIME类型
        model_supports_files: 分析模型是否支持直接上传文件

    Returns:
        ReferenceFile对象，包含原始文件或提取的内容
    """
    suffix = Path(name).suffix.lower()
    file_size = len(data)

    # 图片文件：始终转为base64（图片本身就是视觉内容，不需要提取）
    if content_type.startswith("image/") or suffix in {".png", ".jpg", ".jpeg", ".webp"}:
        data_url, size = image_data_url_from_bytes(data)
        return ReferenceFile(
            name=name,
            file_type="image",
            extracted_images=[data_url],
            file_size=file_size,
            summary=f"{name}: 图片参考，{size}",
        )

    # PDF文件：始终提取文字，避免把整份 PDF 转成大量图片发给中转站。
    if suffix == ".pdf" or content_type == "application/pdf":
        summary, text, warnings, page_count = extract_pdf_text_reference(name, data)
        warning_text = "；".join(warnings)
        return ReferenceFile(
            name=name,
            file_type="pdf",
            text_content=text[:MAX_REFERENCE_CONTEXT_LENGTH],
            extracted_images=[],
            file_size=file_size,
            summary=(
                f"{name}: PDF文件，{page_count or '未知'}页，已提取文字"
                + (f"（{warning_text}）" if warning_text else "")
            )
            if text
            else summary,
        )

    # PPTX文件：类似处理
    if suffix == ".pptx":
        if model_supports_files:
            # 策略A：保留原始PPTX，不提取
            return ReferenceFile(
                name=name,
                file_type="pptx",
                raw_data=data,
                mime_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                file_size=file_size,
                summary=f"{name}: PPTX文件（完整文件，AI直接分析）",
            )
        else:
            # 策略B：提取内容（降级）
            text, images, warnings = process_pptx_reference(name, data, image_slots=6)
            return ReferenceFile(
                name=name,
                file_type="pptx",
                text_content=text,
                extracted_images=images,
                file_size=file_size,
                summary=f"{name}: PPTX文件（模型不支持文件，已提取文字和{len(images)}张图片）",
            )

    # 文本文件：直接读取内容
    if suffix in {".txt", ".md"} or content_type.startswith("text/"):
        try:
            text = data.decode("utf-8")
        except UnicodeDecodeError:
            text = data.decode("gb18030", errors="ignore")
        return ReferenceFile(
            name=name,
            file_type="text",
            text_content=text[:12000],
            file_size=file_size,
            summary=f"{name}: 文本文件",
        )

    # 不支持的类型
    raise ValueError(f"不支持的文件类型: {suffix}")


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


def likely_transient_image_error(exc: Exception) -> bool:
    text = str(exc).lower()
    return any(
        marker in text
        for marker in (
            "ssleoferror",
            "unexpected_eof",
            "eof occurred",
            "connection aborted",
            "connection reset",
            "remote end closed",
            "max retries exceeded",
            "read timed out",
            "timeout",
            "502",
            "503",
            "504",
            "upstream",
        )
    )


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
            "model": slide.get("model") or api["model"],
            "page_reference_image_count": len(slide["reference_images"]),
            "global_reference_image_count": len(deck.get("global_reference_images", [])),
            "final_prompt": compose_prompt(
                deck["global_style"],
                slide["prompt"],
                index,
                len(deck["slides"]),
                deck["language_instruction"],
                deck.get("global_prompt_addendum", ""),
                deck.get("reference_context", ""),
            ),
            "reference_images": [
                *deck.get("global_reference_images", []),
                *slide["reference_images"],
            ],
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
                "global_prompt_addendum": deck.get("global_prompt_addendum", ""),
                "reference_context": deck.get("reference_context", ""),
                "global_reference_image_count": len(deck.get("global_reference_images", [])),
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
        last_exc: Exception | None = None
        for attempt in range(1, IMAGE_PAGE_RETRY_ATTEMPTS + 1):
            with JOBS_LOCK:
                JOBS[job_id]["slides"][index - 1]["status"] = "generating"
                JOBS[job_id]["slides"][index - 1]["error"] = ""
                if attempt > 1:
                    JOBS[job_id]["message"] = (
                        f"第 {index} 页连接不稳定，正在重试 {attempt}/{IMAGE_PAGE_RETRY_ATTEMPTS}"
                    )
            try:
                image_bytes = request_image(
                    base_url=api["base_url"],
                    api_key=api["api_key"],
                    model=record.get("model") or api["model"],
                    prompt=record["final_prompt"],
                    size=api["size"],
                    quality=api["quality"],
                    protocol=api.get("protocol", "auto"),
                    reference_images=record["reference_images"],
                )
                break
            except Exception as exc:  # noqa: BLE001
                last_exc = exc
                if attempt >= IMAGE_PAGE_RETRY_ATTEMPTS or not likely_transient_image_error(exc):
                    raise
                wait_time = min(30, 3 * attempt)
                print(
                    f"[WARNING] Slide {index} transient image error, "
                    f"retrying {attempt + 1}/{IMAGE_PAGE_RETRY_ATTEMPTS} after {wait_time}s: {exc}"
                )
                time.sleep(wait_time)
        else:
            assert last_exc is not None
            raise last_exc
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
    global_reference_images = validate_reference_images(
        payload.get("global_reference_images"),
        "全局参考图片",
        MAX_GLOBAL_REFERENCE_IMAGES,
    )
    global_prompt_addendum = clean_text(
        payload.get("global_prompt_addendum") or payload.get("additional_instructions"),
        6000,
    )
    reference_context = clean_text(
        payload.get("reference_context"),
        MAX_REFERENCE_CONTEXT_LENGTH,
    )

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
        references = validate_reference_images(
            raw.get("reference_images"),
            f"第 {index} 页参考图片",
            MAX_REFERENCE_IMAGES_PER_SLIDE,
        )
        if len(global_reference_images) + len(references) > MAX_TOTAL_REFERENCE_IMAGES_PER_SLIDE:
            raise ValueError(
                f"第 {index} 页全局和本页参考图片合计最多 "
                f"{MAX_TOTAL_REFERENCE_IMAGES_PER_SLIDE} 张"
            )
        model = clean_text(raw.get("model"), 200)
        if not prompt and not references:
            continue
        slides.append({"prompt": prompt, "reference_images": references, "model": model})

    if not slides:
        raise ValueError("请至少输入一页提示词")
    if len(slides) > MAX_SLIDES:
        raise ValueError(f"一次最多生成 {MAX_SLIDES} 页")

    api_key = (
        clean_text(payload.get("image_api_key"))
        or clean_text(payload.get("api_key"))
        or os.getenv("OPENAI_IMAGE_API_KEY", "").strip()
        or os.getenv("OPENAI_API_KEY", "").strip()
    )
    if not api_key:
        raise ValueError("请输入 Image 生成 API Key，或设置 OPENAI_IMAGE_API_KEY / OPENAI_API_KEY 环境变量")

    try:
        concurrency = max(1, min(int(payload.get("concurrency") or DEFAULT_IMAGE_CONCURRENCY), MAX_CONCURRENCY))
    except (TypeError, ValueError):
        raise ValueError("并发页数必须是数字") from None
    protocol = clean_text(payload.get("protocol"), 20) or "auto"
    if protocol not in {"auto", "images", "chat"}:
        raise ValueError("无效的 API 协议")
    style = resolve_style(payload)
    language, language_name, language_instruction = resolve_language(payload)
    deck = {
        "deck_name": clean_text(payload.get("deck_name"), 200) or "ppt-images",
        "style_preset": style["id"],
        "style_name": style["name"],
        "global_style": style["prompt"],
        "language": language,
        "language_name": language_name,
        "language_instruction": language_instruction,
        "global_prompt_addendum": global_prompt_addendum,
        "reference_context": reference_context,
        "global_reference_images": global_reference_images,
        "slides": slides,
    }
    api = {
        "api_key": api_key,
        "base_url": (
            clean_text(payload.get("image_base_url"), 1000)
            or clean_text(payload.get("base_url"), 1000)
            or os.getenv("OPENAI_IMAGE_BASE_URL", "").strip()
            or os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
        ),
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
    default_base_url = os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
    image_base_url = os.getenv("OPENAI_IMAGE_BASE_URL", "").strip() or default_base_url

    # Analysis model configuration
    analysis_provider = os.getenv("ANALYSIS_MODEL_PROVIDER", "").strip() or PROVIDER_OPENAI
    analysis_config = get_provider_config(analysis_provider)

    if analysis_provider == PROVIDER_ANTHROPIC:
        prompt_base_url = os.getenv("ANTHROPIC_BASE_URL", "").strip() or analysis_config["default_base_url"]
        analysis_model = os.getenv("ANALYSIS_MODEL_NAME", "").strip() or os.getenv("OPENAI_TEXT_MODEL", "claude-3-5-sonnet-20241022")
        prompt_key_configured = bool(
            os.getenv("ANTHROPIC_API_KEY", "").strip()
            or os.getenv("OPENAI_TEXT_API_KEY", "").strip()
            or os.getenv("OPENAI_API_KEY", "").strip()
        )
    else:  # OpenAI
        prompt_base_url = os.getenv("OPENAI_TEXT_BASE_URL", "").strip() or default_base_url
        analysis_model = os.getenv("OPENAI_TEXT_MODEL", "gpt-5.5")
        prompt_key_configured = bool(
            os.getenv("OPENAI_TEXT_API_KEY", "").strip()
            or os.getenv("OPENAI_API_KEY", "").strip()
        )

    image_key_configured = bool(
        os.getenv("OPENAI_IMAGE_API_KEY", "").strip()
        or os.getenv("OPENAI_API_KEY", "").strip()
    )

    # Check model capabilities for file upload
    analysis_capabilities = check_model_capabilities(analysis_model, analysis_provider)

    return jsonify(
        {
            "base_url": default_base_url,
            "image_base_url": image_base_url,
            "prompt_base_url": prompt_base_url,
            "model": os.getenv("OPENAI_IMAGE_MODEL", "gpt-image-2"),
            "prompt_model": analysis_model,
            "protocol": os.getenv("OPENAI_IMAGE_PROTOCOL", "auto"),
            "api_key_configured": bool(os.getenv("OPENAI_API_KEY", "").strip()),
            "image_api_key_configured": image_key_configured,
            "prompt_api_key_configured": prompt_key_configured,
            # Analysis model provider configuration
            "analysis_provider": analysis_provider,
            "analysis_providers": {
                provider_id: {
                    "id": provider_id,
                    "name": config["name"],
                    "default_base_url": config["default_base_url"],
                    "default_models": config["default_models"],
                }
                for provider_id, config in PROVIDER_PRESETS.items()
            },
            "analysis_supports_files": analysis_capabilities["supports_files"],
            "analysis_max_file_size": analysis_capabilities["max_file_size"],
            "styles": public_styles(),
            "limits": {
                "max_slides": MAX_SLIDES,
                "max_global_reference_images": MAX_GLOBAL_REFERENCE_IMAGES,
                "max_reference_uploads": MAX_REFERENCE_UPLOADS,
                "max_reference_file_mb": MAX_REFERENCE_FILE_BYTES // 1_000_000,
            },
        }
    )


@app.post("/api/settings/test")
def test_api_settings() -> Response:
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify({"error": "请求内容必须是 JSON"}), 400
    try:
        target, base_url, api_key, provider = resolve_api_settings(payload)
        models = request_model_list(base_url, api_key, provider)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:  # noqa: BLE001 - surface API/proxy errors in UI
        return jsonify({"error": str(exc)}), 502
    return jsonify(
        {
            "ok": True,
            "target": target,
            "provider": provider,
            "base_url": normalize_api_base(base_url),
            "model_count": len(models),
            "models_preview": models[:12],
        }
    )


@app.post("/api/settings/models")
def list_api_models() -> Response:
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify({"error": "请求内容必须是 JSON"}), 400
    try:
        target, base_url, api_key, provider = resolve_api_settings(payload)
        models = request_model_list(base_url, api_key, provider)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    except Exception as exc:  # noqa: BLE001 - surface API/proxy errors in UI
        return jsonify({"error": str(exc)}), 502
    return jsonify(
        {
            "target": target,
            "provider": provider,
            "base_url": normalize_api_base(base_url),
            "models": models,
        }
    )


@app.post("/api/settings/capabilities")
def check_model_capabilities_api() -> Response:
    """Check model capabilities including file upload support."""
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify({"error": "请求内容必须是 JSON"}), 400

    model_name = clean_text(payload.get("model"), 200)
    if not model_name:
        return jsonify({"error": "请提供模型名称"}), 400

    provider = clean_text(payload.get("provider"), 20) or PROVIDER_OPENAI
    if provider not in SUPPORTED_PROVIDERS:
        provider = detect_provider_from_base_url(
            payload.get("base_url", payload.get("prompt_base_url", ""))
        )

    try:
        capabilities = check_model_capabilities(model_name, provider)
        return jsonify(
            {
                "model": model_name,
                "provider": provider,
                **capabilities,
            }
        )
    except Exception as exc:  # noqa: BLE001
        return jsonify({"error": f"模型能力检测失败: {exc}"}), 500


@app.get("/api/styles")
def list_styles() -> Response:
    return jsonify({"styles": public_styles()})


@app.post("/api/styles")
def save_style() -> Response:
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify({"error": "请求内容必须是 JSON"}), 400
    name = clean_text(payload.get("name"), 80)
    prompt = clean_text(payload.get("prompt"), 4000)
    style_id = clean_text(payload.get("id"), 80)
    if not name:
        return jsonify({"error": "请填写风格名称"}), 400
    if not prompt:
        return jsonify({"error": "请填写风格提示词"}), 400

    styles = load_custom_styles()
    existing_ids = {style["id"] for style in styles} | set(STYLE_PRESETS)
    if style_id and style_id in STYLE_PRESETS:
        return jsonify({"error": "内置风格不能被覆盖"}), 400
    if not style_id:
        base = f"user-{style_slug(name)}"
        style_id = base
        suffix = 2
        while style_id in existing_ids:
            style_id = f"{base}-{suffix}"
            suffix += 1

    updated = {"id": style_id, "name": name, "prompt": prompt}
    replaced = False
    for index, style in enumerate(styles):
        if style["id"] == style_id:
            styles[index] = updated
            replaced = True
            break
    if not replaced:
        if len(styles) >= MAX_CUSTOM_STYLES:
            return jsonify({"error": f"最多保存 {MAX_CUSTOM_STYLES} 个自定义风格"}), 400
        styles.append(updated)
    save_custom_styles(styles)
    return jsonify({"style": {**updated, "builtin": False}, "styles": public_styles()})


@app.delete("/api/styles/<style_id>")
def delete_style(style_id: str) -> Response:
    if style_id in STYLE_PRESETS:
        return jsonify({"error": "内置风格不能删除"}), 400
    styles = load_custom_styles()
    next_styles = [style for style in styles if style["id"] != style_id]
    if len(next_styles) == len(styles):
        return jsonify({"error": "风格不存在"}), 404
    save_custom_styles(next_styles)
    return jsonify({"styles": public_styles()})


@app.post("/api/references")
def upload_references() -> Response:
    """
    上传参考文件

    现在会根据分析模型的能力选择处理策略：
    - 如果模型支持文件：保留原始文件
    - 如果模型不支持：提取文字和图片
    """
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "请上传参考文件"}), 400
    if len(files) > MAX_REFERENCE_UPLOADS:
        return jsonify({"error": f"一次最多上传 {MAX_REFERENCE_UPLOADS} 个参考文件"}), 400

    # 获取当前配置的分析模型（从请求参数或环境变量）
    prompt_model = (
        clean_text(request.args.get("prompt_model"), 200)
        or os.getenv("OPENAI_TEXT_MODEL", "").strip()
        or "gpt-5.5"
    )
    prompt_base_url = (
        clean_text(request.args.get("prompt_base_url"), 1000)
        or os.getenv("OPENAI_TEXT_BASE_URL", "").strip()
        or "https://api.openai.com/v1"
    )

    # 检查模型能力
    provider = detect_provider_from_base_url(prompt_base_url)
    capabilities = check_model_capabilities(prompt_model, provider)
    model_supports_files = capabilities.get("supports_files", False)

    reference_files: list[ReferenceFile] = []
    warnings: list[str] = []

    for uploaded in files:
        name = safe_filename(uploaded.filename or "reference", "reference")
        data = uploaded.read(MAX_REFERENCE_FILE_BYTES + 1)

        if len(data) > MAX_REFERENCE_FILE_BYTES:
            warnings.append(f"{name}: 文件超过 {MAX_REFERENCE_FILE_BYTES // 1_000_000}MB，已跳过")
            continue

        try:
            ref_file = process_reference_file_v2(
                name=name,
                data=data,
                content_type=uploaded.mimetype or "",
                model_supports_files=model_supports_files,
            )
            reference_files.append(ref_file)
        except Exception as exc:  # noqa: BLE001
            warnings.append(f"{name}: 处理失败 - {exc}")

    # 构建返回结果
    items = [
        {
            "name": ref.name,
            "file_type": ref.file_type,
            "summary": ref.summary,
            "has_raw_file": ref.raw_data is not None,
            "image_count": len(ref.extracted_images) if ref.extracted_images else 0,
        }
        for ref in reference_files
    ]

    # 分离原始文件和提取内容
    raw_files = [
        {
            "name": ref.name,
            "data": base64.b64encode(ref.raw_data).decode() if ref.raw_data else None,
            "mime_type": ref.mime_type,
        }
        for ref in reference_files
        if ref.raw_data
    ]

    extracted_text = "\n\n".join(ref.text_content for ref in reference_files if ref.text_content)

    extracted_images = []
    for ref in reference_files:
        if ref.extracted_images:
            extracted_images.extend(ref.extracted_images)

    reference_context = text_preview(extracted_text, MAX_REFERENCE_CONTEXT_LENGTH)

    return jsonify(
        {
            "items": items,
            "raw_files": raw_files,  # 新增：原始文件
            "reference_context": reference_context,
            "reference_images": extracted_images[:MAX_GLOBAL_REFERENCE_IMAGES],
            "warnings": warnings,
            "model_supports_files": model_supports_files,  # 告诉前端使用了哪种策略
        }
    )


@app.post("/api/design/prompts")
def design_prompts() -> Response:
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify({"error": "请求内容必须是 JSON"}), 400
    brief = clean_text(payload.get("instruction") or payload.get("brief"), 12_000)
    if not brief:
        return jsonify({"error": "请先输入你的 PPT 生成指令"}), 400
    api_key = (
        clean_text(payload.get("prompt_api_key"))
        or clean_text(payload.get("api_key"))
        or os.getenv("OPENAI_TEXT_API_KEY", "").strip()
        or os.getenv("OPENAI_API_KEY", "").strip()
    )
    if not api_key:
        return jsonify({"error": "请输入 GPT 分析 API Key，或设置 OPENAI_TEXT_API_KEY / OPENAI_API_KEY 环境变量"}), 400
    try:
        slide_count = max(1, min(int(payload.get("slide_count") or 8), MAX_SLIDES))
        style = resolve_style(payload)
        _language, _language_name, language_instruction = resolve_language(payload)
        reference_images = validate_reference_images(
            payload.get("global_reference_images"),
            "全局参考图片",
            MAX_GLOBAL_REFERENCE_IMAGES,
        )
    except (TypeError, ValueError) as exc:
        return jsonify({"error": str(exc)}), 400

    prompt_model = clean_text(payload.get("prompt_model"), 200) or os.getenv(
        "OPENAI_TEXT_MODEL",
        "gpt-5.5",
    )
    base_url = (
        clean_text(payload.get("prompt_base_url"), 1000)
        or clean_text(payload.get("base_url"), 1000)
        or os.getenv("OPENAI_TEXT_BASE_URL", "").strip()
        or os.getenv(
            "OPENAI_BASE_URL",
            "https://api.openai.com/v1",
        )
    )
    deck_name = clean_text(payload.get("deck_name"), 200)
    global_prompt_addendum = clean_text(
        payload.get("global_prompt_addendum") or payload.get("additional_instructions"),
        6000,
    )
    reference_context = clean_text(
        payload.get("reference_context"),
        MAX_REFERENCE_CONTEXT_LENGTH,
    )

    # 新增：获取原始文件
    reference_files = payload.get("reference_files", [])
    print(f"\n{'='*60}")
    print(f"[DEBUG] /api/design/prompts received:")
    print(f"  - brief: {brief[:100]}...")
    print(f"  - slide_count: {slide_count}")
    print(f"  - reference_files: {len(reference_files)} files")
    if reference_files:
        for i, f in enumerate(reference_files):
            print(f"    [{i}] name={f.get('name')}, mime={f.get('mime_type')}, size={len(f.get('data', ''))} bytes (base64)")
    print(f"  - reference_images: {len(reference_images)} images")
    print(f"  - reference_context: {len(reference_context)} chars")
    print(f"{'='*60}\n")

    messages = build_prompt_design_messages(
        brief=brief,
        deck_name=deck_name,
        slide_count=slide_count,
        style=style,
        language_instruction=language_instruction,
        global_prompt_addendum=global_prompt_addendum,
        reference_context=reference_context,
        reference_images=reference_images,
        reference_files=reference_files,  # 新增：传递原始文件
    )
    try:
        text = request_chat_text_api(base_url, api_key, prompt_model, messages)
        data = parse_json_object_from_text(text)
        designed = normalize_designed_deck(data, deck_name)
    except Exception as exc:  # noqa: BLE001 - surface model/API errors in the UI
        return jsonify({"error": str(exc)}), 502
    return jsonify(
        {
            **designed,
            "model": prompt_model,
            "style_preset": style["id"],
            "style_name": style["name"],
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


SINGLES_DIR = OUTPUT_ROOT / "_singles"
SINGLES_DIR.mkdir(parents=True, exist_ok=True)


def _resolve_api_and_deck(payload: dict[str, Any]) -> tuple[dict[str, Any], dict[str, Any]]:
    """Shared validation: returns (deck_context, api)."""
    api_key = (
        clean_text(payload.get("image_api_key"))
        or clean_text(payload.get("api_key"))
        or os.getenv("OPENAI_IMAGE_API_KEY", "").strip()
        or os.getenv("OPENAI_API_KEY", "").strip()
    )
    if not api_key:
        raise ValueError("请输入 Image 生成 API Key，或设置 OPENAI_IMAGE_API_KEY / OPENAI_API_KEY 环境变量")
    protocol = clean_text(payload.get("protocol"), 20) or "auto"
    if protocol not in {"auto", "images", "chat"}:
        raise ValueError("无效的 API 协议")
    style = resolve_style(payload)
    _language, _language_name, language_instruction = resolve_language(payload)
    global_reference_images = validate_reference_images(
        payload.get("global_reference_images"),
        "全局参考图片",
        MAX_GLOBAL_REFERENCE_IMAGES,
    )
    deck = {
        "style_preset": style["id"],
        "global_style": style["prompt"],
        "language_instruction": language_instruction,
        "global_prompt_addendum": clean_text(
            payload.get("global_prompt_addendum") or payload.get("additional_instructions"),
            6000,
        ),
        "reference_context": clean_text(
            payload.get("reference_context"),
            MAX_REFERENCE_CONTEXT_LENGTH,
        ),
        "global_reference_images": global_reference_images,
    }
    api = {
        "api_key": api_key,
        "base_url": (
            clean_text(payload.get("image_base_url"), 1000)
            or clean_text(payload.get("base_url"), 1000)
            or os.getenv("OPENAI_IMAGE_BASE_URL", "").strip()
            or os.getenv("OPENAI_BASE_URL", "https://api.openai.com/v1")
        ),
        "model": clean_text(payload.get("model"), 200)
        or os.getenv("OPENAI_IMAGE_MODEL", "gpt-image-2"),
        "size": clean_text(payload.get("size"), 50) or "2048x1152",
        "quality": clean_text(payload.get("quality"), 30) or "high",
        "protocol": protocol,
    }
    return deck, api


def _get_image_data_url(url_or_path: str) -> str:
    """Convert an internal image URL/path to a base64 data URL string."""
    # Already a data URL
    if url_or_path.startswith("data:"):
        return url_or_path

    # Internal single image: /api/singles/<filename>
    m = re.match(r"/api/singles/(single-[a-f0-9]+\.png)", url_or_path)
    if m:
        path = SINGLES_DIR / m.group(1)
        if path.exists():
            b64 = base64.b64encode(path.read_bytes()).decode()
            return f"data:image/png;base64,{b64}"

    # Internal job image: /api/jobs/<job_id>/images/<filename>
    m = re.match(r"/api/jobs/([^/]+)/images/(slide-\d{2}\.png)", url_or_path)
    if m:
        path = OUTPUT_ROOT / m.group(1) / "images" / m.group(2)
        if path.exists():
            b64 = base64.b64encode(path.read_bytes()).decode()
            return f"data:image/png;base64,{b64}"

    # External URL — download it
    if url_or_path.startswith(("https://", "http://")):
        try:
            r = requests.get(url_or_path, timeout=60, verify=False)
            r.raise_for_status()
            b64 = base64.b64encode(r.content).decode()
            ct = r.headers.get("content-type", "image/png")
            if "jpeg" in ct or "jpg" in ct:
                return f"data:image/jpeg;base64,{b64}"
            return f"data:image/png;base64,{b64}"
        except Exception:
            pass

    raise ValueError("无法解析已有图片地址，请确认图片已成功生成")


@app.post("/api/export/existing")
def export_existing() -> Response:
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify({"error": "请求内容必须是 JSON"}), 400

    slides_raw = payload.get("slides")
    if not isinstance(slides_raw, list):
        return jsonify({"error": "页面数据格式无效"}), 400
    if not slides_raw:
        return jsonify({"error": "没有可导出的页面"}), 400

    slides: list[dict[str, str]] = []
    missing: list[int] = []
    for index, raw in enumerate(slides_raw, start=1):
        if not isinstance(raw, dict):
            return jsonify({"error": f"第 {index} 页数据格式无效"}), 400
        image_url = clean_text(raw.get("image_url"), MAX_REFERENCE_IMAGE_DATA_URL_LENGTH)
        if not image_url:
            missing.append(index)
        slides.append(
            {
                "title": clean_text(raw.get("title"), 200) or f"第 {index} 页",
                "prompt": clean_text(raw.get("prompt"), 20_000),
                "image_url": image_url,
            }
        )
    if missing:
        return jsonify({"error": f"第 {', '.join(map(str, missing))} 页还没有生成图片，无法直接导出"}), 400

    deck_name = clean_text(payload.get("deck_name"), 200) or "image2-ppt"
    job_id = f"{time.strftime('%Y%m%d-%H%M%S')}-export-{uuid.uuid4().hex[:8]}"
    job_dir = OUTPUT_ROOT / job_id
    images_dir = job_dir / "images"
    images_dir.mkdir(parents=True, exist_ok=True)

    job = {
        "id": job_id,
        "status": "packing",
        "message": "正在打包已有页面",
        "total": len(slides),
        "completed": 0,
        "failed": 0,
        "errors": [],
        "slides": [
            {"index": index, "status": "queued"}
            for index in range(1, len(slides) + 1)
        ],
    }
    with JOBS_LOCK:
        JOBS[job_id] = job

    try:
        prompt_records: list[dict[str, Any]] = []
        for index, slide in enumerate(slides, start=1):
            data_url = _get_image_data_url(slide["image_url"])
            image_bytes = image_data_url_to_png_bytes(data_url)
            image_path = images_dir / f"slide-{index:02d}.png"
            image_path.write_bytes(image_bytes)
            public_image_url = f"/api/jobs/{job_id}/images/{image_path.name}"
            prompt_records.append(
                {
                    "index": index,
                    "title": slide["title"],
                    "page_prompt": slide["prompt"],
                    "source_image_url": slide["image_url"],
                    "export_image_url": public_image_url,
                }
            )
            with JOBS_LOCK:
                JOBS[job_id]["completed"] += 1
                JOBS[job_id]["slides"][index - 1] = {
                    "index": index,
                    "status": "completed",
                    "image_url": public_image_url,
                }

        (job_dir / "prompts.json").write_text(
            json.dumps(
                {
                    "deck_name": deck_name,
                    "exported_from_existing_images": True,
                    "slides": prompt_records,
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        exports = create_exports(job_dir, deck_name)
    except Exception as exc:  # noqa: BLE001 - surface export errors in UI
        message = f"已有页面导出失败: {exc}"
        update_job(
            job_id,
            status="failed",
            message=message,
            errors=[message],
        )
        with JOBS_LOCK:
            return jsonify(public_job(JOBS[job_id])), 502

    update_job(
        job_id,
        status="completed",
        message="已有页面已打包，可以下载 ZIP / PDF / PPTX",
        download_url=f"/api/jobs/{job_id}/download",
        downloads={key: f"/api/jobs/{job_id}/download/{key}" for key in exports},
        export_files=exports,
    )
    with JOBS_LOCK:
        return jsonify(public_job(JOBS[job_id]))


@app.post("/api/generate/refine")
def generate_refine() -> Response:
    """Refine an existing slide image based on modification instructions."""
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify({"error": "请求内容必须是 JSON"}), 400
    try:
        deck, api = _resolve_api_and_deck(payload)
    except (TypeError, ValueError) as exc:
        return jsonify({"error": str(exc)}), 400

    slide_index = int(payload.get("index", 1))
    slide_total = int(payload.get("total", 1))
    page_prompt = clean_text(payload.get("prompt"))
    refine_instruction = clean_text(payload.get("refine_instruction"))
    if not refine_instruction:
        return jsonify({"error": "请填写修改要求"}), 400
    existing_image_url = clean_text(payload.get("image_url"))
    if not existing_image_url:
        return jsonify({"error": "请先生成图片后再进行调整"}), 400

    # Convert the existing image to a data URL for reference
    try:
        existing_data_url = _get_image_data_url(existing_image_url)
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400

    # Build the refine prompt
    refine_prompt = f"""
REFINE an existing presentation slide image. Do NOT redesign from scratch.
Only modify the specific elements described in the instructions below.
Keep everything else identical — background, layout, fonts, colors, spacing,
and any unmentioned content must stay the same.

Original slide content and intent:
{page_prompt or "See attached reference image."}

Deck-level supplementary instructions:
{deck.get("global_prompt_addendum") or "None."}

Modification instructions (apply ONLY these changes):
{refine_instruction}

Output one refined slide image reflecting only the requested changes.
Aspect ratio: 16:9 landscape. Presentation-ready quality.
""".strip()

    # Always use Chat Completions (needs image input)
    image_id = uuid.uuid4().hex
    image_path = SINGLES_DIR / f"single-{image_id}.png"
    try:
        image_bytes = request_chat_api(
            base_url=api["base_url"],
            api_key=api["api_key"],
            model=api["model"],
            prompt=refine_prompt,
            size=api["size"],
            quality=api["quality"],
            reference_images=[existing_data_url],
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 502

    image_path.write_bytes(image_bytes)
    return jsonify({
        "image_url": f"/api/singles/{image_path.name}",
        "status": "completed",
    })


@app.post("/api/generate/single")
def generate_single() -> Response:
    """Generate one slide synchronously and return its image URL."""
    payload = request.get_json(silent=True)
    if not isinstance(payload, dict):
        return jsonify({"error": "请求内容必须是 JSON"}), 400
    try:
        deck, api = _resolve_api_and_deck(payload)
    except (TypeError, ValueError) as exc:
        return jsonify({"error": str(exc)}), 400

    slide_index = int(payload.get("index", 1))
    slide_total = int(payload.get("total", 1))
    page_prompt = clean_text(payload.get("prompt"))
    if not page_prompt:
        return jsonify({"error": "请填写本页提示词"}), 400

    try:
        page_references = validate_reference_images(
            payload.get("reference_images"),
            "本页参考图片",
            MAX_REFERENCE_IMAGES_PER_SLIDE,
        )
    except ValueError as exc:
        return jsonify({"error": str(exc)}), 400
    references = [*deck.get("global_reference_images", []), *page_references]
    if len(references) > MAX_TOTAL_REFERENCE_IMAGES_PER_SLIDE:
        return jsonify(
            {
                "error": (
                    "全局和本页参考图片合计最多 "
                    f"{MAX_TOTAL_REFERENCE_IMAGES_PER_SLIDE} 张"
                )
            }
        ), 400

    final_prompt = compose_prompt(
        deck["global_style"],
        page_prompt,
        slide_index,
        slide_total,
        deck["language_instruction"],
        deck.get("global_prompt_addendum", ""),
        deck.get("reference_context", ""),
    )

    image_id = uuid.uuid4().hex
    image_path = SINGLES_DIR / f"single-{image_id}.png"
    try:
        image_bytes = request_image(
            base_url=api["base_url"],
            api_key=api["api_key"],
            model=api["model"],
            prompt=final_prompt,
            size=api["size"],
            quality=api["quality"],
            protocol=api["protocol"],
            reference_images=references if references else None,
        )
    except Exception as exc:
        return jsonify({"error": str(exc)}), 502

    image_path.write_bytes(image_bytes)
    return jsonify({
        "image_url": f"/api/singles/{image_path.name}",
        "status": "completed",
    })


@app.get("/api/singles/<filename>")
def serve_single(filename: str):
    if not re.fullmatch(r"single-[a-f0-9]+\.png", filename):
        return jsonify({"error": "图片文件名无效"}), 400
    return send_from_directory(str(SINGLES_DIR), filename)


INDEX_HTML_PATH = ROOT / "templates" / "index.html"


if __name__ == "__main__":
    port = int(os.getenv("PORT", "7860"))
    print(f"Image2 PPT 图片批量生成器已启动: http://127.0.0.1:{port}")
    app.run(host="127.0.0.1", port=port, debug=False, threaded=True)
