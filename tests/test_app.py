from __future__ import annotations

import base64
import io
import json
import zipfile

import app
import pytest
from pptx import Presentation


PNG_BYTES = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMC"
    "AO+/p9sAAAAASUVORK5CYII="
)


class FakeSession:
    def get(self, *_args, **_kwargs):
        raise AssertionError("URL download should not be used for Base64 responses")


class FakeDownloadResponse:
    content = b"downloaded-image"

    def raise_for_status(self):
        return None


class FakeDownloadSession:
    def get(self, *_args, **_kwargs):
        return FakeDownloadResponse()


class DeniedResponse:
    ok = False
    status_code = 403
    text = '{"error":{"message":"Image generation is not enabled for this group"}}'


class DeniedSession:
    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return None

    def post(self, *_args, **_kwargs):
        return DeniedResponse()


class TransientResponse:
    def __init__(self, status_code, text="", payload=None):
        self.status_code = status_code
        self.text = text
        self._payload = payload or {}
        self.ok = 200 <= status_code < 300

    def json(self):
        return self._payload


class TransientThenSuccessSession:
    calls = 0

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return None

    def post(self, *_args, **_kwargs):
        type(self).calls += 1
        if type(self).calls < 3:
            return TransientResponse(502, "upstream error")
        return TransientResponse(
            200,
            payload={"data": [{"b64_json": base64.b64encode(b"ok").decode("ascii")}]},
        )

    def get(self, *_args, **_kwargs):
        raise AssertionError("URL download should not be used for Base64 responses")


class AlwaysUpstreamFailureSession:
    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return None

    def post(self, *_args, **_kwargs):
        return TransientResponse(
            502,
            '{"error":{"message":"upstream image connection failed, please retry later"}}',
        )


class ChatSuccessSession:
    endpoints = []
    payloads = []

    def __enter__(self):
        return self

    def __exit__(self, *_args):
        return None

    def post(self, endpoint, *_args, **kwargs):
        type(self).endpoints.append(endpoint)
        type(self).payloads.append(kwargs["json"])
        encoded = base64.b64encode(b"chat-ok").decode("ascii")
        return TransientResponse(
            200,
            payload={
                "choices": [
                    {
                        "message": {
                            "content": f"data:image/png;base64,{encoded}"
                        }
                    }
                ]
            },
        )

    def get(self, *_args, **_kwargs):
        raise AssertionError("URL download should not be used for Base64 responses")


def test_compose_prompt_contains_global_and_page_prompt() -> None:
    prompt = app.compose_prompt(
        "Minimal technology style",
        "Show a workflow",
        2,
        5,
        "Use English for all visible slide text.",
    )
    assert "Minimal technology style" in prompt
    assert "Show a workflow" in prompt
    assert "slide 2 of 5" in prompt
    assert "Use English for all visible slide text." in prompt
    assert "Do not invent brands" in prompt


def test_compose_prompt_includes_deck_addendum_and_reference_context() -> None:
    prompt = app.compose_prompt(
        "Minimal business style",
        "Show pricing options",
        1,
        3,
        "Use English.",
        "Use restrained charts and avoid stock photos.",
        "Uploaded PDF uses a dense board-report layout.",
    )
    assert "Use restrained charts" in prompt
    assert "Uploaded PDF uses a dense board-report layout." in prompt


def test_extract_image_bytes_from_base64() -> None:
    expected = b"image-bytes"
    payload = {"data": [{"b64_json": base64.b64encode(expected).decode("ascii")}]}
    assert app.extract_image_bytes(payload, FakeSession()) == expected


def test_extract_chat_image_bytes_from_data_url() -> None:
    expected = b"chat-image"
    encoded = base64.b64encode(expected).decode("ascii")
    payload = {
        "choices": [
            {
                "message": {
                    "content": f"![generated](data:image/png;base64,{encoded})"
                }
            }
        ]
    }
    assert app.extract_chat_image_bytes(payload, FakeSession()) == expected


def test_extract_chat_image_bytes_from_message_images_url() -> None:
    payload = {
        "choices": [
            {
                "message": {
                    "images": [
                        {
                            "type": "image_url",
                            "image_url": {"url": "https://example.com/image.png"},
                        }
                    ]
                }
            }
        ]
    }
    assert (
        app.extract_chat_image_bytes(payload, FakeDownloadSession())
        == b"downloaded-image"
    )


def test_normalize_endpoint_supports_chat_protocol() -> None:
    assert (
        app.normalize_endpoint("https://example.com", "chat")
        == "https://example.com/v1/chat/completions"
    )
    assert (
        app.normalize_endpoint("https://example.com/v1/images/generations", "chat")
        == "https://example.com/v1/chat/completions"
    )


def test_request_image_translates_group_permission_error(monkeypatch) -> None:
    monkeypatch.setattr(app.requests, "Session", DeniedSession)
    with pytest.raises(RuntimeError, match="所属分组未开启图片生成权限"):
        app.request_image(
            "https://example.com/v1/images/generations",
            "test-key",
            "gpt-image-2",
            "test prompt",
            "1024x1024",
            "low",
        )


def test_request_image_retries_transient_errors(monkeypatch) -> None:
    TransientThenSuccessSession.calls = 0
    monkeypatch.setattr(app.requests, "Session", TransientThenSuccessSession)
    monkeypatch.setattr(app.time, "sleep", lambda _seconds: None)
    result = app.request_image(
        "https://example.com/v1/images/generations",
        "test-key",
        "gpt-image-2",
        "test prompt",
        "1024x1024",
        "low",
    )
    assert result == b"ok"
    assert TransientThenSuccessSession.calls == 3


def test_request_image_translates_upstream_failure(monkeypatch) -> None:
    monkeypatch.setattr(app.requests, "Session", AlwaysUpstreamFailureSession)
    monkeypatch.setattr(app.time, "sleep", lambda _seconds: None)
    with pytest.raises(RuntimeError, match="图片上游连接失败"):
        app.request_image(
            "https://example.com/v1/images/generations",
            "test-key",
            "gpt-image-2",
            "test prompt",
            "1024x1024",
            "low",
        )


def test_request_image_chat_protocol_uses_model_and_messages(monkeypatch) -> None:
    ChatSuccessSession.endpoints = []
    ChatSuccessSession.payloads = []
    monkeypatch.setattr(app.requests, "Session", ChatSuccessSession)
    result = app.request_image(
        "https://example.com",
        "test-key",
        "gpt-image-2",
        "test prompt",
        "1024x1024",
        "low",
        protocol="chat",
    )
    assert result == b"chat-ok"
    assert ChatSuccessSession.endpoints == [
        "https://example.com/v1/chat/completions"
    ]
    assert ChatSuccessSession.payloads[0]["model"] == "gpt-image-2"
    assert ChatSuccessSession.payloads[0]["messages"][0]["role"] == "user"
    assert "test prompt" in ChatSuccessSession.payloads[0]["messages"][0]["content"]


def test_request_image_reference_uses_chat_multimodal_content(monkeypatch) -> None:
    ChatSuccessSession.endpoints = []
    ChatSuccessSession.payloads = []
    monkeypatch.setattr(app.requests, "Session", ChatSuccessSession)
    reference = "data:image/png;base64," + base64.b64encode(b"reference").decode("ascii")
    result = app.request_image(
        "https://example.com",
        "test-key",
        "gpt-image-2",
        "test prompt",
        "1024x1024",
        "low",
        protocol="auto",
        reference_images=[reference],
    )
    assert result == b"chat-ok"
    content = ChatSuccessSession.payloads[0]["messages"][0]["content"]
    assert isinstance(content, list)
    assert content[0]["type"] == "text"
    assert content[1] == {"type": "image_url", "image_url": {"url": reference}}


def test_validate_payload_maps_style_preset_and_reference_images(monkeypatch) -> None:
    reference = "data:image/png;base64," + base64.b64encode(b"reference").decode("ascii")
    deck, api = app.validate_payload(
        {
            "api_key": "test-key",
            "style_preset": "minimal-business",
            "language": "en",
            "slides": [{"prompt": "Slide one", "reference_images": [reference]}],
        }
    )
    assert deck["style_name"] == "极简商务"
    assert deck["language_name"] == "English"
    assert "Use English" in deck["language_instruction"]
    assert deck["slides"][0]["reference_images"] == [reference]
    assert api["protocol"] == "auto"


def test_validate_payload_requires_custom_language_requirement() -> None:
    with pytest.raises(ValueError, match="自定义语言要求"):
        app.validate_payload(
            {
                "api_key": "test-key",
                "language": "custom",
                "slides": [{"prompt": "Slide one", "reference_images": []}],
            }
        )


def test_validate_payload_accepts_custom_language_requirement() -> None:
    deck, _api = app.validate_payload(
        {
            "api_key": "test-key",
            "language": "custom",
            "custom_language_requirement": "Use bilingual English and Chinese text.",
            "slides": [{"prompt": "Slide one", "reference_images": []}],
        }
    )
    assert deck["language_name"] == "自定义要求"
    assert deck["language_instruction"] == "Use bilingual English and Chinese text."


def test_validate_payload_accepts_custom_style() -> None:
    deck, _api = app.validate_payload(
        {
            "api_key": "test-key",
            "style_preset": "custom",
            "custom_style": "Cyberpunk blue and purple glassmorphism keynote style.",
            "slides": [{"prompt": "Slide one", "reference_images": []}],
        }
    )
    assert deck["style_name"] == "自定义风格"
    assert deck["global_style"] == "Cyberpunk blue and purple glassmorphism keynote style."


def test_validate_payload_accepts_saved_style_and_global_references(tmp_path, monkeypatch) -> None:
    style_library = tmp_path / "style-library.json"
    style_library.write_text(
        json.dumps(
            {
                "styles": [
                    {
                        "id": "user-board-report",
                        "name": "董事会报告",
                        "prompt": "Quiet executive board report style.",
                    }
                ]
            }
        ),
        encoding="utf-8",
    )
    monkeypatch.setattr(app, "STYLE_LIBRARY_PATH", style_library)
    reference = "data:image/png;base64," + base64.b64encode(b"reference").decode("ascii")
    deck, _api = app.validate_payload(
        {
            "api_key": "test-key",
            "style_preset": "user-board-report",
            "global_prompt_addendum": "Avoid decorative gradients.",
            "reference_context": "Reference deck uses compact tables.",
            "global_reference_images": [reference],
            "slides": [
                {
                    "prompt": "Slide one",
                    "model": "gpt-image-2-custom",
                    "reference_images": [],
                }
            ],
        }
    )
    assert deck["style_name"] == "董事会报告"
    assert deck["global_style"] == "Quiet executive board report style."
    assert deck["global_prompt_addendum"] == "Avoid decorative gradients."
    assert deck["reference_context"] == "Reference deck uses compact tables."
    assert deck["global_reference_images"] == [reference]
    assert deck["slides"][0]["model"] == "gpt-image-2-custom"


def test_validate_payload_requires_custom_style() -> None:
    with pytest.raises(ValueError, match="自定义风格"):
        app.validate_payload(
            {
                "api_key": "test-key",
                "style_preset": "custom",
                "slides": [{"prompt": "Slide one", "reference_images": []}],
            }
        )


def test_validate_payload_allows_concurrency_up_to_20() -> None:
    _deck, api = app.validate_payload(
        {
            "api_key": "test-key",
            "concurrency": 20,
            "slides": [{"prompt": "Slide one", "reference_images": []}],
        }
    )
    assert api["concurrency"] == 20


def test_validate_payload_clamps_concurrency_above_20() -> None:
    _deck, api = app.validate_payload(
        {
            "api_key": "test-key",
            "concurrency": 99,
            "slides": [{"prompt": "Slide one", "reference_images": []}],
        }
    )
    assert api["concurrency"] == app.MAX_CONCURRENCY


def test_health_endpoint() -> None:
    client = app.app.test_client()
    response = client.get("/health")
    assert response.status_code == 200
    assert response.get_json() == {"ok": True}


def test_style_library_api_saves_and_deletes_style(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(app, "STYLE_LIBRARY_PATH", tmp_path / "style-library.json")
    client = app.app.test_client()

    created = client.post(
        "/api/styles",
        json={"name": "Investor Brief", "prompt": "Clean investor pitch style."},
    )
    assert created.status_code == 200
    style = created.get_json()["style"]
    assert style["name"] == "Investor Brief"
    assert style["builtin"] is False

    listed = client.get("/api/styles")
    assert listed.status_code == 200
    assert any(item["id"] == style["id"] for item in listed.get_json()["styles"])

    deleted = client.delete(f"/api/styles/{style['id']}")
    assert deleted.status_code == 200
    assert not any(item["id"] == style["id"] for item in deleted.get_json()["styles"])


def test_design_prompts_endpoint_returns_model_slides(monkeypatch) -> None:
    monkeypatch.setattr(
        app,
        "request_chat_text_api",
        lambda *_args, **_kwargs: json.dumps(
            {
                "deck_name": "Generated Deck",
                "summary": "A concise plan.",
                "slides": [{"title": "Cover", "prompt": "Create a polished cover."}],
            }
        ),
    )
    client = app.app.test_client()
    response = client.post(
        "/api/design/prompts",
        json={
            "api_key": "test-key",
            "instruction": "Create a product launch deck.",
            "slide_count": 1,
            "prompt_model": "gpt-5.5",
        },
    )
    assert response.status_code == 200
    payload = response.get_json()
    assert payload["deck_name"] == "Generated Deck"
    assert payload["model"] == "gpt-5.5"
    assert payload["slides"] == [{"title": "Cover", "prompt": "Create a polished cover."}]


def test_reference_upload_accepts_image() -> None:
    client = app.app.test_client()
    response = client.post(
        "/api/references",
        data={
            "files": (io.BytesIO(PNG_BYTES), "reference.png"),
        },
        content_type="multipart/form-data",
    )
    assert response.status_code == 200
    payload = response.get_json()
    assert payload["items"][0]["name"] == "reference.png"
    assert payload["items"][0]["image_count"] == 1
    assert payload["reference_images"][0].startswith("data:image/")


def test_generate_requires_key(monkeypatch) -> None:
    monkeypatch.delenv("OPENAI_API_KEY", raising=False)
    client = app.app.test_client()
    response = client.post(
        "/api/generate",
        json={
            "api_key": "",
            "page_prompts": ["Slide one"],
        },
    )
    assert response.status_code == 400


def test_generation_job_writes_export_files(tmp_path, monkeypatch) -> None:
    monkeypatch.setattr(app, "OUTPUT_ROOT", tmp_path)
    monkeypatch.setattr(app, "request_image", lambda **_kwargs: PNG_BYTES)
    job_id = "test-job"
    with app.JOBS_LOCK:
        app.JOBS[job_id] = {
            "id": job_id,
            "status": "queued",
            "message": "任务已创建",
            "total": 2,
            "completed": 0,
            "failed": 0,
            "errors": [],
            "slides": [
                {"index": 1, "status": "queued"},
                {"index": 2, "status": "queued"},
            ],
        }

    app.run_generation_job(
        job_id,
        {
            "deck_name": "test-deck",
            "style_preset": "modern-tech",
            "style_name": "现代科技",
            "global_style": "Minimal",
            "language": "en",
            "language_name": "English",
            "language_instruction": "Use English for all visible slide text.",
            "slides": [
                {"prompt": "Slide one", "reference_images": []},
                {"prompt": "Slide two", "reference_images": []},
            ],
        },
        {
            "api_key": "test-key",
            "base_url": "https://example.com",
            "model": "gpt-image-2",
            "size": "1024x1024",
            "quality": "low",
            "concurrency": 2,
            "protocol": "auto",
        },
    )

    with app.JOBS_LOCK:
        assert app.JOBS[job_id]["status"] == "completed"
        assert app.JOBS[job_id]["completed"] == 2
        assert app.JOBS[job_id]["downloads"] == {
            "zip": f"/api/jobs/{job_id}/download/zip",
            "pdf": f"/api/jobs/{job_id}/download/pdf",
            "pptx": f"/api/jobs/{job_id}/download/pptx",
        }
    assert (tmp_path / job_id / "images" / "slide-01.png").exists()
    zip_path = tmp_path / job_id / "test-deck-images.zip"
    pdf_path = tmp_path / job_id / "test-deck-images.pdf"
    pptx_path = tmp_path / job_id / "test-deck-image-only.pptx"
    assert zip_path.exists()
    assert pdf_path.exists()
    assert pptx_path.exists()
    with zipfile.ZipFile(zip_path) as archive:
        assert sorted(archive.namelist()) == [
            "prompts.json",
            "slide-01.png",
            "slide-02.png",
        ]
    presentation = Presentation(pptx_path)
    assert len(presentation.slides) == 2
