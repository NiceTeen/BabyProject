from __future__ import annotations

import base64
import binascii
import re
from collections.abc import Iterator
from dataclasses import dataclass
from datetime import datetime
from io import BytesIO
from pathlib import Path
from threading import Event
from typing import Any
from urllib.parse import urlparse

import requests
from PIL import Image, ImageOps, UnidentifiedImageError


IMAGE_AI_BASE_URL = "https://newapi.keep-sport.cn/v1"
IMAGE_AI_MODEL = "gpt-image-2"
MAX_INPUT_BYTES = 50 * 1024 * 1024
MAX_OUTPUT_BYTES = 100 * 1024 * 1024
REQUEST_TIMEOUT = (30, 600)


@dataclass(frozen=True)
class GeneratedImage:
    content: bytes
    suffix: str


class ImageAiCancelled(Exception):
    pass


def validate_api_key(api_key: str) -> str:
    api_key = str(api_key).strip()
    if not api_key:
        raise ValueError("请先填写并保存图片 AI API Key。")
    return api_key


def generate_image(
    api_key: str,
    prompt: str,
    *,
    cancel_event: Event | None = None,
    session: requests.Session | None = None,
) -> GeneratedImage:
    api_key = validate_api_key(api_key)
    prompt = _validated_prompt(prompt)
    _check_cancelled(cancel_event)

    client = session or requests.Session()
    close_session = session is None
    try:
        response = client.post(
            f"{IMAGE_AI_BASE_URL}/images/generations",
            headers={"Authorization": f"Bearer {api_key}"},
            json={
                "model": IMAGE_AI_MODEL,
                "prompt": prompt,
                "size": "auto",
                "quality": "high",
                "output_format": "png",
            },
            timeout=REQUEST_TIMEOUT,
        )
        return _response_image(response, client, cancel_event)
    except requests.Timeout as exc:
        raise RuntimeError("图片 AI 请求超时，请稍后重试。") from exc
    except requests.RequestException as exc:
        raise RuntimeError(f"图片 AI 网络请求失败：{exc}") from exc
    finally:
        if close_session:
            client.close()


def edit_image(
    api_key: str,
    source_path: Path,
    prompt: str,
    *,
    cancel_event: Event | None = None,
    session: requests.Session | None = None,
) -> GeneratedImage:
    api_key = validate_api_key(api_key)
    prompt = _validated_prompt(prompt)
    _check_cancelled(cancel_event)
    filename, content, content_type = _prepare_input_image(source_path)
    _check_cancelled(cancel_event)

    client = session or requests.Session()
    close_session = session is None
    try:
        response = client.post(
            f"{IMAGE_AI_BASE_URL}/images/edits",
            headers={"Authorization": f"Bearer {api_key}"},
            data={
                "model": IMAGE_AI_MODEL,
                "prompt": prompt,
                "size": "auto",
                "quality": "high",
                "output_format": "png",
            },
            files={"image": (filename, content, content_type)},
            timeout=REQUEST_TIMEOUT,
        )
        return _response_image(response, client, cancel_event)
    except requests.Timeout as exc:
        raise RuntimeError("图片 AI 请求超时，请稍后重试。") from exc
    except requests.RequestException as exc:
        raise RuntimeError(f"图片 AI 网络请求失败：{exc}") from exc
    finally:
        if close_session:
            client.close()


def save_generated_image(
    generated: GeneratedImage,
    output_dir: Path,
    task_id: int,
    source_path: Path | None = None,
) -> Path:
    output_dir.mkdir(parents=True, exist_ok=True)
    timestamp = datetime.now().astimezone().strftime("%Y%m%d_%H%M%S")
    source_part = (
        _safe_filename_part(source_path.stem, "图片") if source_path else "文生图"
    )
    target = output_dir / f"{source_part}__AI__{timestamp}_{task_id}{generated.suffix}"
    temporary = target.with_name(f".{target.name}.part")
    try:
        temporary.write_bytes(generated.content)
        temporary.replace(target)
    except OSError:
        temporary.unlink(missing_ok=True)
        raise
    return target


def _validated_prompt(prompt: str) -> str:
    prompt = str(prompt).strip()
    if not prompt:
        raise ValueError("生图提示词不能为空。")
    return prompt


def _response_image(
    response: requests.Response,
    session: requests.Session,
    cancel_event: Event | None,
) -> GeneratedImage:
    _check_cancelled(cancel_event)
    if not response.ok:
        raise RuntimeError(_api_error_message(response))
    try:
        payload = response.json()
    except ValueError as exc:
        raise RuntimeError("图片 AI 接口没有返回有效的 JSON 数据。") from exc
    result = _extract_image_result(payload, session, cancel_event)
    _check_cancelled(cancel_event)
    return result


def _prepare_input_image(path: Path) -> tuple[str, bytes, str]:
    path = Path(path).resolve()
    if not path.is_file():
        raise ValueError("待修改的图片文件不存在。")
    try:
        content = path.read_bytes()
    except OSError as exc:
        raise ValueError(f"无法读取待修改图片：{exc}") from exc
    if not content:
        raise ValueError("待修改的图片文件为空。")
    if len(content) > MAX_INPUT_BYTES:
        raise ValueError("待修改图片超过 50 MB，无法上传。")

    try:
        with Image.open(BytesIO(content)) as image:
            image.verify()
        with Image.open(BytesIO(content)) as image:
            image_format = str(image.format or "").upper()
    except (UnidentifiedImageError, OSError) as exc:
        raise ValueError("待修改文件不是有效图片。") from exc

    content_types = {
        "JPEG": "image/jpeg",
        "PNG": "image/png",
        "WEBP": "image/webp",
    }
    if image_format in content_types:
        return path.name, content, content_types[image_format]

    try:
        with Image.open(BytesIO(content)) as image:
            converted = ImageOps.exif_transpose(image)
            if converted.mode not in {"RGB", "RGBA"}:
                converted = converted.convert(
                    "RGBA" if "A" in converted.getbands() else "RGB"
                )
            buffer = BytesIO()
            converted.save(buffer, format="PNG")
            content = buffer.getvalue()
    except (UnidentifiedImageError, OSError, ValueError) as exc:
        raise ValueError("图片格式无法转换为 AI 接口支持的格式。") from exc
    if len(content) > MAX_INPUT_BYTES:
        raise ValueError("转换后的图片超过 50 MB，无法上传。")
    return f"{path.stem}.png", content, "image/png"


def _extract_image_result(
    payload: Any,
    session: requests.Session,
    cancel_event: Event | None,
) -> GeneratedImage:
    if not isinstance(payload, dict):
        raise RuntimeError("图片 AI 接口返回的数据结构无效。")
    items = list(_image_result_items(payload))
    for item in items:
        b64_json = _first_string(item, "b64_json", "base64", "image_base64")
        if b64_json:
            return _validated_output(_decode_base64_image(b64_json))

        url = _image_result_url(item)
        if url:
            content = (
                _decode_base64_image(url)
                if url.startswith("data:")
                else _download_result_image(session, url, cancel_event)
            )
            return _validated_output(content)

    error_message = _payload_error_message(payload)
    if error_message:
        raise RuntimeError(f"图片 AI 接口返回错误：{error_message}")
    fields = _payload_field_summary(payload)
    if not items:
        raise RuntimeError(
            f"图片 AI 接口本次没有返回图片数据{fields}，"
            "可能是中转站或上游服务临时异常，请重试。"
        )
    raise RuntimeError(f"图片 AI 接口没有返回可识别的图片内容{fields}。")


def _image_result_items(payload: dict[str, Any]) -> Iterator[dict[str, Any]]:
    queue: list[tuple[Any, int]] = [(payload, 0)]
    seen: set[int] = set()
    wrapper_keys = ("data", "result", "images", "output")
    image_keys = {
        "b64_json",
        "base64",
        "image_base64",
        "url",
        "image_url",
        "image",
    }
    while queue:
        value, depth = queue.pop(0)
        if isinstance(value, (dict, list)):
            identity = id(value)
            if identity in seen:
                continue
            seen.add(identity)
        if isinstance(value, dict):
            if image_keys.intersection(value):
                yield value
            if depth < 3:
                for key in wrapper_keys:
                    nested = value.get(key)
                    if isinstance(nested, (dict, list)):
                        queue.append((nested, depth + 1))
        elif isinstance(value, list) and depth < 3:
            queue.extend((item, depth + 1) for item in value)


def _first_string(item: dict[str, Any], *keys: str) -> str:
    for key in keys:
        value = item.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
    image = item.get("image")
    if (
        isinstance(image, str)
        and image.strip()
        and not image.startswith(("http://", "https://"))
    ):
        return image.strip()
    return ""


def _image_result_url(item: dict[str, Any]) -> str:
    for key in ("url", "image_url"):
        value = item.get(key)
        if isinstance(value, str) and value.strip():
            return value.strip()
        if isinstance(value, dict):
            nested = value.get("url")
            if isinstance(nested, str) and nested.strip():
                return nested.strip()
    image = item.get("image")
    if isinstance(image, str) and image.strip().startswith(
        ("http://", "https://", "data:")
    ):
        return image.strip()
    return ""


def _payload_error_message(payload: dict[str, Any]) -> str:
    for key in ("error", "message", "msg", "detail", "reason", "error_message"):
        message = _error_value_text(payload.get(key))
        if message:
            return _sanitize_api_message(message)
    return ""


def _error_value_text(value: Any) -> str:
    if isinstance(value, str):
        return value.strip()
    if isinstance(value, dict):
        for key in ("message", "msg", "detail", "reason", "error", "description"):
            message = _error_value_text(value.get(key))
            if message:
                return message
        for key in ("code", "type"):
            message = _error_value_text(value.get(key))
            if message:
                return message
    if isinstance(value, list):
        for item in value:
            message = _error_value_text(item)
            if message:
                return message
    return ""


def _sanitize_api_message(message: str) -> str:
    message = re.sub(r"sk-[A-Za-z0-9_-]+", "[已隐藏 API Key]", message)
    message = re.sub(
        r"(?i)Bearer\s+[A-Za-z0-9._~+/=-]+",
        "Bearer [已隐藏凭据]",
        message,
    )
    return message[:1000]


def _payload_field_summary(payload: dict[str, Any]) -> str:
    fields = sorted(str(key)[:60] for key in payload)[:12]
    return f"（响应字段：{', '.join(fields)}）" if fields else ""


def _decode_base64_image(value: str) -> bytes:
    value = value.strip()
    if value.startswith("data:"):
        _header, separator, value = value.partition(",")
        if not separator:
            raise RuntimeError("图片 AI 接口返回的 data URL 无效。")
    if len(value) > (MAX_OUTPUT_BYTES * 4 // 3) + 8:
        raise RuntimeError("图片 AI 返回的图片超过 100 MB。")
    try:
        content = base64.b64decode(value, validate=True)
    except (binascii.Error, ValueError) as exc:
        raise RuntimeError("图片 AI 接口返回的 base64 图片无效。") from exc
    if len(content) > MAX_OUTPUT_BYTES:
        raise RuntimeError("图片 AI 返回的图片超过 100 MB。")
    return content


def _download_result_image(
    session: requests.Session,
    url: str,
    cancel_event: Event | None,
) -> bytes:
    parsed = urlparse(url)
    if parsed.scheme not in {"http", "https"} or not parsed.netloc:
        raise RuntimeError("图片 AI 接口返回了无效的图片地址。")
    try:
        with session.get(url, timeout=REQUEST_TIMEOUT, stream=True) as response:
            if not response.ok:
                raise RuntimeError(f"下载 AI 图片失败：HTTP {response.status_code}")
            chunks: list[bytes] = []
            total = 0
            for chunk in response.iter_content(64 * 1024):
                _check_cancelled(cancel_event)
                if not chunk:
                    continue
                total += len(chunk)
                if total > MAX_OUTPUT_BYTES:
                    raise RuntimeError("图片 AI 返回的图片超过 100 MB。")
                chunks.append(chunk)
            return b"".join(chunks)
    except requests.Timeout as exc:
        raise RuntimeError("下载 AI 图片超时，请稍后重试。") from exc
    except requests.RequestException as exc:
        raise RuntimeError(f"下载 AI 图片失败：{exc}") from exc


def _validated_output(content: bytes) -> GeneratedImage:
    if not content:
        raise RuntimeError("图片 AI 接口返回了空图片。")
    if len(content) > MAX_OUTPUT_BYTES:
        raise RuntimeError("图片 AI 返回的图片超过 100 MB。")
    try:
        with Image.open(BytesIO(content)) as image:
            image.verify()
        with Image.open(BytesIO(content)) as image:
            image_format = str(image.format or "").upper()
    except (UnidentifiedImageError, OSError) as exc:
        raise RuntimeError("图片 AI 接口返回的内容不是有效图片。") from exc
    suffix = {"JPEG": ".jpg", "PNG": ".png", "WEBP": ".webp"}.get(image_format)
    if suffix is None:
        raise RuntimeError(
            f"图片 AI 返回了不支持的图片格式：{image_format or '未知'}。"
        )
    return GeneratedImage(content, suffix)


def _api_error_message(response: requests.Response) -> str:
    message = ""
    try:
        payload = response.json()
    except ValueError:
        payload = None
    if isinstance(payload, dict):
        message = _payload_error_message(payload)
    if message:
        return f"图片 AI 请求失败（HTTP {response.status_code}）：{message}"
    return f"图片 AI 请求失败：HTTP {response.status_code}"


def _safe_filename_part(value: str, fallback: str) -> str:
    value = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "_", value).strip(" ._")
    return value[:80] or fallback


def _check_cancelled(cancel_event: Event | None) -> None:
    if cancel_event is not None and cancel_event.is_set():
        raise ImageAiCancelled("图片 AI 任务已取消。")
