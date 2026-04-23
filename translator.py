"""
translator.py — Dịch tiếng Việt sang tiếng Anh cho email notification.

Gọi 9router (proxy OpenAI-compatible) để dịch. Có các đặc tính:

- Batch: dịch nhiều field cùng lúc trong 1 API call (tiết kiệm token + latency).
- Cache in-memory: request giống nhau không tốn API lần 2.
- Fallback an toàn: nếu API fail, trả về text VN gốc + log warning,
  KHÔNG raise exception để không chặn việc gửi email.

Env vars cần set:
    ROUTER_BASE_URL    (mặc định: http://127.0.0.1:20128/v1)
    ROUTER_API_KEY     (bắt buộc, nếu trống → skip dịch, return gốc)
    ROUTER_MODEL       (mặc định: if/glm-4.7)
    ROUTER_TIMEOUT     (mặc định: 20 giây)
"""

from __future__ import annotations

import json
import logging
import os
from typing import Dict, Iterable

import requests

log = logging.getLogger(__name__)

# ─── Config ──────────────────────────────────────────────────────────────────

_DEFAULT_BASE_URL = "http://127.0.0.1:20128/v1"
_DEFAULT_MODEL    = "if/glm-4.7"
_DEFAULT_TIMEOUT  = 20

# Cache in-memory: key = frozenset((field, text)) → dict
_CACHE: dict[frozenset, Dict[str, str]] = {}
_CACHE_MAX = 256  # giới hạn để không nuốt RAM


def _cfg():
    return {
        "base_url": os.getenv("ROUTER_BASE_URL", _DEFAULT_BASE_URL).rstrip("/"),
        "api_key":  os.getenv("ROUTER_API_KEY", "").strip(),
        "model":    os.getenv("ROUTER_MODEL", _DEFAULT_MODEL),
        "timeout":  int(os.getenv("ROUTER_TIMEOUT", _DEFAULT_TIMEOUT)),
    }


# ─── Prompt ──────────────────────────────────────────────────────────────────

_SYSTEM_PROMPT = (
    "You are a professional Vietnamese-to-English translator specializing in "
    "IT incident and maintenance notifications for a cloud services company. "
    "Translate each field naturally and concisely to business English. "
    "Keep technical terms (VM, IP, CPU, vServer, GPU, DNS, firewall, load balancer, "
    "database, network, switch, router, storage, backup) unchanged. "
    "Preserve the formal tone suitable for customer notifications. "
    "Do NOT add explanations, greetings, or extra text — ONLY output the JSON object "
    "with the exact same keys as the input, values being the English translations."
)


def _build_user_prompt(fields: Dict[str, str]) -> str:
    """Prompt yêu cầu model dịch và trả JSON với cùng keys."""
    return (
        "Translate the following Vietnamese values to English. "
        "Return ONLY a JSON object (no markdown, no code fence) with the same keys:\n\n"
        f"{json.dumps(fields, ensure_ascii=False, indent=2)}"
    )


# ─── Core ────────────────────────────────────────────────────────────────────

def translate_fields(fields: Dict[str, str]) -> Dict[str, str]:
    """Dịch 1 dict {field_name: text_vi} sang {field_name: text_en}.

    - Bỏ qua field có text rỗng/whitespace (giữ nguyên chuỗi rỗng).
    - Nếu API không reach / fail → log warning + trả về text VN gốc cho mỗi field.
    - Có cache theo cặp (field_name, text).
    """
    if not fields:
        return {}

    # Loại bỏ field rỗng
    non_empty = {k: v.strip() for k, v in fields.items() if v and v.strip()}
    result = {k: fields.get(k, "") for k in fields}  # default = gốc
    if not non_empty:
        return result

    # Check cache trước
    cache_key = frozenset(non_empty.items())
    if cache_key in _CACHE:
        log.info("translator: cache hit for %d field(s)", len(non_empty))
        result.update(_CACHE[cache_key])
        return result

    cfg = _cfg()
    if not cfg["api_key"]:
        log.warning("translator: ROUTER_API_KEY empty, skipping translation (returning VN text)")
        result.update(non_empty)
        return result

    try:
        translated = _call_router(non_empty, cfg)
    except Exception as e:
        log.warning("translator: call failed (%s), falling back to VN text", e)
        result.update(non_empty)
        return result

    # Với field nào model không trả về → fallback VN
    merged = {k: translated.get(k) or v for k, v in non_empty.items()}
    result.update(merged)

    # Lưu cache (với lru-ish: xóa entry cũ nhất nếu quá limit)
    if len(_CACHE) >= _CACHE_MAX:
        # python 3.7+: dict giữ insertion order → pop oldest
        _CACHE.pop(next(iter(_CACHE)))
    _CACHE[cache_key] = merged

    return result


def _call_router(fields: Dict[str, str], cfg: dict) -> Dict[str, str]:
    """Gọi 9router chat completions. Trả về dict đã dịch hoặc raise."""
    url = f"{cfg['base_url']}/chat/completions"
    headers = {
        "Content-Type":  "application/json",
        "Authorization": f"Bearer {cfg['api_key']}",
    }
    body = {
        "model": cfg["model"],
        "messages": [
            {"role": "system", "content": _SYSTEM_PROMPT},
            {"role": "user",   "content": _build_user_prompt(fields)},
        ],
        "temperature": 0.2,  # dịch cần ổn định, không sáng tạo
        "stream": False,
    }

    log.info("translator: POST %s model=%s fields=%d", url, cfg["model"], len(fields))
    r = requests.post(url, headers=headers, json=body, timeout=cfg["timeout"])
    r.raise_for_status()
    data = r.json()

    # OpenAI format: data["choices"][0]["message"]["content"]
    try:
        content = data["choices"][0]["message"]["content"].strip()
    except (KeyError, IndexError, TypeError) as e:
        raise RuntimeError(f"Unexpected response shape: {data}") from e

    # Đề phòng model thích wrap bằng markdown ```json ... ```
    content = _strip_code_fence(content)

    try:
        parsed = json.loads(content)
    except json.JSONDecodeError as e:
        raise RuntimeError(f"Model did not return valid JSON: {content[:200]}") from e

    if not isinstance(parsed, dict):
        raise RuntimeError(f"Expected JSON object, got {type(parsed).__name__}")

    # Ép tất cả giá trị về string
    return {k: str(v) for k, v in parsed.items()}


def _strip_code_fence(s: str) -> str:
    """Xoá bỏ ```json ... ``` hoặc ``` ... ``` nếu model trả về."""
    s = s.strip()
    if s.startswith("```"):
        # Bỏ dòng đầu (```json hoặc ```) và dòng cuối ```
        lines = s.splitlines()
        if lines and lines[0].startswith("```"):
            lines = lines[1:]
        if lines and lines[-1].strip() == "```":
            lines = lines[:-1]
        s = "\n".join(lines).strip()
    return s


# ─── Convenience helpers ─────────────────────────────────────────────────────

def is_configured() -> bool:
    """True nếu có API key để thử gọi — endpoint reachable hay không không check."""
    return bool(os.getenv("ROUTER_API_KEY", "").strip())


def clear_cache() -> None:
    _CACHE.clear()
