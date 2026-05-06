"""
anthropic_provider.py — Anthropic Claude provider (HTTP, no SDK dependency).

API docs: https://docs.anthropic.com/en/api/messages

Rate limits:
  HTTP 429 — rate limited (requests per minute / tokens per minute exceeded)
  HTTP 529 — Anthropic service overloaded — treat as rate-limit for AIMD purposes
"""
import random
import time
from typing import Any, Dict, Optional

import requests

from .base import RateLimitError, SYSTEM_INSTRUCTION, parse_llm_json

_API_URL = "https://api.anthropic.com/v1/messages"
_API_VERSION = "2023-06-01"


def call_anthropic(
    model: str,
    prompt: str,
    api_key: str,
    params: Optional[Dict[str, Any]] = None,
    max_retries: int = 5,
    base_backoff: float = 2.0,
) -> Dict[str, Any]:
    """Call the Anthropic Messages API and return a normalised screening result."""
    temperature = 0.2
    max_tokens = 1024
    if params:
        try:
            temperature = max(0.0, min(1.0, float(params["temperature"])))
        except (KeyError, TypeError, ValueError):
            pass
        try:
            max_tokens = max(256, min(4096, int(params["max_tokens"])))
        except (KeyError, TypeError, ValueError):
            pass

    headers = {
        "x-api-key": api_key,
        "anthropic-version": _API_VERSION,
        "content-type": "application/json",
    }
    body: Dict[str, Any] = {
        "model": model,
        "max_tokens": max_tokens,
        "system": SYSTEM_INSTRUCTION,
        "messages": [{"role": "user", "content": prompt}],
        "temperature": temperature,
    }

    r = None
    for attempt in range(1, max_retries + 1):
        r = requests.post(_API_URL, headers=headers, json=body, timeout=120)
        if r.status_code == 200:
            break
        if r.status_code in (429, 529):
            # 429 = rate limit; 529 = service overloaded — both trigger AIMD decrease
            ra = r.headers.get("Retry-After", "")
            try:
                retry_after = float(ra)
            except (TypeError, ValueError):
                retry_after = 0.0
            raise RateLimitError(
                f"Anthropic HTTP {r.status_code} (Retry-After={retry_after:.1f}s)",
                retry_after=retry_after,
            )
        if r.status_code in (500, 502, 503, 504):
            jitter = random.uniform(0, 0.3)
            time.sleep(min(base_backoff * (2 ** (attempt - 1)) + jitter, 30.0))
            continue
        break

    if r is None or r.status_code != 200:
        raise RuntimeError(
            f"Anthropic error {r.status_code if r else 'no_response'}: "
            f"{r.text[:300] if r else ''}"
        )

    data = r.json()
    content: Optional[str] = None
    try:
        content = data["content"][0]["text"].strip()
    except Exception:
        try:
            for block in (data.get("content") or []):
                if isinstance(block, dict) and block.get("type") == "text":
                    txt = block.get("text", "").strip()
                    if txt:
                        content = txt
                        break
        except Exception:
            pass

    if not content:
        raise RuntimeError(f"Invalid Anthropic response structure: {str(data)[:200]}")

    return parse_llm_json(content)
