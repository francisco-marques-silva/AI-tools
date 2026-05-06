"""
openai_provider.py — OpenAI (ChatGPT / GPT) provider.

Supports:
  - Responses API      (gpt-5* family):   POST /v1/responses
  - Chat Completions   (all other GPT):   POST /v1/chat/completions
"""
import random
import time
from collections import deque
from typing import Any, Dict, Optional

import requests

from .base import RateLimitError, SYSTEM_INSTRUCTION, parse_llm_json

_REASONING_PREFIX = "gpt-5"


def call_openai(
    model: str,
    prompt: str,
    api_key: str,
    params: Optional[Dict[str, Any]] = None,
    max_retries: int = 5,
    base_backoff: float = 1.0,
) -> Dict[str, Any]:
    is_reasoning = model.startswith(_REASONING_PREFIX)
    headers = {"Authorization": f"Bearer {api_key}", "Content-Type": "application/json"}

    if is_reasoning:
        url = "https://api.openai.com/v1/responses"
        body: Dict[str, Any] = {"model": model, "input": prompt}
        if params:
            if params.get("reasoning_effort"):
                body["reasoning"] = {"effort": params["reasoning_effort"]}
            if params.get("verbosity"):
                body["text"] = {"verbosity": params["verbosity"]}
    else:
        url = "https://api.openai.com/v1/chat/completions"
        body = {
            "model": model,
            "messages": [
                {"role": "system", "content": SYSTEM_INSTRUCTION},
                {"role": "user", "content": prompt},
            ],
        }
        if params and "temperature" in params:
            body["temperature"] = float(params["temperature"])

    r = None
    for attempt in range(1, max_retries + 1):
        r = requests.post(url, headers=headers, json=body, timeout=60)
        if r.status_code == 200:
            break
        if r.status_code == 429:
            ra = r.headers.get("Retry-After", "")
            try:
                retry_after = float(ra)
            except (TypeError, ValueError):
                retry_after = 0.0
            raise RateLimitError(
                f"OpenAI HTTP 429 (Retry-After={retry_after:.1f}s)",
                retry_after=retry_after,
            )
        if r.status_code in (500, 502, 503, 504):
            jitter = random.uniform(0, 0.25)
            time.sleep(min(base_backoff * (2 ** (attempt - 1)) + jitter, 20.0))
            continue
        break

    if r is None or r.status_code != 200:
        raise RuntimeError(
            f"OpenAI error {r.status_code if r else 'no_response'}: "
            f"{r.text[:200] if r else ''}"
        )

    data = r.json()
    content: Optional[str] = None

    if is_reasoning:
        resp_obj = data.get("response") or data
        ot = resp_obj.get("output_text") or data.get("output_text")
        if isinstance(ot, str) and ot.strip():
            content = ot.strip()
        if not content:
            try:
                for item in (resp_obj.get("output") or data.get("output") or []):
                    for part in (item.get("content") or []):
                        if isinstance(part, dict):
                            txt = None
                            if isinstance(part.get("text"), dict):
                                txt = part.get("text", {}).get("value")
                            elif isinstance(part.get("text"), str):
                                txt = part.get("text")
                            if isinstance(txt, str) and txt.strip():
                                content = txt.strip()
                                break
                    if content:
                        break
            except Exception:
                pass
    else:
        try:
            content = data["choices"][0]["message"]["content"].strip()
        except Exception:
            try:
                content = data["choices"][0].get("text", "").strip()
            except Exception:
                pass

    if not content:
        q: deque = deque([data])
        while q and not content:
            node = q.popleft()
            if isinstance(node, dict):
                ot = node.get("output_text")
                if isinstance(ot, str) and ot.strip():
                    content = ot.strip()
                    break
                txt = node.get("text")
                if isinstance(txt, str) and txt.strip():
                    content = txt.strip()
                    break
                for v in node.values():
                    q.append(v)
            elif isinstance(node, list):
                q.extend(node)

    if not content:
        raise RuntimeError(f"Invalid OpenAI response structure (keys={list(data.keys())[:10]})")

    return parse_llm_json(content)
