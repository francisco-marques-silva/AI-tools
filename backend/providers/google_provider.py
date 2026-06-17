"""
google_provider.py — Google Gemini provider (REST, no SDK dependency).

API docs: https://ai.google.dev/api/generate-content
PDF input: https://ai.google.dev/gemini-api/docs/document-processing

Uses the v1beta generateContent endpoint with:
  - systemInstruction for the JSON-only instruction
  - responseMimeType: "application/json" for structured output
  - inlineData for inline PDF attachment (when pdf_bytes is provided)
"""
import base64
import random
import time
from typing import Any, Dict, Optional

import requests

from .base import RateLimitError, SYSTEM_INSTRUCTION, parse_llm_json

_API_BASE = "https://generativelanguage.googleapis.com/v1beta/models"


def call_google(
    model: str,
    prompt: str,
    api_key: str,
    params: Optional[Dict[str, Any]] = None,
    max_retries: int = 5,
    base_backoff: float = 1.0,
    pdf_bytes: Optional[bytes] = None,
) -> Dict[str, Any]:
    """Call Google Gemini generateContent API and return a normalised screening result.

    When `pdf_bytes` is provided, the PDF is attached as an inlineData part.
    Gemini reads the PDF directly (preserves layout, tables, figures).
    """
    temperature = 0.2
    max_output_tokens = 1024
    if params:
        try:
            temperature = max(0.0, min(2.0, float(params["temperature"])))
        except (KeyError, TypeError, ValueError):
            pass
        try:
            max_output_tokens = max(256, min(8192, int(params["max_output_tokens"])))
        except (KeyError, TypeError, ValueError):
            pass

    url = f"{_API_BASE}/{model}:generateContent?key={api_key}"
    headers = {"Content-Type": "application/json"}

    parts: list = []
    if pdf_bytes:
        pdf_b64 = base64.standard_b64encode(pdf_bytes).decode("ascii")
        parts.append({
            "inlineData": {"mimeType": "application/pdf", "data": pdf_b64},
        })
    parts.append({"text": prompt})

    body: Dict[str, Any] = {
        "systemInstruction": {"parts": [{"text": SYSTEM_INSTRUCTION}]},
        "contents": [{"parts": parts}],
        "generationConfig": {
            "temperature": temperature,
            "maxOutputTokens": max_output_tokens,
            "responseMimeType": "application/json",
        },
    }

    r = None
    request_timeout = 300 if pdf_bytes else 120
    for attempt in range(1, max_retries + 1):
        r = requests.post(url, headers=headers, json=body, timeout=request_timeout)
        if r.status_code == 200:
            break
        if r.status_code == 429:
            ra = r.headers.get("Retry-After", "")
            try:
                retry_after = float(ra)
            except (TypeError, ValueError):
                retry_after = 0.0
            raise RateLimitError(
                f"Google Gemini HTTP 429 quota exceeded (Retry-After={retry_after:.1f}s)",
                retry_after=retry_after,
            )
        if r.status_code in (500, 502, 503, 504):
            jitter = random.uniform(0, 0.25)
            time.sleep(min(base_backoff * (2 ** (attempt - 1)) + jitter, 20.0))
            continue
        break

    if r is None or r.status_code != 200:
        raise RuntimeError(
            f"Google Gemini error {r.status_code if r else 'no_response'}: "
            f"{r.text[:300] if r else ''}"
        )

    data = r.json()
    content: Optional[str] = None
    try:
        content = data["candidates"][0]["content"]["parts"][0]["text"].strip()
    except Exception:
        try:
            for candidate in (data.get("candidates") or []):
                for part in (candidate.get("content", {}).get("parts") or []):
                    if isinstance(part, dict):
                        txt = part.get("text", "").strip()
                        if txt:
                            content = txt
                            break
                if content:
                    break
        except Exception:
            pass

    if not content:
        raise RuntimeError(f"Invalid Gemini response structure: {str(data)[:200]}")

    return parse_llm_json(content)
