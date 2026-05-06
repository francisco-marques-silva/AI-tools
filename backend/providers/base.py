"""
base.py — Shared types and response-parsing utilities for all LLM providers.
"""
import json
from typing import Any, Dict, List


SYSTEM_INSTRUCTION = (
    "You are a systematic review screening assistant. "
    "Return ONLY strict JSON and nothing else — no markdown fences, no prose."
)


class RateLimitError(RuntimeError):
    """Raised by any provider when the API signals rate-limiting (HTTP 429 / 529)."""
    def __init__(self, message: str, retry_after: float = 0.0):
        super().__init__(message)
        self.retry_after = retry_after


def _coerce_eval(arr: Any) -> List[Dict[str, str]]:
    out: List[Dict[str, str]] = []
    if isinstance(arr, str):
        try:
            arr = json.loads(arr)
        except Exception:
            arr = []
    if isinstance(arr, list):
        for it in arr:
            if isinstance(it, dict):
                crit = str(it.get("criterion", "")).strip()
                status = str(it.get("status", "")).strip().lower()
                if status not in {"met", "unclear", "unmet"}:
                    status = "unclear"
                if crit:
                    out.append({"criterion": crit, "status": status})
            elif isinstance(it, (list, tuple)) and len(it) >= 2:
                crit = str(it[0]).strip()
                status = str(it[1]).strip().lower()
                if status not in {"met", "unclear", "unmet"}:
                    status = "unclear"
                if crit:
                    out.append({"criterion": crit, "status": status})
    return out


def parse_llm_json(raw: str) -> Dict[str, Any]:
    """Strip code fences (if any) and parse JSON from LLM output into a normalised dict."""
    content = raw.strip()
    if content.startswith("```"):
        content = content.strip("`")
        if "\n" in content:
            content = content.split("\n", 1)[1].strip()

    try:
        parsed = json.loads(content)
    except Exception:
        raise RuntimeError(f"JSON parse error – raw content snippet: {content[:120]!r}")

    decision = str(parsed.get("decision", "")).strip().lower()
    rationale = str(parsed.get("rationale", "")).strip()

    inc_eval = _coerce_eval(
        parsed.get("inclusion_evaluation")
        or parsed.get("inclusionEvaluations")
        or parsed.get("inclusion")
        or []
    )
    exc_eval = _coerce_eval(
        parsed.get("exclusion_evaluation")
        or parsed.get("exclusionEvaluations")
        or parsed.get("exclusion")
        or []
    )

    if decision not in {"include", "exclude", "maybe"}:
        decision = "maybe"
    if len(rationale.split()) > 12:
        rationale = " ".join(rationale.split()[:12])
    if not rationale:
        rationale = "insufficient information"

    return {
        "decision": decision,
        "rationale": rationale,
        "inclusion_evaluation": inc_eval,
        "exclusion_evaluation": exc_eval,
    }
