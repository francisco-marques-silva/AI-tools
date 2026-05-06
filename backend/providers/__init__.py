"""
providers — Multi-provider LLM dispatch for the screening backend.

Supported providers:
  "openai"    → OpenAI (ChatGPT / GPT-4, GPT-5 families)
  "anthropic" → Anthropic (Claude families)
  "google"    → Google (Gemini families)
"""
from .base import RateLimitError, SYSTEM_INSTRUCTION, parse_llm_json
from .openai_provider import call_openai
from .anthropic_provider import call_anthropic
from .google_provider import call_google

__all__ = [
    "RateLimitError",
    "SYSTEM_INSTRUCTION",
    "parse_llm_json",
    "call_openai",
    "call_anthropic",
    "call_google",
    "call_llm",
]

_PROVIDERS = {
    "openai":    call_openai,
    "anthropic": call_anthropic,
    "google":    call_google,
}

_PROVIDER_ENV_KEY = {
    "openai":    "OPENAI_API_KEY",
    "anthropic": "ANTHROPIC_API_KEY",
    "google":    "GOOGLE_API_KEY",
}


def call_llm(
    provider: str,
    model: str,
    prompt: str,
    api_key: str,
    params=None,
    max_retries: int = 5,
    base_backoff: float = 1.0,
):
    """Route an LLM screening call to the correct provider.

    Returns a normalised dict:
        {"decision": "include"|"exclude"|"maybe", "rationale": str,
         "inclusion_evaluation": [...], "exclusion_evaluation": [...]}
    """
    fn = _PROVIDERS.get(provider)
    if fn is None:
        raise ValueError(
            f"Unknown provider {provider!r}. Valid options: {sorted(_PROVIDERS)}"
        )
    return fn(
        model=model,
        prompt=prompt,
        api_key=api_key,
        params=params,
        max_retries=max_retries,
        base_backoff=base_backoff,
    )


def env_key_for(provider: str) -> str:
    """Return the environment variable name holding the API key for a provider."""
    return _PROVIDER_ENV_KEY.get(provider, "OPENAI_API_KEY")
