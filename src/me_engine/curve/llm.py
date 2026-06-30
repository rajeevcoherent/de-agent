"""A thin LLM client returning validated JSON, plus a DuckDuckGo evidence helper.

The client is intentionally minimal: one chat call that requests a JSON object and
parses it. It targets the OpenAI-compatible schema that both OpenAI and OpenRouter
expose, so the same code path serves either provider.
"""
from __future__ import annotations

import json
import os
from dataclasses import dataclass

import httpx

from .cache import JsonCache
from .config import AgentConfig


class LLMError(RuntimeError):
    pass


@dataclass(frozen=True, slots=True)
class Evidence:
    claim: str
    source: str
    confidence: float


class LLMClient:
    """OpenAI-compatible chat client that returns parsed JSON objects."""

    def __init__(self, config: AgentConfig, cache: JsonCache | None = None) -> None:
        self._config = config
        self._cache = cache if cache is not None else JsonCache()


    def complete_json(self, system: str, user: str) -> dict:
      
        if not self._config.is_online:
            raise LLMError("no API key configured; cannot call the model")

        cache_key = JsonCache.key(self._config.model, system, user)

        # ==========================================================
        # DEBUG MODE: FORCE DISABLE CACHE
        # ==========================================================
        cached = None

        # print("[DEBUG] cached is None =", cached is None)

        # if cached is not None:
            # print("[LLM] CACHE HIT")
            # return cached

        # print(f"[LLM] API CALL -> {self._config.base_url}")
        # print(f"[LLM] MODEL    -> {self._config.model}")

        payload = {
            "model": self._config.model,
            "messages": [
                {"role": "system", "content": system},
                {"role": "user", "content": user},
            ],
            "response_format": {"type": "json_object"},
            "temperature": 0.2,
        }

        headers = {
            "Authorization": f"Bearer {self._config.auth_token}"
        }

        with httpx.Client(timeout=self._config.request_timeout) as client:
            resp = client.post(
                f"{self._config.base_url}/chat/completions",
                json=payload,
                headers=headers,
            )

        # print("[LLM] STATUS   ->", resp.status_code)

        if resp.status_code >= 400:
            raise LLMError(
                f"LLM HTTP {resp.status_code}: {resp.text[:300]}"
            )

        content = resp.json()["choices"][0]["message"]["content"]

        try:
            parsed = json.loads(content)
        except json.JSONDecodeError as exc:
            raise LLMError(
                f"model did not return valid JSON: {exc}"
            ) from exc

        # ==========================================================
        # DEBUG MODE: DO NOT WRITE CACHE
        # ==========================================================
        # self._cache.put(cache_key, parsed)

        return parsed

    # def complete_json(self, system: str, user: str) -> dict:
    #     if not self._config.is_online:
    #         raise LLMError("no API key configured; cannot call the model")

    #     cache_key = JsonCache.key(self._config.model, system, user)
    #     cached = self._cache.get(cache_key)

    #     # DEBUG 1: Cache check
    #     if cached is not None:
    #         print("[LLM] CACHE HIT")
    #         return cached

    #     # DEBUG 2: API call info
    #     print(f"[LLM] API CALL -> {self._config.base_url}")
    #     print(f"[LLM] MODEL    -> {self._config.model}")

    #     payload = {
    #         "model": self._config.model,
    #         "messages": [
    #             {"role": "system", "content": system},
    #             {"role": "user", "content": user},
    #         ],
    #         "response_format": {"type": "json_object"},
    #         "temperature": 0.2,
    #     }

    #     headers = {
    #         "Authorization": f"Bearer {self._config.auth_token}"
    #     }

    #     with httpx.Client(timeout=self._config.request_timeout) as client:
    #         resp = client.post(
    #             f"{self._config.base_url}/chat/completions",
    #             json=payload,
    #             headers=headers
    #         )

    #     # DEBUG 3: Response status
    #     print("[LLM] STATUS   ->", resp.status_code)

    #     if resp.status_code >= 400:
    #         raise LLMError(f"LLM HTTP {resp.status_code}: {resp.text[:300]}")

    #     content = resp.json()["choices"][0]["message"]["content"]


class EvidenceGatherer:
    """Fetches a few web snippets to ground the agent's reasoning.

    Disabled by default (ME_ENABLE_WEB_EVIDENCE=1 to enable) because public DDG
    endpoints rate-limit aggressively and stall batch runs; the agents reason from
    their data-derived priors when web evidence is off.
    """

    def __init__(self, max_results: int = 5, enabled: bool | None = None) -> None:
        self._max_results = max_results
        self._enabled = (os.environ.get("ME_ENABLE_WEB_EVIDENCE") == "1"
                         if enabled is None else enabled)

    def search(self, query: str) -> list[str]:
        if not self._enabled:
            return []
        ddgs_cls = self._import_ddgs()
        if ddgs_cls is None:
            return []
        try:
            with ddgs_cls() as ddgs:
                hits = ddgs.text(query, max_results=self._max_results)
            return [f"{h.get('title', '')}: {h.get('body', '')}" for h in hits]
        except Exception:
            return []   # evidence is best-effort; the agent degrades gracefully

    @staticmethod
    def _import_ddgs():
        try:
            from ddgs import DDGS
            return DDGS
        except ImportError:
            try:
                from duckduckgo_search import DDGS
                return DDGS
            except ImportError:
                return None
    
