"""Runtime configuration for the agent layer.

Keys and model choices are read from the environment / a local .env file so no
secret is ever hard-coded. OpenRouter is the default gateway (it can reach many
models); the OpenAI key is a direct fallback. If neither is present the agent
runs in offline mode and falls back to the canonical shape.
"""
from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path


def _load_dotenv(path: Path) -> None:
    """Minimal .env loader (no dependency); ignores comments and blanks."""
    if not path.exists():
        return
    for line in path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        os.environ.setdefault(key.strip(), value.strip().strip('"').strip("'"))


@dataclass(frozen=True, slots=True)
class AgentConfig:
    openrouter_key: str | None
    openai_key: str | None
    model: str
    base_url: str
    request_timeout: float

    @property
    def is_online(self) -> bool:
        return bool(self.openrouter_key or self.openai_key)

    @property
    def uses_openrouter(self) -> bool:
        return bool(self.openrouter_key)

    @classmethod
    def load(cls, project_root: Path | str = ".") -> "AgentConfig":
        _load_dotenv(Path(project_root) / ".env")
        openrouter = os.environ.get("OPENROUTER_API_KEY")
        openai = os.environ.get("OPENAI_API_KEY")
        if openrouter:
            base = "https://openrouter.ai/api/v1"
            default_model = "openai/gpt-4o-mini"
        else:
            base = "https://api.openai.com/v1"
            default_model = "gpt-4o-mini"
        return cls(
            openrouter_key=openrouter,
            openai_key=openai,
            model=os.environ.get("ME_CURVE_MODEL", default_model),
            base_url=os.environ.get("ME_LLM_BASE_URL", base),
            request_timeout=float(os.environ.get("ME_LLM_TIMEOUT", "60")),
        )

    @property
    def auth_token(self) -> str | None:
        return self.openrouter_key or self.openai_key
