"""Environment configuration via pydantic-settings.

Mirrors the env contract of the Node backend (see backend/.env.example): the
OpenRouter key lives only here, and MODEL / REASONING_EFFORT are swappable
without code changes. A missing key does NOT crash boot — it fails the request
(matching the Node server's lazy-client behavior).
"""

from pydantic_settings import BaseSettings, SettingsConfigDict


class Settings(BaseSettings):
    # protected_namespaces=() so the `model` field doesn't collide with pydantic's
    # reserved `model_` namespace and emit a warning.
    model_config = SettingsConfigDict(
        env_file=".env", env_file_encoding="utf-8", extra="ignore", protected_namespaces=()
    )

    openrouter_api_key: str = ""
    model: str = "openai/gpt-oss-20b"
    # gpt-oss exposes low|medium|high; for models without reasoning OpenRouter ignores it.
    reasoning_effort: str = "high"
    openrouter_base_url: str = "https://openrouter.ai/api/v1"
    # Upper bound on completion tokens per round (reasoning + visible output).
    max_output_tokens: int = 32000
    port: int = 3001
    # Comma-separated origins allowed to call this backend (the add-in dev server).
    allowed_origins: str = "https://localhost:3000"

    @property
    def allowed_origins_list(self) -> list[str]:
        return [o.strip() for o in self.allowed_origins.split(",") if o.strip()]


settings = Settings()
