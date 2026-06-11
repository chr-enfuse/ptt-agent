"""Capture the exact JSON body PydanticAI sends to OpenRouter, to diff against
the working Node request shape."""

import asyncio
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import httpx

CAPTURED = {}


async def _log_request(request: httpx.Request) -> None:
    if request.content:
        try:
            CAPTURED["body"] = json.loads(request.content)
        except Exception:
            CAPTURED["body"] = request.content.decode("utf-8", "replace")
    CAPTURED["url"] = str(request.url)


async def main() -> None:
    from pydantic_ai import Agent
    from pydantic_ai.models.openai import OpenAIChatModel, OpenAIChatModelSettings
    from pydantic_ai.providers.openai import OpenAIProvider

    from app.config import settings

    http_client = httpx.AsyncClient(event_hooks={"request": [_log_request]})
    model = OpenAIChatModel(
        settings.model,
        provider=OpenAIProvider(
            base_url=settings.openrouter_base_url,
            api_key=settings.openrouter_api_key or "MISSING",
            http_client=http_client,
        ),
    )
    msettings = OpenAIChatModelSettings(
        max_tokens=settings.max_output_tokens,
        extra_body={
            "reasoning": {"effort": settings.reasoning_effort},
            "provider": {"require_parameters": True},
        },
    )
    agent = Agent(model, system_prompt="You are a test.", model_settings=msettings)
    try:
        await agent.run("Reply with exactly: pong")
        print("RUN OK")
    except Exception as err:
        print("RUN ERROR:", str(err)[:300])
    print("URL:", CAPTURED.get("url"))
    print("BODY:", json.dumps(CAPTURED.get("body"), indent=2)[:3000])


asyncio.run(main())
