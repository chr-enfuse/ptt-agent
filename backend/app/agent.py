"""PydanticAI agent wiring — the Python port of the model/provider setup, system
prompt, and TOOLS registration from backend/src/routes/chat.ts.

The model talks to OpenRouter through the OpenAI-compatible provider. OpenRouter's
`reasoning` and `provider` passthroughs aren't first-class OpenAI params, so they
ride in `extra_body` per request (plan §7.1). A thin WrapperModel strips gpt-oss
"harmony" artifacts from tool names before PydanticAI dispatches them (plan §7.5).
"""

from __future__ import annotations

import logging
from contextlib import asynccontextmanager

from pydantic_ai import Agent
from pydantic_ai.messages import ToolCallPart
from pydantic_ai.models.openai import OpenAIChatModel, OpenAIChatModelSettings
from pydantic_ai.models.wrapper import WrapperModel
from pydantic_ai.providers.openai import OpenAIProvider

from .config import settings
from .session import ChatSession
from .tools import register_tools

logger = logging.getLogger("deck-agent.agent")

# Ported verbatim from chat.ts — including the tool-use discipline rules (§7.5).
SYSTEM_PROMPT = """You are a financial-deck assistant embedded in a PowerPoint task pane.
You help analysts build client-facing financial presentations from branded templates and
uploaded Excel workbooks.

Workbooks: the user uploads .xlsx files through the task pane. Each upload is announced in
the conversation with its file_id and sheet inventory. To answer anything about workbook
contents, use the tools — list_workbook_structure to orient yourself, then read_excel_range
for the exact cells you need.

Templates and the deck: branded .pptx templates live in a template store, and you can edit
the deck open in PowerPoint. To build slides:
1. list_templates, then load_template for the placeholder inventory (named shapes per slide,
   each marked text or number).
2. read_excel_range for every workbook cell that will feed a placeholder.
3. insert_template_slides to append the template's slides to the open deck — it returns the
   new slides' ids.
4. apply_slide_content per slide to write text into the named placeholder shapes.
Use get_deck_state to inspect the open deck (slide ids, shape names, current text) whenever
you are unsure of its contents.

Using tools (read carefully):
- Invoke tools ONLY through the function-calling interface. Never write a tool call as JSON,
  a code block, or prose in your reply, and never show the user tool names, arguments, or
  channel markers. The user sees a short activity chip for each real call — your visible text
  is for explanations and results only.
- When the request is clear, act immediately. Do not ask the user to confirm or to pick when
  the choice is unambiguous — if only one template exists, or the user already named it, just
  load and insert it. Chain the necessary tool calls yourself rather than narrating a plan and
  stopping. Only ask a clarifying question when genuinely blocked (e.g. no workbook uploaded).

Hard rules:
- Financial figures must come from workbook cells, never from your own memory or arithmetic.
  When you state or place a number, quote it exactly as returned by read_excel_range
  (prefer the cell's "text" display form) and name its source cell (e.g. "Revenue FY24:
  1,234,500 — Sheet1!C7"). If you have not read a cell, do not guess its value.
- Every apply_slide_content fill whose text contains a workbook figure must include its
  source {file_id, sheet, cell}.
- Never rescale or convert a figure yourself (e.g. dollars to $M). Place the value as read,
  or ask the user how to proceed — verified transforms arrive in a later milestone.

(Deterministic verification of placements arrives in the next milestone; placements are not
yet blocked on mismatch.)"""


def _sanitize_tool_name(name: str) -> str:
    """Strip gpt-oss "harmony" artifacts that can leak into a tool name, e.g.
    "insert_template_slides<|channel|>commentary" -> "insert_template_slides"."""
    return name.split("<|")[0].strip() or name.strip()


def _sanitize_response_parts(response) -> None:
    for part in response.parts:
        if isinstance(part, ToolCallPart):
            cleaned = _sanitize_tool_name(part.tool_name)
            if cleaned != part.tool_name:
                logger.info("sanitized tool name %r -> %r", part.tool_name, cleaned)
                part.tool_name = cleaned


class _SanitizingModel(WrapperModel):
    """Cleans harmony artifacts from tool names wherever the model emits them —
    both the non-streaming response and the final accumulated streamed response
    that PydanticAI dispatches tools from. Defensive: never let the hardening
    itself break a run if PydanticAI internals shift."""

    async def request(self, messages, model_settings, model_request_parameters):
        response = await super().request(messages, model_settings, model_request_parameters)
        try:
            _sanitize_response_parts(response)
        except Exception as err:  # pragma: no cover - defensive
            logger.warning("tool-name sanitization (request) skipped: %s", err)
        return response

    @asynccontextmanager
    async def request_stream(self, messages, model_settings, model_request_parameters, run_context=None):
        async with super().request_stream(
            messages, model_settings, model_request_parameters, run_context
        ) as stream:
            try:
                _orig_get = stream.get

                def _sanitized_get():
                    response = _orig_get()
                    _sanitize_response_parts(response)
                    return response

                stream.get = _sanitized_get  # shadow the instance method
            except Exception as err:  # pragma: no cover - defensive
                logger.warning("tool-name sanitization (stream) skipped: %s", err)
            yield stream


# OpenRouter passthroughs ride in extra_body, mirroring the Node request shape:
#   reasoning: { effort }            — gpt-oss low|medium|high (ignored by non-reasoning models)
#   provider: { require_parameters } — only route to providers that honor `tools`,
#                                      so gpt-oss isn't served by a backend that drops
#                                      tool calling and narrates the call as text.
MODEL_SETTINGS = OpenAIChatModelSettings(
    # NB: we send the token cap as `max_tokens` via extra_body rather than the
    # `max_tokens` model setting. PydanticAI maps that setting to the newer
    # `max_completion_tokens` param, which no gpt-oss provider advertises — so
    # with `provider.require_parameters` it filters out every endpoint and 404s.
    # The Node backend sends `max_tokens`; matching that keeps gpt-oss routable.
    # We also deliberately do NOT set parallel_tool_calls for the same reason;
    # sequential execution is enforced by ChatSession.tool_lock (see tools.py).
    extra_headers={"X-Title": "PowerPoint Financial-Deck Agent"},
    extra_body={
        "max_tokens": settings.max_output_tokens,
        "reasoning": {"effort": settings.reasoning_effort},
        "provider": {"require_parameters": True},
    },
)

_base_model = OpenAIChatModel(
    settings.model,
    provider=OpenAIProvider(
        base_url=settings.openrouter_base_url,
        # A missing key must fail the request, not crash boot (matches Node's lazy
        # client). OpenAIProvider requires a non-empty string, so pass a placeholder.
        api_key=settings.openrouter_api_key or "MISSING_OPENROUTER_API_KEY",
    ),
)

model = _SanitizingModel(_base_model)

agent = Agent(
    model,
    deps_type=ChatSession,
    system_prompt=SYSTEM_PROMPT,
    model_settings=MODEL_SETTINGS,
)

register_tools(agent)
