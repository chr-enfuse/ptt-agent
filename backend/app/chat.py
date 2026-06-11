"""POST /api/chat (SSE) and POST /api/chat/client-result — the Python port of
backend/src/routes/chat.ts.

The agent runs in a background task (so a tool blocking on the pane never stalls
frame delivery); the SSE generator independently drains the session queue and
writes frames in exactly the Node format: `event: <name>\\ndata: <json>\\n\\n`.
"""

from __future__ import annotations

import asyncio
import json
import logging
import time
import uuid

from fastapi import APIRouter, Request
from fastapi.responses import JSONResponse, StreamingResponse
from pydantic_ai import Agent
from pydantic_ai.messages import (
    ModelRequest,
    ModelResponse,
    PartDeltaEvent,
    PartStartEvent,
    TextPart,
    TextPartDelta,
    UserPromptPart,
)
from pydantic_ai.usage import UsageLimits

from .agent import agent
from .relay import ToolOutcome, fail_request_calls, has_pending, settle_client_call
from .session import ChatSession

logger = logging.getLogger("deck-agent.chat")

router = APIRouter()

# Safety valve on the tool-use loop — one user turn never runs more rounds than this.
MAX_LOOP_ITERATIONS = 16


def _build_history(messages: list[dict]) -> tuple[str, list]:
    """Last message is the prompt; the rest become PydanticAI message history.
    The system prompt lives on the Agent, never in the history."""
    history: list = []
    for m in messages[:-1]:
        content = str(m.get("content", ""))
        if m.get("role") == "user":
            history.append(ModelRequest(parts=[UserPromptPart(content=content)]))
        else:
            history.append(ModelResponse(parts=[TextPart(content=content)]))
    prompt = str(messages[-1].get("content", ""))
    return prompt, history


def _finish_reason(run) -> str | None:
    """The final round's finish reason, as a plain string (the pane ignores its
    value, but we mirror the Node `done.stopReason` shape)."""
    try:
        fr = run.result.response.finish_reason
    except AttributeError:
        return None
    if fr is None:
        return None
    return getattr(fr, "value", None) or str(fr)


def _text_delta(event) -> str | None:
    """Visible text deltas only — reasoning/thinking parts are not streamed to the
    pane (mirrors the Node loop, which streamed only delta.content)."""
    if isinstance(event, PartStartEvent) and isinstance(event.part, TextPart):
        return event.part.content
    if isinstance(event, PartDeltaEvent) and isinstance(event.delta, TextPartDelta):
        return event.delta.content_delta
    return None


async def _run_agent(prompt: str, history: list, session: ChatSession, req_id: str) -> None:
    """Drive the agent loop, pushing text/tool/client_tool/done/error frames onto
    the session queue. Tools emit their own frames via ctx.deps (see tools.py)."""
    req_start = time.perf_counter()
    finish_reason = "stop"
    try:
        async with agent.iter(
            prompt,
            message_history=history,
            deps=session,
            usage_limits=UsageLimits(request_limit=MAX_LOOP_ITERATIONS),
        ) as run:
            iteration = 0
            async for node in run:
                if session.client_gone.is_set():
                    return
                if Agent.is_model_request_node(node):
                    round_start = time.perf_counter()
                    text_chars = 0
                    async with node.stream(run.ctx) as stream:
                        async for event in stream:
                            delta = _text_delta(event)
                            if delta:
                                text_chars += len(delta)
                                await session.emit("text", {"text": delta})
                    iteration += 1
                    logger.info(
                        "[chat %s] round %d%s (llm %dms)",
                        req_id,
                        iteration,
                        f": {text_chars} chars text" if text_chars else "",
                        int((time.perf_counter() - round_start) * 1000),
                    )
            finish_reason = _finish_reason(run) or "stop"
        logger.info(
            "[chat %s] done after %d round(s) in %dms: finish=%s",
            req_id,
            iteration,
            int((time.perf_counter() - req_start) * 1000),
            finish_reason,
        )
        await session.emit("done", {"stopReason": finish_reason})
    except asyncio.CancelledError:
        raise
    except Exception as err:
        message = str(err)
        logger.error("[chat %s] error: %s", req_id, message)
        await session.emit("error", {"message": message})
    finally:
        await session.out.put(None)  # sentinel -> generator closes


@router.post("")
@router.post("/")
async def chat(request: Request):
    """POST /api/chat — streams Server-Sent Events for one user turn."""
    try:
        body = await request.json()
    except Exception:
        body = None
    incoming = (body or {}).get("messages") if isinstance(body, dict) else None
    if not isinstance(incoming, list) or len(incoming) == 0:
        return JSONResponse(
            status_code=400, content={"error": "body must include a non-empty `messages` array"}
        )

    session = ChatSession()
    prompt, history = _build_history(incoming)

    req_id = uuid.uuid4().hex[:8]
    from .config import settings

    logger.info(
        "[chat %s] start: model=%s, effort=%s, %d message(s)",
        req_id,
        settings.model,
        settings.reasoning_effort,
        len(incoming),
    )

    async def event_generator():
        task = asyncio.create_task(_run_agent(prompt, history, session, req_id))
        try:
            while True:
                frame = await session.out.get()
                if frame is None:
                    break
                event, data = frame
                # Compact separators match Node's JSON.stringify byte-for-byte (§10 diff).
                yield f"event: {event}\ndata: {json.dumps(data, separators=(',', ':'))}\n\n"
        finally:
            # Client disconnected (or stream ended): stop the loop and settle any
            # relay calls this request opened — mirrors the Node res.on("close").
            session.client_gone.set()
            fail_request_calls(session, "the task pane disconnected")
            if not task.done():
                task.cancel()

    return StreamingResponse(
        event_generator(),
        media_type="text/event-stream",
        headers={"Cache-Control": "no-cache", "Connection": "keep-alive"},
    )


@router.post("/client-result")
async def client_result(request: Request):
    """POST /api/chat/client-result — the pane posts each relayed tool's outcome here."""
    try:
        body = await request.json()
    except Exception:
        body = {}
    if not isinstance(body, dict):
        body = {}

    call_id = body.get("callId")
    if not isinstance(call_id, str) or not has_pending(call_id):
        return JSONResponse(status_code=404, content={"error": "unknown or expired callId"})

    content = body.get("content")
    content_str = content if isinstance(content, str) else json.dumps(content)
    settle_client_call(call_id, ToolOutcome(content=content_str, is_error=bool(body.get("isError"))))
    return {"ok": True}
