"""Client-tool relay registry (plan §8).

PowerPoint tools can only run in the Office.js runtime, so a client tool emits a
`client_tool` SSE frame and waits here for the task pane to execute it and POST
the outcome back to /api/chat/client-result. The pending-call registry is
process-level and keyed by callId — exactly like the Node `pendingClientCalls`
Map — so the /client-result handler can settle a call without knowing which
request opened it.
"""

from __future__ import annotations

import asyncio
import uuid
from dataclasses import dataclass

from .session import ChatSession
from .templates import TemplateToolError, get_template_base64

CLIENT_TOOL_TIMEOUT_S = 120

# Tools that must execute in the Office.js runtime — relayed to the task pane.
CLIENT_TOOLS = {"get_deck_state", "insert_template_slides", "apply_slide_content"}


@dataclass
class ToolOutcome:
    content: str
    is_error: bool


# callId -> the future the awaiting client tool is blocked on.
_pending: dict[str, "asyncio.Future[ToolOutcome]"] = {}


def settle_client_call(call_id: str, outcome: ToolOutcome) -> bool:
    """Resolve a pending relay call exactly once. Returns False if unknown/expired."""
    fut = _pending.pop(call_id, None)
    if fut is None:
        return False
    if not fut.done():
        fut.set_result(outcome)
    return True


def has_pending(call_id: str) -> bool:
    return call_id in _pending


async def execute_client_tool(name: str, payload: dict, session: ChatSession) -> ToolOutcome:
    """Emit the `client_tool` frame and block until the pane returns a result.

    Server-side data joins the payload here: the template .pptx travels in this
    event only — never into model context.
    """
    payload = dict(payload)
    if name == "insert_template_slides":
        try:
            payload["base64"] = get_template_base64(str(payload.get("template_id")))
        except TemplateToolError as err:
            return ToolOutcome(content=f"Error: {err}", is_error=True)

    loop = asyncio.get_running_loop()
    fut: "asyncio.Future[ToolOutcome]" = loop.create_future()
    call_id = str(uuid.uuid4())
    _pending[call_id] = fut
    session.request_call_ids.add(call_id)

    await session.emit("client_tool", {"callId": call_id, "name": name, "payload": payload})

    try:
        outcome = await asyncio.wait_for(fut, timeout=CLIENT_TOOL_TIMEOUT_S)
    except asyncio.TimeoutError:
        _pending.pop(call_id, None)
        outcome = ToolOutcome(
            content=(
                f"Error: the task pane did not return a result for {name} "
                f"within {CLIENT_TOOL_TIMEOUT_S}s"
            ),
            is_error=True,
        )
    finally:
        session.request_call_ids.discard(call_id)
    return outcome


def fail_request_calls(session: ChatSession, message: str) -> None:
    """On client disconnect, settle every relay call this request opened (plan §8)."""
    for call_id in list(session.request_call_ids):
        settle_client_call(call_id, ToolOutcome(content=f"Error: {message}", is_error=True))
