"""Per-request chat session — the glue that lets the SSE generator and the tool
functions share state (plan §7.2).

The agent runs in a background task and pushes frames onto `out`; the SSE
generator drains `out` and writes them to the response. Client/Office.js tools
push a `client_tool` frame and then block on a future (kept in the process-level
relay registry, keyed by callId) until the pane POSTs its result back.
"""

from __future__ import annotations

import asyncio
from dataclasses import dataclass, field

# A frame is (event_name, data_dict); None is the sentinel that closes the stream.
Frame = tuple[str, dict] | None


@dataclass
class ChatSession:
    out: "asyncio.Queue[Frame]" = field(default_factory=asyncio.Queue)
    # callIds of relay calls opened by this request — settled with an error if the
    # pane disconnects mid-turn.
    request_call_ids: set[str] = field(default_factory=set)
    client_gone: asyncio.Event = field(default_factory=asyncio.Event)
    # Serializes tool execution so frames never interleave — matches the Node
    # for-loop and the pane's assumption that no frames arrive between a
    # client_tool and its result. (We can't use OpenRouter parallel_tool_calls:
    # sending that param trips `provider.require_parameters` and 404s gpt-oss.)
    tool_lock: asyncio.Lock = field(default_factory=asyncio.Lock)

    async def emit(self, event: str, data: dict) -> None:
        await self.out.put((event, data))
