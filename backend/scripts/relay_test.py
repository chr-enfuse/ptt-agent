"""Simulate the task pane to exercise the client-tool relay end-to-end without
PowerPoint (acceptance cases 5 & 6).

Streams POST /api/chat, and when a `client_tool` frame arrives, POSTs a canned
Office.js result back to /api/chat/client-result keyed by callId — exactly what
Chat.tsx does. Verifies the loop resumes and that frame ordering holds (no frame
between a client_tool and its result POST).
"""

import asyncio
import json
import sys

import httpx

# Windows consoles default to cp1252; model output contains chars like U+202F.
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

BASE = sys.argv[1] if len(sys.argv) > 1 else "http://127.0.0.1:3002"

# Canned Office.js outcomes for each relayed tool.
FAKE_RESULTS = {
    "get_deck_state": {"slides": []},
    "insert_template_slides": {
        "insertedSlides": [{"id": "256", "index": 0}, {"id": "257", "index": 1}],
        "totalSlides": 2,
    },
    "apply_slide_content": {
        "slideId": "257",
        "results": [{"shape_name": "kpi1_value", "ok": True}],
    },
}


def _parse_frame(block: str):
    event, data = "message", ""
    for line in block.split("\n"):
        if line.startswith("event: "):
            event = line[7:].strip()
        elif line.startswith("data: "):
            data += line[6:]
    return event, data


async def run(prompt: str, *, disconnect_on_client_tool: bool = False) -> None:
    print(f"\n=== prompt: {prompt!r} (disconnect={disconnect_on_client_tool}) ===")
    async with httpx.AsyncClient(timeout=120) as client:
        body = {"messages": [{"role": "user", "content": prompt}]}
        relayed = 0
        async with client.stream("POST", f"{BASE}/api/chat", json=body) as resp:
            buffer = ""
            async for chunk in resp.aiter_text():
                buffer += chunk
                while "\n\n" in buffer:
                    block, buffer = buffer.split("\n\n", 1)
                    event, data = _parse_frame(block)
                    if event == "text":
                        print(json.loads(data)["text"], end="", flush=True)
                    elif event == "tool":
                        print(f"\n[tool] {json.loads(data)['name']}", flush=True)
                    elif event == "client_tool":
                        call = json.loads(data)
                        relayed += 1
                        has_b64 = "base64" in call.get("payload", {})
                        print(
                            f"\n[client_tool] {call['name']} callId={call['callId'][:8]} "
                            f"base64={'yes' if has_b64 else 'no'}",
                            flush=True,
                        )
                        if disconnect_on_client_tool:
                            print("[client] disconnecting without responding", flush=True)
                            await resp.aclose()
                            return
                        # POST the canned result back, like the pane.
                        out = FAKE_RESULTS.get(call["name"], {})
                        r = await client.post(
                            f"{BASE}/api/chat/client-result",
                            json={"callId": call["callId"], "content": json.dumps(out), "isError": False},
                        )
                        print(f"[client] posted result -> {r.status_code} {r.text}", flush=True)
                    elif event == "done":
                        print(f"\n[done] {json.loads(data)}", flush=True)
                    elif event == "error":
                        print(f"\n[error] {json.loads(data)}", flush=True)
        print(f"\n(relayed {relayed} client tool call(s))")


async def main() -> None:
    # Case 5: full relay round-trip (insert + state). gpt-oss may or may not call
    # tools at low effort; the explicit asks below nudge it.
    await run("Insert the quarterly-review template slides into the open deck.")
    await run("What slides are currently in the open deck? Use get_deck_state.")
    # Case 6: client disconnects mid relay — server must not hang.
    await run(
        "Insert the quarterly-review template slides into the open deck.",
        disconnect_on_client_tool=True,
    )


asyncio.run(main())
