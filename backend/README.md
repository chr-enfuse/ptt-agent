# backend-py — Python/PydanticAI port of the deck-agent backend

A drop-in replacement for the Node `backend/` service, built on **FastAPI +
PydanticAI**, driving the same OpenRouter chat + tool-use loop. The PowerPoint
add-in (`addin/`) is untouched: the endpoints and SSE event shapes are identical,
and the server listens on **port 3001** (where the pane expects it).

Status: **feature-complete, parity-verified headlessly** (see Acceptance below).
The Node backend remains the default until a human runs the full end-to-end deck
build in PowerPoint and flips the cutover.

## Run

```powershell
cd backend-py
# one-time: create the venv + install deps
py -m venv .venv ; .\.venv\Scripts\python -m pip install -e .
# copy env and set the key
copy .env.example .env   # then edit OPENROUTER_API_KEY
# run on :3001 (stop the Node backend first — only one backend per port)
.\.venv\Scripts\python -m uvicorn app.main:app --host 0.0.0.0 --port 3001
```

The OpenRouter key lives only here. A missing key warns at boot and fails the
request (it does not crash), matching the Node behavior. Model and reasoning
effort are env-configurable (`MODEL`, `REASONING_EFFORT`); see `.env.example`.

## Layout

| File | Role |
|---|---|
| `app/main.py` | FastAPI app, CORS, `/healthz`, router mounting |
| `app/config.py` | pydantic-settings env config |
| `app/agent.py` | PydanticAI Agent: OpenRouter model, system prompt, tool registration, gpt-oss hardening |
| `app/chat.py` | `POST /api/chat` (SSE) + `POST /api/chat/client-result`; the relay glue |
| `app/session.py` | per-request `ChatSession`: out-queue + tool-lock + disconnect event |
| `app/relay.py` | client-tool relay registry (callId → future) |
| `app/tools.py` | the 7 agent tools (4 backend + 3 client/Office.js) |
| `app/excel.py` | openpyxl readers + the Excel **display-text formatter** |
| `app/workbooks.py` | upload/structure endpoints + in-process workbook store |
| `app/templates.py` | template store (reads repo-root `templates.json`) |
| `templates/` | committed `.pptx` artifacts (copied from `backend/templates/`) |
| `scripts/` | verification helpers (workbook gen, relay sim, request probe) |

## Design notes (the non-obvious bits)

- **SSE relay (the make-or-break piece).** The agent runs in a background task and
  pushes frames onto `ChatSession.out`; the SSE generator drains the queue
  independently. This decoupling is essential: a client tool emits its
  `client_tool` frame and then *blocks* awaiting the pane's result — if the agent
  loop and the frame writer were the same coroutine, the frame would never flush
  and the relay would deadlock. The pane assumes **no frame arrives between a
  `client_tool` and its `/client-result` POST**, so all tool execution is
  serialized through `ChatSession.tool_lock`.

- **`max_tokens`, not `max_completion_tokens`.** PydanticAI's `max_tokens` model
  setting serializes to OpenAI's newer `max_completion_tokens`, which no gpt-oss
  OpenRouter provider advertises — so with `provider.require_parameters: true` it
  filters out every endpoint and 404s. We send `max_tokens` via `extra_body`
  instead (top-level, exactly like the Node OpenAI SDK). Same reason we don't set
  `parallel_tool_calls`.

- **OpenRouter passthroughs** (`reasoning`, `provider.require_parameters`) ride in
  `extra_body`; they aren't first-class OpenAI params.

- **gpt-oss tool-name sanitization.** `_SanitizingModel` (a `WrapperModel`) strips
  `<|channel|>…` harmony artifacts from tool names before PydanticAI dispatches,
  on both the streaming and non-streaming paths. Defensive — never breaks a run.

- **Excel display text DIVERGES from Node (intentional).** exceljs `cell.text`
  actually returns the *raw* number (`13456250`) for the common financial formats;
  the Python backend formats them properly (`13,456,250`, `$13,456,250`, `15.3%`)
  via `app/excel.py::format_cell_text`. This was a deliberate, approved decision —
  the Node↔Python SSE diff therefore matches only for non-numeric cases. `raw` and
  `numberFormat` still carry the unformatted value + format code.

## Acceptance (all passing, headless)

1. `GET /healthz` → `{ok:true, service:"deck-agent-backend"}`.
2. Plain chat streams `text` then `done` — **byte-identical SSE to the Node backend**.
3. Backend tool loop: "list the templates" → `event: tool list_templates` → final text.
4. Workbook upload → `read_excel_range` returns formatted `text` for currency/thousands/percent.
5. Client relay round-trips through `/client-result` (insert with base64 in payload, get_deck_state).
6. Client disconnect mid-relay → clean `Error: the task pane disconnected`, no hang (120s timeout in code).
7. Per-round timing + tool-result diagnostics logged.

`scripts/relay_test.py` simulates the pane to exercise 5 & 6 without PowerPoint.

## Remaining for cutover

- Full end-to-end deck build from the actual task pane in PowerPoint (needs a human).
- Flip the run command to `backend-py` on :3001; update root `README.md`/`CLAUDE.md`.
- Retire `backend/` only after a green end-to-end run. Rollback = run Node `backend/` again.
