# PowerPoint Financial-Deck Agent

A conversational agent embedded directly in a PowerPoint task pane. Analysts chat with the agent
to build client-facing financial presentations: they pick a branded `.pptx` template, upload an
Excel workbook, and the agent maps workbook cells into slide placeholders. A **deterministic
verifier** guarantees that every number placed on a slide is exactly equal to its source cell —
numbers are never typed by the model, only mapped by it.

---

## Repository layout

```
glr/
├── addin/      # Frontend — Office Add-in task pane (React + TypeScript + Office.js)
├── backend/    # Backend — Python/FastAPI service (PydanticAI + OpenRouter + openpyxl)
└── README.md   # You are here
```

---

## How the system works

```
┌────────────────── PowerPoint (Windows desktop / web) ──────────────────┐
│  Task pane add-in  (addin/ — served from https://localhost:3000)        │
│   • Chat UI (streams agent replies)                                     │
│   • Office.js calls that read/write the open deck                       │
│   • (later) verification report panel: value → source cell audit        │
└───────────────┬─────────────────────────────────────────────────────────┘
                │ HTTPS fetch (SSE for chat)
                ▼
┌──────────────────── Backend (backend/ — http://localhost:3001) ─────────┐
│   • Holds OPENROUTER_API_KEY (never shipped to the task pane)          │
│   • PydanticAI tool-use loop; streams the configured OpenRouter model   │
│   • Parses uploaded .xlsx deterministically with openpyxl               │
│   • Template store + client-tool relay to the task pane                 │
└───────────────┬──────────────────────────────────────────────────────────┘
                │ OpenAI-compatible chat-completions API
                ▼
        OpenRouter (openrouter.ai/api/v1)
```

### Frontend — `addin/`

An **Office Add-in** scaffolded with the Yeoman generator for Office Add-ins (`yo office`,
React + TypeScript template). Office Add-ins are web apps that Office hosts inside a task pane;
they manipulate the open document through the [Office.js](https://learn.microsoft.com/office/dev/add-ins/)
JavaScript API.

Key pieces:

| File | Purpose |
|---|---|
| `manifest.xml` | Tells PowerPoint about the add-in: ID, name, ribbon button, and that the UI lives at `https://localhost:3000/taskpane.html`. Sideloading registers this manifest with desktop PowerPoint. |
| `src/taskpane/index.tsx` | Entry point — waits for `Office.onReady()`, then renders the React app. |
| `src/taskpane/components/App.tsx` | Shell: header + chat + an Office.js sanity-check widget (inserts text into the current slide). |
| `src/taskpane/components/Chat.tsx` | The chat UI. POSTs the conversation to the backend's `/api/chat` and consumes the **Server-Sent Events** stream, appending text deltas to the last assistant bubble as they arrive. |
| `src/taskpane/taskpane.ts` | Office.js helpers (currently `insertText` — writes a text box onto the selected slide via `PowerPoint.run`). Client-executed agent tools (`get_deck_state`, `apply_slide_content`, …) will live here. |
| `webpack.config.js` | Dev server on **https://localhost:3000** (Office requires HTTPS even for localhost — handled by a locally-trusted dev certificate). |

Dev workflow: `npm start` runs `office-addin-debugging start manifest.xml`, which (1) starts the
webpack dev server, (2) registers the manifest with desktop PowerPoint ("sideloading"), and
(3) launches PowerPoint with the add-in available on the Home ribbon. Edits to `src/` hot-reload.

> **Why the task pane matters architecturally:** Office.js can only run inside the Office host.
> Tools that touch the open deck must execute here; everything else (Excel parsing, verification,
> the model tool-use loop) runs on the backend. The backend relays PowerPoint tool calls to the
> task pane over the SSE channel and waits for the result POSTed back to `/api/chat/client-result`.

### Backend — `backend/`

A **Python + FastAPI** service driven by **PydanticAI**. It exists so the OpenRouter API key never
reaches the browser, and so numbers can be extracted and verified by *code* rather than generated
by the model. It runs from a local virtualenv (see setup below) and listens on `:3001`.

| File | Purpose |
|---|---|
| `app/main.py` | FastAPI wiring: CORS (allows only the add-in origin), `/healthz`, router mounting. |
| `app/config.py` | `pydantic-settings` env config (`OPENROUTER_API_KEY`, `MODEL`, `REASONING_EFFORT`, …). |
| `app/agent.py` | The PydanticAI `Agent`: OpenRouter model wiring, system prompt, tool registration, and the gpt-oss hardening (tool-name sanitization, `max_tokens` via `extra_body`). Default model **`openai/gpt-oss-20b`** (`MODEL`/`REASONING_EFFORT` env-configurable). |
| `app/chat.py` | `POST /api/chat` (SSE: `text` / `tool` / `client_tool` / `done` / `error`) + `POST /api/chat/client-result` — the relay glue. PowerPoint tools are relayed to the task pane and their results POSTed back. |
| `app/tools.py` | The 7 agent tools — 4 backend (workbook/template reads) + 3 client/Office.js (relayed to the pane). |
| `app/excel.py` | openpyxl readers + the Excel display-text formatter. |
| `app/workbooks.py` | `POST /api/workbooks` (multipart `.xlsx` → in-memory store) + `GET /api/workbooks/:fileId/structure`. |
| `app/templates.py` | Template store; reads the repo-root `templates.json` + base64s the `.pptx` for the relay. |
| `.env.example` | Template for required environment variables. |

API surface:

| Endpoint | Description |
|---|---|
| `GET /healthz` | Liveness check → `{ok: true, service: "deck-agent-backend"}` |
| `POST /api/chat` | `{messages: {role, content}[]}` → SSE stream of `event: text` deltas (plus `tool` / `client_tool`), then `event: done` |
| `POST /api/chat/client-result` | `{callId, content, isError?}` — the task pane returns a relayed PowerPoint tool's outcome |
| `POST /api/workbooks` | multipart field `file` (`.xlsx`) → `{fileId, filename, sheets[]}` |
| `GET /api/workbooks/:fileId/structure` | Sheets + used ranges + first-row headers |

### The anti-hallucination core (where this is heading)

The agent's job is **mapping, not transcribing**: tools return exact cell values as structured
data, every proposed placement must carry a `{workbook, sheet, cell, rawValue, …}` provenance
reference, and before anything is written to a slide, backend code re-reads the cited cell and
asserts exact equality. An LLM "semantic judge" reviews labels/units/periods but can never edit a
number.

---

## Setup for a new contributor

### Prerequisites

- **Windows** with **desktop PowerPoint** (Microsoft 365) — the add-in also runs in PowerPoint
  on the web, but the dev sideload flow below targets Windows desktop
- **Node.js LTS ≥ 20** (`winget install OpenJS.NodeJS.LTS`) — for the add-in dev server
- **Python ≥ 3.12** (`winget install Python.Python.3.12`) — for the backend
- **Git**
- An **OpenRouter API key** (for the chat backend)

### 1. Clone and install

```powershell
git clone <repo-url> glr
cd glr

cd backend ; py -m venv .venv ; .\.venv\Scripts\python -m pip install -e . ; cd ..
cd addin   ; npm install ; cd ..
```

### 2. Configure the backend

```powershell
cd backend
copy .env.example .env
# edit .env and set:
#   OPENROUTER_API_KEY=sk-or-...
#   MODEL=openai/gpt-oss-20b   (optional — this is the default)
```

`.env` is gitignored — the key must never be committed and never appears in add-in code.

### 3. Run the backend

```powershell
cd backend
.\.venv\Scripts\python -m uvicorn app.main:app --host 127.0.0.1 --port 3001
```

Sanity check: `curl http://localhost:3001/healthz` → `{"ok":true,...}`.

### 4. Run the add-in (sideloads into PowerPoint)

```powershell
cd addin
npm start
```

First run only: a Windows dialog asks to install the **"Developer CA for Microsoft Office
Add-ins"** certificate — click **Yes**. Office requires the task pane to be served over HTTPS,
and this locally-generated CA makes `https://localhost:3000` trusted. (If the dialog was missed
or the pane shows a security error, run `npx office-addin-dev-certs install` inside `addin/`
and restart.)

`npm start` then starts the dev server, registers `manifest.xml` with PowerPoint, and launches
PowerPoint. Click **Show Task Pane** on the **Home** ribbon.

### 5. Verify it works

1. In the task pane, type a message → you should see a streamed reply from the model
   (requires the backend from step 3 to be running).
2. Use the text box under *"Office.js sanity check"* → **Insert text** writes into the current
   slide, proving the Office.js bridge works with no backend at all.

### Stopping

```powershell
# in addin/ — stops debugging and unregisters the sideloaded add-in
npm run stop

# backend: Ctrl+C the uvicorn process
```

---

## Common commands

| Task | Directory | Command |
|---|---|---|
| Backend dev server | `backend/` | `.\.venv\Scripts\python -m uvicorn app.main:app --port 3001` |
| Backend dev server (auto-reload) | `backend/` | `.\.venv\Scripts\python -m uvicorn app.main:app --port 3001 --reload` |
| Install / update backend deps | `backend/` | `.\.venv\Scripts\python -m pip install -e .` |
| Dev server + sideload PowerPoint | `addin/` | `npm start` |
| Dev server only (no sideload) | `addin/` | `npm run dev-server` |
| Build add-in bundle | `addin/` | `npm run build` (prod) / `npm run build:dev` |
| Validate the manifest | `addin/` | `npm run validate` |
| Lint | `addin/` | `npm run lint` / `npm run lint:fix` |
| Stop & remove sideloaded add-in | `addin/` | `npm run stop` |

---

## Troubleshooting

| Symptom | Fix |
|---|---|
| Task pane is blank / shows a certificate error | `cd addin ; npx office-addin-dev-certs install`, then restart `npm start`. Verify with `npx office-addin-dev-certs verify`. |
| `EADDRINUSE: port 3000` | A previous dev server is still running. Kill it: `Get-NetTCPConnection -LocalPort 3000 -State Listen \| % { Stop-Process -Id $_.OwningProcess -Force }` |
| Chat shows "backend responded …" error | Backend isn't running or `.env` is missing the API key. Start it with the uvicorn command in `backend/` and check the terminal for the `OPENROUTER_API_KEY` warning. |
| Add-in button missing in PowerPoint | Re-run `npm start` in `addin/` (re-sideloads the manifest), or `npm run stop` then `npm start` for a clean re-register. |
| Office.js API errors on older Office builds | The PowerPoint JS API set varies by build. Target a current Microsoft 365 channel. |

---

## Security notes

- `OPENROUTER_API_KEY` lives only in `backend/.env` (gitignored); the task pane authenticates to
  the backend, never to OpenRouter.
- Uploaded workbooks are held in memory only (v1 scaffolding); `*.xlsx` / `*.pptx` are gitignored
  so client financial data can't be committed accidentally.
- CORS on the backend is restricted to the add-in dev origin (`https://localhost:3000`).
