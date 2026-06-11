"""FastAPI app — the Python port of backend/src/index.ts.

Listens on PORT (3001 by default, where the add-in expects it), allows the
add-in dev origin via CORS, and serves no static files (the task pane comes only
from the webpack dev server, exactly as before).
"""

from __future__ import annotations

import logging

from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware

from .chat import router as chat_router
from .config import settings
from .workbooks import router as workbooks_router

logging.basicConfig(level=logging.INFO, format="%(message)s")
logger = logging.getLogger("deck-agent")

app = FastAPI(title="deck-agent-backend")

# The task pane is served from the add-in dev server (https://localhost:3000).
app.add_middleware(
    CORSMiddleware,
    allow_origins=settings.allowed_origins_list,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.get("/healthz")
async def healthz() -> dict[str, object]:
    return {"ok": True, "service": "deck-agent-backend"}


app.include_router(chat_router, prefix="/api/chat")
app.include_router(workbooks_router, prefix="/api/workbooks")


@app.on_event("startup")
async def _startup() -> None:
    logger.info("deck-agent-backend listening on http://localhost:%d", settings.port)
    if not settings.openrouter_api_key:
        logger.warning(
            "WARNING: OPENROUTER_API_KEY is not set — /api/chat will fail. "
            "Copy .env.example to .env and set it."
        )


def main() -> None:
    import uvicorn

    uvicorn.run("app.main:app", host="0.0.0.0", port=settings.port, log_level="info")


if __name__ == "__main__":
    main()
