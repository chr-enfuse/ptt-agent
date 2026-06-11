"""Agent tools — the Python port of the TOOLS array + executeBackendTool /
executeClientTool from backend/src/routes/chat.ts.

PydanticAI derives each tool's JSON schema from its type hints and docstring, so
the hand-written JSON Schemas are gone. Backend tools emit a `tool` frame and run
deterministic code; client/Office.js tools relay through `execute_client_tool`
(which emits the `client_tool` frame and blocks on the pane). Tool errors are
returned as `Error: ...` strings so the model can recover, mirroring the Node
`ToolOutcome.isError` content.
"""

from __future__ import annotations

import json
import logging
import time

from pydantic import BaseModel, Field
from pydantic_ai import Agent, RunContext

from . import excel, templates
from .relay import execute_client_tool
from .session import ChatSession

logger = logging.getLogger("deck-agent.tools")

MAX_CELLS = excel.MAX_CELLS_PER_READ


def _preview(s: str, max_len: int = 300) -> str:
    one_line = " ".join(s.split())
    return one_line if len(one_line) <= max_len else f"{one_line[:max_len]}…[+{len(one_line) - max_len}]"


# --- apply_slide_content argument models (drive the JSON schema + validation) ---


class FillSource(BaseModel):
    file_id: str = Field(description="Uploaded workbook file_id.")
    sheet: str = Field(description="Worksheet name.")
    cell: str = Field(description='Cell address in A1 notation, e.g. "C7".')


class Fill(BaseModel):
    shape_name: str = Field(description="Placeholder shape name, from load_template or get_deck_state.")
    text: str = Field(description="Exact text to place in the shape.")
    source: FillSource | None = Field(
        default=None,
        description=(
            "Provenance of the figure in `text` — the workbook cell it was read from. "
            "Required whenever the text contains a workbook figure."
        ),
    )


def register_tools(agent: Agent) -> None:
    """Attach all seven tools to the agent."""

    # ----- backend tools (deterministic, run in-process) -----------------------

    @agent.tool
    async def list_workbook_structure(ctx: RunContext[ChatSession], file_id: str) -> str:
        """List the sheets of an uploaded Excel workbook: name, row/column counts, and
        first-row headers. Call this first to orient yourself in a workbook before
        reading cell ranges.

        Args:
            file_id: The file_id of an uploaded workbook, as announced in the conversation.
        """
        return await _run_backend(
            ctx,
            "list_workbook_structure",
            {"file_id": file_id},
            lambda: excel.list_workbook_structure(file_id),
        )

    @agent.tool
    async def read_excel_range(ctx: RunContext[ChatSession], file_id: str, sheet: str, range: str) -> str:
        """Read exact cell values from an uploaded Excel workbook. Returns structured
        cells {addr, raw, formula?, numberFormat, text} where raw is the stored value
        (for formula cells, the cached computed result) and text is the display string.
        Call this whenever the user asks about workbook contents — never answer from
        memory. Max 400 cells per call; read large sheets in chunks.

        Args:
            file_id: The file_id of an uploaded workbook, as announced in the conversation.
            sheet: Worksheet name, exactly as returned by list_workbook_structure.
            range: Cell or range in A1 notation, e.g. "B7" or "A1:C10".
        """
        return await _run_backend(
            ctx,
            "read_excel_range",
            {"file_id": file_id, "sheet": sheet, "range": range},
            lambda: excel.read_excel_range(file_id, sheet, range),
        )

    @agent.tool
    async def list_templates(ctx: RunContext[ChatSession]) -> str:
        """List the branded PowerPoint templates available in the template store: id,
        name, description, and slide count. Call load_template for a template's
        placeholder inventory."""
        return await _run_backend(ctx, "list_templates", {}, templates.list_templates)

    @agent.tool
    async def load_template(ctx: RunContext[ChatSession], template_id: str) -> str:
        """Load a template's full placeholder inventory: each slide's named placeholder
        shapes with a description and kind (text | number). Use this to plan which
        workbook cell feeds which placeholder before inserting slides. Placeholders of
        kind number must be filled from workbook cells with a source reference.

        Args:
            template_id: Template id, as returned by list_templates.
        """
        return await _run_backend(
            ctx, "load_template", {"template_id": template_id}, lambda: templates.load_template(template_id)
        )

    # ----- client/Office.js tools (relayed to the task pane) -------------------

    @agent.tool
    async def get_deck_state(ctx: RunContext[ChatSession]) -> str:
        """Read the deck currently open in PowerPoint: each slide's id, index, and
        shapes ({name, text}). Executes in the task pane via Office.js."""
        return await _run_client(ctx, "get_deck_state", {})

    @agent.tool
    async def insert_template_slides(ctx: RunContext[ChatSession], template_id: str) -> str:
        """Append all slides of a branded template to the deck open in PowerPoint.
        Returns the new slides' ids and indexes — use those ids with
        apply_slide_content. Executes in the task pane via Office.js.

        Args:
            template_id: Template id, as returned by list_templates.
        """
        return await _run_client(ctx, "insert_template_slides", {"template_id": template_id})

    @agent.tool
    async def apply_slide_content(ctx: RunContext[ChatSession], slide_id: str, fills: list[Fill]) -> str:
        """Write text into named placeholder shapes on one slide of the open deck. Each
        fill sets one shape's full text. For every fill containing a financial figure,
        quote the figure exactly as returned by read_excel_range and include the source
        cell it was read from. Returns per-fill success. Executes in the task pane via
        Office.js.

        Args:
            slide_id: Slide id, as returned by insert_template_slides or get_deck_state.
            fills: One entry per placeholder shape to fill on this slide.
        """
        payload = {
            "slide_id": slide_id,
            "fills": [f.model_dump(exclude_none=True) for f in fills],
        }
        return await _run_client(ctx, "apply_slide_content", payload)


async def _run_backend(ctx: RunContext[ChatSession], name: str, tool_input: dict, fn) -> str:
    """Run a deterministic backend tool under the session tool-lock: emit the `tool`
    frame, run the code, and return JSON on success or an `Error: ...` string on a
    recoverable tool error (mirrors Node executeBackendTool)."""
    async with ctx.deps.tool_lock:
        await ctx.deps.emit("tool", {"name": name, "input": tool_input})
        start = time.perf_counter()
        try:
            content = json.dumps(fn())
            is_error = False
        except (excel.ExcelToolError, templates.TemplateToolError) as err:
            content = f"Error: {err}"
            is_error = True
        ms = int((time.perf_counter() - start) * 1000)
        logger.info("  ← %s %s(%dms) %s", name, "ERROR " if is_error else "", ms, _preview(content))
        return content


async def _run_client(ctx: RunContext[ChatSession], name: str, payload: dict) -> str:
    """Relay a client tool to the pane under the session tool-lock (so its
    client_tool frame and result never interleave with other tools' frames)."""
    async with ctx.deps.tool_lock:
        start = time.perf_counter()
        outcome = await execute_client_tool(name, payload, ctx.deps)
        ms = int((time.perf_counter() - start) * 1000)
        logger.info(
            "  ← %s %s(%dms) %s", name, "ERROR " if outcome.is_error else "", ms, _preview(outcome.content)
        )
        return outcome.content
