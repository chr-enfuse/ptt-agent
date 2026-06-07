import { Router } from "express";
import Anthropic from "@anthropic-ai/sdk";
import { ExcelToolError, listWorkbookStructure, readExcelRange, MAX_CELLS_PER_READ } from "../excel.js";

export const chatRouter = Router();

const client = new Anthropic(); // reads ANTHROPIC_API_KEY from the environment

const MODEL = "claude-opus-4-8";

/** Safety valve on the tool-use loop — one user turn never runs more rounds than this. */
const MAX_LOOP_ITERATIONS = 16;

// Keep this byte-stable: it sits at the front of the prompt prefix so it can be
// served from the prompt cache. Volatile, per-request context goes in messages.
const SYSTEM_PROMPT = `You are a financial-deck assistant embedded in a PowerPoint task pane.
You help analysts build client-facing financial presentations from branded templates and
uploaded Excel workbooks.

Workbooks: the user uploads .xlsx files through the task pane. Each upload is announced in
the conversation with its file_id and sheet inventory. To answer anything about workbook
contents, use the tools — list_workbook_structure to orient yourself, then read_excel_range
for the exact cells you need.

Hard rule: financial figures must come from workbook cells, never from your own memory or
arithmetic. When you state a number, quote it exactly as returned by read_excel_range and
name its source cell (e.g. "Revenue FY24: 1,234,500 — Sheet1!C7"). If you have not read a
cell, do not guess its value.

(Template fill and verified slide placement arrive in later milestones; for now you read
workbooks and converse.)`;

// Tool definitions render ahead of the system prompt in the prompt prefix — keep them
// byte-stable too (the cache breakpoint on the system block covers tools + system).
const TOOLS: Anthropic.Tool[] = [
  {
    name: "list_workbook_structure",
    description:
      "List the sheets of an uploaded Excel workbook: name, row/column counts, and first-row headers. " +
      "Call this first to orient yourself in a workbook before reading cell ranges.",
    input_schema: {
      type: "object",
      properties: {
        file_id: {
          type: "string",
          description: "The file_id of an uploaded workbook, as announced in the conversation.",
        },
      },
      required: ["file_id"],
    },
  },
  {
    name: "read_excel_range",
    description:
      "Read exact cell values from an uploaded Excel workbook. Returns structured cells " +
      "{addr, raw, formula?, numberFormat, text} where raw is the stored value (for formula " +
      "cells, the cached computed result) and text is the display string. Call this whenever " +
      `the user asks about workbook contents — never answer from memory. Max ${MAX_CELLS_PER_READ} cells per call; ` +
      "read large sheets in chunks.",
    input_schema: {
      type: "object",
      properties: {
        file_id: {
          type: "string",
          description: "The file_id of an uploaded workbook, as announced in the conversation.",
        },
        sheet: {
          type: "string",
          description: "Worksheet name, exactly as returned by list_workbook_structure.",
        },
        range: {
          type: "string",
          description: 'Cell or range in A1 notation, e.g. "B7" or "A1:C10".',
        },
      },
      required: ["file_id", "sheet", "range"],
    },
  },
];

interface ToolOutcome {
  content: string;
  isError: boolean;
}

function executeTool(name: string, input: Record<string, unknown>): ToolOutcome {
  try {
    switch (name) {
      case "list_workbook_structure":
        return { content: JSON.stringify(listWorkbookStructure(String(input.file_id))), isError: false };
      case "read_excel_range":
        return {
          content: JSON.stringify(
            readExcelRange(String(input.file_id), String(input.sheet), String(input.range)),
          ),
          isError: false,
        };
      default:
        return { content: `unknown tool "${name}"`, isError: true };
    }
  } catch (err) {
    if (err instanceof ExcelToolError) {
      return { content: `Error: ${err.message}`, isError: true };
    }
    throw err;
  }
}

/**
 * POST /api/chat
 * Body: { messages: Anthropic.MessageParam[] }
 * Responds with Server-Sent Events:
 *   event: text   data: {"text": "..."}                 (incremental text deltas)
 *   event: tool   data: {"name": "...", "input": {...}} (a tool call is executing)
 *   event: done   data: {"stopReason": "..."}           (stream finished)
 *   event: error  data: {"message": "..."}
 */
chatRouter.post("/", async (req, res) => {
  const messages = req.body?.messages as Anthropic.MessageParam[] | undefined;
  if (!Array.isArray(messages) || messages.length === 0) {
    res.status(400).json({ error: "body must include a non-empty `messages` array" });
    return;
  }

  res.setHeader("Content-Type", "text/event-stream");
  res.setHeader("Cache-Control", "no-cache");
  res.setHeader("Connection", "keep-alive");
  res.flushHeaders();

  const send = (event: string, data: unknown) => {
    res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
  };

  try {
    // Tool-use loop: stream each round to the client; when the model requests tools,
    // execute them backend-side, feed the results back, and continue until end of turn.
    const loopMessages: Anthropic.MessageParam[] = [...messages];

    for (let iteration = 0; ; iteration++) {
      if (iteration >= MAX_LOOP_ITERATIONS) {
        send("error", { message: `tool-use loop exceeded ${MAX_LOOP_ITERATIONS} iterations` });
        return;
      }

      const stream = client.messages.stream({
        model: MODEL,
        max_tokens: 64000,
        thinking: { type: "adaptive" },
        output_config: { effort: "high" },
        tools: TOOLS,
        system: [
          {
            type: "text",
            text: SYSTEM_PROMPT,
            cache_control: { type: "ephemeral" },
          },
        ],
        messages: loopMessages,
      });

      stream.on("text", (delta) => send("text", { text: delta }));

      const message = await stream.finalMessage();

      if (message.stop_reason === "tool_use") {
        loopMessages.push({ role: "assistant", content: message.content });

        const toolResults: Anthropic.ToolResultBlockParam[] = [];
        for (const block of message.content) {
          if (block.type !== "tool_use") continue;
          send("tool", { name: block.name, input: block.input });
          const { content, isError } = executeTool(block.name, block.input as Record<string, unknown>);
          toolResults.push({
            type: "tool_result",
            tool_use_id: block.id,
            content,
            ...(isError ? { is_error: true } : {}),
          });
        }
        loopMessages.push({ role: "user", content: toolResults });
        continue;
      }

      if (message.stop_reason === "pause_turn") {
        // Server-side pause — re-send to let the model resume where it left off.
        loopMessages.push({ role: "assistant", content: message.content });
        continue;
      }

      send("done", { stopReason: message.stop_reason });
      return;
    }
  } catch (err) {
    const message = err instanceof Anthropic.APIError ? `${err.status}: ${err.message}` : String(err);
    console.error("chat error:", message);
    send("error", { message });
  } finally {
    res.end();
  }
});
