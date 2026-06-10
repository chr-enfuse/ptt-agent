import { Router } from "express";
import { randomUUID } from "node:crypto";
import OpenAI from "openai";
import { ExcelToolError, listWorkbookStructure, readExcelRange, MAX_CELLS_PER_READ } from "../excel.js";
import { TemplateToolError, listTemplates, loadTemplate, getTemplateBase64 } from "../templates.js";

export const chatRouter = Router();

// The chat is served through OpenRouter's OpenAI-compatible API. The key and
// model both come from the environment so the model is swappable without code
// changes (see .env.example). Constructed lazily: the SDK throws when the key is
// absent, and we want a missing key to fail the request — not crash boot.
let client: OpenAI | null = null;
function getClient(): OpenAI {
  if (!client) {
    client = new OpenAI({
      apiKey: process.env.OPENROUTER_API_KEY,
      baseURL: process.env.OPENROUTER_BASE_URL ?? "https://openrouter.ai/api/v1",
      // Optional OpenRouter attribution headers (used for their app dashboard).
      defaultHeaders: { "X-Title": "PowerPoint Financial-Deck Agent" },
    });
  }
  return client;
}

const MODEL = process.env.MODEL ?? "openai/gpt-oss-20b";

// OpenRouter's unified reasoning control. gpt-oss exposes low|medium|high effort;
// for models without reasoning OpenRouter simply ignores it.
const REASONING_EFFORT = (process.env.REASONING_EFFORT ?? "high") as "low" | "medium" | "high";

// Upper bound on completion tokens per round (reasoning + visible output).
const MAX_TOKENS = Number(process.env.MAX_OUTPUT_TOKENS ?? 32000);

/** Safety valve on the tool-use loop — one user turn never runs more rounds than this. */
const MAX_LOOP_ITERATIONS = 16;

// Chat-completions request shape plus OpenRouter's `reasoning` passthrough,
// which the stock OpenAI types don't model.
type ChatParams = OpenAI.Chat.Completions.ChatCompletionCreateParamsStreaming & {
  reasoning?: { effort: "low" | "medium" | "high" };
};

const SYSTEM_PROMPT = `You are a financial-deck assistant embedded in a PowerPoint task pane.
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
yet blocked on mismatch.)`;

const TOOLS: OpenAI.Chat.Completions.ChatCompletionTool[] = [
  {
    type: "function",
    function: {
      name: "list_workbook_structure",
      description:
        "List the sheets of an uploaded Excel workbook: name, row/column counts, and first-row headers. " +
        "Call this first to orient yourself in a workbook before reading cell ranges.",
      parameters: {
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
  },
  {
    type: "function",
    function: {
      name: "read_excel_range",
      description:
        "Read exact cell values from an uploaded Excel workbook. Returns structured cells " +
        "{addr, raw, formula?, numberFormat, text} where raw is the stored value (for formula " +
        "cells, the cached computed result) and text is the display string. Call this whenever " +
        `the user asks about workbook contents — never answer from memory. Max ${MAX_CELLS_PER_READ} cells per call; ` +
        "read large sheets in chunks.",
      parameters: {
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
  },
  {
    type: "function",
    function: {
      name: "list_templates",
      description:
        "List the branded PowerPoint templates available in the template store: id, name, " +
        "description, and slide count. Call load_template for a template's placeholder inventory.",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "load_template",
      description:
        "Load a template's full placeholder inventory: each slide's named placeholder shapes with " +
        "a description and kind (text | number). Use this to plan which workbook cell feeds which " +
        "placeholder before inserting slides. Placeholders of kind number must be filled from " +
        "workbook cells with a source reference.",
      parameters: {
        type: "object",
        properties: {
          template_id: {
            type: "string",
            description: "Template id, as returned by list_templates.",
          },
        },
        required: ["template_id"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "get_deck_state",
      description:
        "Read the deck currently open in PowerPoint: each slide's id, index, and shapes " +
        "({name, text}). Executes in the task pane via Office.js.",
      parameters: { type: "object", properties: {} },
    },
  },
  {
    type: "function",
    function: {
      name: "insert_template_slides",
      description:
        "Append all slides of a branded template to the deck open in PowerPoint. Returns the new " +
        "slides' ids and indexes — use those ids with apply_slide_content. Executes in the task " +
        "pane via Office.js.",
      parameters: {
        type: "object",
        properties: {
          template_id: {
            type: "string",
            description: "Template id, as returned by list_templates.",
          },
        },
        required: ["template_id"],
      },
    },
  },
  {
    type: "function",
    function: {
      name: "apply_slide_content",
      description:
        "Write text into named placeholder shapes on one slide of the open deck. Each fill sets " +
        "one shape's full text. For every fill containing a financial figure, quote the figure " +
        "exactly as returned by read_excel_range and include the source cell it was read from. " +
        "Returns per-fill success. Executes in the task pane via Office.js.",
      parameters: {
        type: "object",
        properties: {
          slide_id: {
            type: "string",
            description: "Slide id, as returned by insert_template_slides or get_deck_state.",
          },
          fills: {
            type: "array",
            description: "One entry per placeholder shape to fill on this slide.",
            items: {
              type: "object",
              properties: {
                shape_name: {
                  type: "string",
                  description: "Placeholder shape name, from load_template or get_deck_state.",
                },
                text: {
                  type: "string",
                  description: "Exact text to place in the shape.",
                },
                source: {
                  type: "object",
                  description:
                    "Provenance of the figure in `text` — the workbook cell it was read from. " +
                    "Required whenever the text contains a workbook figure.",
                  properties: {
                    file_id: { type: "string", description: "Uploaded workbook file_id." },
                    sheet: { type: "string", description: "Worksheet name." },
                    cell: { type: "string", description: 'Cell address in A1 notation, e.g. "C7".' },
                  },
                  required: ["file_id", "sheet", "cell"],
                },
              },
              required: ["shape_name", "text"],
            },
          },
        },
        required: ["slide_id", "fills"],
      },
    },
  },
];

/** Tools that must execute in the Office.js runtime — relayed to the task pane over SSE. */
const CLIENT_TOOLS = new Set(["get_deck_state", "insert_template_slides", "apply_slide_content"]);

interface ToolOutcome {
  content: string;
  isError: boolean;
}

function executeBackendTool(name: string, input: Record<string, unknown>): ToolOutcome {
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
      case "list_templates":
        return { content: JSON.stringify(listTemplates()), isError: false };
      case "load_template":
        return { content: JSON.stringify(loadTemplate(String(input.template_id))), isError: false };
      default:
        return { content: `unknown tool "${name}"`, isError: true };
    }
  } catch (err) {
    if (err instanceof ExcelToolError || err instanceof TemplateToolError) {
      return { content: `Error: ${err.message}`, isError: true };
    }
    throw err;
  }
}

// ---------------------------------------------------------------------------
// Client-tool relay: PowerPoint tools can only run in the Office.js runtime,
// so the loop emits a `client_tool` SSE event and waits here for the task pane
// to execute it and POST the outcome back to /api/chat/client-result.
// ---------------------------------------------------------------------------

const CLIENT_TOOL_TIMEOUT_MS = 120_000;

const pendingClientCalls = new Map<string, (outcome: ToolOutcome) => void>();

/** Settle a pending relay call exactly once, clearing its bookkeeping. */
function settleClientCall(callId: string, outcome: ToolOutcome) {
  const resolve = pendingClientCalls.get(callId);
  if (!resolve) return;
  pendingClientCalls.delete(callId);
  resolve(outcome);
}

function executeClientTool(
  name: string,
  input: Record<string, unknown>,
  send: (event: string, data: unknown) => void,
  requestCallIds: Set<string>,
): Promise<ToolOutcome> {
  // Payload the pane needs to execute the call. Server-side data joins it here:
  // the template .pptx travels in this event only — never into model context.
  const payload: Record<string, unknown> = { ...input };
  if (name === "insert_template_slides") {
    try {
      payload.base64 = getTemplateBase64(String(input.template_id));
    } catch (err) {
      if (err instanceof TemplateToolError) {
        return Promise.resolve({ content: `Error: ${err.message}`, isError: true });
      }
      throw err;
    }
  }

  return new Promise((resolve) => {
    const callId = randomUUID();
    const timer = setTimeout(() => {
      settleClientCall(callId, {
        content: `Error: the task pane did not return a result for ${name} within ${CLIENT_TOOL_TIMEOUT_MS / 1000}s`,
        isError: true,
      });
    }, CLIENT_TOOL_TIMEOUT_MS);
    pendingClientCalls.set(callId, (outcome) => {
      clearTimeout(timer);
      requestCallIds.delete(callId);
      resolve(outcome);
    });
    requestCallIds.add(callId);
    send("client_tool", { callId, name, payload });
  });
}

/** Parse a streamed tool-call's accumulated argument string into an input object. */
function parseToolArgs(raw: string): Record<string, unknown> {
  if (!raw || !raw.trim()) return {};
  return JSON.parse(raw) as Record<string, unknown>;
}

/**
 * POST /api/chat/client-result
 * Body: { callId: string, content: string, isError?: boolean }
 * The task pane posts each relayed tool call's outcome here.
 */
chatRouter.post("/client-result", (req, res) => {
  const { callId, content, isError } = (req.body ?? {}) as Record<string, unknown>;
  if (typeof callId !== "string" || !pendingClientCalls.has(callId)) {
    res.status(404).json({ error: "unknown or expired callId" });
    return;
  }
  settleClientCall(callId, {
    content: typeof content === "string" ? content : JSON.stringify(content ?? null),
    isError: Boolean(isError),
  });
  res.json({ ok: true });
});

/**
 * POST /api/chat
 * Body: { messages: { role: "user" | "assistant"; content: string }[] }
 * Responds with Server-Sent Events:
 *   event: text   data: {"text": "..."}                 (incremental text deltas)
 *   event: tool   data: {"name": "...", "input": {...}} (a tool call is executing)
 *   event: done   data: {"stopReason": "..."}           (stream finished)
 *   event: error  data: {"message": "..."}
 */
chatRouter.post("/", async (req, res) => {
  const incoming = req.body?.messages as { role: "user" | "assistant"; content: string }[] | undefined;
  if (!Array.isArray(incoming) || incoming.length === 0) {
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

  // Relay calls opened by this request; settled with an error if the pane disconnects.
  // clientGone stops the loop early instead of streaming rounds to a dead socket.
  const requestCallIds = new Set<string>();
  let clientGone = false;
  res.on("close", () => {
    clientGone = true;
    for (const callId of [...requestCallIds]) {
      settleClientCall(callId, { content: "Error: the task pane disconnected", isError: true });
    }
  });

  try {
    // Tool-use loop: stream each round to the client; when the model requests tools,
    // execute them backend-side, feed the results back, and continue until end of turn.
    // The system prompt leads; the rest is the conversation as the pane sees it.
    const loopMessages: OpenAI.Chat.Completions.ChatCompletionMessageParam[] = [
      { role: "system", content: SYSTEM_PROMPT },
      ...incoming.map((m) => ({ role: m.role, content: m.content })),
    ];

    for (let iteration = 0; ; iteration++) {
      if (clientGone) return;
      if (iteration >= MAX_LOOP_ITERATIONS) {
        send("error", { message: `tool-use loop exceeded ${MAX_LOOP_ITERATIONS} iterations` });
        return;
      }

      const params: ChatParams = {
        model: MODEL,
        max_tokens: MAX_TOKENS,
        messages: loopMessages,
        tools: TOOLS,
        tool_choice: "auto",
        stream: true,
        reasoning: { effort: REASONING_EFFORT },
      };
      const stream = await getClient().chat.completions.create(params);

      // Reassemble the assistant turn from streamed deltas: visible text plus any
      // tool calls (whose id/name/arguments arrive split across chunks by index).
      let text = "";
      const toolCallsByIndex: { id: string; name: string; args: string }[] = [];
      let finishReason: string | null = null;

      for await (const chunk of stream) {
        const choice = chunk.choices[0];
        if (!choice) continue;
        const delta = choice.delta;
        if (delta?.content) {
          text += delta.content;
          send("text", { text: delta.content });
        }
        for (const tc of delta?.tool_calls ?? []) {
          const slot = (toolCallsByIndex[tc.index] ??= { id: "", name: "", args: "" });
          if (tc.id) slot.id = tc.id;
          if (tc.function?.name) slot.name = tc.function.name;
          if (tc.function?.arguments) slot.args += tc.function.arguments;
        }
        if (choice.finish_reason) finishReason = choice.finish_reason;
      }

      const toolCalls = toolCallsByIndex.filter(Boolean);

      if (toolCalls.length > 0) {
        loopMessages.push({
          role: "assistant",
          content: text,
          tool_calls: toolCalls.map((t) => ({
            id: t.id,
            type: "function",
            function: { name: t.name, arguments: t.args },
          })),
        });

        for (const t of toolCalls) {
          let outcome: ToolOutcome;
          let input: Record<string, unknown>;
          try {
            input = parseToolArgs(t.args);
          } catch {
            // Model emitted malformed JSON arguments — report back so it can retry.
            loopMessages.push({
              role: "tool",
              tool_call_id: t.id,
              content: `Error: could not parse arguments for ${t.name} as JSON`,
            });
            continue;
          }
          if (CLIENT_TOOLS.has(t.name)) {
            // The client_tool event both renders the activity chip and carries the call.
            outcome = await executeClientTool(t.name, input, send, requestCallIds);
          } else {
            send("tool", { name: t.name, input });
            outcome = executeBackendTool(t.name, input);
          }
          loopMessages.push({
            role: "tool",
            tool_call_id: t.id,
            content: outcome.content,
          });
        }
        continue;
      }

      send("done", { stopReason: finishReason });
      return;
    }
  } catch (err) {
    const message = err instanceof OpenAI.APIError ? `${err.status}: ${err.message}` : String(err);
    console.error("chat error:", message);
    send("error", { message });
  } finally {
    res.end();
  }
});
