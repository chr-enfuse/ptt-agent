import { Router } from "express";
import Anthropic from "@anthropic-ai/sdk";

export const chatRouter = Router();

const client = new Anthropic(); // reads ANTHROPIC_API_KEY from the environment

const MODEL = "claude-opus-4-8";

// Keep this byte-stable: it sits at the front of the prompt prefix so it can be
// served from the prompt cache. Volatile, per-request context goes in messages.
const SYSTEM_PROMPT = `You are a financial-deck assistant embedded in a PowerPoint task pane.
You help analysts build client-facing financial presentations from branded templates and
uploaded Excel workbooks.

Hard rule: you never type out financial figures yourself. Numbers are extracted from
workbook cells by deterministic tools and verified by code before they reach a slide.
Your job is mapping — choosing which cells feed which placeholders — and conversation.

(Tooling for Excel reading, template fill, and verification arrives in later milestones;
for now you converse and explain what you will be able to do.)`;

/**
 * POST /api/chat
 * Body: { messages: Anthropic.MessageParam[] }
 * Responds with Server-Sent Events:
 *   event: text   data: {"text": "..."}        (incremental text deltas)
 *   event: done   data: {"stopReason": "..."}  (stream finished)
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
    const stream = client.messages.stream({
      model: MODEL,
      max_tokens: 64000,
      thinking: { type: "adaptive" },
      output_config: { effort: "high" },
      system: [
        {
          type: "text",
          text: SYSTEM_PROMPT,
          cache_control: { type: "ephemeral" },
        },
      ],
      messages,
    });

    stream.on("text", (delta) => send("text", { text: delta }));

    const finalMessage = await stream.finalMessage();
    send("done", { stopReason: finalMessage.stop_reason });
  } catch (err) {
    const message = err instanceof Anthropic.APIError ? `${err.status}: ${err.message}` : String(err);
    console.error("chat error:", message);
    send("error", { message });
  } finally {
    res.end();
  }
});
