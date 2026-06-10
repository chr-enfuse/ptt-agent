import "dotenv/config";
import express from "express";
import cors from "cors";
import { chatRouter } from "./routes/chat.js";
import { workbooksRouter } from "./routes/workbooks.js";

const app = express();
const PORT = Number(process.env.PORT ?? 3001);

// The task pane is served from the add-in dev server (https://localhost:3000).
app.use(
  cors({
    origin: process.env.ALLOWED_ORIGINS?.split(",") ?? ["https://localhost:3000"],
  }),
);
app.use(express.json({ limit: "10mb" }));

app.get("/healthz", (_req, res) => {
  res.json({ ok: true, service: "deck-agent-backend" });
});

app.use("/api/chat", chatRouter);
app.use("/api/workbooks", workbooksRouter);

app.listen(PORT, () => {
  console.log(`deck-agent-backend listening on http://localhost:${PORT}`);
  if (!process.env.OPENROUTER_API_KEY) {
    console.warn("WARNING: OPENROUTER_API_KEY is not set — /api/chat will fail. Copy .env.example to .env and set it.");
  }
});
