import * as React from "react";
import { useState, useRef, useEffect } from "react";
import { makeStyles, tokens, Button, Textarea, Text, Spinner } from "@fluentui/react-components";
import { Send24Regular, ArrowUpload24Regular } from "@fluentui/react-icons";

/* global fetch, TextDecoder, FormData, File, HTMLDivElement, HTMLInputElement */

// The backend holds the Anthropic API key and runs the agent tool-use loop.
const BACKEND_URL = "http://localhost:3001";

export interface ChatMessage {
  role: "user" | "assistant";
  content: string;
  /**
   * chat   — a regular conversation turn (sent to the backend)
   * upload — a workbook-upload announcement (sent to the backend so the model sees the file_id)
   * tool   — display-only tool-activity entry (never sent to the backend)
   */
  kind: "chat" | "upload" | "tool";
}

interface UploadedSheet {
  name: string;
  rowCount: number;
  columnCount: number;
}

const useStyles = makeStyles({
  root: {
    display: "flex",
    flexDirection: "column",
    flexGrow: 1,
    minHeight: "0",
    padding: "8px",
    gap: "8px",
  },
  history: {
    flexGrow: 1,
    overflowY: "auto",
    display: "flex",
    flexDirection: "column",
    gap: "8px",
    paddingRight: "4px",
  },
  bubbleUser: {
    alignSelf: "flex-end",
    maxWidth: "85%",
    backgroundColor: tokens.colorBrandBackground2,
    borderRadius: "8px",
    padding: "8px 10px",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
  bubbleAssistant: {
    alignSelf: "flex-start",
    maxWidth: "85%",
    backgroundColor: tokens.colorNeutralBackground3,
    borderRadius: "8px",
    padding: "8px 10px",
    whiteSpace: "pre-wrap",
    wordBreak: "break-word",
  },
  chipUpload: {
    alignSelf: "center",
    maxWidth: "95%",
    backgroundColor: tokens.colorPaletteGreenBackground1,
    color: tokens.colorPaletteGreenForeground1,
    borderRadius: "8px",
    padding: "4px 10px",
    fontSize: tokens.fontSizeBase200,
    wordBreak: "break-word",
  },
  chipTool: {
    alignSelf: "flex-start",
    maxWidth: "95%",
    color: tokens.colorNeutralForeground3,
    fontSize: tokens.fontSizeBase200,
    fontStyle: "italic",
    paddingLeft: "4px",
    wordBreak: "break-word",
  },
  composer: {
    display: "flex",
    gap: "4px",
    alignItems: "flex-end",
  },
  input: {
    flexGrow: 1,
  },
  error: {
    color: tokens.colorPaletteRedForeground1,
  },
});

/** POST the conversation to the backend and consume the SSE response. */
async function streamChat(
  messages: { role: "user" | "assistant"; content: string }[],
  onText: (delta: string) => void,
  onTool: (name: string, input: Record<string, unknown>) => void,
): Promise<void> {
  const res = await fetch(`${BACKEND_URL}/api/chat`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ messages }),
  });
  if (!res.ok || !res.body) {
    throw new Error(`backend responded ${res.status} — is it running on ${BACKEND_URL}?`);
  }

  const reader = res.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";

  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });

    let sep: number;
    while ((sep = buffer.indexOf("\n\n")) !== -1) {
      const frame = buffer.slice(0, sep);
      buffer = buffer.slice(sep + 2);

      let event = "message";
      let data = "";
      for (const line of frame.split("\n")) {
        if (line.startsWith("event: ")) event = line.slice(7).trim();
        else if (line.startsWith("data: ")) data += line.slice(6);
      }
      if (!data) continue;
      if (event === "text") onText(JSON.parse(data).text);
      else if (event === "tool") {
        const tool = JSON.parse(data);
        onTool(tool.name, tool.input ?? {});
      } else if (event === "error") throw new Error(JSON.parse(data).message);
      // "done" carries the stop reason; nothing to do for now.
    }
  }
}

/** Compact human-readable label for a tool-activity chip. */
function describeToolCall(name: string, input: Record<string, unknown>): string {
  if (name === "read_excel_range") return `Reading ${String(input.sheet)}!${String(input.range)}…`;
  if (name === "list_workbook_structure") return "Inspecting workbook structure…";
  return `Running ${name}…`;
}

const Chat: React.FC = () => {
  const styles = useStyles();
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [draft, setDraft] = useState("");
  const [busy, setBusy] = useState(false);
  const [uploading, setUploading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const historyRef = useRef<HTMLDivElement>(null);
  const fileInputRef = useRef<HTMLInputElement>(null);

  useEffect(() => {
    historyRef.current?.scrollTo({ top: historyRef.current.scrollHeight });
  }, [messages]);

  /** The conversation as the backend sees it — tool-activity chips are display-only. */
  const toApiMessages = (history: ChatMessage[]) =>
    history
      .filter((m) => m.kind !== "tool" && m.content.length > 0)
      .map((m) => ({ role: m.role, content: m.content }));

  const uploadWorkbook = async (file: File) => {
    setUploading(true);
    setError(null);
    try {
      const form = new FormData();
      form.append("file", file);
      const res = await fetch(`${BACKEND_URL}/api/workbooks`, { method: "POST", body: form });
      const body = await res.json();
      if (!res.ok) {
        throw new Error(body?.error ?? `upload failed (${res.status})`);
      }
      const sheets = (body.sheets as UploadedSheet[])
        .map((s) => `${s.name} (${s.rowCount}×${s.columnCount})`)
        .join(", ");
      // Announced as a user turn so the model learns the file_id from the conversation.
      setMessages((prev) => [
        ...prev,
        {
          role: "user",
          kind: "upload",
          content: `Uploaded workbook "${body.filename}" (file_id: ${body.fileId}). Sheets: ${sheets}.`,
        },
      ]);
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setUploading(false);
    }
  };

  const send = async () => {
    const content = draft.trim();
    if (!content || busy) return;

    const history: ChatMessage[] = [...messages, { role: "user", content, kind: "chat" }];
    setMessages([...history, { role: "assistant", content: "", kind: "chat" }]);
    setDraft("");
    setBusy(true);
    setError(null);

    try {
      await streamChat(
        toApiMessages(history),
        (delta) => {
          setMessages((prev) => {
            const next = [...prev];
            const last = next[next.length - 1];
            if (last.role === "assistant" && last.kind === "chat") {
              next[next.length - 1] = { ...last, content: last.content + delta };
            } else {
              next.push({ role: "assistant", content: delta, kind: "chat" });
            }
            return next;
          });
        },
        (name, input) => {
          setMessages((prev) => {
            const next = [...prev];
            // Drop a trailing empty assistant bubble so the chip doesn't sit under a placeholder.
            const last = next[next.length - 1];
            if (last.role === "assistant" && last.kind === "chat" && last.content === "") {
              next.pop();
            }
            next.push({ role: "assistant", kind: "tool", content: describeToolCall(name, input) });
            return next;
          });
        },
      );
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
      // Remove a dangling empty assistant placeholder (e.g. on error before any text).
      setMessages((prev) => {
        const last = prev[prev.length - 1];
        if (last && last.role === "assistant" && last.kind === "chat" && last.content === "") {
          return prev.slice(0, -1);
        }
        return prev;
      });
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.history} ref={historyRef}>
        {messages.length === 0 && (
          <Text size={200}>
            Upload an Excel workbook and ask the agent about it — every figure it reports is read from the workbook by
            deterministic tools. (Template fill and verified slide placement arrive in later milestones.)
          </Text>
        )}
        {messages.map((m, i) => {
          if (m.kind === "upload") {
            return (
              <div key={i} className={styles.chipUpload}>
                {m.content}
              </div>
            );
          }
          if (m.kind === "tool") {
            return (
              <div key={i} className={styles.chipTool}>
                {m.content}
              </div>
            );
          }
          return (
            <div key={i} className={m.role === "user" ? styles.bubbleUser : styles.bubbleAssistant}>
              {m.content || (busy && i === messages.length - 1 ? "…" : "")}
            </div>
          );
        })}
        {error && (
          <Text size={200} className={styles.error}>
            {error}
          </Text>
        )}
      </div>
      <div className={styles.composer}>
        <input
          ref={fileInputRef}
          type="file"
          accept=".xlsx"
          style={{ display: "none" }}
          onChange={(e) => {
            const file = e.target.files?.[0];
            e.target.value = "";
            if (file) void uploadWorkbook(file);
          }}
        />
        <Button
          appearance="secondary"
          icon={uploading ? <Spinner size="tiny" /> : <ArrowUpload24Regular />}
          disabled={uploading}
          title="Upload Excel workbook (.xlsx)"
          onClick={() => fileInputRef.current?.click()}
        />
        <Textarea
          className={styles.input}
          value={draft}
          placeholder="Message the deck agent…"
          resize="vertical"
          onChange={(_e, data) => setDraft(data.value)}
          onKeyDown={(e) => {
            if (e.key === "Enter" && !e.shiftKey) {
              e.preventDefault();
              void send();
            }
          }}
        />
        <Button appearance="primary" icon={<Send24Regular />} disabled={busy || !draft.trim()} onClick={send} />
      </div>
    </div>
  );
};

export default Chat;
