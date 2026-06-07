import * as React from "react";
import { useState, useRef, useEffect } from "react";
import { makeStyles, tokens, Button, Textarea, Text } from "@fluentui/react-components";
import { Send24Regular } from "@fluentui/react-icons";

/* global fetch, TextDecoder, HTMLDivElement */

// The backend holds the Anthropic API key and runs the agent loop.
const BACKEND_URL = "http://localhost:3001";

export interface ChatMessage {
  role: "user" | "assistant";
  content: string;
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
async function streamChat(messages: ChatMessage[], onText: (delta: string) => void): Promise<void> {
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
      else if (event === "error") throw new Error(JSON.parse(data).message);
      // "done" carries the stop reason; nothing to do for now.
    }
  }
}

const Chat: React.FC = () => {
  const styles = useStyles();
  const [messages, setMessages] = useState<ChatMessage[]>([]);
  const [draft, setDraft] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const historyRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    historyRef.current?.scrollTo({ top: historyRef.current.scrollHeight });
  }, [messages]);

  const send = async () => {
    const content = draft.trim();
    if (!content || busy) return;

    const history: ChatMessage[] = [...messages, { role: "user", content }];
    setMessages([...history, { role: "assistant", content: "" }]);
    setDraft("");
    setBusy(true);
    setError(null);

    try {
      await streamChat(history, (delta) => {
        setMessages((prev) => {
          const next = [...prev];
          const last = next[next.length - 1];
          next[next.length - 1] = { ...last, content: last.content + delta };
          return next;
        });
      });
    } catch (e) {
      setError(e instanceof Error ? e.message : String(e));
    } finally {
      setBusy(false);
    }
  };

  return (
    <div className={styles.root}>
      <div className={styles.history} ref={historyRef}>
        {messages.length === 0 && (
          <Text size={200}>
            Ask the agent about building a financial deck. (Excel upload, templates, and verified number placement
            arrive in later milestones.)
          </Text>
        )}
        {messages.map((m, i) => (
          <div key={i} className={m.role === "user" ? styles.bubbleUser : styles.bubbleAssistant}>
            {m.content || (busy && i === messages.length - 1 ? "…" : "")}
          </div>
        ))}
        {error && (
          <Text size={200} className={styles.error}>
            {error}
          </Text>
        )}
      </div>
      <div className={styles.composer}>
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
