import { useMemo, useState } from "react";
import { ChatThread } from "./components/ChatThread";
import { Composer } from "./components/Composer";
import { PromptChips } from "./components/PromptChips";
import { runAgent } from "./services/agentClient";
import { listWorkbookMeta } from "./services/excel";
import type { ChatMessage, VisibleMessage, WorkbookHint } from "./types";

const WELCOME =
  "Hi, I'm Crunched — your AI analyst in Excel. I can help you with things like:\n\nError checking and fixing models\nBuilding financial or business models\nAnalyzing and comparing scenario models\n\nWhat should we work on first?";

function newId(): string {
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

export default function App() {
  const [agentMessages, setAgentMessages] = useState<ChatMessage[]>([]);
  const [visible, setVisible] = useState<VisibleMessage[]>([
    { id: "welcome", role: "assistant", text: WELCOME },
  ]);
  const [status, setStatus] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const conversationLabel = useMemo(() => {
    return agentMessages.some((message) => message.role === "user" && typeof message.content === "string")
      ? "Working conversation"
      : "No new conversation";
  }, [agentMessages]);

  async function send(text: string) {
    setError(null);
    setBusy(true);
    setVisible((current) => [...current, { id: newId(), role: "user", text }]);
    const nextHistory: ChatMessage[] = [...agentMessages, { role: "user", content: text }];
    try {
      let hint: WorkbookHint | undefined;
      try {
        const meta = await listWorkbookMeta();
        hint = meta.sheets.map((sheet) => sheet.name);
      } catch {
        hint = undefined;
      }
      const result = await runAgent(nextHistory, hint, setStatus);
      setAgentMessages(result.messages);
      setVisible((current) => [...current, { id: newId(), role: "assistant", text: result.text }]);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setError(message);
      setAgentMessages(nextHistory);
    } finally {
      setBusy(false);
      setStatus(null);
    }
  }

  return (
    <div className="shell">
      <header className="masthead">
        <div>
          <p className="eyebrow">Excel analyst</p>
          <h1>Crunched</h1>
        </div>
        <p className="conversation-label">{conversationLabel}</p>
        <div className={`rule ${busy ? "rule-busy" : ""}`} />
      </header>
      <ChatThread messages={visible} status={status} />
      {error ? <div className="banner">{error}</div> : null}
      <PromptChips disabled={busy} onPick={send} />
      <Composer disabled={busy} onSend={send} />
    </div>
  );
}
