import { useMemo, useState } from "react";
import { ChatThread } from "./components/ChatThread";
import { Composer } from "./components/Composer";
import { PromptChips } from "./components/PromptChips";
import { initialVisible, showPromptChips } from "./demoPrompts";
import { runAgent } from "./services/agentClient";
import { listWorkbookMeta } from "./services/excel";
import { toolCardsFromMessages } from "./toolCards";
import type { ChatMessage, VisibleMessage, WorkbookHint } from "./types";

function newId(): string {
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

export default function App() {
  const [agentMessages, setAgentMessages] = useState<ChatMessage[]>([]);
  const [visible, setVisible] = useState<VisibleMessage[]>(initialVisible);
  const [status, setStatus] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  const conversationLabel = useMemo(() => {
    return agentMessages.some((message) => message.role === "user" && typeof message.content === "string")
      ? "Working conversation"
      : "No new conversation";
  }, [agentMessages]);

  function resetChat() {
    setAgentMessages([]);
    setVisible(initialVisible());
    setStatus(null);
    setError(null);
    setBusy(false);
  }

  async function send(text: string) {
    setError(null);
    setBusy(true);
    setVisible((current) => [...current, { id: newId(), kind: "text", role: "user", text }]);
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
      const cards = toolCardsFromMessages(result.messages, nextHistory.length);
      setAgentMessages(result.messages);
      setVisible((current) => [
        ...current,
        ...cards.map((card) => ({ kind: "tool" as const, ...card })),
        { id: newId(), kind: "text", role: "assistant", text: result.text },
      ]);
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
        <div className="masthead-row">
          <p className="conversation-label">{conversationLabel}</p>
          {showPromptChips(visible) ? null : (
            <button type="button" className="new-chat" disabled={busy} onClick={resetChat}>
              New chat
            </button>
          )}
        </div>
        <div className={`rule ${busy ? "rule-busy" : ""}`} />
      </header>
      <ChatThread messages={visible} status={status} />
      {error ? <div className="banner">{error}</div> : null}
      {showPromptChips(visible) ? <PromptChips disabled={busy} onPick={send} /> : null}
      <Composer disabled={busy} onSend={send} />
    </div>
  );
}
