import { useEffect, useMemo, useRef, useState, useSyncExternalStore } from "react";
import { ChatThread } from "./components/ChatThread";
import { Composer } from "./components/Composer";
import { PromptChips } from "./components/PromptChips";
import { initialVisible, showPromptChips } from "./demoPrompts";
import { buildExplainFormulaPrompt } from "./formulaExplainer";
import { runAgent } from "./services/agentClient";
import {
  canUndo,
  clearUndoStack,
  getSelectedFormula,
  listWorkbookMeta,
  subscribeUndoStack,
  undoLastWrite,
  watchSelection,
} from "./services/excel";
import { clearConversation, loadConversation, saveConversation } from "./storage";
import { toolCardsFromMessages } from "./toolCards";
import type { ChatMessage, VisibleMessage, WorkbookHint } from "./types";
import type { WritePreview } from "./writeConfirm";

function newId(): string {
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

export default function App() {
  const [agentMessages, setAgentMessages] = useState<ChatMessage[]>([]);
  const [visible, setVisible] = useState<VisibleMessage[]>(initialVisible);
  const [status, setStatus] = useState<string | null>(null);
  const [error, setError] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [selection, setSelection] = useState<string | null>(null);
  const [focusToken, setFocusToken] = useState(0);
  const writeResolvers = useRef(new Map<string, (apply: boolean) => void>());
  const canUndoNow = useSyncExternalStore(subscribeUndoStack, canUndo);

  useEffect(() => {
    const unsubscribe = watchSelection(setSelection);
    return unsubscribe;
  }, []);

  useEffect(() => {
    let cancelled = false;
    async function restore() {
      try {
        const meta = await listWorkbookMeta();
        if (cancelled) return;
        const restored = loadConversation(meta.sheets.map((sheet) => sheet.name));
        if (restored) {
          setAgentMessages(restored.agentMessages);
          setVisible(restored.visible);
        }
      } catch {
        // Workbook not available yet (e.g. browser preview).
      }
    }
    restore();
    return () => {
      cancelled = true;
    };
  }, []);

  const conversationLabel = useMemo(() => {
    return agentMessages.some((message) => message.role === "user" && typeof message.content === "string")
      ? "Working conversation"
      : "New conversation";
  }, [agentMessages]);

  function persist(sheetNames: string[], messages: ChatMessage[], vis: VisibleMessage[]) {
    try {
      saveConversation(sheetNames, messages, vis);
    } catch {
      // Storage errors are non-fatal.
    }
  }

  function resetChat() {
    setAgentMessages([]);
    setVisible(initialVisible());
    setStatus(null);
    setError(null);
    setBusy(false);
    setFocusToken((token) => token + 1);
    writeResolvers.current.forEach((resolve) => resolve(false));
    writeResolvers.current.clear();
    clearUndoStack();
    listWorkbookMeta()
      .then((meta) => clearConversation(meta.sheets.map((sheet) => sheet.name)))
      .catch(() => {
        /* ignore */
      });
  }

  function confirmWrite(preview: WritePreview): Promise<boolean> {
    const id = newId();
    return new Promise((resolve) => {
      writeResolvers.current.set(id, resolve);
      setVisible((current) => [
        ...current,
        {
          id,
          kind: "write_confirm",
          sheet: preview.sheet,
          address: preview.address,
          values: preview.values,
          status: "pending",
        },
      ]);
    });
  }

  function decideWrite(id: string, apply: boolean) {
    const resolve = writeResolvers.current.get(id);
    writeResolvers.current.delete(id);
    setVisible((current) =>
      current.map((message) =>
        message.kind === "write_confirm" && message.id === id
          ? { ...message, status: apply ? "applied" : "declined" }
          : message
      )
    );
    resolve?.(apply);
  }

  async function send(text: string) {
    setError(null);
    setBusy(true);
    let latestVisible: VisibleMessage[] = [];
    setVisible((current) => {
      latestVisible = [...current, { id: newId(), kind: "text", role: "user", text }];
      return latestVisible;
    });
    const nextHistory: ChatMessage[] = [...agentMessages, { role: "user", content: text }];
    let sheetNames: string[] = [];
    try {
      let hint: WorkbookHint | undefined;
      try {
        const meta = await listWorkbookMeta();
        sheetNames = meta.sheets.map((sheet) => sheet.name);
        hint = sheetNames;
      } catch {
        hint = undefined;
      }
      const result = await runAgent(nextHistory, hint, {
        onStatus: setStatus,
        confirmWrite,
      });
      const cards = toolCardsFromMessages(result.messages, nextHistory.length);
      setAgentMessages(result.messages);
      setVisible((current) => {
        latestVisible = [
          ...current,
          ...cards.map((card) => ({ kind: "tool" as const, ...card })),
          { id: newId(), kind: "text", role: "assistant", text: result.text },
        ];
        return latestVisible;
      });
      if (sheetNames.length > 0) {
        persist(sheetNames, result.messages, latestVisible);
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setError(message);
      setAgentMessages(nextHistory);
    } finally {
      setBusy(false);
      setStatus(null);
    }
  }

  async function handleUndo() {
    if (!canUndo()) return;
    setBusy(true);
    try {
      const restored = await undoLastWrite();
      if (restored) {
        setVisible((current) => [
          ...current,
          {
            id: newId(),
            kind: "text",
            role: "system",
            text: `Undo: restored ${restored.address} on ${restored.sheet}`,
          },
        ]);
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setError(message);
    } finally {
      setBusy(false);
    }
  }

  async function handleExplainFormula() {
    setError(null);
    setBusy(true);
    try {
      const formulaInfo = await getSelectedFormula();
      if (!formulaInfo) {
        setError("Select a cell containing a formula first.");
        setBusy(false);
        return;
      }
      const prompt = buildExplainFormulaPrompt(formulaInfo.formula);
      await send(prompt);
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setError(message);
      setBusy(false);
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
          <div className="masthead-actions">
            <button
              type="button"
              className="undo-button"
              disabled={busy || !canUndoNow}
              onClick={handleUndo}
              title="Undo last AI write"
            >
              Undo
            </button>
            <button type="button" className="text-action" disabled={busy} onClick={resetChat}>
              New chat
            </button>
          </div>
        </div>
        <div className={`rule ${busy ? "rule-busy" : ""}`} />
      </header>
      <ChatThread messages={visible} status={status} onWriteDecision={decideWrite} />
      {error ? <div className="banner">{error}</div> : null}
      {selection ? (
        <p className="selection-pill" title="Say “this selection” and Crunched will read this range">
          Crunched can see {selection}
        </p>
      ) : null}
      {showPromptChips(visible) ? <PromptChips disabled={busy} onPick={send} /> : null}
      <Composer disabled={busy} onSend={send} focusToken={focusToken} />
      {selection ? (
        <div className="explain-formula-bar">
          <button
            type="button"
            className="explain-formula-button"
            disabled={busy}
            onClick={handleExplainFormula}
            title="Explain the formula in the selected cell"
          >
            Explain formula in {selection}
          </button>
        </div>
      ) : null}
    </div>
  );
}
