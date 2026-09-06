import { useEffect, useMemo, useRef, useState } from "react";
import { ChatThread } from "./components/ChatThread";
import { Composer } from "./components/Composer";
import { Eli5Panel } from "./components/Eli5Panel";
import { GuidedTour } from "./components/GuidedTour";
import { PromptChips } from "./components/PromptChips";
import { initialVisible, showPromptChips } from "./demoPrompts";
import { eli5InputFromVisible, explainLikeFive } from "./eli5";
import { runAgent } from "./services/agentClient";
import {
  canUndo,
  clearUndoStack,
  getSelectedFormula,
  listWorkbookMeta,
  undoLastWrite,
  watchSelection,
} from "./services/excel";
import { parseSuggestions } from "./suggestions";
import { clearConversation, loadConversation, saveConversation } from "./storage";
import { toolCardsFromMessages } from "./toolCards";
import { hasSeenTour, markTourSeen, TOUR_TARGETS } from "./tour";
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
  const [selection, setSelection] = useState<string | null>(null);
  const [tourOpen, setTourOpen] = useState(false);
  const [eli5Open, setEli5Open] = useState(false);
  const [focusToken, setFocusToken] = useState(0);
  const showTourRef = useRef<HTMLButtonElement | null>(null);
  // Bump this to force re-render of the undo button when the stack changes.
  const [undoToken, setUndoToken] = useState(0);

  useEffect(() => {
    const unsubscribe = watchSelection(setSelection);
    return unsubscribe;
  }, []);

  // Starts from the pane, not Office.onReady — a browser preview still gets the tour.
  useEffect(() => {
    if (!hasSeenTour()) {
      setTourOpen(true);
    }
  }, []);

  // Try to restore a persisted conversation for this workbook.
  useEffect(() => {
    let cancelled = false;
    async function restore() {
      try {
        const meta = await listWorkbookMeta();
        if (cancelled) return;
        const sheetNames = meta.sheets.map((s) => s.name);
        const restored = loadConversation(sheetNames);
        if (restored) {
          setAgentMessages(restored.agentMessages);
          setVisible(restored.visible);
        }
      } catch {
        // Workbook not available yet (e.g. browser preview) — ignore.
      }
    }
    restore();
    return () => {
      cancelled = true;
    };
  }, []);

  const eli5Text = useMemo(() => explainLikeFive(eli5InputFromVisible(visible)), [visible]);

  const conversationLabel = useMemo(() => {
    return agentMessages.some((message) => message.role === "user" && typeof message.content === "string")
      ? "Working conversation"
      : "New conversation";
  }, [agentMessages]);

  async function persist(sheetNames: string[], messages: ChatMessage[], vis: VisibleMessage[]) {
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
    clearUndoStack();
    setUndoToken((t) => t + 1);
    // Clear persisted conversation for this workbook.
    listWorkbookMeta()
      .then((meta) => clearConversation(meta.sheets.map((s) => s.name)))
      .catch(() => {
        /* ignore */
      });
  }

  async function send(text: string) {
    setError(null);
    setBusy(true);
    setVisible((current) => [...current, { id: newId(), kind: "text", role: "user", text }]);
    const nextHistory: ChatMessage[] = [...agentMessages, { role: "user", content: text }];
    let sheetNames: string[] = [];
    try {
      // Office.js lives in the pane. The backend never opens the xlsx —
      // it only gets these sheet names as a hint so Claude can pick a tool.
      let hint: WorkbookHint | undefined;
      try {
        const meta = await listWorkbookMeta();
        sheetNames = meta.sheets.map((sheet) => sheet.name);
        hint = sheetNames;
      } catch {
        hint = undefined;
      }
      const result = await runAgent(nextHistory, hint, setStatus);
      const cards = toolCardsFromMessages(result.messages, nextHistory.length);
      setAgentMessages(result.messages);

      // Parse follow-up suggestions from assistant text
      const parsed = parseSuggestions(result.text);
      const newVisible: VisibleMessage[] = [
        ...cards.map((card) => ({ kind: "tool" as const, ...card })),
        {
          id: newId(),
          kind: "text",
          role: "assistant",
          text: parsed?.text ?? result.text,
          suggestions: parsed?.suggestions,
        },
      ];
      setVisible((current) => [...current, ...newVisible]);

      // Persist the conversation.
      if (sheetNames.length > 0) {
        await persist(sheetNames, result.messages, [...visible, { id: newId(), kind: "text", role: "user", text }, ...newVisible]);
      }
    } catch (err) {
      const message = err instanceof Error ? err.message : String(err);
      setError(message);
      setAgentMessages(nextHistory);
    } finally {
      setBusy(false);
      setStatus(null);
      setUndoToken((t) => t + 1);
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
      setUndoToken((t) => t + 1);
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
      const prompt = `Explain this Excel formula in plain English (what it does, step by step): ${formulaInfo.formula}`;
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
              ref={showTourRef}
              type="button"
              className="text-action"
              aria-haspopup="dialog"
              onClick={() => {
                setEli5Open(false);
                setTourOpen(true);
              }}
            >
              Show tour
            </button>
            <button
              type="button"
              className="text-action"
              aria-expanded={eli5Open}
              aria-controls="eli5-panel"
              onClick={() => setEli5Open((open) => !open)}
            >
              Explain like I&apos;m 5
            </button>
            <button
              type="button"
              className="undo-button"
              disabled={busy || !canUndo()}
              onClick={handleUndo}
              title="Undo last AI write"
            >
              Undo
            </button>
            <button
              type="button"
              className="text-action"
              data-tour={TOUR_TARGETS.newChat}
              disabled={busy}
              onClick={resetChat}
            >
              New chat
            </button>
          </div>
        </div>
        <div className={`rule ${busy ? "rule-busy" : ""}`} />
      </header>
      {eli5Open ? (
        <div id="eli5-panel">
          <Eli5Panel text={eli5Text} onClose={() => setEli5Open(false)} />
        </div>
      ) : null}
      <ChatThread messages={visible} status={status} onOptionSelect={send} optionsDisabled={busy} />
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
      <GuidedTour
        open={tourOpen}
        onClose={(reason) => {
          if (reason === "skip" || reason === "done") {
            markTourSeen();
          }
          setTourOpen(false);
          showTourRef.current?.focus();
        }}
      />
    </div>
  );
}
