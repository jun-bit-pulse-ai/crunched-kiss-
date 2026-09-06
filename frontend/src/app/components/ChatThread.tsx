import { useEffect, useRef } from "react";
import type { VisibleMessage } from "../types";
import { writePreviewSummary } from "../writeConfirm";
import { displayToolName } from "../toolCards";

type ChatThreadProps = {
  messages: VisibleMessage[];
  status?: string | null;
  onWriteDecision?: (id: string, apply: boolean) => void;
};

export function ChatThread({ messages, status, onWriteDecision }: ChatThreadProps) {
  const endRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    endRef.current?.scrollIntoView({ block: "end" });
  }, [messages.length, status]);

  return (
    <div className="thread" aria-live="polite">
      {messages.map((message) =>
        message.kind === "write_confirm" ? (
          <article key={message.id} className="write-confirm" aria-label={`Write ${writePreviewSummary(message)}`}>
            <p className="write-confirm-title">
              Crunched wants to write {writePreviewSummary(message)}
            </p>
            <pre className="write-confirm-values">{JSON.stringify(message.values, null, 2)}</pre>
            {message.status === "pending" && onWriteDecision ? (
              <div className="write-confirm-actions">
                <button type="button" className="option-button" onClick={() => onWriteDecision(message.id, true)}>
                  Apply
                </button>
                <button type="button" className="option-button" onClick={() => onWriteDecision(message.id, false)}>
                  Don&apos;t write
                </button>
              </div>
            ) : (
              <p className="write-confirm-status">
                {message.status === "applied" ? "Applied" : "Not written"}
              </p>
            )}
          </article>
        ) : message.kind === "tool" ? (
          <article
            key={message.id}
            className={`tool-card${message.error ? " tool-card-error" : ""}`}
            aria-label={
              message.error
                ? `Tool error: ${displayToolName(message.name)}`
                : `Tool: ${displayToolName(message.name)}`
            }
          >
            <span className="tool-card-name">{displayToolName(message.name)}</span>
            <span className="tool-card-summary">{message.summary}</span>
          </article>
        ) : (
          <article key={message.id} className={`bubble bubble-${message.role}`}>
            <p className="bubble-text">{message.text}</p>
          </article>
        )
      )}
      {status ? (
        <p className="status">
          <span className="status-spinner" aria-hidden="true" />
          {status}
        </p>
      ) : null}
      <div ref={endRef} />
    </div>
  );
}
