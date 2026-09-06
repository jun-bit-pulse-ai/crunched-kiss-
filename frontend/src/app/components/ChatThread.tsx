import { useEffect, useRef } from "react";
import type { VisibleMessage } from "../types";
import { Markdown } from "./Markdown";
import { displayToolName } from "../toolCards";
import { ClarifyingQuestion, parseOptions } from "./ClarifyingQuestion";

type ChatThreadProps = {
  messages: VisibleMessage[];
  status?: string | null;
  /** Called with the picked option's text (e.g. "A) Budget") when a clarifying-question button is clicked. */
  onOptionSelect?: (optionText: string) => void;
  /** Disable option buttons while a request is already in flight. */
  optionsDisabled?: boolean;
};

export function ChatThread({ messages, status, onOptionSelect, optionsDisabled }: ChatThreadProps) {
  const endRef = useRef<HTMLDivElement | null>(null);

  // Key on length + status, not array identity, so this does not fire every render.
  useEffect(() => {
    endRef.current?.scrollIntoView({ block: "end" });
  }, [messages.length, status]);

  return (
    <div className="thread" aria-live="polite">
      {messages.map((message) =>
        message.kind === "tool" ? (
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
        ) : message.role === "assistant" && onOptionSelect && parseOptions(message.text) ? (
          <article key={message.id} className="bubble bubble-assistant">
            <ClarifyingQuestion text={message.text} disabled={optionsDisabled} onSelect={onOptionSelect} />
          </article>
        ) : (
          <article key={message.id} className={`bubble bubble-${message.role}`}>
            <Markdown text={message.text} />
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
