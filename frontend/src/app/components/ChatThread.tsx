import { useEffect, useRef } from "react";
import { speakerLabel, toolCardLabel } from "../a11y";
import type { VisibleMessage } from "../types";
import { Markdown } from "./Markdown";

type ChatThreadProps = {
  messages: VisibleMessage[];
  status?: string | null;
};

export function ChatThread({ messages, status }: ChatThreadProps) {
  const endRef = useRef<HTMLDivElement | null>(null);

  // Key on length + status, not array identity, so this does not fire every render.
  useEffect(() => {
    endRef.current?.scrollIntoView({ block: "end" });
  }, [messages.length, status]);

  return (
    <section className="thread" id="chat-thread" tabIndex={-1} aria-labelledby="chat-heading">
      <h2 id="chat-heading" className="sr-only">
        Conversation
      </h2>
      <div className="thread-log" role="log" aria-live="polite" aria-relevant="additions" aria-atomic="false">
        {messages.map((message) =>
          message.kind === "tool" ? (
            <article
              key={message.id}
              className={`tool-card${message.error ? " tool-card-error" : ""}`}
              aria-label={toolCardLabel(message.name, message.summary, message.error)}
            >
              <span className="tool-card-name">{message.error ? `Error · ${message.name}` : message.name}</span>
              <span className="tool-card-summary">{message.summary}</span>
            </article>
          ) : (
            <article key={message.id} className={`bubble bubble-${message.role}`}>
              <h3 className="sr-only">{speakerLabel(message.role)}</h3>
              <Markdown text={message.text} />
            </article>
          )
        )}
      </div>
      {status ? (
        <p className="status" role="status">
          <span className="status-spinner" aria-hidden="true" />
          {status}
        </p>
      ) : null}
      <div ref={endRef} />
    </section>
  );
}
