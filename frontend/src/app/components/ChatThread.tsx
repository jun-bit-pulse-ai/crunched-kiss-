import { useEffect, useRef } from "react";
import type { VisibleMessage } from "../types";

type ChatThreadProps = {
  messages: VisibleMessage[];
  status?: string | null;
};

export function ChatThread({ messages, status }: ChatThreadProps) {
  const endRef = useRef<HTMLDivElement | null>(null);

  useEffect(() => {
    endRef.current?.scrollIntoView({ block: "end" });
  }, [messages, status]);

  return (
    <div className="thread" aria-live="polite">
      {messages.map((message) =>
        message.kind === "tool" ? (
          <article
            key={message.id}
            className={`tool-card${message.error ? " tool-card-error" : ""}`}
            aria-label={message.error ? `Tool error ${message.name}` : `Tool ${message.name}`}
          >
            <span className="tool-card-name">{message.name}</span>
            <span className="tool-card-summary">{message.summary}</span>
          </article>
        ) : (
          <article key={message.id} className={`bubble bubble-${message.role}`}>
            <p>{message.text}</p>
          </article>
        )
      )}
      {status ? <p className="status">{status}</p> : null}
      <div ref={endRef} />
    </div>
  );
}
