import { useEffect, useRef } from "react";
import type { VisibleMessage } from "../types";

type ChatThreadProps = {
  messages: VisibleMessage[];
  status?: string | null;
};

export function ChatThread({ messages, status }: ChatThreadProps) {
  const endRef = useRef<HTMLDivElement>(null);

  // Pin the view to the newest message. Without this the thread stays where it
  // was, so every reply after the first lands below the fold and the pane looks
  // frozen even though the round trip finished.
  useEffect(() => {
    endRef.current?.scrollIntoView({ block: "end" });
  }, [messages.length, status]);

  return (
    <div className="thread" aria-live="polite">
      {messages.map((message) => (
        <article key={message.id} className={`bubble bubble-${message.role}`}>
          <p>{message.text}</p>
        </article>
      ))}
      {status ? <p className="status">{status}</p> : null}
      <div ref={endRef} />
    </div>
  );
}
