import type { VisibleMessage } from "../types";

type ChatThreadProps = {
  messages: VisibleMessage[];
  status?: string | null;
};

export function ChatThread({ messages, status }: ChatThreadProps) {
  return (
    <div className="thread" aria-live="polite">
      {messages.map((message) => (
        <article key={message.id} className={`bubble bubble-${message.role}`}>
          <p>{message.text}</p>
        </article>
      ))}
      {status ? <p className="status">{status}</p> : null}
    </div>
  );
}
