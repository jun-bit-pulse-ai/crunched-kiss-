import { useState } from "react";

type ComposerProps = {
  disabled?: boolean;
  onSend: (text: string) => void;
};

export function Composer({ disabled, onSend }: ComposerProps) {
  const [value, setValue] = useState("");

  function submit() {
    const text = value.trim();
    if (!text || disabled) {
      return;
    }
    onSend(text);
    setValue("");
  }

  return (
    <form
      className="composer"
      aria-label="Compose a message"
      onSubmit={(event) => {
        event.preventDefault();
        submit();
      }}
    >
      <div className="composer-field">
        <label className="sr-only" htmlFor="composer-input">
          Message
        </label>
        <textarea
          id="composer-input"
          value={value}
          disabled={disabled}
          rows={2}
          placeholder="Ask the model…"
          aria-describedby="composer-hint"
          onChange={(event) => setValue(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === "Enter" && !event.shiftKey) {
              event.preventDefault();
              submit();
            }
          }}
        />
        <p id="composer-hint" className="sr-only">
          Enter to send. Shift+Enter for a new line.
        </p>
      </div>
      <button type="submit" disabled={disabled || value.trim().length === 0} aria-label="Send message">
        Send
      </button>
    </form>
  );
}
