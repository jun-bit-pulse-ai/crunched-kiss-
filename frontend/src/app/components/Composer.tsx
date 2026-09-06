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
      onSubmit={(event) => {
        event.preventDefault();
        submit();
      }}
    >
      <textarea
        value={value}
        disabled={disabled}
        rows={2}
        placeholder="Ask the model…"
        aria-label="Message"
        onChange={(event) => setValue(event.target.value)}
        onKeyDown={(event) => {
          if (event.key === "Enter" && !event.shiftKey) {
            event.preventDefault();
            submit();
          }
        }}
      />
      <button type="submit" disabled={disabled || value.trim().length === 0}>
        Send
      </button>
    </form>
  );
}
