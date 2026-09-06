import { useEffect, useRef, useState } from "react";

type ComposerProps = {
  disabled?: boolean;
  onSend: (text: string) => void;
  /** Bump this (e.g. on mount and after "New chat") to move keyboard focus into the textarea. */
  focusToken?: number;
};

export function Composer({ disabled, onSend, focusToken }: ComposerProps) {
  const [value, setValue] = useState("");
  const textareaRef = useRef<HTMLTextAreaElement | null>(null);

  useEffect(() => {
    textareaRef.current?.focus();
  }, [focusToken]);

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
      data-tour="composer"
      onSubmit={(event) => {
        event.preventDefault();
        submit();
      }}
    >
      <textarea
        ref={textareaRef}
        value={value}
        disabled={disabled}
        rows={2}
        placeholder="Ask Crunched about this workbook…"
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
