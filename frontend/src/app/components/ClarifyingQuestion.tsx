export type Option = {
  letter: string;
  text: string;
};

const OPTION_REGEX = /^([A-D])\)\s+(.+)$/gm;

/**
 * Splits a "question, then A)/B)/C) options" reply into its parts. Returns null when the
 * text doesn't have at least two lettered options, so callers can fall back to plain text.
 */
export function parseOptions(text: string): { question: string; options: Option[] } | null {
  const options: Option[] = [];
  let match: RegExpExecArray | null;
  while ((match = OPTION_REGEX.exec(text)) !== null) {
    options.push({ letter: match[1], text: match[2].trim() });
  }
  if (options.length < 2) {
    return null;
  }

  const firstOptionIndex = text.search(OPTION_REGEX);
  const question = text.slice(0, firstOptionIndex).trim();
  return { question, options };
}

type ClarifyingQuestionProps = {
  text: string;
  disabled?: boolean;
  onSelect: (optionText: string) => void;
};

export function ClarifyingQuestion({ text, disabled, onSelect }: ClarifyingQuestionProps) {
  const parsed = parseOptions(text);
  if (!parsed) {
    return null;
  }

  return (
    <div className="clarifying-question">
      {parsed.question ? <p className="clarifying-question-text">{parsed.question}</p> : null}
      <div className="clarifying-question-options">
        {parsed.options.map((option) => (
          <button
            key={option.letter}
            type="button"
            className="option-button"
            disabled={disabled}
            onClick={() => onSelect(`${option.letter}) ${option.text}`)}
          >
            <span className="option-letter">{option.letter}</span>
            <span className="option-text">{option.text}</span>
          </button>
        ))}
      </div>
    </div>
  );
}
