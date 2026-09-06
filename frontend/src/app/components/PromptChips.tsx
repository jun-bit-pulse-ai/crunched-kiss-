type PromptChipsProps = {
  disabled?: boolean;
  onPick: (text: string) => void;
};

const CHIPS = [
  "Error check Budget model",
  "Check assumptions in Revenue model",
  "Create comparison dashboard",
];

export function PromptChips({ disabled, onPick }: PromptChipsProps) {
  return (
    <div className="chips">
      {CHIPS.map((chip) => (
        <button key={chip} type="button" className="chip" disabled={disabled} onClick={() => onPick(chip)}>
          {chip}
        </button>
      ))}
    </div>
  );
}
