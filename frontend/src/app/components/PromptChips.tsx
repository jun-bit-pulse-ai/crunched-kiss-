import { DEMO_CHIPS } from "../demoPrompts";

type PromptChipsProps = {
  disabled?: boolean;
  onPick: (text: string) => void;
};

export function PromptChips({ disabled, onPick }: PromptChipsProps) {
  return (
    <div className="chips">
      {DEMO_CHIPS.map((chip) => (
        <button key={chip} type="button" className="chip" disabled={disabled} onClick={() => onPick(chip)}>
          {chip}
        </button>
      ))}
    </div>
  );
}
