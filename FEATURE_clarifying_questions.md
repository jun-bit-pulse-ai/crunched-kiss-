# Feature Design: Clarifying Questions with Multiple Choice

> **Implementation note:** shipped as designed, with three deviations from this doc:
> the system-prompt addition lives directly in `backend/app/agent.py` (there is no
> separate `prompts.py` in this project), the option-button styling uses this app's
> existing dark "ledger" theme tokens instead of the light-orange palette sketched
> below (a light card would clash with the rest of the pane), and the test file uses
> this repo's mocha/`assert` convention rather than vitest.

## Problem
When Claude is unsure about the user's intent, it currently writes a text question and the user must type a free-form answer. This is slower and more error-prone than clicking a button.

## Solution
When Claude detects ambiguity, it generates a clarifying question with 2-4 multiple-choice options. The frontend parses these options and renders them as clickable buttons. The user's click sends the selected option back as a message.

---

## How It Works

### 1. Claude Detects Uncertainty

Added to `SYSTEM_PROMPT` in `backend/app/prompts.py`:

```python
CLARIFYING_QUESTION_PROMPT = """
If the user's request is genuinely ambiguous or could match multiple things in the workbook,
ask ONE clarifying question with 2-4 multiple-choice options.

Format exactly as:
> I'd like to help, but I need a bit more clarity:
> 
> **Which sheet would you like to work with?**
> 
> A) Budget — the 7-row financial model
> B) Data — the 5,000-row metrics table
> C) A new sheet

The user can click a button to respond. Keep options short (under 60 chars each).
Never ask a clarifying question if the request is clear.
"""
```

### 2. Frontend Parses Options

New file: `frontend/src/app/components/ClarifyingQuestion.tsx`

```tsx
import * as React from "react";

interface Option {
  letter: string;
  text: string;
}

interface ClarifyingQuestionProps {
  text: string;
  onSelect: (optionText: string) => void;
}

const OPTION_REGEX = /^([A-D])\)\s+(.+)$/gm;

export function parseOptions(text: string): { question: string; options: Option[] } | null {
  const options: Option[] = [];
  let match;
  while ((match = OPTION_REGEX.exec(text)) !== null) {
    options.push({ letter: match[1], text: match[2].trim() });
  }
  if (options.length < 2) return null;
  
  // Extract question text (everything before the first option)
  const firstOptionIndex = text.search(OPTION_REGEX);
  const question = text.slice(0, firstOptionIndex).trim();
  return { question, options };
}

export function ClarifyingQuestion({ text, onSelect }: ClarifyingQuestionProps) {
  const parsed = parseOptions(text);
  if (!parsed) return <div className="bubble assistant">{text}</div>;
  
  return (
    <div className="clarifying-question">
      <div className="question-text">{parsed.question}</div>
      <div className="options">
        {parsed.options.map((opt) => (
          <button
            key={opt.letter}
            className="option-button"
            onClick={() => onSelect(`${opt.letter}) ${opt.text}`)}
          >
            <span className="option-letter">{opt.letter}</span>
            <span className="option-text">{opt.text}</span>
          </button>
        ))}
      </div>
    </div>
  );
}
```

### 3. Frontend Renders Instead of Plain Text

In `frontend/src/app/components/ChatThread.tsx`, replace the text bubble rendering:

```tsx
// Before:
<div className="bubble assistant">{block.text}</div>

// After:
<ClarifyingQuestion 
  text={block.text} 
  onSelect={(option) => onSend(option)} 
/>
```

### 4. Styles

Append to `frontend/src/app/styles.css`:

```css
.clarifying-question {
  background: #fff7ed;
  border: 1px solid #fed7aa;
  border-radius: 10px;
  padding: 12px;
  margin: 4px 0;
}

.question-text {
  font-weight: 600;
  margin-bottom: 10px;
  color: #7c2d12;
}

.options {
  display: flex;
  flex-direction: column;
  gap: 6px;
}

.option-button {
  display: flex;
  align-items: center;
  gap: 10px;
  padding: 8px 12px;
  border: 1px solid #fed7aa;
  border-radius: 8px;
  background: white;
  cursor: pointer;
  text-align: left;
  transition: background 0.15s;
}

.option-button:hover {
  background: #ffedd5;
}

.option-letter {
  font-weight: 600;
  color: #ea580c;
  min-width: 20px;
}

.option-text {
  color: #431407;
}
```

---

## Example Flow

**User:** "Add a total row"

**Claude:** 
> I'd like to help, but I need a bit more clarity:
> 
> **Which sheet should I add the total row to?**
> 
> A) Budget
> B) Data

*[User clicks "A) Budget"]*

**Sent to Claude:** "A) Budget"

**Claude:** `read_range(Budget, A1:D6)` → `write_range(Budget, A7, [["Total", "=SUM(B2:B6)", ...]])`

---

## Edge Cases

| Scenario | Behavior |
|---|---|
| User types answer instead of clicking | Works fine — text is sent as normal |
| Claude doesn't format with A) B) C) | Falls back to plain text bubble |
| Options exceed 4 | Only first 4 are rendered as buttons |
| Option text is very long | CSS truncates with ellipsis |

---

## Files to Modify

| File | Change |
|---|---|
| `backend/app/prompts.py` | Add clarifying question instruction to SYSTEM_PROMPT |
| `frontend/src/app/components/ClarifyingQuestion.tsx` | New component: parser + renderer |
| `frontend/src/app/components/ChatThread.tsx` | Import and use ClarifyingQuestion |
| `frontend/src/app/styles.css` | Add styles for option buttons |
| `frontend/test/clarifyingQuestion.test.ts` | New tests for parser |

---

## Tests

```ts
// frontend/test/clarifyingQuestion.test.ts
import { describe, expect, it } from "vitest";
import { parseOptions } from "../src/app/components/ClarifyingQuestion";

describe("parseOptions", () => {
  it("extracts question and options", () => {
    const text = "Which sheet?\n\nA) Budget\nB) Data\nC) New sheet";
    const result = parseOptions(text);
    expect(result?.question).toBe("Which sheet?");
    expect(result?.options).toEqual([
      { letter: "A", text: "Budget" },
      { letter: "B", text: "Data" },
      { letter: "C", text: "New sheet" },
    ]);
  });

  it("returns null for plain text without options", () => {
    expect(parseOptions("This is just a normal answer.")).toBeNull();
  });

  it("returns null for single option", () => {
    expect(parseOptions("Pick one:\n\nA) Only option")).toBeNull();
  });
});
```

---

## Effort Estimate

- Backend prompt change: 5 minutes
- Frontend component + parser: 20 minutes
- Styles: 10 minutes
- Tests: 15 minutes
- Integration + verification: 10 minutes

**Total: ~1 hour**
