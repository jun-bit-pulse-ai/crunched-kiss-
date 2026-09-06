import { WELCOME } from "./demoPrompts";
import type { VisibleMessage } from "./types";

/**
 * Kid-simple names for the Excel tools the pane already runs.
 * Keep these in sync with excel.ts / tools.py names — do not rename the tools.
 */
export const TOOL_ELI5: Record<string, string> = {
  list_workbook_meta:
    "I peeked at the sheet names and how big they are. I did not read every cell.",
  read_range: "I looked at a little box of cells, not the whole spreadsheet.",
  write_range: "I wrote a number or a recipe into a cell.",
  get_selection: "I looked at the cells you had clicked.",
  find: "I hunted for a word on the sheet, like finding a name on a list.",
};

export type Eli5Tool = {
  name: string;
  error: boolean;
};

export type Eli5Input = {
  assistantText: string | null;
  tools: Eli5Tool[];
};

const WELCOME_ELI5 =
  "I am Crunched. I sit in the side of Excel and help with spreadsheets. I can check mistakes, build number stories, and peek at your sheets. Ask me something and I will show my work.";

export function kidSentenceForTool(name: string, error = false): string {
  if (error) {
    const happy = TOOL_ELI5[name];
    if (happy) {
      return `I tried, but that peek did not work. (${name})`;
    }
    return `I tried a helper called ${name}, but it did not work.`;
  }
  return TOOL_ELI5[name] ?? `I used a helper called ${name}.`;
}

/** Pull the last user turn's tool cards and the assistant reply that followed. */
export function eli5InputFromVisible(messages: VisibleMessage[]): Eli5Input {
  let lastUser = -1;
  for (let index = messages.length - 1; index >= 0; index -= 1) {
    const message = messages[index];
    if (message.kind === "text" && message.role === "user") {
      lastUser = index;
      break;
    }
  }

  const slice = lastUser >= 0 ? messages.slice(lastUser + 1) : messages;
  const tools: Eli5Tool[] = [];
  let assistantText: string | null = null;

  for (const message of slice) {
    if (message.kind === "tool") {
      tools.push({ name: message.name, error: message.error });
      continue;
    }
    if (message.kind === "text" && message.role === "assistant") {
      assistantText = message.text;
    }
  }

  return { assistantText, tools };
}

/** Strip Markdown marks and keep the first couple of sentences. */
export function stripToKidProse(text: string, maxChars = 240): string {
  const plain = text
    .replace(/```[\s\S]*?```/g, " ")
    .replace(/^#{1,6}\s+/gm, "")
    .replace(/\*\*([^*]+)\*\*/g, "$1")
    .replace(/`([^`]+)`/g, "$1")
    .replace(/\*([^*\n]+)\*/g, "$1")
    .replace(/\[([^\]]+)\]\([^)]+\)/g, "$1")
    .replace(/^\s*[-*+]\s+/gm, "")
    .replace(/\s+/g, " ")
    .trim();

  if (!plain) {
    return "";
  }

  const sentences = (plain.match(/[^.!?]+[.!?]+|[^.!?]+$/g) ?? [plain]).map((part) => part.trim());
  let clipped = sentences.slice(0, 2).join(" ").trim();
  if (clipped.length > maxChars) {
    clipped = `${clipped.slice(0, maxChars - 1).trimEnd()}…`;
  }
  return clipped;
}

export function explainLikeFive(input: Eli5Input): string {
  const lines: string[] = [];

  if (input.tools.length > 0) {
    lines.push("Here is what I did with the spreadsheet:");
    for (const tool of input.tools) {
      lines.push(`• ${kidSentenceForTool(tool.name, tool.error)}`);
    }
  }

  if (input.assistantText === WELCOME) {
    if (lines.length === 0) {
      return WELCOME_ELI5;
    }
    lines.push(WELCOME_ELI5);
    return lines.join("\n");
  }

  const prose = input.assistantText ? stripToKidProse(input.assistantText) : "";
  if (prose) {
    lines.push(input.tools.length > 0 ? `Then I said: ${prose}` : prose);
  }

  if (lines.length === 0) {
    return "I have not done anything yet. Ask me a question and I will explain it in tiny words.";
  }

  return lines.join("\n");
}
