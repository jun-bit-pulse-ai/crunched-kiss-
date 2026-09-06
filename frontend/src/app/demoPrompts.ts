import type { VisibleMessage } from "./types";

export const WELCOME =
  "Hi, I'm Crunched — your AI analyst in Excel. I can help you with things like:\n\nError checking and fixing models\nBuilding financial or business models\nAnalyzing and comparing scenario models\n\nWhat should we work on first?";

export const DEMO_CHIPS = [
  "How big is this workbook?",
  "Check the Budget sheet for errors",
  "Fix the hard-coded Gross profit in Budget!D4",
] as const;

export function initialVisible(): VisibleMessage[] {
  return [{ id: "welcome", kind: "text", role: "assistant", text: WELCOME }];
}

export function showPromptChips(messages: VisibleMessage[]): boolean {
  return !messages.some((message) => message.kind === "text" && message.role === "user");
}
