import { truncateHistory } from "../conversation";
import type { ChatMessage, ChatResponse, ToolCall, WorkbookHint } from "../types";
import { dispatchExcelTool } from "./excel";

export const MAX_TOOL_ROUNDS = 8;
export const BACKEND_URL = "https://localhost:8000";

function asToolUseMessage(toolCalls: ToolCall[]): ChatMessage {
  return {
    role: "assistant",
    content: toolCalls.map((call) => ({
      type: "tool_use",
      id: call.id,
      name: call.name,
      input: call.input,
    })),
  };
}

async function postChat(
  messages: ChatMessage[],
  workbookHint?: WorkbookHint,
  forceText = false
): Promise<ChatResponse> {
  const response = await fetch(`${BACKEND_URL}/chat`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      messages: truncateHistory(messages),
      workbook_hint: workbookHint,
      force_text: forceText,
    }),
  });
  if (!response.ok) {
    throw new Error(`Backend returned ${response.status}`);
  }
  return (await response.json()) as ChatResponse;
}

export async function runAgent(
  messages: ChatMessage[],
  workbookHint?: WorkbookHint,
  onStatus?: (status: string) => void
): Promise<{ messages: ChatMessage[]; text: string }> {
  const next = [...messages];

  for (let round = 0; round < MAX_TOOL_ROUNDS; round += 1) {
    onStatus?.(round === 0 ? "Thinking…" : `Workbook step ${round}…`);
    const result = await postChat(next, workbookHint, false);
    if (result.type === "error") {
      throw new Error(result.message);
    }
    if (result.type === "message") {
      next.push({ role: "assistant", content: result.text });
      return { messages: next, text: result.text };
    }

    next.push(asToolUseMessage(result.tool_calls));
    const toolResults = [];
    for (const call of result.tool_calls) {
      onStatus?.(`${call.name}`);
      try {
        const output = await dispatchExcelTool(call.name, call.input);
        toolResults.push({
          type: "tool_result" as const,
          tool_use_id: call.id,
          content: JSON.stringify(output),
        });
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        toolResults.push({
          type: "tool_result" as const,
          tool_use_id: call.id,
          content: message,
          is_error: true,
        });
      }
    }
    next.push({ role: "user", content: toolResults });
  }

  onStatus?.("Wrapping up…");
  next.push({
    role: "user",
    content: "Tool round limit reached. Reply with what you know so far. Do not call more tools.",
  });
  const forced = await postChat(next, workbookHint, true);
  if (forced.type === "error") {
    throw new Error(forced.message);
  }
  if (forced.type !== "message") {
    throw new Error("Model kept requesting tools after the round limit");
  }
  next.push({ role: "assistant", content: forced.text });
  return { messages: next, text: forced.text };
}
