import { CHAT_PATH } from "../apiPaths";
import { truncateHistory } from "../conversation";
import type { ChatMessage, ChatResponse, ToolCall, WorkbookHint } from "../types";
import { WRITE_DECLINED, writePreview, type WritePreview } from "../writeConfirm";
import { dispatchExcelTool } from "./excel";

export const MAX_TOOL_ROUNDS = 8;

// Mirrors backend/app/limits.py's MAX_BODY_BYTES. The backend rejects an oversized
// body with 413 before reading it; webpack then reports an opaque 500. Check here
// so a too-large request never leaves the browser.
export const MAX_REQUEST_BYTES = 1_048_576;

/** UTF-8 byte length of a request body, matching how the backend measures Content-Length. */
export function byteLength(body: string): number {
  return new TextEncoder().encode(body).length;
}

export type RunAgentOptions = {
  onStatus?: (status: string) => void;
  confirmWrite?: (preview: WritePreview) => Promise<boolean>;
  dispatch?: (name: string, input: Record<string, unknown>) => Promise<unknown>;
  chat?: (
    messages: ChatMessage[],
    workbookHint?: WorkbookHint,
    forceText?: boolean
  ) => Promise<ChatResponse>;
};

/** Turn a raw HTTP status into a message a spreadsheet user (not a developer) can act on. */
export function friendlyHttpError(status: number): string {
  if (status === 404) {
    return "Crunched couldn't reach the backend — is `./scripts/dev-backend.sh` running?";
  }
  if (status === 401 || status === 403) {
    return "Crunched's Anthropic API key looks invalid. Check ANTHROPIC_API_KEY in .env and restart the backend.";
  }
  if (status >= 500) {
    return "Crunched's backend hit an error processing that request. Check the backend terminal for details and try again.";
  }
  return `Crunched couldn't complete that request (server said ${status}). Try again in a moment.`;
}

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
  const body = JSON.stringify({
    messages: truncateHistory(messages),
    workbook_hint: workbookHint,
    force_text: forceText,
  });
  if (byteLength(body) > MAX_REQUEST_BYTES) {
    throw new Error(
      "That request is too large to send — try a shorter message or ask about a smaller range."
    );
  }

  let response: Response;
  try {
    response = await fetch(CHAT_PATH, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body,
    });
  } catch {
    throw new Error("Crunched couldn't reach the backend — is `./scripts/dev-backend.sh` running?");
  }
  if (!response.ok) {
    throw new Error(friendlyHttpError(response.status));
  }
  return (await response.json()) as ChatResponse;
}

function resolveOptions(
  onStatusOrOptions?: ((status: string) => void) | RunAgentOptions
): RunAgentOptions {
  if (typeof onStatusOrOptions === "function") {
    return { onStatus: onStatusOrOptions };
  }
  return onStatusOrOptions ?? {};
}

export async function runAgent(
  messages: ChatMessage[],
  workbookHint?: WorkbookHint,
  onStatusOrOptions?: ((status: string) => void) | RunAgentOptions
): Promise<{ messages: ChatMessage[]; text: string }> {
  const options = resolveOptions(onStatusOrOptions);
  const chat = options.chat ?? postChat;
  const dispatch = options.dispatch ?? dispatchExcelTool;
  const next = [...messages];

  const finish = (result: ChatResponse) => {
    if (result.type !== "message") {
      throw new Error("Model kept requesting tools after the round limit");
    }
    next.push({ role: "assistant", content: result.text });
    return { messages: next, text: result.text };
  };

  for (let round = 0; round < MAX_TOOL_ROUNDS; round += 1) {
    options.onStatus?.(round === 0 ? "Thinking…" : `Workbook step ${round}…`);
    const result = await chat(next, workbookHint, false);
    if (result.type === "error") {
      throw new Error(result.message);
    }
    if (result.type === "message") {
      return finish(result);
    }

    next.push(asToolUseMessage(result.tool_calls));
    const toolResults = [];
    for (const call of result.tool_calls) {
      options.onStatus?.(`${call.name}`);
      try {
        if (call.name === "write_range" && options.confirmWrite) {
          const allowed = await options.confirmWrite(writePreview(call.input));
          if (!allowed) {
            toolResults.push({
              type: "tool_result" as const,
              tool_use_id: call.id,
              content: WRITE_DECLINED,
              is_error: true,
            });
            continue;
          }
        }
        const output = await dispatch(call.name, call.input);
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

  options.onStatus?.("Wrapping up…");
  next.push({
    role: "user",
    content: "Tool round limit reached. Reply with what you know so far. Do not call more tools.",
  });
  const forced = await chat(next, workbookHint, true);
  if (forced.type === "error") {
    throw new Error(forced.message);
  }
  return finish(forced);
}
