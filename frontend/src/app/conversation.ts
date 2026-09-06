import type { ChatMessage, ContentBlock, ToolResultBlock } from "./types";

export const MAX_HISTORY_MESSAGES = 12;

function isToolResult(block: ContentBlock): block is ToolResultBlock {
  return block.type === "tool_result";
}

export function stubOldToolResults(messages: ChatMessage[]): ChatMessage[] {
  const lastToolResultIndex = [...messages]
    .reverse()
    .findIndex((message) => Array.isArray(message.content) && message.content.some(isToolResult));
  const keepFrom =
    lastToolResultIndex === -1 ? messages.length : messages.length - 1 - lastToolResultIndex;

  return messages.map((message, index) => {
    if (index >= keepFrom || !Array.isArray(message.content)) {
      return message;
    }
    return {
      ...message,
      content: message.content.map((block) =>
        isToolResult(block) ? { ...block, content: "[omitted tool result]" } : block
      ),
    };
  });
}

/**
 * A history may only begin on a plain user turn. The API rejects a request
 * whose first message is an assistant turn, and rejects a `tool_result` whose
 * matching `tool_use` is not in the history.
 */
function opensAConversation(message: ChatMessage): boolean {
  if (message.role !== "user") {
    return false;
  }
  return !Array.isArray(message.content) || !message.content.some(isToolResult);
}

export function truncateHistory(messages: ChatMessage[]): ChatMessage[] {
  if (messages.length === 0) {
    return [];
  }

  // Cutting a fixed number of messages off the end splits tool_use/tool_result
  // pairs and can open the window on an assistant turn, both of which the API
  // rejects. One "error-check this model" request is eight tool rounds, so the
  // naive slice broke well inside a normal conversation. Cut on a clean user
  // turn instead: everything after one is self-consistent.
  const starts: number[] = [];
  for (let index = 0; index < messages.length; index += 1) {
    if (opensAConversation(messages[index])) {
      starts.push(index);
    }
  }
  if (starts.length === 0) {
    return stubOldToolResults(messages);
  }

  // Newest clean turn that still leaves a full window of context. A long tool
  // loop contains no clean turn, so the nearest cut can be far back; going over
  // budget is cheap because older tool payloads are stubbed out below, whereas
  // cutting to the last turn would drop the context a follow-up question needs.
  const roomy = starts.filter((index) => messages.length - index >= MAX_HISTORY_MESSAGES);
  const start = roomy.length > 0 ? roomy[roomy.length - 1] : starts[0];
  return stubOldToolResults(messages.slice(start));
}
