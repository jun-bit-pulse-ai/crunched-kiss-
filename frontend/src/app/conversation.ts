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

export function truncateHistory(messages: ChatMessage[]): ChatMessage[] {
  if (messages.length === 0) {
    return [];
  }
  const latest = messages[messages.length - 1];
  const window = messages.slice(-MAX_HISTORY_MESSAGES);
  if (window[window.length - 1] !== latest) {
    return stubOldToolResults([...window, latest]);
  }
  return stubOldToolResults(window);
}
