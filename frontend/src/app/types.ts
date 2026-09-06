export type Role = "user" | "assistant";

export type ToolResultBlock = {
  type: "tool_result";
  tool_use_id: string;
  content: string;
  is_error?: boolean;
};

export type ToolUseBlock = {
  type: "tool_use";
  id: string;
  name: string;
  input: Record<string, unknown>;
};

export type ContentBlock = ToolResultBlock | ToolUseBlock;

export type ChatMessage = {
  role: Role;
  content: string | ContentBlock[];
};

export type ToolCall = {
  id: string;
  name: string;
  input: Record<string, unknown>;
};

export type ChatResponse =
  | { type: "tool_calls"; tool_calls: ToolCall[] }
  | { type: "message"; text: string }
  | { type: "error"; message: string };

export type VisibleText = {
  id: string;
  kind: "text";
  role: "user" | "assistant" | "system";
  text: string;
  suggestions?: string[];
};

export type VisibleTool = {
  id: string;
  kind: "tool";
  name: string;
  summary: string;
  error: boolean;
};

export type VisibleMessage = VisibleText | VisibleTool;

export type WorkbookHint = string[];
