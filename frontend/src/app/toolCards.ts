import type { ChatMessage, ContentBlock, ToolResultBlock, ToolUseBlock } from "./types";

export type ToolCard = {
  id: string;
  name: string;
  summary: string;
  error: boolean;
};

const TOOL_DISPLAY_NAMES: Record<string, string> = {
  list_workbook_meta: "Scanned workbook",
  read_range: "Read range",
  write_range: "Wrote range",
  get_selection: "Read selection",
  find: "Searched",
};

/** Human label for a tool card. `name` stays the raw tool id elsewhere (tests, aria-labels). */
export function displayToolName(name: string): string {
  return TOOL_DISPLAY_NAMES[name] ?? name;
}

type SheetMetaLike = {
  name?: unknown;
  usedRange?: { rowCount?: unknown; columnCount?: unknown } | null;
};

function parsePayload(content: string): unknown {
  try {
    return JSON.parse(content);
  } catch {
    return content;
  }
}

function asRecord(value: unknown): Record<string, unknown> | null {
  return value !== null && typeof value === "object" && !Array.isArray(value)
    ? (value as Record<string, unknown>)
    : null;
}

function asString(value: unknown): string | undefined {
  return typeof value === "string" && value.length > 0 ? value : undefined;
}

function asNumber(value: unknown): number | undefined {
  return typeof value === "number" && Number.isFinite(value) ? value : undefined;
}

/** "· truncated" suffix when a read/search came back partial, so it's visible without asking the model to mention it. */
function truncatedSuffix(payload: Record<string, unknown> | null): string {
  return payload?.truncated === true ? " · truncated" : "";
}

function sheetSize(sheet: SheetMetaLike): string | undefined {
  const name = asString(sheet.name);
  if (!name) {
    return undefined;
  }
  const rows = asNumber(sheet.usedRange?.rowCount);
  const cols = asNumber(sheet.usedRange?.columnCount);
  if (rows === undefined || cols === undefined) {
    return name;
  }
  return `${name} ${rows}×${cols}`;
}

function formulasDiffer(payload: Record<string, unknown>): boolean {
  const values = payload.values;
  const formulas = payload.formulas;
  if (!Array.isArray(values) || !Array.isArray(formulas)) {
    return false;
  }
  return JSON.stringify(values) !== JSON.stringify(formulas);
}

function a1Address(input: Record<string, unknown>, payload: Record<string, unknown> | null): string {
  const fromPayload = asString(payload?.address);
  if (fromPayload) {
    if (fromPayload.includes("!")) {
      return fromPayload;
    }
    const sheet = asString(payload?.sheet) ?? asString(input.sheet);
    return sheet ? `${sheet}!${fromPayload}` : fromPayload;
  }
  const sheet = asString(input.sheet);
  const address = asString(input.address);
  if (sheet && address) {
    return address.includes("!") ? address : `${sheet}!${address}`;
  }
  return address ?? sheet ?? "";
}

export function summarizeTool(
  name: string,
  input: Record<string, unknown>,
  resultContent: string,
  isError = false
): { name: string; summary: string; error: boolean } {
  if (isError) {
    return { name, summary: resultContent, error: true };
  }

  const parsed = parsePayload(resultContent);
  const payload = asRecord(parsed);

  if (name === "list_workbook_meta") {
    const sheets = Array.isArray(payload?.sheets) ? (payload.sheets as SheetMetaLike[]) : [];
    const parts = sheets.map(sheetSize).filter((part): part is string => Boolean(part));
    return { name, summary: parts.join(" · ") || "no sheets", error: false };
  }

  if (name === "read_range") {
    const address = a1Address(input, payload);
    const formulaTag = payload && formulasDiffer(payload) ? " · formulas" : "";
    return {
      name,
      summary: `${address}${formulaTag}${truncatedSuffix(payload)}`,
      error: false,
    };
  }

  if (name === "write_range") {
    return { name, summary: a1Address(input, payload), error: false };
  }

  if (name === "find") {
    const query = asString(payload?.query) ?? asString(input.query) ?? "";
    const matches = asNumber(payload?.matches) ?? 0;
    const label = matches === 1 ? "1 match" : `${matches} matches`;
    return { name, summary: `“${query}” · ${label}${truncatedSuffix(payload)}`, error: false };
  }

  if (name === "get_selection") {
    return { name, summary: `${a1Address(input, payload)}${truncatedSuffix(payload)}`, error: false };
  }

  return { name, summary: typeof parsed === "string" ? parsed : JSON.stringify(parsed), error: false };
}

function isToolUse(block: ContentBlock): block is ToolUseBlock {
  return block.type === "tool_use";
}

function isToolResult(block: ContentBlock): block is ToolResultBlock {
  return block.type === "tool_result";
}

export function toolCardsFromMessages(messages: ChatMessage[], fromIndex = 0): ToolCard[] {
  const cards: ToolCard[] = [];
  const pending = new Map<string, { name: string; input: Record<string, unknown> }>();

  for (const message of messages.slice(fromIndex)) {
    if (!Array.isArray(message.content)) {
      continue;
    }
    for (const block of message.content) {
      if (isToolUse(block)) {
        pending.set(block.id, { name: block.name, input: block.input });
        continue;
      }
      if (!isToolResult(block)) {
        continue;
      }
      const call = pending.get(block.tool_use_id);
      if (!call) {
        continue;
      }
      pending.delete(block.tool_use_id);
      const summarized = summarizeTool(call.name, call.input, block.content, Boolean(block.is_error));
      cards.push({ id: block.tool_use_id, ...summarized });
    }
  }

  return cards;
}
