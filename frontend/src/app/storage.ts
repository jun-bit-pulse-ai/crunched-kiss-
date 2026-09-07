import type { ChatMessage, VisibleMessage } from "./types";

const STORAGE_KEY = "crunched_conversations";
const SCHEMA_VERSION = 1;
const MAX_STORED_WORKBOOKS = 5;
const MAX_STORED_JSON_CHARS = 500_000;
const DANGEROUS_KEYS = new Set(["__proto__", "constructor", "prototype"]);

type PersistedConversation = {
  version: number;
  agentMessages: ChatMessage[];
  visible: VisibleMessage[];
  savedAt: string;
};

type StorageData = {
  version: number;
  conversations: Record<string, PersistedConversation>;
  lru: string[]; // workbook keys in order of most-recently-used
};

export type ConversationStore = {
  getItem(key: string): string | null;
  setItem(key: string, value: string): void;
};

function defaultStore(): ConversationStore | null {
  try {
    if (typeof localStorage !== "undefined") {
      return localStorage;
    }
  } catch {
    // Private mode or a non-browser test runner.
  }
  return null;
}

function emptyConversations(): Record<string, PersistedConversation> {
  return Object.create(null) as Record<string, PersistedConversation>;
}

function emptyData(): StorageData {
  return { version: SCHEMA_VERSION, conversations: emptyConversations(), lru: [] };
}

export function isSafeStorageKey(key: string): boolean {
  return key.length > 0 && !DANGEROUS_KEYS.has(key);
}

function isChatMessage(value: unknown): value is ChatMessage {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return false;
  }
  const rec = value as Record<string, unknown>;
  if (rec.role !== "user" && rec.role !== "assistant") {
    return false;
  }
  return typeof rec.content === "string" || Array.isArray(rec.content);
}

function isVisibleMessage(value: unknown): value is VisibleMessage {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return false;
  }
  const rec = value as Record<string, unknown>;
  if (typeof rec.id !== "string") {
    return false;
  }
  if (rec.kind === "text") {
    return (
      (rec.role === "user" || rec.role === "assistant" || rec.role === "system") &&
      typeof rec.text === "string"
    );
  }
  if (rec.kind === "tool") {
    return typeof rec.name === "string" && typeof rec.summary === "string" && typeof rec.error === "boolean";
  }
  if (rec.kind === "write_confirm") {
    return (
      typeof rec.sheet === "string" &&
      typeof rec.address === "string" &&
      (rec.status === "pending" || rec.status === "applied" || rec.status === "declined")
    );
  }
  return false;
}

function isPersistedConversation(value: unknown): value is PersistedConversation {
  if (!value || typeof value !== "object" || Array.isArray(value)) {
    return false;
  }
  const rec = value as Record<string, unknown>;
  return (
    rec.version === SCHEMA_VERSION &&
    typeof rec.savedAt === "string" &&
    Array.isArray(rec.agentMessages) &&
    rec.agentMessages.every(isChatMessage) &&
    Array.isArray(rec.visible) &&
    rec.visible.every(isVisibleMessage)
  );
}

function parseStorage(raw: string): StorageData {
  const parsed: unknown = JSON.parse(raw);
  if (!parsed || typeof parsed !== "object" || Array.isArray(parsed)) {
    return emptyData();
  }
  const rec = parsed as Record<string, unknown>;
  if (rec.version !== SCHEMA_VERSION) {
    return emptyData();
  }
  if (!rec.conversations || typeof rec.conversations !== "object" || Array.isArray(rec.conversations)) {
    return emptyData();
  }
  if (!Array.isArray(rec.lru) || !rec.lru.every((key) => typeof key === "string")) {
    return emptyData();
  }
  const conversations = emptyConversations();
  for (const [key, value] of Object.entries(rec.conversations as Record<string, unknown>)) {
    if (!isSafeStorageKey(key) || !isPersistedConversation(value)) {
      continue;
    }
    conversations[key] = value;
  }
  return {
    version: SCHEMA_VERSION,
    conversations,
    lru: rec.lru.filter(
      (key) => isSafeStorageKey(key) && Object.prototype.hasOwnProperty.call(conversations, key)
    ),
  };
}

function readStorage(store: ConversationStore | null): StorageData {
  try {
    const raw = store?.getItem(STORAGE_KEY);
    if (!raw) {
      return emptyData();
    }
    return parseStorage(raw);
  } catch {
    return emptyData();
  }
}

function writeStorage(store: ConversationStore | null, data: StorageData): void {
  try {
    const raw = JSON.stringify(data);
    if (raw.length > MAX_STORED_JSON_CHARS) {
      return;
    }
    store?.setItem(STORAGE_KEY, raw);
  } catch {
    // Storage full or private mode — silently fail.
  }
}

/**
 * Identify the workbook a conversation belongs to.
 *
 * Sheet names alone are not an identity: every blank workbook is called
 * "Sheet1", so two of them shared a conversation, and adding a sheet to a real
 * model silently orphaned its history. Prefer the document URL, which Excel
 * gives us for any saved file, and keep sheet names only as the fallback for a
 * workbook that has never been saved. The prefixes stop the two namespaces
 * colliding.
 */
export function workbookKey(sheetNames: string[], documentUrl?: string | null): string {
  if (documentUrl) {
    return `url:${documentUrl}`;
  }
  return `sheets:${[...sheetNames].sort().join("\n")}`;
}

export function saveConversation(
  sheetNames: string[],
  agentMessages: ChatMessage[],
  visible: VisibleMessage[],
  store: ConversationStore | null = defaultStore(),
  documentUrl?: string | null
): void {
  const key = workbookKey(sheetNames, documentUrl);
  if (!isSafeStorageKey(key)) {
    return;
  }
  const data = readStorage(store);

  data.conversations[key] = {
    version: SCHEMA_VERSION,
    agentMessages,
    visible,
    savedAt: new Date().toISOString(),
  };

  // Update LRU: remove existing, push to front.
  data.lru = data.lru.filter((item) => item !== key);
  data.lru.unshift(key);

  // Evict oldest if over limit.
  while (data.lru.length > MAX_STORED_WORKBOOKS) {
    const oldest = data.lru.pop();
    if (oldest) {
      delete data.conversations[oldest];
    }
  }

  writeStorage(store, data);
}

export function loadConversation(
  sheetNames: string[],
  store: ConversationStore | null = defaultStore(),
  documentUrl?: string | null
): { agentMessages: ChatMessage[]; visible: VisibleMessage[] } | null {
  const key = workbookKey(sheetNames, documentUrl);
  if (!isSafeStorageKey(key)) {
    return null;
  }
  const data = readStorage(store);
  const conv = Object.prototype.hasOwnProperty.call(data.conversations, key)
    ? data.conversations[key]
    : undefined;
  if (!conv || conv.version !== SCHEMA_VERSION) {
    return null;
  }
  return { agentMessages: conv.agentMessages, visible: conv.visible };
}

export function clearConversation(
  sheetNames: string[],
  store: ConversationStore | null = defaultStore(),
  documentUrl?: string | null
): void {
  const key = workbookKey(sheetNames, documentUrl);
  if (!isSafeStorageKey(key)) {
    return;
  }
  const data = readStorage(store);
  delete data.conversations[key];
  data.lru = data.lru.filter((item) => item !== key);
  writeStorage(store, data);
}
