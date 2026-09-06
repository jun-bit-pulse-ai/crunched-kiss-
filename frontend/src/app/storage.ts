import type { ChatMessage, VisibleMessage } from "./types";

const STORAGE_KEY = "crunched_conversations";
const SCHEMA_VERSION = 1;
const MAX_STORED_WORKBOOKS = 5;

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

function emptyData(): StorageData {
  return { version: SCHEMA_VERSION, conversations: {}, lru: [] };
}

function readStorage(store: ConversationStore | null): StorageData {
  try {
    const raw = store?.getItem(STORAGE_KEY);
    if (!raw) {
      return emptyData();
    }
    const parsed = JSON.parse(raw) as StorageData;
    if (parsed.version !== SCHEMA_VERSION) {
      return emptyData();
    }
    return parsed;
  } catch {
    return emptyData();
  }
}

function writeStorage(store: ConversationStore | null, data: StorageData): void {
  try {
    store?.setItem(STORAGE_KEY, JSON.stringify(data));
  } catch {
    // Storage full or private mode — silently fail.
  }
}

/** Build a workbook key from sheet names (fallback when file name is unavailable). */
export function workbookKey(sheetNames: string[]): string {
  return [...sheetNames].sort().join("\n");
}

export function saveConversation(
  sheetNames: string[],
  agentMessages: ChatMessage[],
  visible: VisibleMessage[],
  store: ConversationStore | null = defaultStore()
): void {
  const key = workbookKey(sheetNames);
  const data = readStorage(store);

  data.conversations[key] = {
    version: SCHEMA_VERSION,
    agentMessages,
    visible,
    savedAt: new Date().toISOString(),
  };

  // Update LRU: remove existing, push to front.
  data.lru = data.lru.filter((k) => k !== key);
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
  store: ConversationStore | null = defaultStore()
): { agentMessages: ChatMessage[]; visible: VisibleMessage[] } | null {
  const key = workbookKey(sheetNames);
  const data = readStorage(store);
  const conv = data.conversations[key];
  if (!conv || conv.version !== SCHEMA_VERSION) {
    return null;
  }
  return { agentMessages: conv.agentMessages, visible: conv.visible };
}

export function clearConversation(
  sheetNames: string[],
  store: ConversationStore | null = defaultStore()
): void {
  const key = workbookKey(sheetNames);
  const data = readStorage(store);
  delete data.conversations[key];
  data.lru = data.lru.filter((k) => k !== key);
  writeStorage(store, data);
}
