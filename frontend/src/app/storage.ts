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

function readStorage(): StorageData {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) {
      return { version: SCHEMA_VERSION, conversations: {}, lru: [] };
    }
    const parsed = JSON.parse(raw) as StorageData;
    if (parsed.version !== SCHEMA_VERSION) {
      return { version: SCHEMA_VERSION, conversations: {}, lru: [] };
    }
    return parsed;
  } catch {
    return { version: SCHEMA_VERSION, conversations: {}, lru: [] };
  }
}

function writeStorage(data: StorageData): void {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
  } catch {
    // Storage full or private mode — silently fail.
  }
}

/** Build a workbook key from sheet names (fallback when file name is unavailable). */
export function workbookKey(sheetNames: string[]): string {
  return sheetNames.sort().join("\n");
}

export function saveConversation(
  sheetNames: string[],
  agentMessages: ChatMessage[],
  visible: VisibleMessage[]
): void {
  const key = workbookKey(sheetNames);
  const data = readStorage();

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

  writeStorage(data);
}

export function loadConversation(
  sheetNames: string[]
): { agentMessages: ChatMessage[]; visible: VisibleMessage[] } | null {
  const key = workbookKey(sheetNames);
  const data = readStorage();
  const conv = data.conversations[key];
  if (!conv || conv.version !== SCHEMA_VERSION) {
    return null;
  }
  return { agentMessages: conv.agentMessages, visible: conv.visible };
}

export function clearConversation(sheetNames: string[]): void {
  const key = workbookKey(sheetNames);
  const data = readStorage();
  delete data.conversations[key];
  data.lru = data.lru.filter((k) => k !== key);
  writeStorage(data);
}
