import assert from "assert";
import {
  clearConversation,
  loadConversation,
  saveConversation,
  workbookKey,
  type ConversationStore,
} from "../src/app/storage";

function memoryStore(): ConversationStore {
  const data = new Map<string, string>();
  return {
    getItem: (key) => data.get(key) ?? null,
    setItem: (key, value) => {
      data.set(key, value);
    },
  };
}

const memory = new Map<string, string>();
(globalThis as { localStorage: Storage }).localStorage = {
  get length() {
    return memory.size;
  },
  clear() {
    memory.clear();
  },
  getItem(key: string) {
    return memory.get(key) ?? null;
  },
  setItem(key: string, value: string) {
    memory.set(key, value);
  },
  removeItem(key: string) {
    memory.delete(key);
  },
  key() {
    return null;
  },
} as Storage;

describe("storage", () => {
  it("round-trips a conversation", () => {
    const store = memoryStore();
    const agentMessages = [{ role: "user" as const, content: "hello" }];
    const visible = [{ id: "1", kind: "text" as const, role: "assistant" as const, text: "hi" }];
    saveConversation(["Budget", "Data"], agentMessages, visible, store);
    const restored = loadConversation(["Data", "Budget"], store);
    assert.deepStrictEqual(restored?.agentMessages, agentMessages);
    assert.deepStrictEqual(restored?.visible, visible);
  });

  it("returns null when no conversation exists", () => {
    assert.strictEqual(loadConversation(["Missing"], memoryStore()), null);
  });

  it("evicts the oldest conversation after 5 workbooks", () => {
    const store = memoryStore();
    for (let i = 1; i <= 6; i += 1) {
      saveConversation([`Sheet${i}`], [{ role: "user" as const, content: String(i) }], [], store);
    }
    assert.strictEqual(loadConversation(["Sheet1"], store), null);
    assert.notStrictEqual(loadConversation(["Sheet6"], store), null);
  });

  it("clears a conversation", () => {
    const store = memoryStore();
    saveConversation(["Budget"], [{ role: "user" as const, content: "x" }], [], store);
    clearConversation(["Budget"], store);
    assert.strictEqual(loadConversation(["Budget"], store), null);
  });

  it("workbookKey is order-independent and does not mutate its input", () => {
    const a = ["B", "A"];
    const b = ["A", "B"];
    assert.strictEqual(workbookKey(a), workbookKey(b));
    assert.deepStrictEqual(a, ["B", "A"]); // sort() must not mutate the caller's array
  });

  it("treats a missing store as a no-op that never throws", () => {
    assert.strictEqual(loadConversation(["X"], null), null);
    saveConversation(["X"], [], [], null); // should not throw
    clearConversation(["X"], null); // should not throw
  });

  it("cannot pollute the prototype through a __proto__ sheet name", () => {
    // The key is namespaced ("sheets:__proto__"), so a hostile sheet name can no
    // longer *be* a dangerous key. Assert the property that matters rather than
    // the old mechanism: nothing leaks onto Object.prototype, and the entry is
    // still readable only by the workbook that wrote it.
    saveConversation(["__proto__"], [{ role: "user", content: "nope" }], []);
    assert.strictEqual(({} as Record<string, unknown>).polluted, undefined);
    assert.strictEqual(Object.getPrototypeOf({}), Object.prototype);
    const mine = loadConversation(["__proto__"]);
    assert.deepStrictEqual(mine?.agentMessages, [{ role: "user", content: "nope" }]);
    assert.strictEqual(loadConversation(["Sheet1"]), null);
  });

  it("still ignores a raw __proto__ key in tampered storage", () => {
    localStorage.setItem(
      "crunched_conversations",
      JSON.stringify({
        version: 1,
        conversations: { __proto__: { version: 1, agentMessages: [], visible: [], savedAt: "x" } },
        lru: ["__proto__"],
      })
    );
    assert.strictEqual(loadConversation([], undefined, "url:whatever"), null);
    assert.strictEqual(Object.getPrototypeOf({}), Object.prototype);
  });

  it("ignores tampered localStorage that is not a conversation blob", () => {
    localStorage.setItem(
      "crunched_conversations",
      JSON.stringify({
        version: 1,
        conversations: { Budget: { version: 1, agentMessages: "boom", visible: [], savedAt: "x" } },
        lru: ["Budget"],
      })
    );
    assert.strictEqual(loadConversation(["Budget"]), null);
  });
});

describe("workbookKey identifies the file, not its shape", () => {
  it("uses the document URL when Excel gives us one", () => {
    assert.strictEqual(
      workbookKey(["Sheet1"], "file:///Users/j/Documents/budget.xlsx"),
      "url:file:///Users/j/Documents/budget.xlsx"
    );
  });

  it("keeps two different files apart even with identical sheet names", () => {
    const a = workbookKey(["Sheet1"], "file:///a/one.xlsx");
    const b = workbookKey(["Sheet1"], "file:///b/two.xlsx");
    assert.notStrictEqual(a, b);
  });

  it("still matches the same file after a sheet is added or renamed", () => {
    const before = workbookKey(["Data", "Budget"], "file:///a/model.xlsx");
    const after = workbookKey(["Data", "Budget", "Scratch"], "file:///a/model.xlsx");
    assert.strictEqual(before, after);
  });

  it("falls back to sheet names for an unsaved workbook", () => {
    assert.strictEqual(workbookKey(["Budget", "Data"], null), "sheets:Budget\nData");
    assert.strictEqual(workbookKey(["Budget", "Data"], ""), "sheets:Budget\nData");
  });

  it("does not confuse a sheet-name key with a URL key", () => {
    assert.notStrictEqual(workbookKey(["Sheet1"], null), workbookKey(["Sheet1"], "file:///a/x.xlsx"));
  });
});

describe("conversations do not leak between workbooks", () => {
  it("keeps two blank workbooks separate", () => {
    const store = memoryStore();
    saveConversation(["Sheet1"], [{ role: "user", content: "workbook A" }], [], store, "file:///a.xlsx");
    const other = loadConversation(["Sheet1"], store, "file:///b.xlsx");
    assert.strictEqual(other, null);
  });
});
