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
});
