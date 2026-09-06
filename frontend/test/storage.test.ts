import assert from "assert";
import { clearConversation, loadConversation, saveConversation, workbookKey } from "../src/app/storage";

describe("storage", () => {
  // Clear localStorage before each test so they don't leak.
  beforeEach(() => {
    localStorage.clear();
  });

  it("round-trips a conversation", () => {
    const agentMessages = [{ role: "user" as const, content: "hello" }];
    const visible = [{ id: "1", kind: "text" as const, role: "assistant" as const, text: "hi" }];
    saveConversation(["Budget", "Data"], agentMessages, visible);
    const restored = loadConversation(["Data", "Budget"]);
    assert.deepStrictEqual(restored?.agentMessages, agentMessages);
    assert.deepStrictEqual(restored?.visible, visible);
  });

  it("returns null when no conversation exists", () => {
    assert.strictEqual(loadConversation(["Missing"]), null);
  });

  it("evicts the oldest conversation after 5 workbooks", () => {
    for (let i = 1; i <= 6; i += 1) {
      saveConversation([`Sheet${i}`], [{ role: "user" as const, content: String(i) }], []);
    }
    assert.strictEqual(loadConversation(["Sheet1"]), null);
    assert.notStrictEqual(loadConversation(["Sheet6"]), null);
  });

  it("clears a conversation", () => {
    saveConversation(["Budget"], [{ role: "user" as const, content: "x" }], []);
    clearConversation(["Budget"]);
    assert.strictEqual(loadConversation(["Budget"]), null);
  });

  it("workbookKey is order-independent", () => {
    assert.strictEqual(workbookKey(["B", "A"]), workbookKey(["A", "B"]));
  });
});
