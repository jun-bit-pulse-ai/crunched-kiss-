import assert from "assert";
import { MAX_HISTORY_MESSAGES, stubOldToolResults, truncateHistory } from "../src/app/conversation";
import type { ChatMessage } from "../src/app/types";

function userText(text: string): ChatMessage {
  return { role: "user", content: text };
}

describe("truncateHistory", () => {
  it("keeps the latest user turn even when history is long", () => {
    const messages: ChatMessage[] = [];
    for (let i = 0; i < 20; i += 1) {
      messages.push(userText(`old ${i}`));
      messages.push({ role: "assistant", content: `reply ${i}` });
    }
    messages.push(userText("latest question"));
    const truncated = truncateHistory(messages);
    assert.ok(truncated.length <= MAX_HISTORY_MESSAGES + 1);
    assert.deepStrictEqual(truncated[truncated.length - 1], userText("latest question"));
  });
});

describe("stubOldToolResults", () => {
  it("replaces older tool payloads with a one-line stub", () => {
    const messages: ChatMessage[] = [
      {
        role: "user",
        content: [
          {
            type: "tool_result",
            tool_use_id: "old",
            content: JSON.stringify({ values: [[1, 2, 3]] }),
          },
        ],
      },
      userText("follow up"),
      {
        role: "user",
        content: [
          {
            type: "tool_result",
            tool_use_id: "new",
            content: JSON.stringify({ values: [["keep"]] }),
          },
        ],
      },
    ];
    const stubbed = stubOldToolResults(messages);
    const first = stubbed[0].content;
    assert.ok(Array.isArray(first));
    assert.ok(first[0].type === "tool_result");
    assert.strictEqual(first[0].content, "[omitted tool result]");
    const last = stubbed[2].content;
    assert.ok(Array.isArray(last));
    assert.ok(last[0].type === "tool_result");
    assert.ok(last[0].content.includes("keep"));
  });
});
