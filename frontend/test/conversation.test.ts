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

function toolRound(id: string): ChatMessage[] {
  return [
    { role: "assistant", content: [{ type: "tool_use", id, name: "read_range", input: {} }] },
    { role: "user", content: [{ type: "tool_result", tool_use_id: id, content: "{}" }] },
  ];
}

function longToolConversation(rounds: number, trailingText: boolean): ChatMessage[] {
  const messages: ChatMessage[] = [userText("check the budget")];
  for (let i = 0; i < rounds; i += 1) {
    messages.push(...toolRound(`t${i}`));
  }
  if (trailingText) {
    messages.push(userText("and the margins?"));
  }
  return messages;
}

function orphanedToolResults(messages: ChatMessage[]): string[] {
  const offered = new Set<string>();
  for (const message of messages) {
    if (!Array.isArray(message.content)) continue;
    for (const block of message.content) {
      if (block.type === "tool_use") offered.add(block.id);
    }
  }
  const orphans: string[] = [];
  for (const message of messages) {
    if (!Array.isArray(message.content)) continue;
    for (const block of message.content) {
      if (block.type === "tool_result" && !offered.has(block.tool_use_id)) {
        orphans.push(block.tool_use_id);
      }
    }
  }
  return orphans;
}

describe("truncateHistory keeps the window valid for the API", () => {
  // The API rejects a history that opens with an assistant turn, and rejects a
  // tool_result whose tool_use was cut away. Eight tool rounds is one ordinary
  // "error-check this model" request, so the plain slice fell foul of both.
  for (const rounds of [6, 7, 8]) {
    for (const trailingText of [false, true]) {
      it(`starts on a user turn with ${rounds} rounds (trailing text: ${trailingText})`, () => {
        const window = truncateHistory(longToolConversation(rounds, trailingText));
        assert.strictEqual(window[0].role, "user");
        assert.ok(
          !Array.isArray(window[0].content) ||
            !window[0].content.some((block) => block.type === "tool_result"),
          "window must not open with a tool_result"
        );
      });

      it(`leaves no orphaned tool_result with ${rounds} rounds (trailing text: ${trailingText})`, () => {
        const window = truncateHistory(longToolConversation(rounds, trailingText));
        assert.deepStrictEqual(orphanedToolResults(window), []);
      });
    }
  }

  it("keeps earlier context when a follow-up question lands after a tool loop", () => {
    const messages = longToolConversation(8, true);
    const window = truncateHistory(messages);
    // The follow-up ("and the margins?") only makes sense with the original
    // question in view, so the window must not collapse to the last message.
    assert.ok(window.length > 1, `window collapsed to ${window.length} message(s)`);
    assert.deepStrictEqual(window[0], userText("check the budget"));
    assert.deepStrictEqual(window[window.length - 1], userText("and the margins?"));
  });

  it("keeps the matching tool_use when the newest message is a tool_result", () => {
    const messages = longToolConversation(8, false);
    const window = truncateHistory(messages);
    const last = window[window.length - 1];
    assert.deepStrictEqual(last, messages[messages.length - 1]);
    assert.deepStrictEqual(orphanedToolResults(window), []);
  });

  it("still caps an ordinary back-and-forth chat", () => {
    const messages: ChatMessage[] = [];
    for (let i = 0; i < 20; i += 1) {
      messages.push(userText(`q ${i}`));
      messages.push({ role: "assistant", content: `a ${i}` });
    }
    messages.push(userText("latest"));
    const window = truncateHistory(messages);
    assert.ok(window.length <= MAX_HISTORY_MESSAGES + 1, `window was ${window.length}`);
    assert.strictEqual(window[0].role, "user");
  });
});
