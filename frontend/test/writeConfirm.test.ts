import assert from "assert";
import { WRITE_DECLINED, writePreview, writePreviewSummary } from "../src/app/writeConfirm";
import { runAgent } from "../src/app/services/agentClient";
import type { ChatResponse, ToolCall } from "../src/app/types";

function toolCalls(calls: ToolCall[]): ChatResponse {
  return { type: "tool_calls", tool_calls: calls };
}

function message(text: string): ChatResponse {
  return { type: "message", text };
}

describe("writePreview", () => {
  it("summarizes sheet and address for the confirm card", () => {
    const preview = writePreview({ sheet: "Budget", address: "D4", values: [["=D2-D3"]] });
    assert.deepStrictEqual(preview, {
      sheet: "Budget",
      address: "D4",
      values: [["=D2-D3"]],
    });
    assert.strictEqual(writePreviewSummary(preview), "Budget!D4");
  });
});

describe("runAgent write confirm", () => {
  it("does not dispatch write_range when the user declines", async () => {
    const dispatched: string[] = [];
    const replies: ChatResponse[] = [
      toolCalls([{ id: "toolu_1", name: "write_range", input: { sheet: "Budget", address: "D4", values: [["=D2-D3"]] } }]),
      message("Write skipped."),
    ];

    const result = await runAgent([{ role: "user", content: "fix D4" }], undefined, {
      confirmWrite: async () => false,
      dispatch: async (name) => {
        dispatched.push(name);
        return { ok: true };
      },
      chat: async () => replies.shift() as ChatResponse,
    });

    assert.deepStrictEqual(dispatched, []);
    const toolResult = result.messages.find(
      (entry) => Array.isArray(entry.content) && entry.content[0] && "type" in entry.content[0] && entry.content[0].type === "tool_result"
    );
    assert.ok(toolResult);
    const block = (toolResult?.content as { content: string; is_error?: boolean }[])[0];
    assert.strictEqual(block.content, WRITE_DECLINED);
    assert.strictEqual(block.is_error, true);
  });

  it("dispatches write_range after the user applies", async () => {
    const dispatched: string[] = [];
    const replies: ChatResponse[] = [
      toolCalls([{ id: "toolu_1", name: "write_range", input: { sheet: "Budget", address: "D4", values: [["=D2-D3"]] } }]),
      message("Done."),
    ];

    await runAgent([{ role: "user", content: "fix D4" }], undefined, {
      confirmWrite: async () => true,
      dispatch: async (name) => {
        dispatched.push(name);
        return { ok: true };
      },
      chat: async () => replies.shift() as ChatResponse,
    });

    assert.deepStrictEqual(dispatched, ["write_range"]);
  });
});
