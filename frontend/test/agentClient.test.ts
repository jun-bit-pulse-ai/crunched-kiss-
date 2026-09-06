import assert from "assert";
import * as excelService from "../src/app/services/excel";
import {
  MAX_REQUEST_BYTES,
  MAX_TOOL_ROUNDS,
  byteLength,
  friendlyHttpError,
  runAgent,
} from "../src/app/services/agentClient";
import type { ChatMessage, ChatResponse } from "../src/app/types";

type FakeFetch = (input: unknown, init?: { body?: string }) => Promise<Response>;

function jsonResponse(body: ChatResponse, status = 200): Response {
  return {
    ok: status >= 200 && status < 300,
    status,
    json: async () => body,
  } as Response;
}

function withFetch(fake: FakeFetch, run: () => Promise<void>): Promise<void> {
  const original = global.fetch;
  global.fetch = fake as typeof fetch;
  return run().finally(() => {
    global.fetch = original;
  });
}

function withDispatch(fake: typeof excelService.dispatchExcelTool, run: () => Promise<void>): Promise<void> {
  const original = excelService.dispatchExcelTool;
  (excelService as { dispatchExcelTool: typeof fake }).dispatchExcelTool = fake;
  return run().finally(() => {
    (excelService as { dispatchExcelTool: typeof fake }).dispatchExcelTool = original;
  });
}

describe("byteLength", () => {
  it("counts UTF-8 bytes, not JS string length", () => {
    assert.strictEqual(byteLength("abc"), 3);
    assert.strictEqual(byteLength("€"), 3); // 1 UTF-16 code unit, 3 UTF-8 bytes
  });
});

describe("friendlyHttpError", () => {
  it("maps common statuses to actionable messages", () => {
    assert.match(friendlyHttpError(404), /backend/i);
    assert.match(friendlyHttpError(401), /API key/i);
    assert.match(friendlyHttpError(500), /error/i);
  });
});

describe("runAgent", () => {
  it("rejects an oversized request before it ever reaches fetch (issue #32)", async () => {
    let fetchCalled = false;
    await withFetch(
      async () => {
        fetchCalled = true;
        return jsonResponse({ type: "message", text: "unused" });
      },
      async () => {
        const huge: ChatMessage[] = [{ role: "user", content: "x".repeat(MAX_REQUEST_BYTES + 1) }];
        await assert.rejects(() => runAgent(huge), /too large/i);
      }
    );
    assert.strictEqual(fetchCalled, false, "fetch must not run for an oversized body");
  });

  it("returns the final text on a plain end_turn reply", async () => {
    await withFetch(
      async () => jsonResponse({ type: "message", text: "42" }),
      async () => {
        const result = await runAgent([{ role: "user", content: "what is 6*7?" }]);
        assert.strictEqual(result.text, "42");
        assert.strictEqual(result.messages.at(-1)?.role, "assistant");
      }
    );
  });

  it("executes a tool call and feeds the result back as a single user message", async () => {
    let call = 0;
    const responses: ChatResponse[] = [
      { type: "tool_calls", tool_calls: [{ id: "t1", name: "list_workbook_meta", input: {} }] },
      { type: "message", text: "done" },
    ];
    await withDispatch(
      async () => ({ sheets: [] }),
      () =>
        withFetch(
          async () => jsonResponse(responses[call++]),
          async () => {
            const result = await runAgent([{ role: "user", content: "how big is this book?" }]);
            assert.strictEqual(result.text, "done");
            const toolResultTurn = result.messages.find(
              (m) => Array.isArray(m.content) && m.content.some((b) => b.type === "tool_result")
            );
            assert.ok(toolResultTurn, "expected one user message carrying the tool_result");
          }
        )
    );
  });

  it("stops calling tools and forces a text reply after MAX_TOOL_ROUNDS", async () => {
    let round = 0;
    await withDispatch(
      async () => ({ ok: true }),
      () =>
        withFetch(
          async () => {
            round += 1;
            if (round > MAX_TOOL_ROUNDS) {
              return jsonResponse({ type: "message", text: "forced" });
            }
            return jsonResponse({
              type: "tool_calls",
              tool_calls: [{ id: `t${round}`, name: "write_range", input: {} }],
            });
          },
          async () => {
            const result = await runAgent([{ role: "user", content: "keep digging forever" }]);
            assert.strictEqual(result.text, "forced");
            // MAX_TOOL_ROUNDS tool rounds, plus one forced round.
            assert.strictEqual(round, MAX_TOOL_ROUNDS + 1);
          }
        )
    );
  });

  it("surfaces a backend error message instead of throwing an opaque failure", async () => {
    await withFetch(
      async () => jsonResponse({ type: "error", message: "ANTHROPIC_API_KEY is not configured" }),
      async () => {
        await assert.rejects(
          () => runAgent([{ role: "user", content: "hi" }]),
          /ANTHROPIC_API_KEY/
        );
      }
    );
  });
});
