import assert from "assert";
import { summarizeTool, toolCardsFromMessages } from "../src/app/toolCards";
import type { ChatMessage } from "../src/app/types";

function toolUse(id: string, name: string, input: Record<string, unknown> = {}): ChatMessage {
  return {
    role: "assistant",
    content: [{ type: "tool_use", id, name, input }],
  };
}

function toolResult(id: string, content: unknown, isError = false): ChatMessage {
  return {
    role: "user",
    content: [
      {
        type: "tool_result",
        tool_use_id: id,
        content: typeof content === "string" ? content : JSON.stringify(content),
        is_error: isError,
      },
    ],
  };
}

describe("summarizeTool", () => {
  it("lists sheet names and used-range sizes for list_workbook_meta", () => {
    const card = summarizeTool(
      "list_workbook_meta",
      {},
      JSON.stringify({
        sheets: [
          { name: "Data", usedRange: { address: "Data!A1:GR5000", rowCount: 5000, columnCount: 200 } },
          { name: "Budget", usedRange: { address: "Budget!A1:D6", rowCount: 6, columnCount: 4 } },
        ],
      })
    );
    assert.strictEqual(card.name, "list_workbook_meta");
    assert.strictEqual(card.error, false);
    assert.strictEqual(card.summary, "Data 5000×200 · Budget 6×4");
  });

  it("names the address for read_range and notes formulas when they differ", () => {
    const card = summarizeTool(
      "read_range",
      { sheet: "Budget", address: "A1:D6" },
      JSON.stringify({
        address: "Budget!A1:D6",
        values: [[1000]],
        formulas: [["=B2-B3"]],
      })
    );
    assert.strictEqual(card.summary, "Budget!A1:D6 · formulas");
  });

  it("names the written cell for write_range", () => {
    const card = summarizeTool(
      "write_range",
      { sheet: "Budget", address: "D4", values: [["=D2-D3"]] },
      JSON.stringify({ ok: true, sheet: "Budget", address: "D4" })
    );
    assert.strictEqual(card.summary, "Budget!D4");
  });

  it("names the query and match count for find", () => {
    const card = summarizeTool(
      "find",
      { query: "Revenue" },
      JSON.stringify({ query: "Revenue", matches: 3, addresses: ["Budget!A2"], truncated: false })
    );
    assert.strictEqual(card.summary, "“Revenue” · 3 matches");
  });

  it("names the selection address", () => {
    const card = summarizeTool(
      "get_selection",
      {},
      JSON.stringify({ sheet: "Budget", address: "Budget!A1:D6" })
    );
    assert.strictEqual(card.summary, "Budget!A1:D6");
  });

  it("flags tool errors and keeps the error text", () => {
    const card = summarizeTool("read_range", { sheet: "Missing" }, "Worksheet Missing not found", true);
    assert.strictEqual(card.error, true);
    assert.strictEqual(card.summary, "Worksheet Missing not found");
  });
});

describe("toolCardsFromMessages", () => {
  it("returns no cards for a text-only turn", () => {
    const cards = toolCardsFromMessages([
      { role: "user", content: "How big is this workbook?" },
      { role: "assistant", content: "Two sheets." },
    ]);
    assert.deepStrictEqual(cards, []);
  });

  it("pairs tool_use with tool_result and keeps order", () => {
    const cards = toolCardsFromMessages([
      { role: "user", content: "How big is this workbook?" },
      toolUse("call-1", "list_workbook_meta"),
      toolResult("call-1", {
        sheets: [{ name: "Data", usedRange: { rowCount: 5000, columnCount: 200 } }],
      }),
      toolUse("call-2", "write_range", { sheet: "Budget", address: "D4" }),
      toolResult("call-2", { ok: true, sheet: "Budget", address: "D4" }),
      { role: "assistant", content: "Data is 5000 by 200." },
    ]);
    assert.strictEqual(cards.length, 2);
    assert.strictEqual(cards[0].id, "call-1");
    assert.strictEqual(cards[0].name, "list_workbook_meta");
    assert.strictEqual(cards[0].summary, "Data 5000×200");
    assert.strictEqual(cards[1].id, "call-2");
    assert.strictEqual(cards[1].name, "write_range");
    assert.strictEqual(cards[1].summary, "Budget!D4");
  });

  it("ignores messages before fromIndex so prior turns are not replayed", () => {
    const messages: ChatMessage[] = [
      { role: "user", content: "first" },
      toolUse("old", "find", { query: "old" }),
      toolResult("old", { query: "old", matches: 1, addresses: ["A1"] }),
      { role: "assistant", content: "found it" },
      { role: "user", content: "write D4" },
      toolUse("new", "write_range", { sheet: "Budget", address: "D4" }),
      toolResult("new", { ok: true, sheet: "Budget", address: "D4" }),
      { role: "assistant", content: "done" },
    ];
    const cards = toolCardsFromMessages(messages, 4);
    assert.strictEqual(cards.length, 1);
    assert.strictEqual(cards[0].id, "new");
    assert.strictEqual(cards[0].name, "write_range");
  });
});
