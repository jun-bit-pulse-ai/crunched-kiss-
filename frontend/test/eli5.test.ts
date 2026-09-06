import assert from "assert";
import { WELCOME } from "../src/app/demoPrompts";
import {
  eli5InputFromVisible,
  explainLikeFive,
  kidSentenceForTool,
  stripToKidProse,
  TOOL_ELI5,
} from "../src/app/eli5";
import type { VisibleMessage } from "../src/app/types";

describe("kidSentenceForTool", () => {
  it("maps list_workbook_meta to a peek, not a full read", () => {
    assert.strictEqual(
      kidSentenceForTool("list_workbook_meta"),
      "I peeked at the sheet names and how big they are. I did not read every cell."
    );
  });

  it("covers every Excel tool the pane already ships", () => {
    assert.ok(TOOL_ELI5.read_range);
    assert.ok(TOOL_ELI5.write_range);
    assert.ok(TOOL_ELI5.get_selection);
    assert.ok(TOOL_ELI5.find);
    assert.match(kidSentenceForTool("read_range"), /little box/i);
    assert.match(kidSentenceForTool("write_range"), /wrote/i);
    assert.match(kidSentenceForTool("get_selection"), /clicked/i);
    assert.match(kidSentenceForTool("find"), /hunted/i);
  });

  it("falls back on an unknown helper name without inventing a tool", () => {
    assert.strictEqual(kidSentenceForTool("not_a_real_tool"), "I used a helper called not_a_real_tool.");
  });

  it("says the peek failed when the card is an error", () => {
    assert.match(kidSentenceForTool("read_range", true), /did not work/);
  });
});

describe("stripToKidProse", () => {
  it("drops Markdown marks and keeps the first two sentences", () => {
    const prose = stripToKidProse("**Budget** is fine. `D4` was hard-coded. Extra detail stays out.");
    assert.strictEqual(prose, "Budget is fine. D4 was hard-coded.");
  });

  it("caps a long single sentence", () => {
    const long = `${"word ".repeat(80).trim()}.`;
    const prose = stripToKidProse(long, 40);
    assert.ok(prose.length <= 40);
    assert.match(prose, /…$/);
  });
});

describe("eli5InputFromVisible", () => {
  it("uses welcome text when nobody has asked yet", () => {
    const input = eli5InputFromVisible([
      { id: "welcome", kind: "text", role: "assistant", text: WELCOME },
    ]);
    assert.strictEqual(input.assistantText, WELCOME);
    assert.deepStrictEqual(input.tools, []);
  });

  it("keeps tools and the reply from the latest user turn only", () => {
    const messages: VisibleMessage[] = [
      { id: "welcome", kind: "text", role: "assistant", text: WELCOME },
      { id: "u1", kind: "text", role: "user", text: "How big?" },
      { id: "t-old", kind: "tool", name: "find", summary: "old", error: false },
      { id: "a1", kind: "text", role: "assistant", text: "Old answer." },
      { id: "u2", kind: "text", role: "user", text: "Write D4" },
      { id: "t-new", kind: "tool", name: "write_range", summary: "Budget!D4", error: false },
      { id: "a2", kind: "text", role: "assistant", text: "I wrote the formula." },
    ];
    const input = eli5InputFromVisible(messages);
    assert.deepStrictEqual(input.tools, [{ name: "write_range", error: false }]);
    assert.strictEqual(input.assistantText, "I wrote the formula.");
  });
});

describe("explainLikeFive", () => {
  it("explains welcome with a canned kid sentence", () => {
    const text = explainLikeFive({ assistantText: WELCOME, tools: [] });
    assert.match(text, /I am Crunched/);
    assert.doesNotMatch(text, /financial or business models/);
  });

  it("lists kid sentences for tools, then restates the answer", () => {
    const text = explainLikeFive({
      assistantText: "The **Data** sheet is 5000 by 200.",
      tools: [{ name: "list_workbook_meta", error: false }],
    });
    assert.match(text, /Here is what I did with the spreadsheet/);
    assert.match(text, /I peeked at the sheet names/);
    assert.match(text, /Then I said: The Data sheet is 5000 by 200\./);
  });

  it("has a fallback when there is nothing to explain", () => {
    const text = explainLikeFive({ assistantText: null, tools: [] });
    assert.match(text, /have not done anything yet/i);
  });
});
