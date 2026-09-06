import assert from "assert";
import { parseInline, parseMarkdown } from "../src/app/markdown";

describe("parseInline", () => {
  it("marks bold spans and keeps the surrounding text", () => {
    assert.deepStrictEqual(parseInline("The answer is **Budget!A4** today"), [
      { type: "text", text: "The answer is " },
      { type: "bold", text: "Budget!A4" },
      { type: "text", text: " today" },
    ]);
  });

  it("handles several bold spans in one line", () => {
    assert.deepStrictEqual(parseInline("**Data**, **Budget**"), [
      { type: "bold", text: "Data" },
      { type: "text", text: ", " },
      { type: "bold", text: "Budget" },
    ]);
  });

  it("marks italic and inline code", () => {
    assert.deepStrictEqual(parseInline("set *now* via `=SUM(A1:A2)`"), [
      { type: "text", text: "set " },
      { type: "italic", text: "now" },
      { type: "text", text: " via " },
      { type: "code", text: "=SUM(A1:A2)" },
    ]);
  });

  it("leaves a lone asterisk alone so formulas survive", () => {
    assert.deepStrictEqual(parseInline("=B2*C2 is 5 * 3"), [
      { type: "text", text: "=B2*C2 is 5 * 3" },
    ]);
  });

  it("returns nothing for an empty string", () => {
    assert.deepStrictEqual(parseInline(""), []);
  });
});

describe("parseMarkdown", () => {
  it("reads headings and their level", () => {
    assert.deepStrictEqual(parseMarkdown("### Workbook Overview"), [
      { type: "heading", level: 3, spans: [{ type: "text", text: "Workbook Overview" }] },
    ]);
  });

  it("caps heading level at 3 so the pane keeps its scale", () => {
    const [block] = parseMarkdown("# Title");
    assert.strictEqual(block.type === "heading" && block.level, 3);
  });

  it("groups consecutive bullets into one list", () => {
    assert.deepStrictEqual(parseMarkdown("- one\n- **two**"), [
      {
        type: "list",
        items: [
          [{ type: "text", text: "one" }],
          [{ type: "bold", text: "two" }],
        ],
      },
    ]);
  });

  it("treats a horizontal rule as its own block", () => {
    const blocks = parseMarkdown("before\n\n---\n\nafter");
    assert.deepStrictEqual(
      blocks.map((b) => b.type),
      ["paragraph", "rule", "paragraph"]
    );
  });

  it("keeps blank-line separated paragraphs apart and joins wrapped lines", () => {
    const blocks = parseMarkdown("one\ntwo\n\nthree");
    assert.strictEqual(blocks.length, 2);
    assert.deepStrictEqual(blocks[0], {
      type: "paragraph",
      spans: [{ type: "text", text: "one two" }],
    });
  });

  it("returns no blocks for empty or whitespace-only text", () => {
    assert.deepStrictEqual(parseMarkdown("   \n\n "), []);
  });
});
