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
        ordered: false,
        start: 1,
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

describe("parseMarkdown tables", () => {
  const table = [
    "| Sheet | Used Range | Rows |",
    "|-------|-----------|------|",
    "| **Data** | A1:GR5000 | 5,000 |",
    "| **Budget** | A1:D6 | 6 |",
  ].join("\n");

  it("reads the header row and the body rows", () => {
    const [block] = parseMarkdown(table);
    assert.strictEqual(block.type, "table");
    if (block.type !== "table") return;
    assert.deepStrictEqual(
      block.header.map((cell) => cell.map((span) => span.text)),
      [["Sheet"], ["Used Range"], ["Rows"]]
    );
    assert.strictEqual(block.rows.length, 2);
    assert.deepStrictEqual(block.rows[0][0], [{ type: "bold", text: "Data" }]);
    assert.deepStrictEqual(block.rows[1][2], [{ type: "text", text: "6" }]);
  });

  it("does not leave the delimiter row as visible text", () => {
    const blocks = parseMarkdown(table);
    assert.strictEqual(blocks.length, 1);
  });

  it("pads a short row so cells stay under the right heading", () => {
    const [block] = parseMarkdown("| A | B |\n|---|---|\n| only |");
    assert.strictEqual(block.type === "table" && block.rows[0].length, 2);
  });

  it("treats a pipe line without a delimiter row as an ordinary paragraph", () => {
    const [block] = parseMarkdown("| not | a table |");
    assert.strictEqual(block.type, "paragraph");
  });

  it("ends the table when normal text resumes", () => {
    const blocks = parseMarkdown(`${table}\nAfter the table.`);
    assert.deepStrictEqual(
      blocks.map((b) => b.type),
      ["table", "paragraph"]
    );
  });
});

describe("parseMarkdown ordered lists", () => {
  it("keeps the numbers the model chose", () => {
    const [block] = parseMarkdown("1. first\n2. **second**");
    assert.strictEqual(block.type, "list");
    if (block.type !== "list") return;
    assert.strictEqual(block.ordered, true);
    assert.strictEqual(block.start, 1);
    assert.deepStrictEqual(block.items[1], [{ type: "bold", text: "second" }]);
  });

  it("keeps bullet lists unordered", () => {
    const [block] = parseMarkdown("- one\n- two");
    assert.strictEqual(block.type === "list" && block.ordered, false);
  });

  it("does not treat a year or a price as a list", () => {
    const [block] = parseMarkdown("2024. was a strong year");
    assert.strictEqual(block.type, "paragraph");
  });
});

/**
 * Spreadsheet prose collides with Markdown syntax. These cases all came from an
 * adversarial pass over real Crunched replies and every one of them regressed
 * the pane before the fixes in markdown.ts.
 */
describe("parseMarkdown against spreadsheet prose", () => {
  it("keeps bold after a multiplication formula on the same line", () => {
    const [block] = parseMarkdown("Cell D4 is hard-coded; =B2*C2 is what it should be, so D4 is **wrong**.");
    assert.strictEqual(block.type, "paragraph");
    const spans = block.type === "paragraph" ? block.spans : [];
    assert.deepStrictEqual(
      spans.filter((span) => span.type === "bold"),
      [{ type: "bold", text: "wrong" }]
    );
  });

  it("renders every bold in a sentence that also contains a formula", () => {
    const [block] = parseMarkdown(
      "**D4 (Gross profit, 2026)** is hard-coded instead of =B2*C2, unlike **B4** and **C4**."
    );
    const spans = block.type === "paragraph" ? block.spans : [];
    assert.deepStrictEqual(
      spans.filter((span) => span.type === "bold").map((span) => span.text),
      ["D4 (Gross profit, 2026)", "B4", "C4"]
    );
  });

  it("keeps an inline code span after a formula star", () => {
    const [block] = parseMarkdown("Use =B2*C2 with `=SUM(A1:A10)*2` to scale.");
    const spans = block.type === "paragraph" ? block.spans : [];
    assert.deepStrictEqual(
      spans.filter((span) => span.type === "code"),
      [{ type: "code", text: "=SUM(A1:A10)*2" }]
    );
  });

  it("still refuses to italicise a star that does not open a word", () => {
    const [block] = parseMarkdown("Costs rose 5 * 3 percent and =B2*C2 holds.");
    const spans = block.type === "paragraph" ? block.spans : [];
    assert.strictEqual(spans.some((span) => span.type === "italic"), false);
  });

  it("absorbs a blank line between table row groups instead of leaking raw pipes", () => {
    const blocks = parseMarkdown("| A | B |\n|---|---|\n| 1 | 2 |\n\n| 3 | 4 |");
    assert.strictEqual(blocks.length, 1);
    assert.strictEqual(blocks[0].type === "table" && blocks[0].rows.length, 2);
  });

  it("ends the table when the blank line is followed by prose", () => {
    const blocks = parseMarkdown("| A |\n|---|\n| 1 |\n\nProse after.");
    assert.strictEqual(blocks.length, 2);
    assert.strictEqual(blocks[1].type, "paragraph");
  });

  it("treats '# of rows' as prose and keeps the # that means 'number of'", () => {
    const [block] = parseMarkdown("# of rows: 5,000");
    assert.strictEqual(block.type, "paragraph");
    const spans = block.type === "paragraph" ? block.spans : [];
    assert.ok(spans.some((span) => span.type === "text" && span.text.startsWith("# of rows")));
  });

  it("still reads a capitalised # line as a heading", () => {
    const [block] = parseMarkdown("# Summary");
    assert.strictEqual(block.type, "heading");
  });

  it("keeps a list whole when a bullet marker arrives with no content", () => {
    const blocks = parseMarkdown("- \n- B4 uses =B2-B3");
    assert.strictEqual(blocks.length, 1);
    assert.strictEqual(blocks[0].type === "list" && blocks[0].items.length, 1);
  });

  it("parses a star-heavy line quickly and without dropping the trailing bold", () => {
    const source = "=A1*B1 ".repeat(600) + "**end**";
    const started = Date.now();
    const [block] = parseMarkdown(source);
    assert.ok(Date.now() - started < 1000, "parseInline must not degrade to a pathological scan");
    const spans = block.type === "paragraph" ? block.spans : [];
    assert.deepStrictEqual(spans.slice(-1), [{ type: "bold", text: "end" }]);
  });
});
