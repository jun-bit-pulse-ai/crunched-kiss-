import assert from "assert";
import { parseOptions } from "../src/app/components/ClarifyingQuestion";

describe("parseOptions", () => {
  it("extracts the question and lettered options", () => {
    const text = "Which sheet?\n\nA) Budget\nB) Data\nC) New sheet";
    const result = parseOptions(text);
    assert.deepStrictEqual(result, {
      question: "Which sheet?",
      options: [
        { letter: "A", text: "Budget" },
        { letter: "B", text: "Data" },
        { letter: "C", text: "New sheet" },
      ],
    });
  });

  it("returns null for plain text without options", () => {
    assert.strictEqual(parseOptions("This is just a normal answer."), null);
  });

  it("returns null for a single option (needs at least two)", () => {
    assert.strictEqual(parseOptions("Pick one:\n\nA) Only option"), null);
  });

  it("caps at four options since the regex only matches A-D", () => {
    const text = "Which?\n\nA) One\nB) Two\nC) Three\nD) Four\nE) Five";
    const result = parseOptions(text);
    assert.strictEqual(result?.options.length, 4);
  });

  it("matches the blockquote format from the system prompt", () => {
    const text = [
      "> **Which sheet would you like to work with?**",
      ">",
      "> A) Budget — the 7-row financial model",
      "> B) Data — the 5,000-row metrics table",
      "> C) A new sheet",
    ].join("\n");
    const result = parseOptions(text);
    assert.deepStrictEqual(result, {
      question: "**Which sheet would you like to work with?**",
      options: [
        { letter: "A", text: "Budget — the 7-row financial model" },
        { letter: "B", text: "Data — the 5,000-row metrics table" },
        { letter: "C", text: "A new sheet" },
      ],
    });
  });

  it("does not treat prose like 'I checked B) and…' as options", () => {
    assert.strictEqual(parseOptions("I checked B) and it looked fine."), null);
  });
});
