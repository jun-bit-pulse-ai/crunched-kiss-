import assert from "assert";
import { parseSuggestions } from "../src/app/suggestions";

describe("parseSuggestions", () => {
  it("extracts suggestions from a 💡 line", () => {
    const text = 'The total is 42.\n\n💡 "Show me the formulas" · "Find errors" · "Add a total row"';
    const result = parseSuggestions(text);
    assert.deepStrictEqual(result, {
      text: "The total is 42.",
      suggestions: ["Show me the formulas", "Find errors", "Add a total row"],
    });
  });

  it("returns null when no 💡 line is present", () => {
    assert.strictEqual(parseSuggestions("Just a normal answer."), null);
  });

  it("returns null for an empty 💡 line", () => {
    assert.strictEqual(parseSuggestions("Answer\n\n💡"), null);
  });

  it("handles a single suggestion", () => {
    const text = 'Done.\n\n💡 "Explain this formula"';
    const result = parseSuggestions(text);
    assert.deepStrictEqual(result, {
      text: "Done.",
      suggestions: ["Explain this formula"],
    });
  });

  it("preserves text without altering it when no suggestions", () => {
    const text = "This is just a plain reply.";
    const result = parseSuggestions(text);
    assert.strictEqual(result, null);
  });
});
