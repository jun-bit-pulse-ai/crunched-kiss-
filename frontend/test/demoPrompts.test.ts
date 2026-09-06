import assert from "assert";
import { DEMO_CHIPS, initialVisible, showPromptChips } from "../src/app/demoPrompts";
import type { VisibleMessage } from "../src/app/types";

describe("DEMO_CHIPS", () => {
  it("matches the 15-minute demo script, not a missing Revenue model", () => {
    assert.deepStrictEqual([...DEMO_CHIPS], [
      "How big is this workbook?",
      "Read Budget!A1:D6 with formulas and list any errors",
      "Fix the hard-coded Gross profit in Budget!D4",
    ]);
  });
});

describe("showPromptChips", () => {
  it("shows chips on the welcome-only thread", () => {
    assert.strictEqual(showPromptChips(initialVisible()), true);
  });

  it("hides chips after the user sends something", () => {
    const afterSend: VisibleMessage[] = [
      ...initialVisible(),
      { id: "u1", kind: "text", role: "user", text: "How big is this workbook?" },
    ];
    assert.strictEqual(showPromptChips(afterSend), false);
  });
});

describe("initialVisible", () => {
  it("is a single assistant welcome bubble", () => {
    const welcome = initialVisible();
    assert.strictEqual(welcome.length, 1);
    assert.strictEqual(welcome[0].kind, "text");
    if (welcome[0].kind === "text") {
      assert.strictEqual(welcome[0].role, "assistant");
      assert.match(welcome[0].text, /Crunched/);
    }
  });
});
