import assert from "assert";
import { buildExplainFormulaPrompt } from "../src/app/formulaExplainer";

describe("buildExplainFormulaPrompt", () => {
  it("asks for a plain-English, step-by-step breakdown of the given formula", () => {
    const prompt = buildExplainFormulaPrompt("=SUM(B2:B10)");
    assert.match(prompt, /plain English/);
    assert.match(prompt, /step by step/);
    assert.ok(prompt.includes("=SUM(B2:B10)"));
  });

  it("passes the formula through verbatim, even a complex nested one", () => {
    const formula = '=IF(VLOOKUP(A1,Data!A:B,2,FALSE)>100,"High","Low")';
    const prompt = buildExplainFormulaPrompt(formula);
    assert.ok(prompt.includes(formula));
  });
});
