import assert from "assert";
import { getSelectedFormula } from "../src/app/services/excel";

describe("getSelectedFormula", () => {
  it("is defined", () => {
    assert.strictEqual(typeof getSelectedFormula, "function");
  });
});
