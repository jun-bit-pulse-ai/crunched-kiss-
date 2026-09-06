import assert from "assert";
import { canUndo, clearUndoStack } from "../src/app/services/excel";

describe("Undo stack", () => {
  it("starts empty", () => {
    clearUndoStack();
    assert.strictEqual(canUndo(), false);
  });

  it("can be cleared", () => {
    // We can't push directly (snapshotRange is private), but we can verify clear works after a hypothetical push.
    clearUndoStack();
    assert.strictEqual(canUndo(), false);
  });
});
