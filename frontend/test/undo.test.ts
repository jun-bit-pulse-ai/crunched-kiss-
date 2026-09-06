import assert from "assert";
import {
  MAX_UNDO_DEPTH,
  canUndo,
  clearUndoStack,
  popSnapshot,
  pushSnapshot,
  undoDepth,
} from "../src/app/undoStack";

function snapshot(address: string) {
  return { sheet: "Sheet1", address, values: [[address]] };
}

describe("undoStack", () => {
  beforeEach(() => clearUndoStack());

  it("starts empty", () => {
    assert.strictEqual(canUndo(), false);
    assert.strictEqual(undoDepth(), 0);
  });

  it("can push and pop a snapshot in LIFO order", () => {
    pushSnapshot(snapshot("A1"));
    pushSnapshot(snapshot("B2"));
    assert.strictEqual(canUndo(), true);
    assert.deepStrictEqual(popSnapshot(), snapshot("B2"));
    assert.deepStrictEqual(popSnapshot(), snapshot("A1"));
    assert.strictEqual(canUndo(), false);
  });

  it("popping an empty stack returns undefined rather than throwing", () => {
    assert.strictEqual(popSnapshot(), undefined);
  });

  it("evicts the oldest snapshot once past MAX_UNDO_DEPTH", () => {
    for (let i = 0; i < MAX_UNDO_DEPTH + 3; i += 1) {
      pushSnapshot(snapshot(`A${i}`));
    }
    assert.strictEqual(undoDepth(), MAX_UNDO_DEPTH);
    // The three oldest (A0, A1, A2) should have been evicted; the most recent
    // pop should be the very last one pushed.
    assert.deepStrictEqual(popSnapshot(), snapshot(`A${MAX_UNDO_DEPTH + 2}`));
  });

  it("can be cleared", () => {
    pushSnapshot(snapshot("A1"));
    clearUndoStack();
    assert.strictEqual(canUndo(), false);
    assert.strictEqual(undoDepth(), 0);
  });
});
