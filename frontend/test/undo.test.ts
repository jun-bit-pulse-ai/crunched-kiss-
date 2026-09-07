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
  return { sheet: "Sheet1", address, formulas: [[address]] };
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

describe("snapshots carry formulas, not computed values", () => {
  beforeEach(() => clearUndoStack());

  it("round-trips a formula so undo can put it back", () => {
    // The bug this guards: snapshotting `.values` captured the number a formula
    // happened to show, so undo replaced "=B4/B2" with e.g. 0.6 and the model
    // stopped being live.
    pushSnapshot({ sheet: "Budget", address: "Budget!B5:D5", formulas: [["=B4/B2", "=C4/C2", "=D4/D2"]] });
    assert.deepStrictEqual(popSnapshot()?.formulas, [["=B4/B2", "=C4/C2", "=D4/D2"]]);
  });

  it("round-trips literal values unchanged", () => {
    // Excel's .formulas returns the literal where there is no formula, so plain
    // numbers and text must survive the same path.
    pushSnapshot({ sheet: "Budget", address: "Budget!D4", formulas: [[1000]] });
    assert.deepStrictEqual(popSnapshot()?.formulas, [[1000]]);
  });
});
