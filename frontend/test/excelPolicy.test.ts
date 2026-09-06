import assert from "assert";
import {
  MAX_READ_CELLS,
  assertWriteValues,
  countCells,
  sliceValuesToCellCap,
} from "../src/app/excelPolicy";

describe("countCells", () => {
  it("multiplies rows by columns", () => {
    assert.strictEqual(countCells(10, 20), 200);
  });

  it("treats negative dimensions as empty", () => {
    assert.strictEqual(countCells(-1, 5), 0);
  });
});

describe("sliceValuesToCellCap", () => {
  it("returns the full grid when under the cap", () => {
    const values = [
      [1, 2],
      [3, 4],
    ];
    const result = sliceValuesToCellCap(values, MAX_READ_CELLS);
    assert.strictEqual(result.truncated, false);
    assert.deepStrictEqual(result.values, values);
    assert.strictEqual(result.totalRows, 2);
    assert.strictEqual(result.totalCols, 2);
  });

  it("keeps complete rows and flags truncation past 2000 cells", () => {
    const values = Array.from({ length: 500 }, (_, row) => [row, row, row, row, row]);
    const result = sliceValuesToCellCap(values, MAX_READ_CELLS);
    assert.strictEqual(result.truncated, true);
    assert.strictEqual(result.totalRows, 500);
    assert.strictEqual(result.totalCols, 5);
    assert.ok(result.values.length * 5 <= MAX_READ_CELLS);
    assert.strictEqual(result.values.length, Math.floor(MAX_READ_CELLS / 5));
  });
});

describe("assertWriteValues", () => {
  it("accepts a non-empty 2D array", () => {
    assert.strictEqual(assertWriteValues([["Hello"]]), true);
  });

  it("rejects a flat array or empty grid", () => {
    assert.strictEqual(assertWriteValues(["Hello"]), false);
    assert.strictEqual(assertWriteValues([]), false);
    assert.strictEqual(assertWriteValues(null), false);
  });
});
