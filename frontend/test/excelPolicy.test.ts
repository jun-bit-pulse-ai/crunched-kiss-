import assert from "assert";
import {
  HEADER_PREVIEW_COLS,
  MAX_FIND_RESULTS,
  MAX_READ_CELLS,
  assertWriteValues,
  countCells,
  headerPreviewWidth,
  limitAddresses,
  sliceValuesToCellCap,
  toHeaderPreview,
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

describe("headerPreviewWidth", () => {
  it("caps the preview at HEADER_PREVIEW_COLS", () => {
    assert.strictEqual(headerPreviewWidth(200), HEADER_PREVIEW_COLS);
  });

  it("uses the sheet width when it is narrower than the cap", () => {
    assert.strictEqual(headerPreviewWidth(4), 4);
  });

  it("never asks Excel for zero or negative columns", () => {
    assert.strictEqual(headerPreviewWidth(0), 0);
    assert.strictEqual(headerPreviewWidth(-3), 0);
  });
});

describe("toHeaderPreview", () => {
  it("stringifies cells and drops trailing blanks", () => {
    assert.deepStrictEqual(toHeaderPreview(["Line item", 2024, null, ""]), [
      "Line item",
      "2024",
    ]);
  });

  it("keeps blanks that sit between real headers", () => {
    assert.deepStrictEqual(toHeaderPreview(["A", null, "C"]), ["A", "", "C"]);
  });

  it("returns an empty list for a blank or missing row", () => {
    assert.deepStrictEqual(toHeaderPreview([null, ""]), []);
    assert.deepStrictEqual(toHeaderPreview(undefined), []);
  });
});

describe("limitAddresses", () => {
  it("passes short lists through untouched", () => {
    const found = ["Budget!A1", "Budget!A9"];
    const result = limitAddresses(found, 12);
    assert.deepStrictEqual(result.addresses, found);
    assert.strictEqual(result.truncated, false);
    assert.strictEqual(result.total, 2);
  });

  it("caps long lists at MAX_FIND_RESULTS and reports the true total", () => {
    const found = Array.from({ length: 900 }, (_, i) => `Data!A${i + 1}`);
    const result = limitAddresses(found);
    assert.strictEqual(result.addresses.length, MAX_FIND_RESULTS);
    assert.strictEqual(result.truncated, true);
    assert.strictEqual(result.total, 900);
    assert.strictEqual(result.addresses[0], "Data!A1");
  });
});
