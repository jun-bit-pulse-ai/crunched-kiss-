import assert from "assert";
import type { CellValue } from "../src/app/excelPolicy";
import { canUndo, clearUndoStack } from "../src/app/undoStack";
import { readRange, undoLastWrite, writeRange } from "../src/app/services/excel";

type LoadedProps = Set<string>;

type FakeRange = {
  address: string;
  rowCount: number;
  columnCount: number;
  values: CellValue[][];
  formulas: CellValue[][];
  loaded: LoadedProps;
  isNullObject?: boolean;
  load: (props: string | string[]) => void;
  getAbsoluteResizedRange: (rows: number, cols: number) => FakeRange;
  getRow?: (index: number) => FakeRange;
};

function rangeOf(
  address: string,
  values: CellValue[][],
  formulas?: CellValue[][]
): FakeRange {
  const loaded: LoadedProps = new Set();
  const formulasGrid = formulas ?? values;
  const self: FakeRange = {
    address,
    rowCount: values.length,
    columnCount: values[0]?.length ?? 0,
    values,
    formulas: formulasGrid,
    loaded,
    load(props) {
      for (const prop of Array.isArray(props) ? props : [props]) {
        loaded.add(prop);
      }
    },
    getAbsoluteResizedRange(rows, cols) {
      if (rows === self.rowCount && cols === self.columnCount) {
        return self;
      }
      return rangeOf(
        address,
        values.slice(0, rows).map((row) => row.slice(0, cols)),
        formulasGrid.slice(0, rows).map((row) => row.slice(0, cols))
      );
    },
  };
  return self;
}

function installExcel(run: (batch: (context: unknown) => Promise<unknown>) => Promise<unknown>) {
  (globalThis as unknown as { Excel: { run: typeof run } }).Excel = { run };
}

describe("excel.ts Office.js policy", () => {
  beforeEach(() => {
    clearUndoStack();
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  it("rejects a million-cell A1 address before Excel.run", async () => {
    let ran = false;
    installExcel(async () => {
      ran = true;
      return {};
    });
    await assert.rejects(() => readRange("Data", "A1:XFD1048576"), /at most/);
    assert.strictEqual(ran, false);
  });

  it("rejects write_range without a sheet name before Excel.run", async () => {
    let ran = false;
    installExcel(async () => {
      ran = true;
      return {};
    });
    await assert.rejects(() => writeRange("  ", "A1", [["x"]]), /sheet name/);
    assert.strictEqual(ran, false);
  });

  it("clamps a huge Office-reported range before loading values", async () => {
    const huge: FakeRange = {
      address: "Data!A1",
      rowCount: 1_000_000,
      columnCount: 5,
      get values(): CellValue[][] {
        throw new Error("values loaded on the unclamped range");
      },
      get formulas(): CellValue[][] {
        throw new Error("formulas loaded on the unclamped range");
      },
      loaded: new Set(),
      load(props) {
        for (const prop of Array.isArray(props) ? props : [props]) {
          if (prop === "values" || prop === "formulas") {
            throw new Error(`loaded ${prop} on the unclamped range`);
          }
          this.loaded.add(prop);
        }
      },
      getAbsoluteResizedRange(rows, cols) {
        const grid = Array.from({ length: rows }, (_, row) =>
          Array.from({ length: cols }, (__, col) => row * cols + col)
        );
        return rangeOf("Data!A1:E400", grid);
      },
    };

    installExcel(async (batch) => {
      const context = {
        workbook: {
          worksheets: {
            getItem(name: string) {
              assert.strictEqual(name, "Data");
              return { getRange: () => huge };
            },
          },
        },
        sync: async () => undefined,
      };
      return batch(context);
    });

    const result = await readRange("Data", "A1");
    assert.strictEqual(result.truncated, true);
    assert.strictEqual(result.total_rows, 1_000_000);
    assert.strictEqual(result.total_cols, 5);
    assert.ok(result.values.length * 5 <= 2000);
    assert.ok(!huge.loaded.has("values"));
  });

  it("snapshots the prior formulas and writes the new ones in one Excel.run", async () => {
    const target = rangeOf("Budget!D4", [["1000"]]);
    let syncs = 0;

    installExcel(async (batch) => {
      const context = {
        workbook: {
          worksheets: {
            getItem(name: string) {
              assert.strictEqual(name, "Budget");
              return {
                getRange(address: string) {
                  assert.strictEqual(address, "D4");
                  return target;
                },
              };
            },
          },
        },
        sync: async () => {
          syncs += 1;
        },
      };
      return batch(context);
    });

    const result = await writeRange("Budget", "D4", [["=D2-D3"]]);
    assert.deepStrictEqual(result, { ok: true, sheet: "Budget", address: "Budget!D4" });
    assert.deepStrictEqual(target.formulas, [["=D2-D3"]]);
    assert.strictEqual(canUndo(), true);
    assert.strictEqual(syncs, 2);
  });

  it("does not push an undo snapshot when the first sync fails", async () => {
    const target = rangeOf("Budget!D4", [["1000"]]);
    installExcel(async (batch) => {
      const context = {
        workbook: {
          worksheets: {
            getItem() {
              return { getRange: () => target };
            },
          },
        },
        sync: async () => {
          throw new Error("Office.js bridge down");
        },
      };
      return batch(context);
    });

    await assert.rejects(() => writeRange("Budget", "D4", [["=D2-D3"]]), /bridge down/);
    assert.strictEqual(canUndo(), false);
    assert.deepStrictEqual(target.formulas, [["1000"]]);
  });

  it("undo restores the snapshotted formulas through Office.js", async () => {
    const cells = new Map<string, FakeRange>([["Budget!D4", rangeOf("Budget!D4", [["1000"]])]]);

    installExcel(async (batch) => {
      const context = {
        workbook: {
          worksheets: {
            getItem(name: string) {
              return {
                getRange(address: string) {
                  const key = address.includes("!") ? address : `${name}!${address}`;
                  const existing = cells.get(key);
                  if (existing) {
                    return existing;
                  }
                  const created = rangeOf(key, [[""]]);
                  cells.set(key, created);
                  return created;
                },
              };
            },
          },
        },
        sync: async () => undefined,
      };
      return batch(context);
    });

    await writeRange("Budget", "D4", [["=D2-D3"]]);
    assert.deepStrictEqual(cells.get("Budget!D4")?.formulas, [["=D2-D3"]]);

    const undone = await undoLastWrite();
    assert.deepStrictEqual(undone, { sheet: "Budget", address: "Budget!D4" });
    // The cell comes back as what it was, not as the number it was showing.
    assert.deepStrictEqual(cells.get("Budget!D4")?.formulas, [["1000"]]);
    assert.strictEqual(canUndo(), false);
  });
});
