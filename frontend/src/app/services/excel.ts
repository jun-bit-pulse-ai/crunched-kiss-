/* global Excel */

import {
  MAX_READ_CELLS,
  SELECTION_PREVIEW_CELLS,
  assertWriteValues,
  sliceValuesToCellCap,
  type CellValue,
} from "../excelPolicy";

export type SheetMeta = {
  name: string;
  usedRange: {
    address: string;
    rowCount: number;
    columnCount: number;
  } | null;
};

export async function listWorkbookMeta(): Promise<{ sheets: SheetMeta[] }> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const usedRanges = sheets.items.map((sheet) => {
      const used = sheet.getUsedRangeOrNullObject();
      used.load(["address", "rowCount", "columnCount"]);
      return used;
    });
    await context.sync();

    return {
      sheets: sheets.items.map((sheet, index) => {
        const used = usedRanges[index];
        return {
          name: sheet.name,
          usedRange: used.isNullObject
            ? null
            : {
                address: used.address,
                rowCount: used.rowCount,
                columnCount: used.columnCount,
              },
        };
      }),
    };
  });
}

export async function readRange(sheet: string, address: string) {
  return Excel.run(async (context) => {
    const range = context.workbook.worksheets.getItem(sheet).getRange(address);
    range.load(["values", "formulas", "address", "rowCount", "columnCount"]);
    await context.sync();
    const values = sliceValuesToCellCap(range.values as CellValue[][], MAX_READ_CELLS);
    const formulas = sliceValuesToCellCap(range.formulas as CellValue[][], MAX_READ_CELLS);
    return {
      address: range.address,
      values: values.values,
      formulas: formulas.values,
      truncated: values.truncated,
      total_rows: values.totalRows,
      total_cols: values.totalCols,
    };
  });
}

export async function writeRange(sheet: string, address: string, values: unknown) {
  if (!assertWriteValues(values)) {
    throw new Error("write_range requires a non-empty 2D values array");
  }
  return Excel.run(async (context) => {
    const range = context.workbook.worksheets.getItem(sheet).getRange(address);
    range.values = values;
    await context.sync();
    return { ok: true, sheet, address };
  });
}

export async function getSelection() {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["address", "values", "worksheet/name", "rowCount", "columnCount"]);
    await context.sync();
    const preview = sliceValuesToCellCap(range.values as CellValue[][], SELECTION_PREVIEW_CELLS);
    return {
      sheet: range.worksheet.name,
      address: range.address,
      preview: preview.values,
      truncated: preview.truncated,
      total_rows: preview.totalRows,
      total_cols: preview.totalCols,
    };
  });
}

export async function dispatchExcelTool(
  name: string,
  input: Record<string, unknown>
): Promise<unknown> {
  switch (name) {
    case "list_workbook_meta":
      return listWorkbookMeta();
    case "read_range":
      return readRange(String(input.sheet ?? ""), String(input.address ?? ""));
    case "write_range":
      return writeRange(String(input.sheet ?? ""), String(input.address ?? ""), input.values);
    case "get_selection":
      return getSelection();
    default:
      throw new Error(`Unknown Excel tool: ${name}`);
  }
}
