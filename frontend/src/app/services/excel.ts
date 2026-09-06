/* global Excel */

import {
  MAX_FIND_RESULTS,
  MAX_READ_CELLS,
  SELECTION_PREVIEW_CELLS,
  assertWriteValues,
  headerPreviewWidth,
  limitAddresses,
  sliceValuesToCellCap,
  toHeaderPreview,
} from "../excelPolicy";
import type { CellValue } from "../excelPolicy";

export type SheetMeta = {
  name: string;
  usedRange: {
    address: string;
    rowCount: number;
    columnCount: number;
  } | null;
  /** First row of the used range, capped at HEADER_PREVIEW_COLS columns. */
  headerPreview: string[];
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

    // Load only the first row, capped in width, so a 200-column sheet stays cheap.
    const headerRows = usedRanges.map((used) => {
      if (used.isNullObject) {
        return null;
      }
      const width = headerPreviewWidth(used.columnCount);
      if (width === 0) {
        return null;
      }
      const header = used.getRow(0).getAbsoluteResizedRange(1, width);
      header.load("values");
      return header;
    });
    await context.sync();

    return {
      sheets: sheets.items.map((sheet, index) => {
        const used = usedRanges[index];
        const header = headerRows[index];
        return {
          name: sheet.name,
          usedRange: used.isNullObject
            ? null
            : {
                address: used.address,
                rowCount: used.rowCount,
                columnCount: used.columnCount,
              },
          headerPreview: header
            ? toHeaderPreview((header.values as CellValue[][])[0])
            : [],
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

export type FindResult = {
  query: string;
  matches: number;
  addresses: string[];
  truncated: boolean;
  searched: string;
};

export async function find(
  query: string,
  sheet?: string,
  matchCase = false,
  completeMatch = false
): Promise<FindResult> {
  if (!query) {
    throw new Error("find requires a non-empty query");
  }
  return Excel.run(async (context) => {
    const worksheets = context.workbook.worksheets;
    worksheets.load("items/name");
    await context.sync();

    const targets = sheet ? [worksheets.getItem(sheet)] : worksheets.items;
    const hits = targets.map((worksheet) => {
      const areas = worksheet.findAllOrNullObject(query, { completeMatch, matchCase });
      areas.load(["cellCount", "areas/items/address"]);
      return areas;
    });
    await context.sync();

    let matches = 0;
    const found: string[] = [];
    for (const areas of hits) {
      if (areas.isNullObject) {
        continue;
      }
      matches += areas.cellCount;
      for (const area of areas.areas.items) {
        found.push(area.address);
      }
    }

    const limited = limitAddresses(found, MAX_FIND_RESULTS);
    return {
      query,
      matches,
      addresses: limited.addresses,
      truncated: limited.truncated,
      searched: sheet ?? "all sheets",
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
    case "find":
      return find(
        String(input.query ?? ""),
        input.sheet === undefined ? undefined : String(input.sheet),
        Boolean(input.match_case),
        Boolean(input.complete_match)
      );
    default:
      throw new Error(`Unknown Excel tool: ${name}`);
  }
}
