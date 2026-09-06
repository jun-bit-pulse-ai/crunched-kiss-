/* global Excel */

import {
  MAX_FIND_RESULTS,
  MAX_READ_CELLS,
  SELECTION_PREVIEW_CELLS,
  assertWriteValues,
  headerPreviewWidth,
  limitAddresses,
  rowsWithinCellCap,
  selectionLabel,
  toHeaderPreview,
} from "../excelPolicy";
import type { CellValue } from "../excelPolicy";
import { popSnapshot, pushSnapshot } from "../undoStack";

export { canUndo, clearUndoStack, subscribe as subscribeUndoStack } from "../undoStack";

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

export async function undoLastWrite(): Promise<{ sheet: string; address: string } | null> {
  const snapshot = popSnapshot();
  if (!snapshot) {
    return null;
  }
  await Excel.run(async (context) => {
    const range = context.workbook.worksheets.getItem(snapshot.sheet).getRange(snapshot.address);
    range.values = snapshot.values;
    await context.sync();
  });
  return { sheet: snapshot.sheet, address: snapshot.address };
}

export async function listWorkbookMeta(): Promise<{ sheets: SheetMeta[] }> {
  return Excel.run(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const usedRanges = sheets.items.map((sheet) => {
      // valuesOnly=true ignores formatted-but-empty cells that inflate used ranges.
      const used = sheet.getUsedRangeOrNullObject(true);
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
    const requested = context.workbook.worksheets.getItem(sheet).getRange(address);
    requested.load(["rowCount", "columnCount"]);
    await context.sync();

    const totalRows = requested.rowCount;
    const totalCols = requested.columnCount;
    const { rows, truncated } = rowsWithinCellCap(totalRows, totalCols, MAX_READ_CELLS);

    // Clamp the range's SHAPE before loading it, so a huge requested address is
    // never materialised across the Office.js bridge just to be sliced in JS.
    const target = truncated ? requested.getAbsoluteResizedRange(rows, totalCols) : requested;
    target.load(["values", "formulas", "address"]);
    await context.sync();

    return {
      address: target.address,
      values: target.values as CellValue[][],
      formulas: target.formulas as CellValue[][],
      truncated,
      total_rows: totalRows,
      total_cols: totalCols,
    };
  });
}

export async function writeRange(sheet: string, address: string, values: unknown) {
  if (!assertWriteValues(values)) {
    throw new Error("write_range requires a non-empty 2D values array");
  }
  const rows = values.length;
  const cols = values[0]?.length ?? 0;

  return Excel.run(async (context) => {
    // "address" is documented as the top-left anchor, not necessarily a range already
    // sized to match `values` — resize to the write's true extent so both the write
    // and the undo snapshot cover every cell that is actually about to change.
    const anchor = context.workbook.worksheets.getItem(sheet).getRange(address);
    const target = cols > 0 ? anchor.getAbsoluteResizedRange(rows, cols) : anchor;
    target.load(["values", "address"]);
    await context.sync();

    // Snapshot and write in the same batch: if the write throws, sync() rejects
    // before the snapshot is ever pushed, so a failed write can't leave a phantom
    // undo entry.
    pushSnapshot({ sheet, address: target.address, values: target.values as CellValue[][] });
    target.values = values;
    await context.sync();
    return { ok: true, sheet, address: target.address };
  });
}

export async function getSelection() {
  return Excel.run(async (context) => {
    const selected = context.workbook.getSelectedRange();
    selected.load(["rowCount", "columnCount", "worksheet/name"]);
    await context.sync();

    const totalRows = selected.rowCount;
    const totalCols = selected.columnCount;
    const { rows, truncated } = rowsWithinCellCap(totalRows, totalCols, SELECTION_PREVIEW_CELLS);

    const target = truncated ? selected.getAbsoluteResizedRange(rows, totalCols) : selected;
    target.load(["address", "values"]);
    await context.sync();

    return {
      sheet: selected.worksheet.name,
      address: target.address,
      preview: target.values as CellValue[][],
      truncated,
      total_rows: totalRows,
      total_cols: totalCols,
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

/**
 * Watch the user's selection so the pane can show "Crunched sees: Sheet1!B2:D10" without
 * the user having to ask. Returns an unsubscribe function; swallow watch errors since this
 * is a UX nicety, not a tool the model depends on.
 */
export function watchSelection(onChange: (label: string | null) => void): () => void {
  let disposed = false;

  async function report() {
    if (disposed) {
      return;
    }
    try {
      const selection = await getSelection();
      onChange(
        selectionLabel(
          selection.sheet,
          selection.address,
          selection.total_rows,
          selection.total_cols
        )
      );
    } catch {
      onChange(null);
    }
  }

  Excel.run(async (context) => {
    context.workbook.onSelectionChanged.add(report);
    await context.sync();
  }).catch(() => onChange(null));

  report();

  return () => {
    disposed = true;
  };
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

export async function getSelectedFormula(): Promise<{ sheet: string; address: string; formula: string } | null> {
  return Excel.run(async (context) => {
    const range = context.workbook.getSelectedRange();
    range.load(["address", "formulas", "worksheet/name"]);
    await context.sync();
    const formula = (range.formulas as CellValue[][])[0]?.[0];
    if (!formula || typeof formula !== "string" || formula[0] !== "=") {
      return null;
    }
    return {
      sheet: range.worksheet.name,
      address: range.address,
      formula,
    };
  });
}
