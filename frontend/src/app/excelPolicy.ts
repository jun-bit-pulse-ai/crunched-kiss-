export const MAX_READ_CELLS = 2000;
export const SELECTION_PREVIEW_CELLS = 50;

export type CellValue = string | number | boolean | null;

export function countCells(rowCount: number, colCount: number): number {
  return Math.max(0, rowCount) * Math.max(0, colCount);
}

/**
 * How many rows of a `totalCols`-wide range fit under `maxCells`. Used to clamp a
 * range's SHAPE before asking Office.js to load it, so a huge range is never
 * materialised just to be sliced afterward in JS.
 */
export function rowsWithinCellCap(
  totalRows: number,
  totalCols: number,
  maxCells: number = MAX_READ_CELLS
): { rows: number; truncated: boolean } {
  const total = countCells(totalRows, totalCols);
  if (total <= maxCells) {
    return { rows: Math.max(0, totalRows), truncated: false };
  }
  const cols = Math.max(totalCols, 1);
  const rows = Math.max(1, Math.floor(maxCells / cols));
  return { rows, truncated: true };
}

export function sliceValuesToCellCap<T>(
  values: T[][],
  maxCells: number = MAX_READ_CELLS
): { values: T[][]; truncated: boolean; totalRows: number; totalCols: number } {
  const totalRows = values.length;
  const totalCols = values[0]?.length ?? 0;
  const { rows, truncated } = rowsWithinCellCap(totalRows, totalCols, maxCells);
  return {
    values: truncated ? values.slice(0, rows) : values,
    truncated,
    totalRows,
    totalCols,
  };
}

export function assertWriteValues(values: unknown): values is CellValue[][] {
  return Array.isArray(values) && values.length > 0 && values.every((row) => Array.isArray(row));
}

export const MAX_FIND_RESULTS = 50;
export const HEADER_PREVIEW_COLS = 20;

/** How many columns of a sheet's first row are worth showing as headers. */
export function headerPreviewWidth(
  columnCount: number,
  cap: number = HEADER_PREVIEW_COLS
): number {
  return Math.max(0, Math.min(columnCount, cap));
}

/** First row to header strings, with trailing blanks dropped so wide sheets stay short. */
export function toHeaderPreview(row: CellValue[] | undefined): string[] {
  if (!row) {
    return [];
  }
  const cells = row.map((cell) => (cell === null || cell === undefined ? "" : String(cell)));
  let end = cells.length;
  while (end > 0 && cells[end - 1] === "") {
    end -= 1;
  }
  return cells.slice(0, end);
}

/** Cap a find result while still reporting how many matches really exist. */
export function limitAddresses(
  addresses: string[],
  max: number = MAX_FIND_RESULTS
): { addresses: string[]; truncated: boolean; total: number } {
  return {
    addresses: addresses.slice(0, max),
    truncated: addresses.length > max,
    total: addresses.length,
  };
}

/**
 * Label for the "Crunched can see this" pill.
 *
 * Office.js already returns a sheet-qualified address, so prefixing the sheet
 * name again produced "Data!Data!L1:N6". Add the sheet only when it is missing,
 * and say the size in words, since a bare "6×3" reads like a cell reference.
 */
export function selectionLabel(
  sheet: string,
  address: string,
  rowCount: number,
  columnCount: number
): string | null {
  if (!address) {
    return null;
  }
  const qualified = address.includes("!") ? address : sheet ? `${sheet}!${address}` : address;
  if (rowCount <= 1 && columnCount <= 1) {
    return qualified;
  }
  const rows = `${rowCount} ${rowCount === 1 ? "row" : "rows"}`;
  const columns = `${columnCount} ${columnCount === 1 ? "column" : "columns"}`;
  return `${qualified} · ${rows} × ${columns}`;
}
