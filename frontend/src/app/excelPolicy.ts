export const MAX_READ_CELLS = 2000;
export const SELECTION_PREVIEW_CELLS = 50;

export type CellValue = string | number | boolean | null;

export function countCells(rowCount: number, colCount: number): number {
  return Math.max(0, rowCount) * Math.max(0, colCount);
}

export function sliceValuesToCellCap<T>(
  values: T[][],
  maxCells: number = MAX_READ_CELLS
): { values: T[][]; truncated: boolean; totalRows: number; totalCols: number } {
  const totalRows = values.length;
  const totalCols = values[0]?.length ?? 0;
  const total = countCells(totalRows, totalCols);
  if (total <= maxCells) {
    return { values, truncated: false, totalRows, totalCols };
  }
  const cols = Math.max(totalCols, 1);
  const maxRows = Math.max(1, Math.floor(maxCells / cols));
  return {
    values: values.slice(0, maxRows),
    truncated: true,
    totalRows,
    totalCols,
  };
}

export function assertWriteValues(values: unknown): values is CellValue[][] {
  return Array.isArray(values) && values.length > 0 && values.every((row) => Array.isArray(row));
}
