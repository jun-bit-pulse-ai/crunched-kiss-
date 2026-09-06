# T3 — Office.js tool executor with size guards (+ pure range helpers and tests)

**Agent model:** claude-sonnet-5. **Estimate:** 55 min. **Depends on:** T1. **Runs in parallel with:** T2, T4, T5, T6.
**Worktree:** `/Users/junseki/Documents/GitHub/ck-T3` on branch `task/T3`.

## Context

You are working in an Excel task-pane add-in (React + TypeScript + Office.js / ExcelApi 1.9+). Read `CLAUDE.md`
first, then `src/agent/contracts.ts` and treat it as FROZEN: implement exactly its `ToolExecutor`, every
input/output interface, and every number in `LIMITS`. You cannot open Excel; the human will run your tools from
the pane's Web Inspector via `window.crunched.run(name, input)` (already wired in `src/excel/devConsole.ts`).
So: keep all address/paging math in pure, unit-tested functions, keep every Office.js call straightforward,
and make every failure a `{ ok: false, error }` the model can act on. Never run `npm start`/`npm install`.

## You own (create or replace exactly these)

- `src/excel/ranges.ts` — pure helpers, no Office.js import
- `src/excel/ranges.test.ts` — vitest
- `src/excel/tools.ts` — `export const executeTool: ToolExecutor` (replaces the T1 stub)

Do NOT touch: `src/agent/**`, `src/taskpane/**`, `src/excel/devConsole.ts`, `webpack.config.js`, `package.json`, `manifest.xml`, `docs/**`.

## 1. `src/excel/ranges.ts` (pure)

```ts
import type { CellValue } from '../agent/contracts';

/** 0-based inclusive rectangle; sheet is null when the address had no sheet prefix. */
export interface ParsedAddress { sheet: string | null; r0: number; c0: number; r1: number; c1: number }

export const MAX_ROWS = 1_048_576;
export const MAX_COLS = 16_384;

/**
 * Accepts: 'A1', 'A1:C10', 'Data!A1:C10', "'My Sheet'!$A$1:$C$10" (unescape '' -> '), 'A:C' (rows 0..MAX_ROWS-1),
 * '3:5' (cols 0..MAX_COLS-1). Normalises so r0<=r1, c0<=c1. Throws Error('Invalid address: <input>') otherwise.
 */
export function parseAddress(a: string): ParsedAddress
export function colToIndex(col: string): number            // 'A'->0, 'Z'->25, 'AA'->26, 'XFD'->16383; case-insensitive
export function indexToCol(i: number): string              // inverse of colToIndex
export function toA1(p: { r0: number; c0: number; r1: number; c1: number }): string  // 'A1:C10', or 'A1' for a single cell
export function quoteSheet(name: string): string            // "'My Sheet'" when it has spaces or any of !'"[]:*?/\\ ; escape ' as ''
export function cellCount(p: ParsedAddress): number
/**
 * Page geometry for a range: rowsPerPage = max(1, floor(maxCells / cols)); totalPages = max(1, ceil(rows / rowsPerPage));
 * firstRowIdx/lastRowIdx are 0-based ABSOLUTE sheet row indexes of the page. Throws Error('page N out of range (0..M)') if page >= totalPages or page < 0.
 */
export function pageBounds(p: ParsedAddress, page: number, maxCells: number): { rowsPerPage: number; totalPages: number; firstRowIdx: number; lastRowIdx: number }
/** undefined/null -> null; strings cut to maxChars (with '…'); numbers/booleans pass through; anything else -> String(v). */
export function clampCell(v: unknown, maxChars: number): CellValue
/** Excel.RangeValueType string -> is it an error cell. */
export function isErrorType(t: unknown): boolean            // t === 'Error'
export function isNumericType(t: unknown): boolean          // t === 'Double' || t === 'Integer'
/** Strip []:*?/\\ and trim to maxChars (LIMITS.SHEET_NAME_CHARS); throws if empty after stripping. */
export function sanitizeSheetName(n: string, maxChars: number): string
/** Validates a rectangular CellValue[][] (>=1 row, all rows same length >=1, cells are string|number|boolean|null); returns {rows, cols} or throws Error with a precise message. */
export function gridShape(values: unknown): { rows: number; cols: number }
```

## 2. `src/excel/ranges.test.ts`

At least 14 vitest cases covering: every function above; `parseAddress` of `'A:C'` gives `r0 0, r1 MAX_ROWS-1`;
`'3:5'` gives `c0 0, c1 MAX_COLS-1`; `"'My Sheet'!$B$2:$D$4"` gives sheet `My Sheet`; `"'O''Brien'!A1"` gives
sheet `O'Brien`; reversed `'C10:A1'` normalises; `'A0'`, `''`, `'1A'` throw; `pageBounds` for a 100001x12 range
with maxCells 2000 gives `rowsPerPage 166`, `totalPages 603`, page 602 covers the last rows; for 5000 cols gives
`rowsPerPage 1`; page 603 throws; `clampCell('x'.repeat(300), 200)` has length 200 and ends with `…`;
`gridShape([[1,2],[3]])` throws; `quoteSheet('Data')` is `Data`; `sanitizeSheetName('Q1/Q2: plan?*', 31)` is `Q1Q2 plan`.

## 3. `src/excel/tools.ts`

Structure:

```ts
import type { ToolExecutor, ToolName, ToolResult, ... } from '../agent/contracts';
import { LIMITS } from '../agent/contracts';
import { ... } from './ranges';

const handlers: { [K in ToolName]: (input: unknown) => Promise<unknown> } = { ... };

export const executeTool: ToolExecutor = async (name, input) => {
  try {
    const h = handlers[name];
    if (!h) return { ok: false, error: `Unknown tool ${name}` };
    return { ok: true, data: await h(input) };
  } catch (e) {
    return { ok: false, error: describeError(e) };
  }
};
```

`describeError(e)`: for `OfficeExtension.Error` return `` `${e.code}: ${e.message}` `` (add `e.debugInfo?.errorLocation`
when present); for plain `Error` return `e.message`; otherwise `String(e)`. Input validation errors are thrown as
`Error('Invalid input: ...')` inside handlers so they also surface as `ok:false`. Validate shapes defensively even
though tool schemas are strict (`typeof input.sheet === 'string'`, etc.).

Shared helpers inside tools.ts:

- `resolveSheet(ctx, parsed, inputSheet)`: `parsed.sheet ?? inputSheet`; empty -> throw `Invalid input: sheet is required`.
  `const ws = ctx.workbook.worksheets.getItemOrNullObject(name); ws.load('name'); await ctx.sync(); if (ws.isNullObject) throw new Error(\`Sheet "${name}" not found. Sheets: ...\`)`
  (load `worksheets.load('items/name')` in the same sync so the error can list existing names).
- `effectiveRange(ctx, ws, localAddress)`: `const target = ws.getRange(localAddress); const eff = target.getIntersectionOrNullObject(ws.getUsedRangeOrNullObject(true)); eff.load('address,rowCount,columnCount'); await ctx.sync();`
  returns `null` when `eff.isNullObject`, else `parseAddress(eff.address)` (Excel returns `Sheet!A1:L100001`).
  Note: if the used range itself is null (empty sheet) `getIntersectionOrNullObject` on a null object still yields a null object; handle both.
- `readPage(ctx, ws, sheetName, eff: ParsedAddress, mode, page): Promise<RangePage>`: `pageBounds(eff, page, LIMITS.READ_PAGE_CELLS)`;
  loop from `firstRowIdx` to `lastRowIdx` in blocks of at most `LIMITS.ROWS_PER_SYNC` rows: `ws.getRangeByIndexes(rowIdx, eff.c0, n, cols)`,
  `load('values')` or `load('values,formulas')` in formulas mode, `await ctx.sync()` per block; concatenate; clamp every
  cell with `clampCell(v, LIMITS.CELL_TEXT_CHARS)`; return `{ sheet, address: toA1({r0:firstRowIdx, c0:eff.c0, r1:lastRowIdx, c1:eff.c1}), mode, page, totalPages, rowsPerPage, firstRow: firstRowIdx + 1, columns: [indexToCol(c0)..indexToCol(c1)], totalRows: eff.r1-eff.r0+1, totalCols: cols, values, formulas? }`.
  (A page is <= 2000 cells so it is usually one block; write the loop anyway, it is what keeps big pages under Excel Online's ~5 MB per-request cap.)

Per tool (each is ONE `Excel.run` unless stated):

- **get_workbook_overview** (2 syncs total regardless of workbook size).
  Sync 1: `ctx.workbook.load('name')`; `const sheets = ctx.workbook.worksheets; sheets.load('items/name,items/visibility')`;
  `ctx.workbook.getActiveWorksheet().load('name')`; `ctx.workbook.getSelectedRange().load('address')`; sync.
  Then for the first `LIMITS.OVERVIEW_MAX_SHEETS` sheets: `const used = ws.getUsedRangeOrNullObject(true); used.load('address,rowCount,columnCount')`.
  Sync 2 (combined with the samples): for each sheet also queue
  `ws.getRangeByIndexes(top, left, Math.min(rowCount, 1 + LIMITS.OVERVIEW_SAMPLE_ROWS), Math.min(columnCount, LIMITS.OVERVIEW_MAX_COLS)).load('values')`
  where `top/left` come from `parseAddress(used.address)` — but `used.address` is only known after a sync, so do: sync A (metadata + used ranges), sync B (samples). Two data syncs plus the first metadata sync is acceptable; state the actual count in your report.
  Build `SheetSummary { name, visible: visibility === 'Visible', usedRange: local address without the sheet prefix (or null), rows, cols, header: row 0 clamped, sample: rows 1..3 clamped }`;
  `omittedSheets = total - described`; `selection = address as returned (with sheet)`; `workbookName = ctx.workbook.name`.
- **read_range**: `parseAddress(input.address)` (throws -> ok:false); resolve sheet; `effectiveRange(ctx, ws, toA1(parsed))`; if null ->
  return a RangePage with `values: []`, `totalRows 0`, `totalPages 1`, `note: 'no used cells in that range'`; else validate `page` is a non-negative
  integer (throw `Invalid input: page must be an integer >= 0`) and `readPage(...)`.
- **search_workbook**: sheets to scan = `[input.sheet]` if non-empty, else active sheet first then the others in workbook order, at most `LIMITS.SEARCH_MAX_SHEETS`.
  Per sheet, STAGED so a common substring on a huge sheet cannot blow the payload:
  1. `const found = ws.findAllOrNullObject(query, { completeMatch, matchCase }); found.load('isNullObject,areaCount,cellCount'); await ctx.sync();` skip if null.
  2. `const areas = found.areas; areas.load('items/address,items/rowCount,items/columnCount'); await ctx.sync();`
  3. For areas in order while `hits.length < LIMITS.SEARCH_MAX_HITS`: if `rowCount*columnCount === 1` queue `area.load('values')`; else queue
     `area.getRangeByIndexes(0, 0, Math.min(rowCount, Math.ceil(remaining / columnCount)), columnCount).load('address,values')`; one sync; push
     `{ sheet, address: <cell address without sheet prefix, computed per cell from the area's top-left via parseAddress + toA1>, value: clampCell(v) }` per cell until the cap.
  Stop scanning further sheets once the cap is reached. `truncated = cap reached`; `sheetsSearched = number of sheets actually scanned`.
  Note: `findAll` matches displayed text, so `'#DIV/0!'` and `'#REF!'` are expected to match error cells; the human verifies this in Excel and `summarize_range` is the fallback for finding errors.
- **get_selection**: `const sel = ctx.workbook.getSelectedRange(); sel.load('address'); sel.worksheet.load('name'); await ctx.sync();`
  parse the address (it has the sheet prefix), `effectiveRange` on the local part (a whole-sheet selection is thereby clipped), and `readPage(..., 'formulas', 0)`.
  Return `{ sheet, address: local selection address, page }`.
- **summarize_range**: address `''` -> the used range (`getUsedRangeOrNullObject(true)`), else `effectiveRange`. If null -> throw `Range has no used cells`.
  Clip to the first `LIMITS.SUMMARY_MAX_COLS` columns and `LIMITS.SUMMARY_MAX_ROWS` rows (`truncated = true` if clipped). Row 0 is the header.
  `rowsPerChunk = Math.max(1, Math.floor(LIMITS.SUMMARY_CHUNK_CELLS / cols))`; for each chunk `ws.getRangeByIndexes(r, c0, n, cols).load('values,valueTypes')`,
  ONE chunk per `ctx.sync()` (do not queue all chunks at once). Accumulate per column: `nonEmpty` (`valueTypes !== 'Empty'`), `numeric` (`isNumericType`),
  `text` (`'String'`), `errors` (`isErrorType`), `min/max/sum` over numeric values, `samples` = first `LIMITS.SUMMARY_SAMPLE_VALUES` distinct non-empty values
  (clamped to 40 chars). `errorCells` = first `LIMITS.SUMMARY_MAX_ERROR_CELLS` error cells as `{ address: toA1 of the cell, value: String(v) }`, `errorCellsTruncated` when more exist.
  Return `RangeSummary { sheet, address: local address of the scanned range, totalRows, totalCols (of the effective range before clipping), rowsScanned, colsScanned, truncated, columns, errorCells, errorCellsTruncated }`.
  Never keep the raw values after a chunk is processed.
- **write_range**: `gridShape(values)` (throws precise errors); `rows*cols <= LIMITS.WRITE_MAX_CELLS` else throw
  `` `Too many cells (${n}); write at most ${LIMITS.WRITE_MAX_CELLS} per call, chunk by rows` ``. Parse address; resolve sheet. If the parsed address is a single
  cell and the grid is bigger: `range = ws.getRange(addr).getResizedRange(rows - 1, cols - 1)`; otherwise the address dimensions must equal the grid
  (else throw with both shapes). Sync 1: `range.load('address,formulas')` -> `previous` (clamped cells; `''` for empty). Then `range.formulas = values`
  (Office semantics: strings starting with `=` become formulas, other values are stored as values, `null` leaves the cell unchanged, `''` clears). Sync 2.
  Sync 3: `range.load('values,valueTypes')`. `errorCells` = every cell with `isErrorType`, address via `toA1` of the absolute cell. `readBack` = the first
  `Math.max(1, Math.floor(LIMITS.WRITE_READBACK_CELLS / cols))` rows of values (clamped); `readBackTruncated` when cut. `cellsWritten = rows*cols`.
- **add_sheet**: `name = sanitizeSheetName(input.name, LIMITS.SHEET_NAME_CHARS)`; `worksheets.getItemOrNullObject(name).load('isNullObject')`; sync;
  if it exists throw `` `Sheet "${name}" already exists` ``; else `const ws = worksheets.add(name); ws.activate(); await ctx.sync(); return { name }`.

Office.js reminders: `Excel.run(async (ctx) => ...)`; only `load` what you need; `*OrNullObject` + `.isNullObject`; indexes 0-based; `Range.address`
includes the sheet prefix; `valueTypes` values are `'Boolean'|'Double'|'Empty'|'Error'|'Integer'|'RichValue'|'String'|'Unknown'`.

## 4. Verify, commit

```bash
npx tsc --noEmit -p tsconfig.json
npm test
git add src/excel/ranges.ts src/excel/ranges.test.ts src/excel/tools.ts && git commit -m "T3: Office.js tool executor with size guards"
```

## Acceptance

Offline (you): `npm test` passes with >= 14 cases in `ranges.test.ts`; `npx tsc --noEmit` passes; `grep -c "startsWith('#')" src/excel/tools.ts` prints 0
(errors are detected via `valueTypes`, never by string prefix); `grep -c 'ROWS_PER_SYNC\|SUMMARY_CHUNK_CELLS' src/excel/tools.ts` >= 2.

In Excel (the human, in the pane's Web Inspector with `fixtures/big-model.xlsx` open; this is the 1:15 gate):
```js
await crunched.run('get_workbook_overview', {})                                                        // ok, 33 sheets (Data, Assumptions, Summary, Scenario_01..30), < 3 s
await crunched.run('read_range', { sheet: 'Data', address: 'A:L', mode: 'values', page: 0 })           // ok, 166 rows, totalPages 603, columns A..L, < 2 s
await crunched.run('read_range', { sheet: 'Data', address: 'A:L', mode: 'formulas', page: 602 })       // ok, last page, formulas present for F/H/L
await crunched.run('search_workbook', { query: '#DIV/0!', sheet: '', matchCase: false, completeMatch: false })  // hit Summary!B4 (if 0 hits, note it: summarize_range is the fallback)
await crunched.run('summarize_range', { sheet: 'Data', address: '' })                                   // ok, 12 columns, rowsScanned 100001, errorCells [], < 15 s (record the time)
await crunched.run('summarize_range', { sheet: 'Summary', address: '' })                                // errorCells include B4 and B5
await crunched.run('get_selection', {})                                                                 // ok for whatever is selected
await crunched.run('write_range', { sheet: 'Summary', address: 'D1', values: [[1, 2], ['=D1+E1', '=1/0']] }) // cellsWritten 4, errorCells [{ address: 'E2', value: '#DIV/0!' }], previous 2x2
await crunched.run('write_range', { sheet: 'Summary', address: 'D1:E2', values: [[null, ''], [null, null]] })   // then read D1:E2: D1 unchanged (1), E1 cleared
await crunched.run('add_sheet', { name: 'Scratch' })                                                    // ok; second call -> ok:false already exists
```

## When done, report

Final message: files; test count; the exact `crunched.run` commands above (so the human can paste them); the number of
`ctx.sync()` calls in `get_workbook_overview`; any Office.js API you were unsure about (name it precisely so the human can
check it first); `CONTRACT CHANGE REQUEST` if any output type could not be honoured.
