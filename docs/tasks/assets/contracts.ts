// src/agent/contracts.ts
// Single source of truth shared by the agent loop (src/agent/loop.ts), the Office.js tool
// executor (src/excel/tools.ts), the UI (src/taskpane/**) and scripts/. This is the ONLY file
// that every parallel task imports. FROZEN after task T1: a needed change is reported as a
// "CONTRACT CHANGE REQUEST" in the agent's final message (see CLAUDE.md) and applied by the
// orchestrator, never by editing this file from a task branch.
import type Anthropic from '@anthropic-ai/sdk';

// ---------- Models (env overrides are applied in src/agent/client.ts) ----------
export const DEFAULT_MODEL = 'claude-sonnet-5';
export const REASON_MODEL = 'claude-opus-5';
export type Effort = 'low' | 'medium' | 'high' | 'xhigh' | 'max';
export const DEFAULT_EFFORT: Effort = 'medium';
export const REASON_EFFORT: Effort = 'high';

// ---------- Size guards: enforced in code by the executor and the loop, repeated to the model in tool descriptions ----------
export const LIMITS = {
  /** Max cells returned by one read_range / get_selection page. */
  READ_PAGE_CELLS: 2000,
  /** Max cells written by one write_range call. */
  WRITE_MAX_CELLS: 5000,
  /** Max cells echoed back (readBack) after a write; errorCells are always computed over the whole written range. */
  WRITE_READBACK_CELLS: 500,
  /** Max hits returned by search_workbook. */
  SEARCH_MAX_HITS: 50,
  /** Max sheets scanned by search_workbook when sheet is ''. */
  SEARCH_MAX_SHEETS: 50,
  /** Max sheets described in get_workbook_overview (the rest are only counted). */
  OVERVIEW_MAX_SHEETS: 100,
  /** Header/sample columns per sheet in the overview. */
  OVERVIEW_MAX_COLS: 50,
  /** Sample data rows per sheet in the overview (after the header row). */
  OVERVIEW_SAMPLE_ROWS: 3,
  /** summarize_range scans at most this many rows ... */
  SUMMARY_MAX_ROWS: 200_000,
  /** ... and this many columns (the first N of the range). */
  SUMMARY_MAX_COLS: 100,
  /** Cells loaded per ctx.sync() while summarizing (keeps each Office.js payload small). */
  SUMMARY_CHUNK_CELLS: 20_000,
  /** Distinct sample values reported per column by summarize_range. */
  SUMMARY_SAMPLE_VALUES: 3,
  /** Error-cell addresses reported by summarize_range. */
  SUMMARY_MAX_ERROR_CELLS: 50,
  /** Any single cell's text is cut to this many chars before it reaches the model. */
  CELL_TEXT_CHARS: 200,
  /** Serialized tool_result hard cap (chars); the loop truncates beyond this. */
  TOOL_RESULT_CHARS: 30_000,
  /** Max model<->tool round trips per user message. */
  MAX_TOOL_ITERATIONS: 25,
  /** Rows per ctx.sync() when reading multi-row blocks. */
  ROWS_PER_SYNC: 500,
  /** Max chars of workbook context appended to a user message. */
  CONTEXT_PREAMBLE_CHARS: 4000,
  /** max_tokens per model call (non-streaming / streaming). */
  MAX_TOKENS: 16_000,
  MAX_TOKENS_STREAMING: 32_000,
  /** Oldest messages are dropped (whole turns only) beyond this many API messages. */
  HISTORY_MAX_MESSAGES: 60,
  /** A single tool call is abandoned (ok:false) after this long. */
  TOOL_TIMEOUT_MS: 60_000,
  /** SDK client request timeout (ms). */
  REQUEST_TIMEOUT_MS: 120_000,
  /** Excel sheet-name limit. */
  SHEET_NAME_CHARS: 31,
} as const;

// ---------- Cell model ----------
export type CellValue = string | number | boolean | null;
export type ReadMode = 'values' | 'formulas';

// ---------- Tool inputs (mirror TOOL_DEFINITIONS exactly) ----------
export type GetWorkbookOverviewInput = Record<string, never>;
export interface ReadRangeInput { sheet: string; address: string; mode: ReadMode; page: number }
/** sheet: '' searches all sheets (active first). */
export interface SearchWorkbookInput { query: string; sheet: string; matchCase: boolean; completeMatch: boolean }
export type GetSelectionInput = Record<string, never>;
/** address: '' summarizes the sheet's whole used range. */
export interface SummarizeRangeInput { sheet: string; address: string }
export interface WriteRangeInput { sheet: string; address: string; values: CellValue[][] }
export interface AddSheetInput { name: string }

export interface ToolInputMap {
  get_workbook_overview: GetWorkbookOverviewInput;
  read_range: ReadRangeInput;
  search_workbook: SearchWorkbookInput;
  get_selection: GetSelectionInput;
  summarize_range: SummarizeRangeInput;
  write_range: WriteRangeInput;
  add_sheet: AddSheetInput;
}
export type ToolName = keyof ToolInputMap;
export const TOOL_NAMES: readonly ToolName[] = [
  'get_workbook_overview', 'read_range', 'search_workbook', 'get_selection', 'summarize_range', 'write_range', 'add_sheet',
];
export const WRITE_TOOLS: ReadonlySet<ToolName> = new Set<ToolName>(['write_range', 'add_sheet']);
export function isToolName(x: string): x is ToolName {
  return (TOOL_NAMES as readonly string[]).includes(x);
}

// ---------- Tool outputs ----------
export interface SheetSummary {
  name: string;
  visible: boolean;
  /** Local address of the used range, e.g. 'A1:L100001'; null when the sheet is empty. */
  usedRange: string | null;
  rows: number;
  cols: number;
  header: CellValue[];
  sample: CellValue[][];
}
export interface WorkbookOverview {
  workbookName?: string;
  activeSheet: string;
  /** Selection address including sheet, e.g. "Data!B4:C9". */
  selection: string;
  sheetCount: number;
  sheets: SheetSummary[];
  omittedSheets: number;
}
export interface RangePage {
  sheet: string;
  /** Local address of THIS page's rows, e.g. 'A1:L166'. */
  address: string;
  mode: ReadMode;
  page: number;
  totalPages: number;
  rowsPerPage: number;
  /** 1-based Excel row number of values[0]. */
  firstRow: number;
  /** Column letters for each column of the page, e.g. ['A','B',...]. */
  columns: string[];
  /** Rows/cols of the whole effective (used-range-clipped) range, not just this page. */
  totalRows: number;
  totalCols: number;
  values: CellValue[][];
  /** Present when mode === 'formulas': same shape as values; non-formula cells hold the value. */
  formulas?: CellValue[][];
  note?: string;
}
export interface SearchHit { sheet: string; address: string; value: CellValue }
export interface SearchResult { query: string; hits: SearchHit[]; truncated: boolean; sheetsSearched: number }
export interface SelectionResult { sheet: string; address: string; page: RangePage }
export interface ColumnSummary {
  column: string;
  header: CellValue;
  nonEmpty: number;
  numeric: number;
  text: number;
  errors: number;
  min?: number;
  max?: number;
  sum?: number;
  samples: CellValue[];
}
export interface RangeSummary {
  sheet: string;
  address: string;
  totalRows: number;
  totalCols: number;
  rowsScanned: number;
  colsScanned: number;
  truncated: boolean;
  columns: ColumnSummary[];
  errorCells: { address: string; value: string }[];
  errorCellsTruncated: boolean;
}
export interface WriteResult {
  sheet: string;
  /** Effective local address written, e.g. 'B2:D10'. */
  address: string;
  cellsWritten: number;
  /** Formulas/values of the range BEFORE the write (full shape, for Undo in the UI; not sent to the model). */
  previous: CellValue[][];
  /** First LIMITS.WRITE_READBACK_CELLS computed values after the write. */
  readBack: CellValue[][];
  readBackTruncated: boolean;
  /** Cells whose valueTypes === 'Error' after the write, over the WHOLE written range. */
  errorCells: { address: string; value: string }[];
}
export interface AddSheetResult { name: string }

export interface ToolOutputMap {
  get_workbook_overview: WorkbookOverview;
  read_range: RangePage;
  search_workbook: SearchResult;
  get_selection: SelectionResult;
  summarize_range: RangeSummary;
  write_range: WriteResult;
  add_sheet: AddSheetResult;
}

export type ToolResult<T = unknown> = { ok: true; data: T } | { ok: false; error: string };
/** Implemented by src/excel/tools.ts. Must NEVER throw: every failure is { ok:false, error }. */
export type ToolExecutor = (name: ToolName, input: unknown) => Promise<ToolResult>;

// ---------- Tool definitions sent to the API ----------
// strict: true => every property listed in `required`, additionalProperties:false, no nullable types,
// no minimum/maximum (unsupported in strict mode; the executor validates instead). Keep ORDER STABLE:
// the tools array is part of the cached prompt prefix.
const noProps = { type: 'object' as const, properties: {}, required: [] as string[], additionalProperties: false };
export const TOOL_DEFINITIONS: Anthropic.Tool[] = [
  {
    name: 'get_workbook_overview',
    strict: true,
    description:
      'Map of the workbook: every sheet with its used range, dimensions, header row and ' +
      `${LIMITS.OVERVIEW_SAMPLE_ROWS} sample rows, plus the active sheet and current selection. ` +
      'Cheap on any workbook size (never loads data beyond the samples). The same map is already in the ' +
      '<workbook_context> block of the first user message; call this again only after sheets were added or if the context is missing.',
    input_schema: noProps,
  },
  {
    name: 'read_range',
    strict: true,
    description:
      `Read one page of values or formulas from a range. Returns at most ${LIMITS.READ_PAGE_CELLS} cells per page ` +
      '(rowsPerPage = floor(2000 / columns)) plus totalPages; whole-column addresses like "A:A" are clipped to the used range. ' +
      'Never try to read a whole large sheet: use search_workbook to locate cells and summarize_range for column statistics, ' +
      'then read small pages. mode "formulas" shows each formula cell as value {=FORMULA}.',
    input_schema: {
      type: 'object',
      additionalProperties: false,
      required: ['sheet', 'address', 'mode', 'page'],
      properties: {
        sheet: { type: 'string', description: 'Worksheet name, e.g. "Data". Ignored if address already has a sheet prefix.' },
        address: { type: 'string', description: 'A1-style address: "A1:L200", "Data!A1:L200", "A:C" (clipped to used range).' },
        mode: { type: 'string', enum: ['values', 'formulas'], description: '"values" = computed values; "formulas" = values plus formula text.' },
        page: { type: 'integer', description: '0-based page index. Start at 0; use totalPages from the result to continue.' },
      },
    },
  },
  {
    name: 'search_workbook',
    strict: true,
    description:
      'Find cells whose displayed text matches a query (labels, numbers, or error text such as "#REF!" or "#DIV/0!"). ' +
      `Returns up to ${LIMITS.SEARCH_MAX_HITS} hits with sheet, address and value. Use it to locate things in large sheets instead of paging.`,
    input_schema: {
      type: 'object',
      additionalProperties: false,
      required: ['query', 'sheet', 'matchCase', 'completeMatch'],
      properties: {
        query: { type: 'string' },
        sheet: { type: 'string', description: `One sheet name, or "" to search all sheets (active first, max ${LIMITS.SEARCH_MAX_SHEETS}).` },
        matchCase: { type: 'boolean' },
        completeMatch: { type: 'boolean', description: 'true = whole-cell match; false = substring.' },
      },
    },
  },
  {
    name: 'get_selection',
    strict: true,
    description:
      `The range the user currently has selected, with values and formulas (first ${LIMITS.READ_PAGE_CELLS} cells; use read_range for more). ` +
      'Use when the user says "this", "here", "the selected cells".',
    input_schema: noProps,
  },
  {
    name: 'summarize_range',
    strict: true,
    description:
      'Per-column statistics of a range WITHOUT returning its data: header, non-empty/numeric/text/error counts, min/max/sum, ' +
      `${LIMITS.SUMMARY_SAMPLE_VALUES} sample values, and the addresses of up to ${LIMITS.SUMMARY_MAX_ERROR_CELLS} error cells. ` +
      `Scans up to ${LIMITS.SUMMARY_MAX_ROWS} rows x ${LIMITS.SUMMARY_MAX_COLS} columns. The first row is treated as the header. ` +
      'This is how to understand a sheet with thousands of rows, and a reliable way to find error cells.',
    input_schema: {
      type: 'object',
      additionalProperties: false,
      required: ['sheet', 'address'],
      properties: {
        sheet: { type: 'string' },
        address: { type: 'string', description: 'Local A1 range such as "A1:L5000", or "" for the whole used range.' },
      },
    },
  },
  {
    name: 'write_range',
    strict: true,
    description:
      'Write a rectangular block of values and/or formulas (strings beginning with "=") into a range. ' +
      `Max ${LIMITS.WRITE_MAX_CELLS} cells per call; chunk larger writes by rows. If address is a single cell it is the top-left corner. ` +
      'null leaves a cell unchanged; "" clears it. Returns computed values read back plus errorCells (#REF!, #DIV/0!, #NAME? ...), ' +
      'so always check errorCells after writing. Prefer formulas over pasted numbers for anything derived; to compute an aggregate over ' +
      'a big range, write a formula (e.g. "=AVERAGE(Data!E2:E100001)") into a scratch cell and read its computed value from readBack.',
    input_schema: {
      type: 'object',
      additionalProperties: false,
      required: ['sheet', 'address', 'values'],
      properties: {
        sheet: { type: 'string' },
        address: { type: 'string', description: '"B2" (top-left) or "B2:D10" (must match the values shape).' },
        values: {
          type: 'array',
          description: 'Rows of cells. "=SUM(B2:B10)" for formulas, numbers for numbers, "" to clear, null to leave unchanged.',
          items: { type: 'array', items: { type: ['string', 'number', 'boolean', 'null'] } },
        },
      },
    },
  },
  {
    name: 'add_sheet',
    strict: true,
    description:
      `Create and activate a new worksheet (max ${LIMITS.SHEET_NAME_CHARS} chars, no []:*?/\\). ` +
      'Use for dashboards, scenario summaries or scratch calculations so existing sheets stay untouched.',
    input_schema: {
      type: 'object',
      additionalProperties: false,
      required: ['name'],
      properties: { name: { type: 'string' } },
    },
  },
];

// ---------- Agent loop contract (implemented by src/agent/loop.ts, consumed by the UI) ----------
export interface AgentTurnOptions {
  /** 'Reason' toggle: use REASON_MODEL + REASON_EFFORT (note: switching model busts the prompt cache). */
  reason: boolean;
  /** 'Agent Mode' toggle: when false, WRITE_TOOLS go through callbacks.confirmWrite first. */
  agentMode: boolean;
  /** client.messages.stream (true) or client.messages.create (false). */
  streaming: boolean;
  signal?: AbortSignal;
}
export interface ToolCallEvent { id: string; name: string; input: unknown }
export interface ToolResultEvent extends ToolCallEvent { result: ToolResult; ms: number }
export interface AgentCallbacks {
  onText(delta: string): void;
  onToolCall(e: ToolCallEvent): void;
  onToolResult(e: ToolResultEvent): void;
  /** Resolve true to execute the write, false to skip it (the model is told the user declined). */
  confirmWrite(e: ToolCallEvent): Promise<boolean>;
}
export interface AgentTurnInput {
  client: Anthropic;
  /** Full API history from previous turns (assistant content blocks + tool_result user messages included). */
  history: Anthropic.MessageParam[];
  userText: string;
  /** Workbook context appended as the LAST text block of the user message, wrapped in <workbook_context>. May be ''. */
  contextPreamble: string;
  options: AgentTurnOptions;
  executeTool: ToolExecutor;
  callbacks: AgentCallbacks;
}
export type TurnStopReason = NonNullable<Anthropic.Message['stop_reason']> | 'aborted' | 'max_iterations';
export interface TurnUsage { inputTokens: number; outputTokens: number; cacheReadTokens: number; cacheWriteTokens: number }
export interface AgentTurnOutput {
  /** New full history (input history + this turn's messages, trimmed). Replace UI state with this. */
  history: Anthropic.MessageParam[];
  finalText: string;
  stopReason: TurnStopReason;
  model: string;
  usage: TurnUsage;
  toolCalls: number;
}
export type RunAgentTurn = (input: AgentTurnInput) => Promise<AgentTurnOutput>;

// ---------- UI display model ----------
export interface TurnMeta { model: string; usage: TurnUsage; toolCalls: number; ms: number }
export type ChatItem =
  | { kind: 'user'; id: string; text: string }
  | { kind: 'assistant'; id: string; text: string; streaming: boolean; meta?: TurnMeta }
  | { kind: 'tool'; id: string; name: string; input: unknown; result?: ToolResult; ms?: number }
  | { kind: 'error'; id: string; text: string };

export const INTRO_TEXT =
  "Hi, I'm Crunched — your AI analyst in Excel. I can help you with things like:\n" +
  '• Error checking and fixing models\n' +
  '• Building financial or business models\n' +
  '• Analyzing and comparing scenario models\n\n' +
  'What should we work on first?';
export const SUGGESTIONS: readonly string[] = [
  'Error check this model',
  'Check the assumptions in this workbook',
  'Create a comparison dashboard across scenarios',
];

// ---------- Pure helpers shared by UI, loop and executor ----------
/** Single-line text of a cell for prompts/TSV: '' for empty, tabs/newlines flattened, optionally cut. */
export function cellToText(v: CellValue | undefined, maxChars: number = LIMITS.CELL_TEXT_CHARS): string {
  if (v === null || v === undefined || v === '') return '';
  const s = String(v).replace(/[\t\r\n]+/g, ' ');
  return s.length > maxChars ? s.slice(0, Math.max(0, maxChars - 1)) + '…' : s;
}

/** Deterministic, compact text rendering of the overview for the <workbook_context> block. */
export function formatOverviewForPrompt(o: WorkbookOverview, maxChars: number = LIMITS.CONTEXT_PREAMBLE_CHARS): string {
  const lines: string[] = [];
  if (o.workbookName) lines.push(`Workbook: ${o.workbookName}`);
  lines.push(`Active sheet: ${o.activeSheet} | Selection: ${o.selection} | Sheets: ${o.sheetCount}`);
  for (const s of o.sheets) {
    lines.push(`- ${s.name}${s.visible ? '' : ' (hidden)'}: ${s.usedRange ?? 'empty'} (${s.rows} rows x ${s.cols} cols)`);
    if (s.header.length) lines.push(`  header: ${s.header.map((c) => cellToText(c, 40)).join(' | ')}`);
    for (const row of s.sample) lines.push(`  row: ${row.map((c) => cellToText(c, 40)).join(' | ')}`);
  }
  if (o.omittedSheets > 0) lines.push(`(+${o.omittedSheets} more sheets not listed; use search_workbook)`);
  const text = lines.join('\n');
  return text.length > maxChars ? text.slice(0, maxChars - 20) + '\n...[truncated]' : text;
}

/** Changes when sheets are added/removed/resized; the UI re-sends the full overview only when this changes. */
export function overviewSignature(o: WorkbookOverview): string {
  return o.sheets.map((s) => `${s.name}:${s.rows}x${s.cols}`).join(';') + `|${o.sheetCount}`;
}

/** One-line context for follow-up turns (what "this"/"here" refers to). */
export function formatSelectionLine(o: WorkbookOverview): string {
  return `Active sheet: ${o.activeSheet} | Selection: ${o.selection}`;
}

function gridToTsv(p: RangePage): string {
  const head = ['', ...p.columns].join('\t');
  const rows = p.values.map((row, r) => {
    const cells = row.map((v, c) => {
      const f = p.formulas?.[r]?.[c];
      const base = cellToText(v);
      return typeof f === 'string' && f.startsWith('=') ? `${base} {${cellToText(f)}}` : base;
    });
    return [String(p.firstRow + r), ...cells].join('\t');
  });
  return [head, ...rows].join('\n');
}

/**
 * Text the model sees for a successful tool call. Pages are rendered as a metadata line followed by a
 * TSV grid (about half the tokens of JSON); write results omit `previous` (UI-only); everything else is JSON.
 */
export function serializeToolData(name: ToolName, data: unknown): string {
  if (name === 'read_range' || name === 'get_selection') {
    const sel = name === 'get_selection' ? (data as SelectionResult) : null;
    const p = sel ? sel.page : (data as RangePage);
    const lastRow = p.firstRow + Math.max(0, p.values.length - 1);
    const meta =
      `${p.sheet}!${p.address} | mode=${p.mode} | page ${p.page + 1}/${p.totalPages} | rows ${p.firstRow}-${lastRow} of ${p.totalRows} | ${p.totalCols} cols` +
      (p.note ? ` | ${p.note}` : '');
    const head = sel ? `selection: ${sel.sheet}!${sel.address}\n` : '';
    return head + meta + '\n' + gridToTsv(p);
  }
  if (name === 'write_range') {
    const { previous, ...rest } = data as WriteResult;
    return JSON.stringify({ ...rest, previousCells: previous.length ? previous.length * previous[0].length : 0 });
  }
  return JSON.stringify(data);
}

/** Truncate a serialized tool result to the hard cap with an explicit marker. */
export function capToolResultText(s: string, max: number = LIMITS.TOOL_RESULT_CHARS): string {
  return s.length > max ? s.slice(0, max) + `\n...[truncated: ${s.length - max} more chars; request a smaller page or a summary]` : s;
}
