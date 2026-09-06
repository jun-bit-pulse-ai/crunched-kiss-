# T4 — Agent loop (streaming manual tool loop), system prompt, history trimming, tests, live smoke

**Agent model:** claude-sonnet-5. **Estimate:** 45 min. **Depends on:** T1. **Runs in parallel with:** T2, T3, T5, T6.
**Worktree:** `/Users/junseki/Documents/GitHub/ck-T4` on branch `task/T4`. The orchestrator copied `.env` into this worktree so `npm run smoke` can reach the API.

## Context

You are working in an Excel task-pane add-in (React + TypeScript). Read `CLAUDE.md` first, then `src/agent/contracts.ts`
and `src/agent/client.ts` (both FROZEN). Implement the agent loop the UI calls. The loop must NOT import anything from
`src/excel` or reference `Excel`/`Office`; the tool executor is injected (`AgentTurnInput.executeTool`), which is what makes
it testable in Node with a fake client and runnable against the live API with a fake workbook. Never run `npm start`/`npm install`.

HARD API FACTS (do not deviate): models `claude-sonnet-5` / `claude-opus-5`; `thinking: { type: 'adaptive' }` only (never
`budget_tokens`, 400); effort via `output_config: { effort }`; never `temperature`; never prefill an assistant message; tool
inputs come from `block.input` (already an object); ALL tool_results of one assistant turn go back in ONE user message; use SDK
types `Anthropic.MessageParam`, `Anthropic.Tool`, `Anthropic.ToolUseBlock`, `Anthropic.ToolResultBlockParam`, `Anthropic.Message`,
`Anthropic.MessageStreamParams`; typed errors `Anthropic.APIUserAbortError`, `Anthropic.RateLimitError`, `Anthropic.APIError`.
Prompt caching: max 4 breakpoints per request; Sonnet 5 caches prefixes >= 1024 tokens (tools + system clear that).

## You own (create or replace exactly these)

- `src/agent/systemPrompt.ts`
- `src/agent/history.ts`, `src/agent/history.test.ts`
- `src/agent/loop.ts` (replaces the T1 stub), `src/agent/loop.test.ts`
- `scripts/agentSmoke.ts`

Do NOT touch: `src/agent/contracts.ts`, `src/agent/client.ts`, `src/excel/**`, `src/taskpane/**`, `webpack.config.js`, `package.json`, `docs/**`.

## 1. `src/agent/systemPrompt.ts`

Export `export const SYSTEM_PROMPT = \`...\`` containing EXACTLY the text below (a frozen string: no dates, nothing per-session, so the
ephemeral cache block hits every turn).

```
You are Crunched, an AI analyst that lives in the Excel task pane. You read and write the user's open workbook through tools. Be concise and precise; cite cells as Sheet!A1.

## Workbook context
The first user message (and any later one where the sheet list changed) ends with a <workbook_context> block: sheet names, used ranges, dimensions, header row, 3 sample rows, the active sheet and the current selection. Later messages carry a one-line context with the active sheet and selection. Trust it for layout; it is not the data.

## Working with workbooks of any size
Workbooks can be huge (hundreds of sheets, a million cells). Never read a whole sheet. Order of operations:
1. Use the context (or get_workbook_overview) to learn the layout.
2. search_workbook to locate labels, values or error text such as #REF!, #DIV/0!, #N/A, #NAME?, #VALUE!.
3. summarize_range for per-column statistics and error-cell addresses of a big range in one call.
4. read_range for small pages only (max 2000 cells; use page and totalPages deliberately). Use mode "formulas" to check logic, "values" to check numbers.
5. For an aggregate over a big range (sum, average, count of errors), write a formula into a scratch cell (a new sheet from add_sheet, or the dashboard) with write_range and read the computed number from readBack. Excel does the math; you spend no tokens on rows.
If a result says truncated, page or narrow the request. Never guess values you did not read.

## Writing rules
- Before changing existing cells, say in one line what you will change and why.
- Prefer formulas over hard-coded numbers for anything derived.
- Put new analyses, dashboards and scratch calculations on a new sheet (add_sheet) unless the user points at a location.
- Max 5000 cells per write_range; chunk larger writes by rows.
- null leaves a cell unchanged; "" clears it.
- After every write, inspect errorCells and readBack and fix what you broke.
- If a write is declined by the user, do not retry it; ask what they prefer.

## Error-checking playbook
Search for error text (#REF!, #DIV/0!, #N/A, #NAME?, #VALUE!, #NUM!) and confirm with summarize_range (errorCells). Read formulas of key columns and look for: hard-coded numbers inside formula columns, inconsistent formulas along a row or column, totals that do not span the full range, references to empty cells, assumptions used without a label. Report as a short list with cell addresses and a proposed fix for each. Fix only when asked, or when the user asked for fixes in Agent Mode.

## Output style
Short answers, bullets, no tables wider than 6 columns. End with a one-line summary of every change made (Sheet!range). Say when something is out of scope (no web access, no external data).
```

## 2. `src/agent/history.ts` (+ tests)

```ts
import type Anthropic from '@anthropic-ai/sdk';
/**
 * Keep at most maxMessages messages by dropping from the FRONT, but only ever cut at a plain user message
 * (a user message whose content is a string or contains no tool_result block), so tool_use/tool_result pairs
 * are never split. Returns the same array when nothing needs trimming.
 */
export function trimHistory(history: Anthropic.MessageParam[], maxMessages: number): Anthropic.MessageParam[]
/** True when the message's content contains no tool_result blocks (string content counts as plain). */
export function isPlainUserMessage(m: Anthropic.MessageParam): boolean
```
`history.test.ts` (>= 3 cases): (a) short history returned unchanged (same reference); (b) a 70-message history with turns of
`[user text, assistant tool_use, user tool_result, assistant text]` trimmed to 60 starts with a plain user message and has <= 60 messages;
(c) trimming never produces a history whose first message is a tool_result user message or an assistant message.

## 3. `src/agent/loop.ts`

Exports `export const runAgentTurn: RunAgentTurn` and `export function buildRequest(history, options): Anthropic.MessageStreamParams`
(pure, for tests). Imports `MODEL_CONFIG` from `./client`, `SYSTEM_PROMPT` from `./systemPrompt`, `trimHistory` from `./history`,
and from `./contracts`: `LIMITS`, `TOOL_DEFINITIONS`, `WRITE_TOOLS`, `isToolName`, `serializeToolData`, `capToolResultText`, types.

Behaviour:

1. `const model = options.reason ? MODEL_CONFIG.reasonModel : MODEL_CONFIG.defaultModel;`
   `const effort = options.reason ? MODEL_CONFIG.reasonEffort : MODEL_CONFIG.defaultEffort;`
2. User message: `content = [{ type: 'text', text: userText }, ...(contextPreamble ? [{ type: 'text', text: \`<workbook_context>\n${contextPreamble}\n</workbook_context>\` }] : [])]`.
   `let history = trimHistory([...input.history, userMessage], LIMITS.HISTORY_MAX_MESSAGES)`.
3. `buildRequest(history, options)` returns
   ```ts
   {
     model, max_tokens: options.streaming ? LIMITS.MAX_TOKENS_STREAMING : LIMITS.MAX_TOKENS,
     system: [{ type: 'text', text: SYSTEM_PROMPT, cache_control: { type: 'ephemeral' } }],
     tools: TOOL_DEFINITIONS,
     thinking: { type: 'adaptive' },
     output_config: { effort },
     messages: withTailBreakpoint(history),
   }
   ```
   `withTailBreakpoint` returns a shallow copy of `history` where the LAST message's content is an array whose last block carries
   `cache_control: { type: 'ephemeral' }` (clone that message; never mutate the stored history). That is 2 breakpoints total (system + moving tail),
   within the limit of 4, and gives every iteration of the tool loop a cache read of the whole previous prefix.
4. Loop for `iteration < LIMITS.MAX_TOOL_ITERATIONS`:
   - Request: if `options.streaming`:
     `const stream = client.messages.stream(params, { signal: options.signal }); stream.on('text', (d) => { text += d; callbacks.onText(d); }); const message = await stream.finalMessage();`
     else `const message = await client.messages.create(params, { signal: options.signal });` and for each `text` block call `callbacks.onText(block.text)`.
   - Accumulate usage: `input_tokens`, `output_tokens`, `cache_read_input_tokens ?? 0`, `cache_creation_input_tokens ?? 0`.
   - `history.push({ role: 'assistant', content: message.content })` (verbatim, thinking blocks included).
   - `stop_reason`:
     - `'tool_use'`: execute tools (below), push ONE `{ role: 'user', content: toolResults }`, continue.
     - `'pause_turn'`: continue (the API asks to be re-sent as is).
     - `'max_tokens'`: if the message contains complete `tool_use` blocks, execute them as for `'tool_use'`; otherwise, ONCE per turn, push
       `{ role: 'user', content: [{ type: 'text', text: 'Continue exactly where you stopped.' }] }` and continue; the second time, stop with stopReason `'max_tokens'`.
     - `'refusal'`, `'end_turn'`, `'stop_sequence'`: stop with that stopReason.
   - Tool execution, SEQUENTIALLY (Office.js is single-threaded): for each `Anthropic.ToolUseBlock`:
     `callbacks.onToolCall({ id, name, input })`; `const t0 = Date.now()`;
     - `!isToolName(name)` -> `result = { ok: false, error: \`Unknown tool ${name}\` }`.
     - else if `WRITE_TOOLS.has(name) && !options.agentMode && !(await callbacks.confirmWrite(...))` ->
       `result = { ok: false, error: 'The user declined this write. Do not retry it; ask what they prefer.' }` with `declined = true`.
     - else `result = await withTimeout(executeTool(name, block.input), LIMITS.TOOL_TIMEOUT_MS)` where a timeout yields
       `{ ok: false, error: \`${name} timed out after ${LIMITS.TOOL_TIMEOUT_MS / 1000}s\` }` (Promise.race; do not cancel the underlying call).
     - `callbacks.onToolResult({ id, name, input, result, ms: Date.now() - t0 })`.
     - `toolResults.push({ type: 'tool_result', tool_use_id: block.id, content: capToolResultText(result.ok ? serializeToolData(name, result.data) : JSON.stringify({ error: result.error })), is_error: !result.ok && !declined })`.
       (A declined write is NOT `is_error`: the model should respond gracefully, not treat it as a failure.)
   - If `options.signal?.aborted` at any point after a request: stop with `'aborted'`.
5. Abort handling: catch `Anthropic.APIUserAbortError` (or `signal.aborted`) -> stopReason `'aborted'`. Before returning, if the last message in
   `history` is an assistant message containing `tool_use` blocks without a following user message, append a user message with one
   `tool_result` per block: `{ type: 'tool_result', tool_use_id, content: 'cancelled by user', is_error: true }` so the transcript stays valid for the next turn.
6. Other `Anthropic.APIError`s propagate (the UI renders `err.status` and `err.message`); non-API errors propagate too.
7. Iterations exhausted -> stopReason `'max_iterations'`.
8. Return `{ history, finalText: text, stopReason, model, usage: { inputTokens, outputTokens, cacheReadTokens, cacheWriteTokens }, toolCalls }`.

Keep `loop.ts` under ~200 lines; small named helpers (`withTailBreakpoint`, `withTimeout`, `closeDanglingToolUse`) are preferred over comments.

## 4. `src/agent/loop.test.ts` (vitest, streaming: false, fake client)

Fake client: `{ messages: { create: vi.fn() } } as unknown as Anthropic`, where `create` returns canned `Anthropic.Message` objects
(`{ id, type: 'message', role: 'assistant', model, content, stop_reason, stop_sequence: null, usage: { input_tokens, output_tokens, cache_read_input_tokens, cache_creation_input_tokens } }`; cast as needed).
Executor: `vi.fn()` returning `{ ok: true, data: {...} }`. Callbacks: `vi.fn()`s; `confirmWrite` resolves as each test needs.

Cases (>= 8):
- (a) text-only turn: `finalText` equals the text, `history.length === input.history.length + 2`, `onText` called once, `stopReason 'end_turn'`.
- (b) one `tool_use` (`read_range`) then `end_turn`: executor called with the parsed input object; the second request's last message is a user message
  with ONE `tool_result` whose `content` contains the serialized page (`mode=values`); `toolCalls === 1`.
- (c) two `tool_use` blocks in one assistant message -> exactly ONE user message containing two `tool_result` blocks, in the same order as the tool_use ids.
- (d) executor returns `{ ok: false, error: 'boom' }` -> the tool_result has `is_error: true` and content contains `boom`.
- (e) `agentMode: false`, `write_range` requested, `confirmWrite` resolves false -> executor NOT called; content mentions "declined"; `is_error` is false.
- (f) a fake that always returns `tool_use` stops after `LIMITS.MAX_TOOL_ITERATIONS` requests with stopReason `'max_iterations'`.
- (g) `buildRequest` output: `thinking.type === 'adaptive'`, `output_config.effort === 'medium'` (and `'high'` + `model === MODEL_CONFIG.reasonModel` when `reason: true`),
  `system[0].cache_control.type === 'ephemeral'`, the last block of the last message has `cache_control`, no `temperature`, `JSON.stringify(params)` does not contain `budget_tokens`,
  and the last text block of the user message starts with `<workbook_context>`.
- (h) an aborted signal after the first request (`AbortController.abort()` inside the executor mock) -> stopReason `'aborted'` and the returned history ends with a user message whose tool_result content is `cancelled by user`.
- (i) executor that never resolves + a tiny timeout (mock `LIMITS.TOOL_TIMEOUT_MS` via `vi.spyOn`/`vi.useFakeTimers`) -> tool_result `is_error` with "timed out". If mocking the constant is awkward, expose an optional internal `toolTimeoutMs` parameter on a non-exported helper and test that helper directly.
- (j) `max_tokens` with no tool_use -> a second request is made whose last message text is `Continue exactly where you stopped.`; a second `max_tokens` ends the turn with stopReason `'max_tokens'`.

## 5. `scripts/agentSmoke.ts` (real API, fake workbook; `npm run smoke`)

```ts
import 'dotenv/config';
import { createNodeClient } from '../src/agent/client';
import { runAgentTurn } from '../src/agent/loop';
import type { ToolExecutor, CellValue, WorkbookOverview, RangePage, ... } from '../src/agent/contracts';
```
Implement `class FakeWorkbook` holding `Map<sheet, Map<'A1', CellValue>>` seeded with sheet `Sales` (A1:B7 = header `Month`,`Revenue` then 6 rows
Jan..Jun with numbers 120, 135, 150, 160, 155, 170) and sheet `Assumptions` (A1:B3 = `Growth`,0.05 / `Tax`,0.25 / `Discount`,0.1), and
`executeTool: ToolExecutor` supporting `get_workbook_overview` (computed from the maps), `read_range` (whole sheet as one page; `mode 'formulas'`
returns stored strings starting with `=` in `formulas`), `search_workbook` (substring over `String(value)`), `write_range` (stores the grid at the
address, single-cell anchor expands; `readBack` = the stored grid; `errorCells: []`; `previous` = old cells), `add_sheet`; `summarize_range` and
`get_selection` return `{ ok: false, error: 'not available in the smoke harness' }`. Keep it ~120 lines; correctness over completeness.

Run two turns with `streaming: true`, `agentMode: true`, `reason: false`, `contextPreamble = formatOverviewForPrompt(overview)` on turn 1 and
`formatSelectionLine(overview)` on turn 2, printing streamed text to stdout and one line per tool call (`name`, ms, ok):
1. `"Which sheets are in this workbook and what is in Sales?"` -> assert the answer mentions `Sales`.
2. `"Put a SUM of Sales!B2:B7 into Sales!B8 with the label \"Total\" in Sales!A8."` -> assert `fake.get('Sales', 'B8')` is a string starting with `=` and `A8 === 'Total'`.
After turn 2 print `usage` and assert `usage.cacheReadTokens > 0` (the tools + system prefix is > 1024 tokens and was written by turn 1 within the 5-minute TTL);
on failure print the usage object and exit 1. Exit 0 with `SMOKE OK` otherwise. Total cost is well under $0.05.

## 6. Verify, commit

```bash
npx tsc --noEmit -p tsconfig.json
npm test
npm run smoke          # needs .env in this worktree; if ANTHROPIC_API_KEY is missing, say so and let the human run it
grep -n 'budget_tokens\|temperature' src/agent/loop.ts        # must print nothing
grep -c cache_control src/agent/loop.ts                       # >= 2 (system + tail)
git add src/agent/systemPrompt.ts src/agent/history.ts src/agent/history.test.ts src/agent/loop.ts src/agent/loop.test.ts scripts/agentSmoke.ts
git commit -m "T4: agent loop, system prompt, history trimming, tests, live smoke"
```

## Acceptance

`npm test` passes with >= 8 loop cases and >= 3 history cases; `npx tsc --noEmit` passes; the two greps above hold; `npm run smoke` prints
`SMOKE OK` with `cacheReadTokens > 0` on turn 2 (run by you if `.env` is present, otherwise by the human immediately after merge).

## When done, report

Final message: files; test counts; the smoke output (streamed text summary, tool calls, usage of both turns); anything in the request
shape the live API rejected (exact 400 message) and what you changed; `CONTRACT CHANGE REQUEST` if any.
