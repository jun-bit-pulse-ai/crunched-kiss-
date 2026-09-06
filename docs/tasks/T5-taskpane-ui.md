# T5 — Task pane chat UI (Crunched look, streaming, tool cards, toggles, confirm card, mock mode)

**Agent model:** claude-sonnet-5. **Estimate:** 45 min. **Depends on:** T1. **Runs in parallel with:** T2, T3, T4, T6.
**Worktree:** `/Users/junseki/Documents/GitHub/ck-T5` on branch `task/T5`.

## Context

You are working in an Excel task-pane add-in (React 18 + TypeScript + webpack 5). Read `CLAUDE.md` first, then
`src/agent/contracts.ts` (FROZEN): it defines `ChatItem`, `TurnMeta`, `INTRO_TEXT`, `SUGGESTIONS`, `AgentCallbacks`,
`RunAgentTurn`, `AgentTurnOutput`, `ToolExecutor`, `ToolResult`, `WorkbookOverview`, `WriteResult`, `formatOverviewForPrompt`,
`overviewSignature`, `formatSelectionLine`. Also read `src/agent/client.ts` (`createClient()`, `MODEL_CONFIG.streaming`) and
`src/excel/devConsole.ts` (`installDevConsole()`), both final. Other agents are writing `src/agent/loop.ts` (exports
`runAgentTurn: RunAgentTurn`) and `src/excel/tools.ts` (exports `executeTool: ToolExecutor`); stubs with those exact
exports exist, so the build passes. Never run `npm start`/`npm install`. Do not add dependencies: plain CSS, no UI library.

The scaffold's `src/taskpane/taskpane.html` has `<div id="container"></div>` and loads office.js from the CDN; do not edit it.

## You own (create or replace exactly these)

- `src/taskpane/index.tsx` (replaces the T1 stub)
- `src/taskpane/mockTurn.ts`
- `src/taskpane/taskpane.css`
- `src/taskpane/components/App.tsx`, `ChatPanel.tsx`, `MessageBubble.tsx`, `ToolCallCard.tsx`, `Composer.tsx`, `ConfirmCard.tsx`

Do NOT touch: `src/agent/**`, `src/excel/**`, `src/taskpane/taskpane.html`, `src/commands/**`, `webpack.config.js`, `package.json`, `manifest.xml`, `docs/**`.

## 1. `src/taskpane/index.tsx`

```tsx
import { createRoot } from 'react-dom/client';
import App from './components/App';
import './taskpane.css';
import { installDevConsole } from '../excel/devConsole';
import { createClient, MODEL_CONFIG } from '../agent/client';
import { runAgentTurn } from '../agent/loop';
import { executeTool } from '../excel/tools';
import { mockExecuteTool, mockRunAgentTurn } from './mockTurn';
```
Mount inside `Office.onReady((info) => ...)`. `const mock = new URLSearchParams(window.location.search).has('mock') || !info.host;`
Mock mode -> `<App runTurn={mockRunAgentTurn} executeTool={mockExecuteTool} client={null} streaming={true} mock />`;
real mode -> `installDevConsole()` then `<App runTurn={runAgentTurn} executeTool={executeTool} client={createClient()} streaming={MODEL_CONFIG.streaming} />`.
Also start a 3-second fallback timer: if `Office.onReady` has not fired by then and `?mock=1` is in the URL, mount in mock mode anyway
(office.js may not resolve outside a host). Guard against double mounting.

## 2. `src/taskpane/mockTurn.ts`

`export const mockExecuteTool: ToolExecutor` returns a canned `WorkbookOverview` for `get_workbook_overview` (3 sheets: `Data` 100001x12
with a header row and 3 sample rows, `Assumptions` 21x3, `Summary` 7x2; `activeSheet 'Summary'`, `selection 'Summary!B4'`), a canned
`WriteResult` for `write_range` (echo the input shape, `previous` of the same shape filled with `''`, `errorCells: []`), and
`{ ok: false, error: 'mock: not implemented' }` otherwise.
`export const mockRunAgentTurn: RunAgentTurn` scripts one realistic turn honouring `signal` (reject/stop on abort) with `setTimeout` delays:
stream "Looking at the Summary sheet." word by word (60 ms/word) -> `onToolCall` `search_workbook` -> 400 ms -> `onToolResult` ok with
2 hits -> `onToolCall` `read_range` -> 500 ms -> `onToolResult` ok -> if `!options.agentMode`: `await callbacks.confirmWrite(write_range event)`
and on decline stream "Skipped the write." else stream "Wrote Summary!B10." -> stream a final 3-bullet answer. Return
`{ history: [...input.history, userMsg, assistantMsg], finalText, stopReason: 'end_turn', model: options.reason ? 'claude-opus-5' : 'claude-sonnet-5', usage: { inputTokens: 2310, outputTokens: 180, cacheReadTokens: input.history.length ? 2100 : 0, cacheWriteTokens: input.history.length ? 0 : 2100 }, toolCalls: 2 }`.

## 3. `App.tsx` (default export `App`)

Props: `{ runTurn: RunAgentTurn; executeTool: ToolExecutor; client: Anthropic | null; streaming: boolean; mock?: boolean }`.
State: `items: ChatItem[]` (initial: one assistant item with `INTRO_TEXT`, `streaming: false`), `history: Anthropic.MessageParam[]`,
`reason` (false), `agentMode` (true), `running`, `abort: AbortController | null`, `pendingConfirm: { e: ToolCallEvent; resolve: (ok: boolean) => void } | null`,
`lastSignature: string | null`.

`send(text)`:
1. Push a user item and a streaming assistant item (`text: ''`). `running = true`; new `AbortController`.
2. Preamble: `const ov = await executeTool('get_workbook_overview', {})`. If ok: `const sig = overviewSignature(ov.data)`;
   `contextPreamble = sig !== lastSignature ? formatOverviewForPrompt(ov.data) : formatSelectionLine(ov.data)`; set `lastSignature = sig`.
   If not ok: `''` (swallow; the model can call the tool itself).
3. `const t0 = Date.now(); const out = await runTurn({ client: client as Anthropic, history, userText: text, contextPreamble, options: { reason, agentMode, streaming, signal }, executeTool, callbacks })`.
   Callbacks: `onText` appends to the CURRENT streaming assistant item; `onToolCall` inserts a tool item and then a NEW streaming assistant item
   after it (so text after a tool call gets its own bubble; drop empty assistant items on completion); `onToolResult` fills `result`/`ms` on the tool
   item with the same id; `confirmWrite` returns a Promise resolved by ConfirmCard through `pendingConfirm`.
4. On success: `history = out.history`; set `meta` on the last non-empty assistant item: `{ model: out.model, usage: out.usage, toolCalls: out.toolCalls, ms: Date.now() - t0 }`;
   `console.debug('[crunched] turn', out.model, out.usage, out.stopReason)`. If `out.stopReason === 'max_iterations'` append an error item
   "Stopped after 25 tool calls; ask me to continue."; `'refusal'` -> "The model declined this request."; `'aborted'` -> "Stopped."; `'max_tokens'` -> "Answer was cut off; ask me to continue."
5. On catch: error item with `err.status ? \`${err.status}: ${err.message}\` : err.message` (do not lose the history: keep the previous `history`).
6. Finally `running = false`, `abort = null`.
Stop button: `abort.abort()`. Undo: `undoWrite(item)` calls `executeTool('write_range', { sheet, address, values: previous })` from a tool item whose
result is an ok `WriteResult` (see ToolCallCard) and appends an error-styled info item "Restored Sheet!range" or the failure.

## 4. Components

- `ChatPanel.tsx`: renders `items` (MessageBubble for user/assistant/error, ToolCallCard for tool), the ConfirmCard when `pendingConfirm`,
  a "Thinking…" pulsing indicator on the streaming assistant item while its text is empty (this covers the silent pause of adaptive thinking),
  and the `SUGGESTIONS` chips under the intro until the first user message (click -> `send(chip)`). Auto-scroll to bottom on change.
- `MessageBubble.tsx`: minimal markdown (paragraphs, lines starting with `- ` or `• ` as bullets, `` `code` `` spans, `**bold**`); blinking caret while
  streaming; for assistant items with `meta`, a small muted footer: `claude-sonnet-5 · 2.3k in (2.1k cached) · 180 out · 2 tools · 4.1 s`.
- `ToolCallCard.tsx`: header `read_range Data!A1:L166 · 312 ms` (label from input: `sheet`+`address`, or `query`, or `name`), status dot
  (running / ok / error), collapsed by default; expanded shows pretty JSON of `input` and of `result.data` or `result.error`, capped at 4000 chars.
  For an ok `write_range` result with non-empty `previous`, show an `Undo` button (calls `onUndo(item)`); disabled while `running`.
- `Composer.tsx`: textarea with placeholder `Ask me anything...` (Enter sends, Shift+Enter newline), Send / Stop button, two pill toggles
  `Reason` and `Agent Mode` with `aria-pressed`, hint under the toggles: Reason on -> `uses claude-opus-5 (slower, pricier)`; Agent Mode off ->
  `writes ask for confirmation`. Disabled while running (Stop stays enabled).
- `ConfirmCard.tsx`: tool name + short input summary (`write_range Summary!B10 (1x2)` / `add_sheet "Dashboard"`) + `Apply` / `Skip` buttons.
- `taskpane.css`: full-height flex column for a ~350 px pane; header `Crunched` + subtitle `your AI analyst in Excel`; scrolling message
  list; user bubbles right/dark, assistant left/light, error items amber, tool cards mono 12 px with chevron; composer pinned at the bottom;
  minimal `prefers-color-scheme: dark` support (backgrounds/text). Use the IBM Plex fonts the html already loads with system fallbacks.

## 5. Verify, commit

```bash
npx tsc --noEmit -p tsconfig.json
npx webpack --mode development && rm -rf dist
git add src/taskpane && git commit -m "T5: task pane chat UI + mock mode"
```
You cannot open a browser. The human verifies mock mode at `https://localhost:3000/taskpane.html?mock=1` in Safari.

## Acceptance

- `npx tsc --noEmit` and `npx webpack --mode development` pass on your branch with the T1 stubs.
- Human, in Safari at `?mock=1`: intro + 3 chips render; clicking a chip fills and sends; text streams word by word; two tool cards flip from
  running to a result with ms and expand to JSON; with Agent Mode off the ConfirmCard appears and Skip/Apply change the answer; Stop mid-stream
  ends the turn with "Stopped."; Reason toggle shows the hint; the usage footer appears under the answer; nothing overflows horizontally at 320 px width.
- After T3/T4 merge, in Excel: a real streamed answer with tool cards and the usage footer (`cached` > 0 on the second message).

## When done, report

Final message: files; the `App` props signature; how `?mock=1` and the no-host fallback work; anything you assumed about `loop.ts`/`tools.ts`
beyond their exported names; `CONTRACT CHANGE REQUEST` if any.
