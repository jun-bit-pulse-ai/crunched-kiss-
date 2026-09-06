# T7 — Integration: run in Excel with the human, fix first-contact bugs, prove the scenarios

**Agent model:** claude-sonnet-5, paired with the human who drives desktop Excel. **Estimate:** 45 min. **Depends on:** T2, T3, T4, T5, T6 all merged into `main`.
**Runs in:** the main working tree `/Users/junseki/Documents/GitHub/crunched-kiss` on `main`. You are the ONLY agent now; no worktree.

## Context

All parallel branches are merged; `npx tsc --noEmit && npm test && npx webpack --mode development` pass on `main`; the dev server is running
(`npm start`, owned by the human); `fixtures/big-model.xlsx` exists. Read `CLAUDE.md`, then `src/agent/contracts.ts`, `src/agent/loop.ts`,
`src/excel/tools.ts`, `src/taskpane/components/App.tsx`, `webpack.config.js`. You may edit ANY source file to fix integration bugs, but keep
changes minimal and local; do not restructure or rename. Never run `npm start`/`npm run stop`: when `webpack.config.js` changes, ask the human
to run `npm run stop; npm start`; for other changes webpack live-reload usually refreshes the pane, otherwise the human closes and reopens the pane
(ribbon button) or, as a last resort, quits Excel, clears `~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/*` and restarts.

`src/agent/contracts.ts` may be changed ONLY if both sides of an interface must change together; say so explicitly in your final message.

## Procedure

### Step 0 — Static gate (2 min)
```bash
npx tsc --noEmit -p tsconfig.json && npm test && npx webpack --mode development && rm -rf dist
grep -rn 'budget_tokens\|temperature' src/agent/ ; echo "(must be empty)"
```

### Step 1 — Tool gate in Web Inspector (already run by the human at 1:15; re-run anything that failed)
Ask the human to open `fixtures/big-model.xlsx`, open the pane, right-click > Inspect Element, and paste the ten `crunched.run(...)`
commands from `docs/tasks/T3-excel-tools.md` (Acceptance). Every command must return `{ ok: true, ... }` except the intentional duplicate
`add_sheet`. Fix `src/excel/tools.ts` / `ranges.ts` until they do. Record the timings of `get_workbook_overview`, `read_range` page 0 and
`summarize_range` on `Data` (the human puts them in the README later).

Null-write probe (Office.js semantics are asserted, not verified): after the `write_range` of `[[1,2],['=D1+E1','=1/0']]` at `Summary!D1`,
run `write_range` `{ address: 'D1:E2', values: [[null, ''], [null, null]] }` then `read_range` `Summary!D1:E2`. Expected: D1 still `1`, E1 empty,
row 2 unchanged. If `null` actually CLEARS cells, change the `write_range` description in `contracts.ts` and the "Writing rules" line in
`systemPrompt.ts` to say so (both strings, nothing else) and report it.

Error-text probe: if `search_workbook` for `#DIV/0!` returned 0 hits, keep the tool (it still finds labels) and confirm `summarize_range` on
`Summary` lists B4 and B5 in `errorCells`; add one sentence to the system prompt's playbook saying search may miss error cells and
summarize_range is authoritative.

### Step 2 — First real chat (10 min)
Human sends `What is in this workbook?` with Agent Mode on, Reason off. Expected: streamed answer within ~10 s naming Data, Assumptions,
Summary and the 30 Scenario sheets, with NO `read_range` of the whole Data sheet (the preamble already carries the map). Diagnose from the
Web Inspector console and Network tab; the human pastes errors to you. Fault table:

| Symptom | Likely cause | Fix |
|---|---|---|
| 401/403 from `/api/anthropic` | proxy hook did not set the key | check `onProxyReq` vs `on.proxyReq` for the installed http-proxy-middleware; `.env` has the key; restart server |
| Request goes to api.anthropic.com directly | `DIRECT_ANTHROPIC_KEY` set, or baseURL wrong | clear the env var; `createClient()` baseURL must be `${origin}/api/anthropic` |
| 400 mentioning `budget_tokens`, `temperature`, prefill | stale request shape | remove the param (`grep -rn` in `src/agent`) |
| 400 on tool schema (`strict`, `additionalProperties`, `minimum`) | schema not strict-compatible | every property in `required`, `additionalProperties:false`, no `minimum`; worst case drop `strict: true` on the offending tool |
| Nothing streams, then the whole answer appears | proxy/dev-server buffering | `compress: false` in devServer; else set `CRUNCHED_STREAMING=false` in `.env` and restart |
| `Thinking…` forever, then a timeout | SSE broken inside WKWebView | `CRUNCHED_STREAMING=false` (loop switches to `messages.create`) |
| `InvalidArgument` on `getRangeByIndexes` | off-by-one between A1 and 0-based indexes | check `parseAddress`/`pageBounds` with the failing address in vitest |
| `ItemNotFound` from `findAll` | non-OrNullObject call | use `findAllOrNullObject` |
| `InvalidOperation` when reading a null intersection | used range empty | handle `isNullObject` before loading properties |
| Cell text like `[object Object]` | rich values | `clampCell` -> `String(v)`; acceptable |
| Answer ignores the workbook / calls overview again | preamble empty | `executeTool('get_workbook_overview')` failed in App; check console |
| Pane blank after an edit | stale bundle | close/reopen pane; else clear the Excel cache and restart |

### Step 3 — Scenarios (25 min). All with `fixtures/big-model.xlsx` open. Tick each in your final message.

1. **Overview**: `What is in this workbook?` -> names the sheets; no read of the whole Data sheet; usage footer shows the model and tokens.
2. **Size**: `Summarize the Data sheet: how many rows, average price, any errors?` -> one `summarize_range` call, answer gives ~100,000 rows and an
   average price near 252; no `read_range` paging (a couple of pages at most).
3. **Error check (formula)**: `Error check the Summary sheet` -> reports `Summary!B4` (#DIV/0!), `Summary!B5` (#REF!), and `Summary!B7` (partial range
   `SUM(Data!F2:F5001)` vs the full column); bonus: notices `Data!F5001` / `Data!F77701` hard-coded 12345 via search for `12345`.
4. **Formula write**: `Write the average of Data!E as a formula into Summary!B10 with the label "Avg price" in A10` -> one `write_range` card,
   `cellsWritten 2`, `errorCells []`, readBack shows the number; Excel shows a formula in B10; the card's Undo restores the empty cells.
5. **Confirm gate**: Agent Mode off, repeat scenario 4 at B11 -> ConfirmCard appears; Skip -> the model does not retry and asks; Apply -> written.
6. **Reason**: Reason on, `Which Scenario sheet has the highest total profit?` -> Network request body `model: "claude-opus-5"`; the footer shows
   `cached 0` on this turn (model switch busts the cache; expected) and the turn completes.
7. **Stop + continuity**: send `Read every page of the Data sheet` -> after 2-3 tool cards press Stop -> "Stopped." -> send `What was the last page you read?`
   -> the model answers (history stayed valid: the dangling tool_use got `cancelled by user` results).
8. **Cache**: on the second message of a conversation (same model) the footer shows `cached` > 0.

If a scenario fails, fix and re-run it. If it still fails after 10 minutes, apply the cut list in `docs/PLAN.md` (in order) and record the cut in `docs/TIMELOG.md`.

### Step 4 — Commit
```bash
npx tsc --noEmit -p tsconfig.json && npm test
git add -A && git commit -m "T7: integration fixes (see message body)" -m "<one line per fix>"
```

## Acceptance

Scenarios 1-5 and 7-8 pass in desktop Excel on the 1.2M-cell fixture; scenario 6 passes or Reason is cut per the cut list; `npm test` green;
the commit message lists every fix; timings recorded for `get_workbook_overview`, `read_range` page, `summarize_range` on Data.

## When done, report

Final message: the list of fixes (file + one line each); the scenario checklist with pass/fail; the three timings; whether `null` leaves cells
unchanged; whether `findAll` matched `#DIV/0!`; any cut applied; anything the README must state as a known limitation.
