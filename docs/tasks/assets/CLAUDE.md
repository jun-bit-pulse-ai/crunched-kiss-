# CLAUDE.md — house rules for every agent working in this repo

This repo is an Excel task-pane add-in (React + TypeScript + Office.js, webpack 5) with a Claude
tool-use agent running inside the pane. Several Claude Code agents work in parallel on disjoint files.
These rules are enforced by the orchestrator at merge time; a violation means your branch is rejected.

## Files
- Edit ONLY the files listed under "You own" in your task prompt. Create no other files.
- NEVER edit: `src/agent/contracts.ts`, `src/agent/client.ts`, `package.json`, `package-lock.json`,
  `manifest.xml`, `tsconfig.json`, `vitest.config.ts`, `CLAUDE.md`, `docs/**`, `.env`, `.env.example`.
- If you believe the contract must change, do NOT change it. Finish everything else, then put a
  section titled `CONTRACT CHANGE REQUEST` in your final message with: the exact diff you need,
  why, and what breaks without it. The orchestrator is the only person who edits the contract.
- Do not add, remove or upgrade npm dependencies. Everything you need is already installed
  (`@anthropic-ai/sdk`, `react`, `vitest`, `exceljs`, `dotenv`, `tsx`, `typescript`, Office.js types).
- Do not delete files you do not own, even if they look unused.

## Processes
- NEVER run `npm start`, `npm run stop`, `npm run dev-server`, `office-addin-debugging`, or anything
  that opens Excel or a browser. The human owns the dev server and Excel.
- NEVER run `npm install`, `npm i`, `npm ci`, `npm uninstall` (T1 is the single exception and says so).
- You may run: `npx tsc --noEmit -p tsconfig.json`, `npm test`, `npx vitest run <file>`,
  `npx webpack --mode development`, `node scripts/...`, `npm run smoke` (only if your task says so).

## Definition of done
- `npx tsc --noEmit -p tsconfig.json` passes for the files you own. If tsc reports errors ONLY in
  files you do not own (another agent's half-finished branch), do not fix them; list them in your
  final message and continue.
- `npm test` passes (your tests plus the existing ones).
- Your final message contains: (1) files created/changed, (2) how you verified, (3) any deviation
  from the task prompt, (4) `CONTRACT CHANGE REQUEST` if needed, (5) anything the human must do
  next (restart the dev server, run a command in Excel's Web Inspector, ...).
- Commit on your task branch when done: `git add <your files> && git commit -m "<Txx>: <title>"`.
  Do not push. Do not merge. Do not touch other branches.

## API facts (do not deviate; the Claude API changed recently)
- Models: `claude-sonnet-5` (default), `claude-opus-5` (Reason toggle). Exact strings, no date suffix.
- Thinking: `thinking: { type: 'adaptive' }` ONLY. `budget_tokens` returns 400. Never send `temperature`.
- Effort: `output_config: { effort: 'low'|'medium'|'high'|'xhigh'|'max' }`.
- No assistant prefill (400). Parse tool inputs from `block.input` (already an object); never string-match.
- All `tool_result` blocks of one assistant turn go back in ONE user message.
- Use SDK types (`Anthropic.MessageParam`, `Anthropic.Tool`, `Anthropic.ToolUseBlock`,
  `Anthropic.ToolResultBlockParam`, `Anthropic.Message`); do not redefine them.
- Strict tools: every property in `required`, `additionalProperties: false`, no `minimum`/`maximum`.
- Prompt caching: `cache_control: { type: 'ephemeral' }` on the system block; tools + system first and stable.

## Office.js reminders
- Every Excel call is `await Excel.run(async (ctx) => { ...; await ctx.sync(); })`. Load only the
  properties you need (`range.load('values,valueTypes')`), batch loads, then one `ctx.sync()`.
- Use `*OrNullObject` variants (`getUsedRangeOrNullObject(true)`, `findAllOrNullObject`,
  `getItemOrNullObject`, `getIntersectionOrNullObject`) and check `.isNullObject` after sync.
- Large ranges: read in row blocks via `getRangeByIndexes(row, col, rowCount, colCount)`; never load a
  whole used range's values. Indexes are 0-based; A1 addresses are 1-based.
- Writing `range.formulas = grid`: strings starting with `=` become formulas, other values are stored
  as plain values, `null` leaves the cell unchanged, `''` clears it.
- Error cells: `valueTypes[r][c] === 'Error'` (not a string prefix check).
- `Range.address` comes back as `Sheet!A1:B2` (sheet quoted with single quotes when it has spaces).
- Office.js is single-threaded per pane: execute tool calls sequentially, never `Promise.all` over `Excel.run`.
