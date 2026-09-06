# T8 — README + docs/ARCHITECTURE.md (setup, reasoning, trade-offs, diagrams, decisions table)

**Agent model:** claude-sonnet-5. **Estimate:** 25 min. **Depends on:** T7 (its final message is your input). **Runs in:** main tree, `main`, alone.

## Context

You are documenting a finished take-home in `/Users/junseki/Documents/GitHub/crunched-kiss`: an Excel task-pane add-in where a chat agent
reads and writes the open workbook, built in 4 hours by one human orchestrating Claude Code agents. Read `CLAUDE.md`, then, before writing
anything: `README.md` (holds the original exercise text, restored in T0; KEEP it verbatim under a final heading `## The exercise`),
`docs/PLAN.md`, `docs/TIMELOG.md`, `src/agent/contracts.ts`, `src/agent/loop.ts`, `src/agent/client.ts`, `src/agent/systemPrompt.ts`,
`src/excel/tools.ts`, `webpack.config.js`, `src/taskpane/components/App.tsx`, `package.json`, `.env.example`, `git log --oneline`, and the T7
report the human pastes into your prompt (fixes, scenario results, timings, known limitations). Do NOT modify any source file, `package.json`,
or `docs/PLAN.md`. Do not invent commands: every command you write must exist in `package.json` scripts or be a plain `npx`/`node` invocation
you verified. Do not claim a feature works unless the T7 report says so.

## You own

- `README.md` (rewrite; keep the exercise text at the bottom)
- `docs/ARCHITECTURE.md` (new)

## README.md (target 250-400 lines, plain and honest, no marketing)

1. **What this is** (2-3 sentences) + `![task pane](docs/screenshot.png)` placeholder line with a `<fill in: screenshot>` marker.
2. **Setup (macOS desktop Excel)**: prerequisites (Node 20-24, Excel for Mac, an Anthropic key); `npm install`; `cp .env.example .env` and paste the key;
   `npm start` (explain the one-time Keychain prompt from office-addin-dev-certs and that it sideloads and launches Excel); where the add-in appears
   (Home > Crunched; or Insert > Add-ins > My Add-ins > Developer Add-ins); `npm run stop`. Also `npm run fixture` (builds `fixtures/big-model.xlsx`,
   1.2M cells; `--rows 10000 --out fixtures/small-model.xlsx` for a lighter one), `npm test`, `npm run smoke` (live API, fake workbook, no Excel).
   Troubleshooting bullets: `npx office-addin-dev-certs verify` / `install --machine`; Safari check of `https://localhost:3000/taskpane.html`;
   clear `~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/*`; `defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true`
   for Inspect Element; `?mock=1` to see the UI without Excel; Excel Online upload as an alternative host (untested).
3. **Architecture**: one paragraph on the core bet (Office.js only exists in the pane, so the tool loop lives in the pane; the only secret is the
   key, held by a ~20-line proxy inside webpack-dev-server; therefore one process, one cert, one language, no mkcert, no CORS, no relay protocol).
   A short mermaid sequence diagram (copy the one from docs/ARCHITECTURE.md, trimmed). A "Why not FastAPI/LangGraph" paragraph: what it would
   have cost in 4 h (second process, second cert, action-relay RPC, session state, second language) and when it becomes right (web search,
   persistence, auth, multi-user, checkpointing) — and note the pane code would not change because the loop already talks to a `baseURL`.
4. **Tools and size guards**: a table built from `contracts.ts` (name, purpose, guard from `LIMITS`); then "How it handles workbooks of any size"
   as a numbered list: context preamble (full map on turn 1 / when sheets change, one line after), paged reads (2000 cells, used-range
   intersection, 500-row syncs), search before read, summarize_range (O(columns) output, 20k-cell chunks), Excel as the calculator (formula into a
   scratch cell), result caps (200-char cells, 30k-char results), write caps + read-back error detection via valueTypes, 25 iterations, history
   trimming. Include the measured timings from T7 on the fixture (`<fill in>` if missing).
5. **Model configuration**: env vars and defaults (claude-sonnet-5, effort medium; Reason toggle -> claude-opus-5, effort high), adaptive thinking,
   `output_config.effort`, why there is no `temperature`/`budget_tokens`/prefill, prompt caching (system + moving tail breakpoint; the footer shows
   cached tokens; a model switch busts the cache by design), streaming flag, request/tool timeouts.
6. **Security notes**: key never in the bundle (`curl -sk https://localhost:3000/taskpane.js | grep -c sk-ant` = 0), why `dangerouslyAllowBrowser`
   is safe here, `DIRECT_ANTHROPIC_KEY` as a documented demo-only escape hatch, production path (deploy the same proxy behind auth and rate limits;
   manifest URLs change; pane code does not), Agent Mode off = confirmation on every write, Undo from the write card (snapshot of the overwritten range).
7. **Process**: contract-first parallel build (link `docs/PLAN.md`, `docs/tasks/`, `docs/TIMELOG.md`), which agent built what, what the human did
   by hand (T0, T7 driving Excel, prompt tuning), earlier drafts in `docs/history/` "superseded because they added a second process".
8. **Trade-offs and cuts**: be specific from the T7 report and `docs/TIMELOG.md`; **What I would do next**: compaction/context editing for long
   sessions, formatting tool (number formats, column widths), a proper undo stack, diff preview before writes, evals for the error-check playbook
   against seeded fixtures, Excel Online CI, backend proxy with per-user auth and audit log, streaming thinking summaries in Reason mode.
9. **Testing**: what is unit-tested (ranges/paging math, loop with a fake client, history trimming, contract serializers), the live smoke, and the
   manual Excel checklist (the eight T7 scenarios).
10. **General thoughts** — a clearly marked section `<fill in: human writes this>`; leave it empty.
11. **The exercise** — the original text, unchanged.

## docs/ARCHITECTURE.md

(a) mermaid `sequenceDiagram` of one user turn: User -> TaskPane -> executeTool(get_workbook_overview) -> runAgentTurn -> proxy (/api/anthropic)
-> Anthropic API -> tool_use -> executeTool -> Office.js -> Excel -> tool_result -> ... -> end_turn -> UI; (b) mermaid `flowchart` of the size
strategy (context -> search -> summarize -> paged read -> write <= 5000 -> read-back/errorCells -> Excel-computed aggregates); (c) a **Decisions**
table with columns Decision / Why / Trade-off / Production path for: loop in the pane; proxy in webpack-dev-server; reuse of the existing scaffold
instead of `yo office`; enforced size guards + summarize_range; Sonnet 5 default with Opus 5 Reason; strict tools without nullable/minimum;
contract-first parallel agents in worktrees; (d) a file map with one line per source file; (e) known limitations from T7.

## Verify

- Every command in README exists: cross-check against `package.json` scripts and `.env.example`.
- `grep -c '```mermaid' docs/ARCHITECTURE.md` >= 2; README renders (no unclosed code fences).
- `git add README.md docs/ARCHITECTURE.md && git commit -m "T8: README and architecture doc"`.

## When done, report

Final message: sections written; every `<fill in>` marker left for the human (with line numbers); any claim you were unsure about and left out.
