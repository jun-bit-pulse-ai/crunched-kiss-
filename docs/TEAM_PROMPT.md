# Crunched KISS — 3-hour trunk sprint, three agents

Send **section A** to all three agents. Then send each agent only its own block from **section B**. Section C is Jun's checklist.

Issues #2–#5 already exist on `jun-bit-pulse-ai/crunched-kiss-` with these owners; this prompt keeps their numbering and adds #6.

---

## A. Common prompt — paste to Kimi, Cursor and Claude Code

You are one of three AI agents (Kimi, Cursor, Claude Code) finishing the Crunched KISS Excel add-in in the next 3 hours. Repo: `jun-bit-pulse-ai/crunched-kiss-`, trunk branch `main`. `FINAL_PLAN.md` is the source of truth. The draft plans are archives, except that `CLAUDE_MACBOOK_PLAN.md` holds already-verified code you may adapt where your issue points at it.

Jun is the integrator. Jun merges every pull request and decides anything that crosses issue boundaries.

You own exactly one GitHub issue. Its "Exclusive files" list is a hard boundary.

**Rules**

1. **Work in your own git worktree**, at the path your issue gives, on the branch your issue names. Never run git commands in `/Users/junseki/Documents/GitHub/crunched-kiss`; that shared checkout is Jun's. Never commit or push to `main`, never force-push, never touch another agent's branch.
2. **Edit only your Exclusive files.** If you need a change elsewhere, do not make it. Comment the exact diff you want on your issue, then continue with the rest of your work.
3. **Never** rename or move files, reformat code you did not change, add dependencies (only issue #4 adds `openpyxl`), or touch `.env`, `certs/`, `FINAL_PLAN.md`, or the draft plans. Do not commit `scripts/big.xlsx`.
4. **One machine, one Excel, two ports.** Only the owner of issue #6 runs `npm run dev-server`, `npm start`, or uvicorn, and only they sideload into Excel. Everyone else proves their work with unit tests. Live verification happens in the integration phase.
5. **Tests before every push:** `cd backend && .venv/bin/pytest -q` and `cd frontend && npm test`. Both must be green. Add a unit test for every pure function you add: mocha with ts-node in `frontend/test`, pytest in `backend/tests`. Office.js code is not unit-tested; keep the Office.js calls in `excel.ts` and the pure logic in `excelPolicy.ts` so the logic is testable.
6. **The tool-name contract** binds `backend/app/tools.py` to `dispatchExcelTool` in `frontend/src/app/services/excel.ts`. Only issue #3 changes it, and changes both sides in one pull request.
7. **Commits:** small, one concern each, `feat|fix|docs|test: ... (#N)`.
8. **Pull request:** branch off `origin/main`. Before opening, and again before Jun merges, run `git fetch origin && git rebase origin/main` and re-run both suites. The body says `Closes #N`, lists the files you touched, and states exactly how you verified. Open a draft PR by T+0:45, mark it ready by T+1:30. Comment your status on the issue at T+0:45 and when the PR is ready.
9. **Blocked more than 15 minutes?** Comment on the issue with what you tried, then move to your next checkbox. Never widen your scope to get unblocked.
10. **Done means:** PR ready, rebased on `origin/main`, both suites green, every changed file inside your list. Then stop. Do not pick up another issue unless Jun assigns it.

**One-time setup in your worktree**

```bash
cd <your worktree path>
cp /Users/junseki/Documents/GitHub/crunched-kiss/.env .env
cd backend && uv venv .venv && uv pip install -r requirements.txt && cd ..
cd frontend && npm ci && cd ..
cd backend && .venv/bin/pytest -q && cd ../frontend && npm test
```

Both suites must be green before you change anything. Baseline right now is 11 backend and 9 frontend tests passing.

**Timeline**, where T is when Jun sends this

| Window | What happens |
|---|---|
| T+0:00 | Jun commits the finished proxy work and pushes `main`. Everyone branches from that commit. |
| T+0:00–1:30 | Build. Draft PR by T+0:45, ready by T+1:30. |
| T+1:30–1:50 | Jun squash-merges in order: #3, then #4, then #5. Each rebased, each green. |
| T+1:50–2:40 | Issue #6: integration and the live demo in Excel on `main`. |
| T+2:40–3:00 | Final README pass, tick FINAL_PLAN, push, submit. |

---

## B. Per-agent blocks

### Cursor — issue #2 is already finished; take issue #5

Issue #2 (proxy `/api`) is **done and committed by Jun** as of T+0:00. The pane now calls the relative `/api/chat`, webpack proxies `/api` to `http://127.0.0.1:8000`, uvicorn runs without TLS, and `backend/app/main.py` serves both the old and the `/api`-prefixed routes. Do not redo it. Your job is issue #5.

**Issue #5 — README is the walkthrough**

Worktree `/Users/junseki/Documents/GitHub/ck-issue-5`, branch `issue-5-readme`.

```bash
git -C /Users/junseki/Documents/GitHub/crunched-kiss worktree add /Users/junseki/Documents/GitHub/ck-issue-5 -b issue-5-readme origin/main
```

Exclusive files: `README.md` only.

- [ ] What it is: two or three sentences.
- [ ] Setup on a Mac, in runnable order: `.env` at the repo root with `ANTHROPIC_API_KEY`; `./scripts/setup-certs.sh`; `./scripts/dev-backend.sh` (plain HTTP on 127.0.0.1:8000, no certificate flags any more); `cd frontend && npm run dev-server`; `npm start` to sideload; and the fallback of copying `manifest.xml` into `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` then restarting Excel.
- [ ] Architecture: the diagram from FINAL_PLAN "Target shape", plus one paragraph on why the agent loop lives in the pane and why there is now a single HTTPS origin.
- [ ] The tool table: `list_workbook_meta` (with header preview), `read_range`, `write_range`, `get_selection`, `find`. Write it for the post-#3 state and say so if #3 has not merged yet.
- [ ] Any workbook size: meta first, `find` to locate labels, the 2,000-cell read cap, 8 tool rounds then `force_text`, and `scripts/make_big_workbook.py` for the 1M-cell fixture.
- [ ] Tests: `cd backend && .venv/bin/pytest -q`, `cd frontend && npm test`, and one line on why Office.js and the React pane are hand-checked instead.
- [ ] What was cut: copy FINAL_PLAN "Out of scope" and add streaming and write-confirm.
- [ ] The 15-minute demo script from FINAL_PLAN, and an honest short time log.

Acceptance: someone with a fresh clone and an API key can follow it start to finish. No code files touched.

### Claude Code — issue #3

**Issue #3 — `find` tool and header preview**

Worktree `/Users/junseki/Documents/GitHub/ck-issue-3`, branch `issue-3-find`.

```bash
git -C /Users/junseki/Documents/GitHub/crunched-kiss worktree add /Users/junseki/Documents/GitHub/ck-issue-3 -b issue-3-find origin/main
```

Exclusive files: `backend/app/tools.py`, `backend/app/agent.py` (the `SYSTEM_PROMPT` string only, never the routing), `backend/tests/test_tools.py`, `frontend/src/app/services/excel.ts`, `frontend/src/app/excelPolicy.ts`, `frontend/test/excelPolicy.test.ts`, and new test files under `frontend/test/`.

- [ ] `tools.py`: add `find` with inputs `query` (string, required), `sheet`, `match_case`, `complete_match` (all optional). The description must say it returns at most 50 addresses and should be used to locate labels before reading. Extend the `list_workbook_meta` description to mention the header preview.
- [ ] `excel.ts`: implement `find` using `worksheet.findAllOrNullObject(query, { completeMatch, matchCase })` across the named sheet or all sheets, loading `cellCount,areas/items/address`; return the total plus up to 50 addresses, or a clear "no matches" line. Register `find` in `dispatchExcelTool`. In `listWorkbookMeta`, add a `headerPreview`: the first row of `getUsedRangeOrNullObject(true)`, capped at 20 columns via `getAbsoluteResizedRange(1, n)`.
- [ ] `excelPolicy.ts`: put the caps and their pure helpers here (`MAX_FIND_RESULTS = 50`, `HEADER_PREVIEW_COLS = 20`, formatting and slicing), and test them in `excelPolicy.test.ts`.
- [ ] `test_tools.py`: assert `find` is in `tool_names()` and that its schema requires `query`.
- [ ] `agent.py`: add one sentence to `SYSTEM_PROMPT` telling the model to use `find` to locate labels and then read only the block it needs.
- [ ] Reference: `CLAUDE_MACBOOK_PLAN.md` Task 4 Step 6 has a verified implementation of both executors on the same Office.js calls, in a different file layout. Adapt it; do not paste its structure.

Acceptance: both suites green, `find` present on both sides of the contract, no server run.

### Kimi — issue #4

**Issue #4 — 1M-cell workbook fixture**

Worktree `/Users/junseki/Documents/GitHub/ck-issue-4`, branch `issue-4-fixture`.

```bash
git -C /Users/junseki/Documents/GitHub/crunched-kiss worktree add /Users/junseki/Documents/GitHub/ck-issue-4 -b issue-4-fixture origin/main
```

Exclusive files: `scripts/make_big_workbook.py` (new), `backend/requirements.txt` (add `openpyxl>=3.1`, nothing else), `.gitignore` (add `scripts/*.xlsx`).

- [ ] The script writes `scripts/big.xlsx` with openpyxl in `write_only=True` mode: sheet `Data`, 5,000 rows by 200 columns, header `Metric_1` through `Metric_200` and a numeric body; sheet `Budget` with header `Line item, 2024, 2025, 2026` and rows Revenue, COGS, Gross profit (`=B2-B3`, `=C2-C3`, then a hard-coded `1000` for 2026 as a planted error), Margin (`=B4/B2` and so on), and `Per unit` dividing by empty cells so it yields `#DIV/0!`.
- [ ] Make the output path independent of the working directory: resolve it from the script's own location, so it works from anywhere.
- [ ] Verify: from a clean venv, `uv pip install -r backend/requirements.txt` then run the script. It should print `wrote scripts/big.xlsx` in roughly 3 seconds and produce a file of about 6 MB. Confirm `git status` does not show `big.xlsx`.
- [ ] Reference: `CLAUDE_MACBOOK_PLAN.md` Task 7 Step 1 has a verified version of this script.

Acceptance: the script runs from a clean venv, the workbook is untracked, backend tests still green.

### Whoever finishes first — issue #6, integration (Jun creates it)

Starts at T+1:50, after Jun has merged #3, #4 and #5. Runs in the shared checkout `/Users/junseki/Documents/GitHub/crunched-kiss` on `main`, and is the only agent allowed to do so.

- [ ] `git checkout main && git pull`; reinstall dependencies if the manifests changed; both suites green.
- [ ] Start the backend and dev server, sideload, generate and open `scripts/big.xlsx`.
- [ ] Run FINAL_PLAN's success criteria and demo script. "How big is this workbook?" must go through meta or `find`, never a full read of `Data`. "Check the Budget sheet for errors" must name the hard-coded `Budget!D4` and the `#DIV/0!` row. One `write_range` of a formula must land in the cell.
- [ ] Fix only small breakages, as direct commits on `main` (`fix: ... (#6)`), tests green before each push. Anything larger becomes a new issue.
- [ ] Tick the now-true FINAL_PLAN checkboxes in one `docs:` commit, and comment on #6 with what passed, what was fixed, and what remains.

---

## C. Jun's checklist

**Before you send anything (about five minutes).** The proxy work for issue #2 is finished but sitting uncommitted in your shared checkout, alongside unrelated edits to `CLAUDE_MACBOOK_PLAN.md`. Commit and push it first, so all three agents branch from the same `main`; otherwise Cursor will redo it and collide with everyone.

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss
git add frontend/webpack.config.js frontend/src/app/services/agentClient.ts frontend/src/app/apiPaths.ts frontend/test/apiPaths.test.ts backend/app/main.py backend/tests/test_http.py scripts/dev-backend.sh FINAL_PLAN.md
git commit -m "feat: proxy API through the add-in origin (#2)"
git push origin main
gh issue close 2 --comment "Proxy landed on main; pane uses /api/chat, uvicorn is plain HTTP."
```

Then create the integration issue and paste each agent's block:

```bash
gh issue create --title "[Integration] Live demo checks on main" --body "Starts after #3, #4 and #5 merge. Owner runs in the shared checkout on main; see TEAM_PROMPT.md section B."
```

**Also worth doing.** `frontend/manifest.xml` still lists an AppDomain for `https://localhost:8000`, which is now unused; harmless, so leave it unless the pane misbehaves. And `backend/app/config.py` defaults to `claude-sonnet-4-6`, the previous generation; the current one is `claude-sonnet-5`. Change it yourself in one commit rather than giving it to an agent, since it is one line in a file nobody owns.

**During the sprint.** Merge in the order #3, #4, #5. Require a rebase and green tests on each. If two PRs somehow touch the same file, merge the smaller one and make the other rebase. Watch for an agent editing outside its list; that is the failure mode this prompt exists to prevent.
