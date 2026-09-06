# Time log (human-filled, 15-minute windows)

Honest record of the 4-hour window. Fill the "Actually did" column as you go; leave "Status" as
`green` / `yellow` (behind, cut list consulted) / `red` (cut applied). Started at: `<HH:MM>`.

| Window | Planned (see docs/PLAN.md) | Actually did | Status |
|---|---|---|---|
| 0:00-0:15 | T0a: archive two-process draft, promote `frontend/` to root, `npm install`, `.env`, Safari cert check | | |
| 0:15-0:30 | T0b: `npm start`, pane opens in desktop Excel, commit; launch T1 | | |
| 0:30-0:45 | Verify + merge T1; create worktrees; launch T2, T6, T3, T4, then T5 | | |
| 0:45-1:00 | Merge T6 (build fixture), merge T2 (restart server, curl + in-pane SDK check) | | |
| 1:00-1:15 | Merge T4 (`npm test`, `npm run smoke`); check T5 UI at `?mock=1` in Safari | | |
| 1:15-1:30 | Merge T3, T5; full check; GATE: all 7 tools return ok in Web Inspector on the fixture | | |
| 1:30-1:45 | T7 integration: first real chat; fault table | | |
| 1:45-2:00 | T7: scenarios 1-3 (overview, summarize, error check) | | |
| 2:00-2:15 | T7: scenarios 4-8 (formula write, confirm gate, Reason, Stop, cache); DECISION POINT; launch T8 | | |
| 2:15-2:30 | Prompt-tune with the 3 chips; record timings for README; T8 running | | |
| 2:30-2:45 | Read README against reality; write "General thoughts" | | |
| 2:45-3:00 | Code-quality pass (loop.ts, tools.ts, App.tsx); optional `npm run smoke` re-run | | |
| 3:00-3:15 | Fresh-clone validation + demo rehearsal on a fresh Excel launch | | |
| 3:15-3:30 | Buffer | | |
| 3:30-3:45 | Final README polish, push, add collaborators markusskagemo + larsgmu | | |
| 3:45-4:00 | Hard stop on code. Walkthrough prep. | | |

Cuts applied (in order, from docs/PLAN.md "Cut list"): `<none>`
