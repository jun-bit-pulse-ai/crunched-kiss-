# Crunched KISS — Final Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:executing-plans or superpowers:subagent-driven-development. Work on `main` only (trunk). Steps use checkbox (`- [ ]`) syntax.

**Goal:** Keep a small, explainable Excel task-pane agent that chats with Claude and reads/writes the live workbook, including huge sheets, and land follow-ups as short commits on `main`.

**Architecture:** The task pane owns the agent loop. FastAPI is one stateless Claude turn. Office.js is the only Excel runtime. The model sees addresses and samples, never a whole used range.

**Tech Stack:** TypeScript, React, Office.js · Python, FastAPI, Anthropic tool-use · desktop Excel + trusted HTTPS.

**Source of truth:** this file. The drafts below are history, not instructions.

| Draft | Keep | Drop |
|---|---|---|
| [CLAUDE_MACBOOK_PLAN.md](CLAUDE_MACBOOK_PLAN.md) | Native `tool_use`, pane-owned loop, `/api` proxy, `find`, 1M-cell demo, hour-2 tracer, TDD on pure logic, cut order, 15-min demo script | Opus-by-default, `yo office` as a hard gate, dual `addin/` rename |
| [IMPLEMENTATION_PLAN.md](IMPLEMENTATION_PLAN.md) | Excel boundary, desktop-first / no Vercel, Mac WEF sideload, trunk (no long-lived branches), success criteria | JSON-in-text actions, one-shot `context` dict, `CORS *`, hardcoded key, four agents racing on the same contract, E2E last |
| [ai-orchestrator-plan.md](ai-orchestrator-plan.md) | Timeouts, tool allowlist, max rounds, cheap-vs-expensive routing as a later idea | Building a Claude Code router; it is a different product |

**Already shipped** on `feat/excel-taskpane-agent` (merge this to `main`): chat pane, `list_workbook_meta` / `read_range` / `write_range` / `get_selection`, 2,000-cell cap, 8-round limit, `.env` API key, backend + frontend unit tests, two HTTPS origins (`office-addin-dev-certs`).

---

## Trunk workflow

All work lands on `main` in small commits. No long-lived feature branches after this merge.

- One concern per commit. Tests green before push.
- If two people must touch the same file, finish that file before the second starts.
- Draft plans stay in the repo as archives. Change behavior only by updating this file, then the code.
- Do not commit `.env`, `certs/*.pem`, `node_modules`, or `.venv`.

```
main  ← only lasting branch
├── frontend/   Office add-in
├── backend/    FastAPI + Anthropic
├── scripts/    certs, icons, later big.xlsx
└── FINAL_PLAN.md
```

---

## Locked decisions

1. **Loop in the pane.** Backend is `POST /chat` (or `/api/chat` after the proxy task). Tools only run inside Excel’s WebView.
2. **Native tool use.** No JSON parsed from prose.
3. **One HTTPS origin for the pane (next).** Webpack proxies `/api` to plain-HTTP uvicorn on `127.0.0.1:8000`. Until that lands, both servers share the `certs/` pair.
4. **Never read wholesale.** Overview/meta first; `read_range` capped at 2,000 cells; add `find` so the model locates labels instead of scanning.
5. **Key on the server.** `ANTHROPIC_API_KEY` in repo-root `.env` only.
6. **Tests where logic is pure.** Policy, history truncation, `/chat` contract (mocked client). Office.js and the React pane are hand-checked in Excel.
7. **Sonnet by default.** Swap via `ANTHROPIC_MODEL`. Opus is optional, not required for the demo.

---

## Target shape

```
Excel WebView  https://localhost:3000
  Chat UI  →  runAgent()  →  excel.ts (Office.js)
       │ fetch /api/chat
       ▼
webpack (HTTPS, trusted)  ──proxy──▶  uvicorn :8000 (HTTP, localhost)
                                          │
                                          ▼
                                    Anthropic tool-use
```

Closed tool list:

| Tool | Role |
|---|---|
| `list_workbook_meta` (shipped) / `get_workbook_overview` (upgrade) | Shape + used ranges; later add header preview |
| `read_range` | Values + formulas; cap 2,000 cells |
| `write_range` | 2D values; validate before `Excel.run` |
| `get_selection` (shipped) | Active range + small preview |
| `find` (next) | Up to 50 addresses for a label |

Hard cap: 8 tool rounds per user send, then `force_text`.

---

## Follow-up tasks on `main`

Do these in order. Each task is one commit. Skip anything already true in the tree.

### Task 1: Merge the add-in onto trunk

- [x] Fast-forward / merge `feat/excel-taskpane-agent` into `main`.
- [x] Confirm `cd backend && .venv/bin/pytest -q` and `cd frontend && npm test`.
- [x] Close the feature branch after it is on `main`.

### Task 2: One origin — proxy `/api`

**Files:** [frontend/webpack.config.js](frontend/webpack.config.js), [frontend/src/app/services/agentClient.ts](frontend/src/app/services/agentClient.ts), [backend/app/main.py](backend/app/main.py), [scripts/dev-backend.sh](scripts/dev-backend.sh)

- [x] Point the pane at relative `/api/chat` (or keep `/chat` and proxy that path).
- [x] Proxy `/api` (or `/chat`, `/health`) to `http://127.0.0.1:8000`.
- [x] Run uvicorn **without** TLS on localhost; only webpack stays HTTPS.
- [x] `curl -sk https://localhost:3000/api/health` returns `{"ok":true}`.
- [x] Commit: `feat: proxy API through the add-in origin`

### Task 3: `find` + richer overview

**Files:** [backend/app/tools.py](backend/app/tools.py), [frontend/src/app/services/excel.ts](frontend/src/app/services/excel.ts), [frontend/test/excelPolicy.test.ts](frontend/test/excelPolicy.test.ts)

- [x] Add `find(query, sheet?)` — max 50 addresses.
- [x] Extend meta/overview with a first-row header preview (cap columns).
- [x] Keep names in sync between Python schemas and `dispatchExcelTool`.
- [x] Commit: `feat: find tool and header preview for large books`

### Task 4: Prove any size

**Files:** Create [scripts/make_big_workbook.py](scripts/make_big_workbook.py)

- [x] Generate a ~1M-cell Data sheet plus a small Budget sheet with a planted formula error.
- [x] Hand-check in Excel: "How big is this workbook?" must not `read_range` the whole Data sheet.
- [x] Commit: `feat: 1M-cell workbook fixture`

### Task 5: README = walkthrough

**Files:** [README.md](README.md)

- [x] Setup: `.env`, certs, `dev-backend.sh`, `npm run start`, Mac WEF fallback (`~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`).
- [x] Diagram, tool list, what was cut, 15-minute demo script (overview → Budget errors → one write).
- [x] Commit: `docs: setup, architecture, demo script`

---

## Out of scope (do not build)

- Multi-tier Claude Code orchestrator as this product
- LangGraph, streaming, conversation persistence, auth, Vercel
- Reason / Agent Mode toggles
- Write-confirm dialogs (first real-user follow-up, not the interview bar)
- Excel Online as the primary host

---

## Success criteria

- [x] `main` has the working add-in; no required long-lived branch
- [x] Task pane opens in desktop Excel over trusted HTTPS
- [x] "What is in A1?" and "Write Hello to B1." work with a live key
- [x] A huge used range does not send the whole range to the model
- [x] Backend tests pass without Excel
- [x] README is enough for the hiring call

---

## 15-minute call

1. Open a large book (or `scripts/big.xlsx` after Task 4). Ask how big it is. Point at one meta/overview call.
2. Error-check a small model sheet. Show `include_formulas` / formula fields.
3. Write one formula. Show `write_range` and the cell.
4. Walk `agentClient.ts` (loop), `agent.py` (one turn), `tools.py` (contract). Why Excel never lives in Python.
5. Name what is deliberately missing.
