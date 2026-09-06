# Crunched KISS

An Excel task-pane agent for a four-hour take-home. You chat in the sidebar; Claude asks for workbook data through tools; Office.js is the only Excel runtime. Large sheets stay usable because the model sees addresses and samples, never a whole used range.

**Key idea:** The AI never holds the spreadsheet. It requests small pieces (a sheet list, one range, a search result), reasons about them, optionally writes back, and replies in plain English. This works on a 1-million-cell workbook as well as a small one — Claude is never shown more than a few thousand cells at a time.

---

## Features

- **Chat interface** inside Excel's task pane — ask questions in natural language
- **Tool-use loop** — Claude reads, searches, and writes cells through structured tools
- **Clarifying questions** — when the model is unsure, it presents 2–4 multiple-choice buttons instead of making you type
- **Tool visibility** — every Excel operation appears as a card in the chat thread so nothing happens invisibly
- **Large-sheet safe** — reads are capped at 2,000 cells; metadata is O(sheets), not O(cells)
- **Live selection** — the pane shows your current Excel selection so you can say "this table"
- **8-round cap** — per-question tool loop hard-limits at 8 rounds, then forces a text answer

---

## Prerequisites

- **macOS** with Apple Silicon or Intel (tested on macOS 14+)
- **Node.js 20–24** and **npm 9–11**
- **Python 3.12+**
- **Microsoft Excel** desktop (Microsoft 365 / 16.x) — the add-in sideloads into desktop Excel only
- **Anthropic API key** (Claude Sonnet or Haiku)
- Either [**mkcert**](https://github.com/FiloSottile/mkcert) or Microsoft's `office-addin-dev-certs` (the setup script uses whichever is present)

---

## Setup (3 terminals)

### Terminal 1 — Environment & Certs

```bash
# Clone and enter the repo
cd crunched-kiss

# Create your API key file (never commit this)
cp .env.example .env
# Edit .env and set:
#   ANTHROPIC_API_KEY=sk-ant-api03-...
#   ANTHROPIC_MODEL=claude-sonnet-4-6   (or claude-sonnet-5-1)

# Generate trusted HTTPS certificates for Excel's WebView
./scripts/setup-certs.sh
```

> **Why HTTPS?** Excel's WebView requires a trusted origin. Webpack serves the pane on `https://localhost:3000`. The backend itself is plain HTTP on `127.0.0.1:8000` — no cert flags needed there.

### Terminal 2 — Backend

```bash
cd backend
python3 -m venv .venv && source .venv/bin/activate
pip install -r requirements.txt
cd ..
./scripts/dev-backend.sh
```

The backend starts on `http://127.0.0.1:8000`. It exposes one endpoint:

| Endpoint | Purpose |
|---|---|
| `POST /api/chat` | One Claude turn — returns `tool_calls` or a `message` |

### Terminal 3 — Frontend & Excel

```bash
cd frontend
npm install
npm run dev-server   # https://localhost:3000 with HMR
```

In a **fourth** terminal (or after `dev-server` is up):

```bash
cd frontend
npm start            # sideloads into Excel
```

Look for **Crunched** on Excel's Home tab, or go to **Insert → Add-ins → My Add-ins**.

---

## Troubleshooting

| Problem | Fix |
|---|---|
| `npm start` does not show the pane | Copy the manifest manually: `mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/ && cp frontend/manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` then restart Excel. |
| `NET::ERR_CERT_AUTHORITY_INVALID` in Excel | Re-run `./scripts/setup-certs.sh` and trust the generated cert in Keychain Access (Always Trust). |
| Backend says `ANTHROPIC_API_KEY is not configured` | Check `.env` exists at repo root and `ANTHROPIC_API_KEY` is set (not commented). |
| `Module not found` on `npm run dev-server` | Run `npm install` again; some devDependencies may have failed on first install. |
| Excel shows a blank pane | Open Safari → Develop menu → inspect the WebView. Check the console for CORS or cert errors. |
| Changes don't appear after git pull | The add-in caches aggressively. Click the `...` on the task pane → **Reload**, or restart Excel. |

---

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│  Excel Desktop (macOS)                                      │
│  ┌─────────────────────────────────────────────────────┐   │
│  │  Task Pane WebView  https://localhost:3000          │   │
│  │  • React + TypeScript                               │   │
│  │  • Office.js (excel.ts) — the only Excel wrapper    │   │
│  │  • runAgent() loops: user → Claude → tool → repeat  │   │
│  └─────────────────────────────────────────────────────┘   │
│                           │                                 │
│                           │ fetch /api/chat                 │
│                           ▼                                 │
│  ┌─────────────────────────────────────────────────────┐   │
│  │  webpack dev-server (HTTPS, trusted cert)           │   │
│  │  proxies /api/* ────────────────────────────────────────┐│
│  └─────────────────────────────────────────────────────┘  ││
└────────────────────────────────────────────────────────────│┘
                                                             │
                                                             ▼
                                               ┌─────────────────────────┐
                                               │  uvicorn :8000 (HTTP)   │
                                               │  FastAPI — one stateless│
                                               │  Claude turn per call   │
                                               │  • agent.py — LLM call  │
                                               │  • tools.py — schemas   │
                                               └─────────────────────────┘
                                                            │
                                                            ▼
                                               ┌─────────────────────────┐
                                               │  Anthropic API          │
                                               │  tool_use → tool_result │
                                               │  loop (max 8 rounds)    │
                                               └─────────────────────────┘
```

**Why stateless?** The backend holds no conversation state. The pane sends the full message history each time. This means you can restart the backend mid-chat and nothing breaks.

**Why single-origin?** Webpack proxies `/api` to uvicorn, so the WebView talks to one HTTPS origin. No CORS headers, no second trusted certificate, no `Access-Control-Allow-Origin` gymnastics.

---

## Tools

Each tool the model uses becomes a visible card in the chat thread. Nothing is invisible.

| Tool | What it does | Limit |
|---|---|---|
| `list_workbook_meta` | Sheet names, used-range addresses, dimensions, header preview | O(sheets), not O(cells) |
| `read_range` | Values and formulas for one address | 2,000 cell cap |
| `write_range` | Writes a 2D array back to the sheet | Validated before `Excel.run` |
| `get_selection` | The user's currently selected range | Small preview only |
| `find` | Locate a label by text search | Max 50 addresses |

**Hard cap:** 8 tool rounds per user message, then the pane forces a text-only reply.

---

## How the clarifying questions work

When the model detects ambiguity, the system prompt instructs it to ask a single multiple-choice question formatted like:

```
Which sheet would you like to work with?

A) Budget — the 7-row financial model
B) Data — the 5,000-row metrics table
C) A new sheet
```

The frontend parses `A) …` `B) …` lines and renders them as clickable buttons. Clicking sends the option text back as a user message. If the model doesn't format with lettered options, it falls back to plain text — no breakage.

---

## Guided tour and Explain like I'm 5

The first time the pane opens in a session, a short skippable tour points at the chat thread, the demo chips, the composer, peek cards, and **New chat**. Press Esc to skip, or use Next / Back. The composer stays usable — the tour never disables it. “Tour seen” lives in `sessionStorage`, so a new Excel session (or a new tab) gets the tour again. **Show tour** in the header replays it for a hiring-call reviewer.

**Explain like I'm 5** (header) restates the latest assistant reply and any peek cards in tiny words. That formatter is local and deterministic — it maps tool names to kid sentences (`list_workbook_meta` → “I peeked at the sheet names… I did not read every cell.”) and strips Markdown from the last answer. It does not call Anthropic, so the backend contract is unchanged.

Neither feature needs Excel to start: the pane mounts without `Office.onReady`, and the tour runs from there.

---

## Development

```bash
# Backend tests (pytest)
cd backend && .venv/bin/pytest -q

# Frontend tests (mocha + ts-node)
cd frontend && npm test

# Generate a large test workbook
python3 scripts/make_big_workbook.py   # → scripts/big.xlsx (gitignored)

# Regenerate add-in icons
python3 scripts/make_icons.py
```

### Key files

| File | Role |
|---|---|
| `frontend/src/app/App.tsx` | Root component: message state, send loop, selection watcher |
| `frontend/src/app/services/agentClient.ts` | Fetch loop: handles tool_calls ↔ tool_result for up to 8 rounds |
| `frontend/src/app/services/excel.ts` | The only Office.js wrapper — all Excel I/O goes through here |
| `frontend/src/app/components/ChatThread.tsx` | Renders messages, tool cards, and clarifying questions |
| `frontend/src/app/toolCards.ts` | Human-readable names and summaries for tool cards |
| `backend/app/agent.py` | One Claude turn: system prompt + tool schema + response parsing |
| `backend/app/tools.py` | Tool definitions, schemas, and the 2,000-cell read policy |
| `backend/app/main.py` | FastAPI app: CORS, `/api/chat`, `/api/health` |

---

## General Thoughts / Design Philosophy

### Why KISS won

The original plan included a multi-tier orchestrator, LangGraph, streaming, auth, Vercel deploy, and separate Agent/Reason modes. All of that was cut. What survived is a single Webpack pane, one FastAPI endpoint, and a loop that fits in two files (`agentClient.ts` and `agent.py`).

**The reason it works:** Excel already has a runtime (Office.js). The AI doesn't need to own Excel — it just needs to ask for small pieces of data. The hard part is not the LLM; it's the policy that prevents the model from requesting a million cells. That policy lives in `tools.py` and is enforced in the pane, not the backend.

### What I'd add next (in order)

1. **Undo stack** — Snapshot cells before `write_range`, expose an Undo button. Removes the fear of AI overwriting data.
2. **Suggested follow-ups** — After each reply, show 2–3 chip buttons ("Show formulas", "Highlight errors"). Reuses the same parser/renderer as clarifying questions.
3. **Conversation persistence** — `localStorage` keyed by workbook name so chats survive reloads.
4. **Formula explainer** — Select a cell → "Explain this formula" → Claude breaks it down.

### What was deliberately left out

| Cut | Why |
|---|---|
| Multi-agent orchestrator | One model turn is enough; Excel is the real agent |
| LangGraph / LangChain | Adds complexity without adding capability for 5 tools |
| Streaming responses | Office.js add-ins are small; full messages are fast enough |
| Auth / login | It's a local desktop add-in; the API key is in `.env` |
| Vercel / cloud deploy | The backend must talk to localhost Excel; cloud doesn't help |
| Write-confirm dialogs | Good idea for real users, but adds UI friction for a demo |
| Excel Online primary | Desktop Excel has the full Office.js API; Online is a subset |

### Trunk-based workflow

This repo was built with trunk-based development: short-lived branches (`fix/ux-gaps`, `feat/markdown-rendering`), squash-merged to `main` within minutes, not days. No long-lived feature branches. GitHub Issues were created as pick-up tasks so multiple agents could work in parallel. It worked well for a 4-hour sprint.

---

## 15-minute demo script

1. **Open a large book** (`scripts/big.xlsx`). Ask "How big is this workbook?" Point at the `list_workbook_meta` card — not a full read of the Data sheet.
2. **Error-check Budget.** "Find errors in the Budget sheet." Show that it finds the planted hard-coded cell and `#DIV/0!` via formulas.
3. **Write one formula.** "Add a total row to Budget." Watch the `write_range` card, then show the cell in Excel.
4. **Show the code.** Walk `agentClient.ts` (the loop), `agent.py` (one turn), `tools.py` (the contract). Emphasize: Excel never lives in Python.
5. **Name what's missing.** Undo, persistence, formula explainer — all future work, not interview bar.

---

## Time log (honest)

| Hour | What happened |
|---|---|
| 1 | Sideload + HTTPS certs. Excel's WebView is picky. |
| 2 | Chat UI, Office.js wrappers, and the first `read_range`. |
| 3 | FastAPI backend, tool-use contract, and pytest suite. |
| 4 | One-origin `/api` proxy, `find` tool, README, and the live demo. |
| After | Markdown rendering (#17), history windowing (#18), tool cards (#14), clarifying questions (#21), UX quick wins (#24). |

---

## License

MIT. Built for a take-home interview, kept as a reference implementation.
