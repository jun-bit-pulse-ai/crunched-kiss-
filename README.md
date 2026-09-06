# Crunched KISS

A 4-hour, code-quality-first Excel task-pane agent. Chat lives in the add-in. Claude proposes tools. Office.js is the only Excel runtime. Large workbooks stay addressable because the model sees **addresses and samples**, never the whole used range.

## Architecture

```
Excel task pane (React + Office.js)
        │  HTTPS
        ▼
FastAPI  POST /chat   (one model turn)
        │
        ▼
Anthropic tool-use
        │
        ▼
Pane executes list_workbook_meta / read_range / write_range / get_selection
        │
        ▼
Tool results go back on the next POST /chat
```

The backend never touches Excel. That avoids COM, workbook uploads, and CORS-to-Excel. It is the same boundary as production Crunched (Office.js frontend, Python agent backend) without LangGraph.

**Rejected alternatives**
- Frontend-only Anthropic calls put the API key in the WebView.
- LangGraph is the real Crunched orchestrator; a single explicit turn is easier to test and walk through in 15 minutes.
- Dumping the used range into the prompt fails the “any size workbook” requirement.

## Tool loop

`POST /chat` is one Claude turn. The pane owns history and Excel round-trips.

- Response is exactly one of `tool_calls`, `message`, or `error`.
- Hard cap: **8 tool rounds** per user send, then `force_text`.
- Reads over **2,000 cells** are truncated in `frontend/src/app/excelPolicy.ts` (enforced) and documented in `backend/app/tools.py`.
- Older tool payloads in history are stubbed so the context window does not grow without bound.

## Setup

Prerequisites: Node 20–24, Python 3.12+, desktop Excel, and either [mkcert](https://github.com/FiloSottile/mkcert) or Microsoft’s `office-addin-dev-certs`.

```bash
# 1. Certificates (Excel WebView refuses HTTP)
brew install mkcert && mkcert -install   # recommended
./scripts/setup-certs.sh                 # mkcert if present, else office-addin-dev-certs

# 2. Icons (already generated; re-run if needed)
python3 scripts/make_icons.py

# 3. Backend
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
cp ../.env.example ../.env               # add ANTHROPIC_API_KEY
../scripts/dev-backend.sh                # HTTPS on https://localhost:8000

# 4. Add-in
cd frontend
npm install
npm run validate
npm test
npm run start                            # webpack HTTPS + sideload into desktop Excel
```

If `npm run start` does not attach to Excel, upload `frontend/manifest.xml` via **Insert → Add-ins → Upload My Add-in**, with `npm run dev-server` already running.

Open a workbook, send “What is in A1?” then “Write Hello to B1.”

## Tests

```bash
cd backend && .venv/bin/pytest -q
cd frontend && npm test
```

Backend tests mock Anthropic. They do not need Excel or a live API key.

## Layout

- `frontend/` — Office add-in. `excel.ts` is the only Office.js wrapper.
- `backend/app/agent.py` — the only LLM call.
- `backend/app/tools.py` — closed tool list and the 2,000-cell policy.
- `certs/` — shared HTTPS pair for webpack and uvicorn (gitignored).

The Yeoman generator (`yo office`) refuses odd-numbered Node (this machine is Node 23). The add-in follows the official [Office-Addin-TaskPane-React](https://github.com/OfficeDev/Office-Addin-TaskPane-React) webpack + manifest layout, stripped of sample ribbon logic.

## What was cut

Streaming, conversation persistence, LangGraph, web research, scenario dashboards, Reason/Agent Mode toggles, Excel Online as the primary host. Formula repair is limited to what `read_range` already returns.

## Walkthrough notes

The interesting design is the Excel boundary, not the chat chrome. Start there: why the backend returns tool calls, why metadata is O(sheets), and why the pane enforces the cell cap even if the model asks for `A:XFD`.
