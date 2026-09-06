# Claude MacBook Plan — Crunched KISS Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** In four hours, ship an Excel task-pane add-in where a user chats with Claude and Claude reads from and writes to the open workbook, including workbooks with a million cells.

**Architecture:** The React task pane owns the agent loop. It sends the full conversation to a stateless FastAPI endpoint that adds the system prompt and tool schemas, calls Claude once, and returns the raw content blocks. When Claude asks for a tool, the task pane runs it through Office.js (the only place Excel is reachable), appends the `tool_result`, and calls the endpoint again until Claude stops. The backend is a pure function; all Excel state lives in Excel.

**Tech Stack:** TypeScript, React, Office.js (yo office scaffold, webpack dev server) · Python 3.12+, FastAPI, `anthropic` SDK 1.x · Claude Opus 5 with adaptive thinking · vitest + pytest.

**Target machine:** this MacBook. Excel for Mac 16.112 is installed, `uv` is on PATH, `mkcert` is not (and is not needed, see Task 1). Node defaults to 23, which the Office generator rejects; Node 24 is installed under nvm and every `npm` command in this plan runs under it (`nvm use 24`).

**Where it builds:** in a git worktree at `/Users/junseki/Documents/GitHub/crunched-kiss-claude` on branch `claude-macbook-plan`, created in Task 0, because the main checkout is on `feat/excel-taskpane-agent` with a separate implementation of the other plan.

---

## 0. Read this first: the decisions that shape everything

The exercise grades reasoning over features. These are the calls, made up front so the four hours are execution only.

| Decision | Choice | Why | What it costs |
|---|---|---|---|
| Who runs the agent loop | **The task pane.** Backend is one stateless `POST /api/chat`. | Tools can only execute inside Excel's WebView, so every tool call round-trips to the pane regardless. Putting the loop where the tools are removes all session state, WebSockets, and "resume a paused graph" code. Trivially unit-testable on both sides. | Conversation history is re-sent each turn (prompt caching makes this cheap). Streaming text is a stretch goal, not v1. |
| How Claude asks for Excel actions | **Native tool use** (`tool_use` / `tool_result` blocks), 4 tools. | Structured, validated inputs; parallel calls; error results the model can recover from. No JSON-in-prose parsing. | Tool schemas live in Python and executors in TypeScript; they must be kept in sync by hand (documented, 4 tools, acceptable). |
| Certificates and origins | **One HTTPS origin.** Webpack dev server (already trusted by `office-addin-dev-certs`) proxies `/api/*` to a plain-HTTP uvicorn on `127.0.0.1:8000`. | Eliminates the second certificate the README warns about and all CORS. The pane calls a relative URL. | Dev-only convenience; in production you would front both with one reverse proxy anyway. |
| "Workbooks of any size" | **Never read wholesale.** `get_workbook_overview` returns shape + header row only; `find` locates labels; `read_range` is capped at 2,000 cells and returns an error with dimensions so the model pages. | Office.js does the heavy lifting in-process; only small text tables reach the model. A 1M-cell sheet costs one overview call. | The model must plan reads. The system prompt teaches it; the cap error reinforces it. |
| Model | `claude-opus-5`, `thinking: adaptive`, `effort: medium`, top-level `cache_control`. Env-overridable. | Current model; medium effort keeps chat latency reasonable while still reasoning about formulas. Caching covers the growing tool-result history. | Opus latency is a few seconds per step. `CLAUDE_MODEL=claude-sonnet-5` is a one-line swap for snappier demos. |
| Tests | Unit tests only where logic is pure: A1 address maths, table formatting, the agent loop, the endpoint (mocked client). Office.js and React are verified by hand in Excel. | Office.js cannot run outside Excel; mocking it proves nothing. Tests go where bugs actually hide. | UI regressions are caught manually. Stated in the README. |
| Order of work | **Tracer bullet by 2:00**: pane → backend → Claude → cell written. Then harden and polish. | Sideloading, certs, and Office.js quirks are the schedule risks. Hit them first, not in hour four. | Early UI is ugly for an hour. |

**How this differs from `IMPLEMENTATION_PLAN.md` (the Kimi 3 plan):** native tool use instead of parsing JSON from free text; a real multi-step loop with full history instead of a single-shot `context` dict; one origin instead of two certificates; integration at hour 2 instead of hour 4; a current model ID (`claude-3-5-sonnet-20241022` is retired); `find` and formula reads for the error-checking use case; the API key from `.env`, never hardcoded; a scripted 1M-cell workbook so "any size" is demonstrated, not claimed.

---

## 1. Architecture

```
 Excel for Mac (WebView, HTTPS only)
 ┌──────────────────────────────────────────────────────────────┐
 │  Task pane  https://localhost:3000/taskpane.html             │
 │  ┌──────────────┐   ┌──────────────┐   ┌──────────────────┐  │
 │  │ Chat UI      │──▶│ agent/loop   │──▶│ agent/tools      │  │
 │  │ (React)      │◀──│ runTurn()    │◀──│ executeTool()    │  │
 │  └──────────────┘   └──────┬───────┘   └────────┬─────────┘  │
 │                            │ fetch /api/chat     │ Office.js  │
 └────────────────────────────┼─────────────────────┼───────────┘
                              │                     ▼
                webpack dev server            Excel workbook
                (HTTPS, trusted cert)         (any size)
                              │ proxy → http://127.0.0.1:8000
                              ▼
 ┌──────────────────────────────────────────────────────────────┐
 │  FastAPI  POST /api/chat  {messages} → {content, stop_reason} │
 │  agent.complete(): system prompt + 4 tool schemas + Claude    │
 └──────────────────────────────┬───────────────────────────────┘
                                ▼
                        Anthropic API (claude-opus-5)
```

One user turn, step by step:

```
user: "Sum the 2024 column into a Total row"
  pane ──POST {messages:[user]}──▶ backend ──▶ Claude
  ◀── content:[tool_use find("2024")], stop_reason:tool_use
  pane runs find via Office.js → "Budget!B1"
  pane ──POST {messages:[user, assistant(tool_use), user(tool_result)]}──▶ …
  ◀── tool_use read_range(Budget, A1:D6)      → pane reads, appends result
  ◀── tool_use write_range(Budget, A7, [["Total","=SUM(B2:B6)",…]]) → pane writes
  ◀── text:"Added a Total row at Budget!A7:D7", stop_reason:end_turn   → loop ends
```

### Tool contract (the only thing both sides must agree on)

| Tool | Input | Returns (plain text) | Cap |
|---|---|---|---|
| `get_workbook_overview` | none | active sheet, selection, per sheet: used-range address, rows × cols, first-row preview | 20 preview columns |
| `read_range` | `sheet`, `address`, `include_formulas?` | tab-separated table with row numbers and column letters; formulas as `value {=formula}` | 2,000 cells, else error with dimensions |
| `write_range` | `sheet`, `address`, `values[][]` | confirmation with final address | 2,000 cells; single-cell address auto-resizes |
| `find` | `query`, `sheet?`, `match_case?`, `complete_match?` | match count + up to 50 addresses | 50 addresses |

---

## 2. File structure

```
crunched-kiss-claude/              ← git worktree, branch claude-macbook-plan (Task 0)
├── README.md                      ← rewritten in Task 9 (setup + thoughts)
├── CLAUDE_MACBOOK_PLAN.md         ← this file
├── .gitignore
├── scripts/
│   └── make_big_workbook.py       ← 1M-cell test workbook (Task 7)
├── addin/                         ← `yo office` output, trimmed
│   ├── manifest.xml
│   ├── package.json               ← + vitest, + @anthropic-ai/sdk (types only)
│   ├── webpack.config.js          ← + proxy /api → 127.0.0.1:8000
│   ├── vitest.config.ts
│   └── src/taskpane/
│       ├── index.tsx              ← Office.onReady → <App/>
│       ├── taskpane.html
│       ├── taskpane.css
│       ├── api/chat.ts            ← fetch("/api/chat")
│       ├── agent/loop.ts          ← runTurn(): the agent loop (pure, tested)
│       ├── agent/loop.test.ts
│       ├── agent/tools.ts         ← name → executor dispatch
│       ├── excel/a1.ts            ← column letters, address parsing (pure, tested)
│       ├── excel/a1.test.ts
│       ├── excel/format.ts        ← values[][] → tab table with A1 headers (pure, tested)
│       ├── excel/format.test.ts
│       ├── excel/client.ts        ← the 4 Office.js executors (hand-verified)
│       └── components/
│           ├── App.tsx            ← state: messages, busy, error
│           ├── MessageList.tsx    ← renders text / tool_use / tool_result blocks
│           ├── Composer.tsx       ← textarea, Enter to send
│           └── SuggestedPrompts.tsx
└── backend/
    ├── requirements.txt
    ├── .env.example
    ├── app/__init__.py
    ├── app/main.py                ← FastAPI: /api/health, /api/chat
    ├── app/agent.py               ← complete(messages): one Claude call
    ├── app/tools.py               ← the 4 tool schemas + MAX_CELLS
    ├── app/prompts.py             ← SYSTEM_PROMPT
    └── tests/test_chat.py
```

---

## 3. Timeline (240 minutes)

| Clock | Task | Milestone |
|---|---|---|
| 0:00–0:05 | Task 0 — worktree, gitignore | |
| 0:05–0:30 | Task 1 — scaffold, trim, sideload | **M1: pane opens in Excel** |
| 0:30–0:55 | Task 2 — backend endpoint + test + real call | curl gets a Claude reply |
| 0:55–1:05 | Task 3 — proxy `/api` through webpack | pane reaches backend |
| 1:05–1:35 | Task 4 — Excel layer (a1, format, client, dispatch) | |
| 1:35–1:55 | Task 5 — agent loop (TDD) | |
| 1:55–2:25 | Task 6 — minimal chat UI, wire it up | **M2: tracer bullet, cell written from chat** |
| 2:25–2:50 | Task 7 — 1M-cell workbook, harden | **M3: any-size proven** |
| 2:50–3:15 | Task 8 — polish (tool cards, chips, errors, reset) | |
| 3:15–3:45 | Task 9 — README, diagram, trade-offs, push | **M4: submittable** |
| 3:45–4:00 | Buffer — run the demo script end to end | |

**If behind schedule, cut in this order:** suggested-prompt chips → collapsible tool-result cards → autoscroll → `find` tool (overview + read still satisfy "any size"). Never cut the README or M2.

---

## 4. Tasks

Conventions: commit after every green step; commands are run from the path shown; `@superpowers:test-driven-development` applies to Tasks 2, 4, 5; `@superpowers:verification-before-completion` applies before every milestone claim.

### Task 0: Isolate this build from the other implementation (0:00–0:05)

The main checkout at `/Users/junseki/Documents/GitHub/crunched-kiss` is on branch `feat/excel-taskpane-agent`, which holds a separate implementation built from `IMPLEMENTATION_PLAN.md` (`frontend/`, `backend/`, `scripts/`, `certs/`, its own README). This plan also uses `backend/`, so it builds in its own git worktree on its own branch, based on the brief commit `2987b87` (which is also where `main` points). Nothing on the other branch is modified or deleted; the two implementations can be run, compared, or submitted independently.

**Files:**
- Create: `/Users/junseki/Documents/GitHub/crunched-kiss-claude/` (worktree), `.gitignore` inside it

- [ ] **Step 1: Look before creating anything**

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss && git status --short && git branch --show-current && git worktree list
```

Expected: the three plan files untracked, current branch `feat/excel-taskpane-agent`, a single worktree. If a `crunched-kiss-claude` worktree is already listed, skip Step 2.

- [ ] **Step 2: Create the worktree and copy this plan into it**

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss
git worktree add /Users/junseki/Documents/GitHub/crunched-kiss-claude -b claude-macbook-plan 2987b87
cp CLAUDE_MACBOOK_PLAN.md /Users/junseki/Documents/GitHub/crunched-kiss-claude/
```

Expected: the new folder contains only `LICENSE`, the brief `README.md`, and this plan. Every path in the rest of this plan is under that folder.

- [ ] **Step 3: Write `.gitignore`**

```gitignore
node_modules/
dist/
.venv/
__pycache__/
*.pyc
.pytest_cache/
.env
certs/
scripts/*.xlsx
.DS_Store
```

- [ ] **Step 4: Commit**

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude
git add .gitignore CLAUDE_MACBOOK_PLAN.md
git commit -m "chore: add gitignore and the implementation plan"
```

---

### Task 1: Scaffold the add-in, trim it, sideload it (0:05–0:30)

**Files:**
- Create: `addin/` (generated), then edit `addin/manifest.xml`, `addin/src/taskpane/index.tsx`, `addin/src/taskpane/components/App.tsx`
- Delete: `addin/src/taskpane/components/Header.tsx`, `HeroList.tsx`, `TextInsertion.tsx`, `addin/src/taskpane/office-document.ts` (names vary slightly by generator version; delete every sample component except `App.tsx`)

- [ ] **Step 1: Generate the project**

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude
source ~/.nvm/nvm.sh && nvm use 24
npm install -g yo generator-office
yo office
```

`generator-office` refuses odd-numbered Node majors and this Mac defaults to Node 23; v24.13.1 is already installed under nvm. Run `nvm use 24` in every new terminal before any `npm` command in this plan, because global installs and native `node_modules` builds are per Node version.

Answer the prompts: project type **Office Add-in Task Pane project using React framework**; script type **TypeScript**; name **addin**; Office application **Excel**; if asked for a manifest type, choose the **add-in only (XML) manifest**. The generator creates `addin/` and runs `npm install`.

- [ ] **Step 2: Set the display name**

In `addin/manifest.xml`, change `<DisplayName DefaultValue="addin" />` to `<DisplayName DefaultValue="Crunched (Claude)" />` (the other branch's add-in is also called Crunched; the suffix tells them apart in Excel's Add-ins menu) and set `<Description DefaultValue="AI analyst in Excel" />`.

- [ ] **Step 3: Replace the sample UI with a hello pane**

`addin/src/taskpane/index.tsx`:

```tsx
import * as React from "react";
import { createRoot } from "react-dom/client";
import App from "./components/App";

/* global document, Office */

const root = createRoot(document.getElementById("container")!);
Office.onReady(() => {
  root.render(<App />);
});
```

`addin/src/taskpane/components/App.tsx` (temporary, replaced in Task 6):

```tsx
import * as React from "react";

export default function App() {
  return <div style={{ padding: 16 }}>Crunched is alive.</div>;
}
```

Delete the other sample components and `office-document.ts`. Leave `@fluentui` packages installed; removing them is not worth the minutes.

- [ ] **Step 4: Start the dev server and sideload**

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude/addin
npm run dev-server
```

In a second terminal:

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude/addin
npm start
```

`npm start` runs `office-addin-debugging`, which installs a locally trusted CA into the macOS keychain (expect a password prompt: accept), copies `manifest.xml` into `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/`, and launches Excel. This is the README's mkcert trick, done for you by the scaffold. mkcert is only needed if this step fails.

Expected: Excel opens; **Home → Add-ins → Crunched** (or the ribbon button the scaffold adds) opens a pane reading "Crunched is alive."

If the pane is blank: `npx office-addin-dev-certs verify`; if that fails, `npx office-addin-dev-certs install`. If the manifest does not appear, create the `wef` folder by hand, copy `manifest.xml` into it, and restart Excel.

- [ ] **Step 5: Enable the pane's web inspector (debugging for the rest of the session)**

```bash
defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true
```

Restart Excel. Right-click inside the pane → **Inspect Element** gives a Safari console for `console.log` and network errors. Right-click → **Reload** refreshes the pane after code changes.

- [ ] **Step 6: Commit** (M1 reached)

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude
git add addin
git commit -m "feat: scaffold Excel task pane add-in with hello pane"
```

---

### Task 2: Backend endpoint (0:30–0:55)

**Files:**
- Create: `backend/requirements.txt`, `backend/.env.example`, `backend/pytest.ini`, `backend/app/__init__.py`, `backend/app/tools.py`, `backend/app/prompts.py`, `backend/app/agent.py`, `backend/app/main.py`
- Test: `backend/tests/test_chat.py`

- [ ] **Step 1: Environment**

`backend/requirements.txt`:

```
anthropic>=1.0
fastapi>=0.115
uvicorn[standard]>=0.30
pytest>=8
httpx>=0.27
openpyxl>=3.1
```

(`httpx` is for FastAPI's `TestClient`; the SDK itself uses `httpx2`, a separate package, so both coexist. `openpyxl` is for Task 7.)

`backend/.env.example`:

```
ANTHROPIC_API_KEY=sk-ant-...
CLAUDE_MODEL=claude-opus-5
CLAUDE_EFFORT=medium
```

The brief's key line is the placeholder `xxxx[will copy and paste later]`, so the key comes from the hiring contact; if the main checkout's `.env` already holds a working key, copy it from there. Never commit `.env`.

`backend/pytest.ini` (so a plain `pytest` run from `backend/` can import `app`):

```ini
[pytest]
pythonpath = .
testpaths = tests
```

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude/backend
uv venv .venv && source .venv/bin/activate && uv pip install -r requirements.txt
cp .env.example .env   # then paste the API key into .env (see below)
touch app/__init__.py
```

- [ ] **Step 2: Write the failing endpoint test**

`backend/tests/test_chat.py`:

```python
from anthropic.types import Message, TextBlock, Usage
from fastapi.testclient import TestClient

from app import agent
from app.main import app


class FakeMessages:
    def __init__(self, reply):
        self.reply = reply
        self.calls = []

    def create(self, **kwargs):
        self.calls.append(kwargs)
        return self.reply


class FakeClient:
    def __init__(self, reply):
        self.messages = FakeMessages(reply)


def reply_with(text: str) -> Message:
    return Message(
        id="msg_1", type="message", role="assistant", model="claude-opus-5",
        content=[TextBlock(type="text", text=text)],
        stop_reason="end_turn", stop_sequence=None,
        usage=Usage(input_tokens=5, output_tokens=2),
    )


def test_chat_passes_history_through_and_returns_content(monkeypatch):
    fake = FakeClient(reply_with("Hello"))
    monkeypatch.setattr(agent, "get_client", lambda: fake)

    r = TestClient(app).post("/api/chat", json={"messages": [{"role": "user", "content": "hi"}]})

    assert r.status_code == 200
    body = r.json()
    assert body["stop_reason"] == "end_turn"
    assert body["content"] == [{"type": "text", "text": "Hello"}]
    sent = fake.messages.calls[0]
    assert sent["messages"] == [{"role": "user", "content": "hi"}]
    assert [t["name"] for t in sent["tools"]] == [
        "get_workbook_overview", "read_range", "write_range", "find",
    ]


def test_chat_rejects_empty_history():
    r = TestClient(app).post("/api/chat", json={"messages": []})
    assert r.status_code == 422
```

- [ ] **Step 3: Run it, confirm it fails**

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude/backend && pytest -q
```

Expected: `ModuleNotFoundError: No module named 'app.agent'` (or similar import error).

- [ ] **Step 4: Tool schemas**

`backend/app/tools.py`:

```python
"""Tool definitions sent to Claude.

The task pane implements one executor per name in addin/src/taskpane/agent/tools.ts.
Keep names, parameters and MAX_CELLS in sync with that file.
"""

MAX_CELLS = 2000

TOOLS = [
    {
        "name": "get_workbook_overview",
        "description": (
            "Return the workbook layout: every sheet with its used-range address, "
            "row/column counts and its first row (a headers preview), plus the active "
            "sheet and the user's current selection. Cheap even on huge workbooks. "
            "Call this first whenever you do not already know the layout."
        ),
        "input_schema": {"type": "object", "properties": {}, "additionalProperties": False},
    },
    {
        "name": "read_range",
        "description": (
            f"Read cell values (and optionally formulas) from one sheet. At most {MAX_CELLS} "
            "cells per call; a larger request fails and reports the range's dimensions so "
            "you can page through it in smaller blocks. Returns a tab-separated table with "
            "row numbers and column letters so you can cite exact cells."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet": {"type": "string", "description": "Worksheet name, e.g. 'Budget'."},
                "address": {"type": "string", "description": "A1-style range on that sheet, e.g. 'A1:F40'."},
                "include_formulas": {
                    "type": "boolean",
                    "description": "Also return formulas as 'value {=formula}'. Use when checking a model for errors.",
                },
            },
            "required": ["sheet", "address"],
            "additionalProperties": False,
        },
    },
    {
        "name": "write_range",
        "description": (
            f"Write a 2-D array of values into a sheet. At most {MAX_CELLS} cells per call. "
            "Strings beginning with '=' are written as formulas; prefer formulas over "
            "hard-coded numbers for derived values. 'address' must either match the "
            "dimensions of 'values' exactly, or be the single top-left cell, in which case "
            "the target is resized to fit. Overwrites existing content without confirmation."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "sheet": {"type": "string"},
                "address": {"type": "string"},
                "values": {
                    "type": "array",
                    "description": "Rows of cells; outer array is rows.",
                    "items": {
                        "type": "array",
                        "items": {"type": ["string", "number", "boolean", "null"]},
                    },
                },
            },
            "required": ["sheet", "address", "values"],
            "additionalProperties": False,
        },
    },
    {
        "name": "find",
        "description": (
            "Search for text across the workbook (or one sheet) and return the matching "
            "cell addresses (at most 50). Use it to locate labels such as 'Revenue' or "
            "'Total' before reading, instead of scanning large ranges."
        ),
        "input_schema": {
            "type": "object",
            "properties": {
                "query": {"type": "string"},
                "sheet": {"type": "string", "description": "Restrict to one sheet. Omit to search every sheet."},
                "match_case": {"type": "boolean"},
                "complete_match": {
                    "type": "boolean",
                    "description": "Match the whole cell content rather than a substring.",
                },
            },
            "required": ["query"],
            "additionalProperties": False,
        },
    },
]
```

- [ ] **Step 5: System prompt**

`backend/app/prompts.py`:

```python
SYSTEM_PROMPT = """You are Crunched, an AI analyst working inside the user's open Excel workbook.
You can inspect and change the workbook only through the provided tools.

Workbooks can be very large. Never try to read everything:
- If you do not know the layout, call get_workbook_overview first.
- Use find to locate labels, then read_range on just the block you need (at most 2000 cells per call). Page through larger blocks with several reads.
- Ask for include_formulas when the user wants errors, assumptions or logic checked; a hard-coded number where a formula belongs is a finding.

When writing:
- Prefer formulas (strings starting with "=") for derived values.
- Do not overwrite data you have not read unless the user explicitly asked for that cell.
- In your final reply say exactly what changed, as Sheet!Range.

Cite cells as Sheet!A1. Be concise: short sentences, no preamble. Ask a clarifying question only when the request is genuinely ambiguous; otherwise make a sensible choice and state it."""
```

- [ ] **Step 6: The one Claude call**

`backend/app/agent.py`:

```python
import os

import anthropic

from .prompts import SYSTEM_PROMPT
from .tools import TOOLS

MODEL = os.environ.get("CLAUDE_MODEL", "claude-opus-5")
EFFORT = os.environ.get("CLAUDE_EFFORT", "medium")

_client: anthropic.Anthropic | None = None


def get_client() -> anthropic.Anthropic:
    """Lazy so tests can swap it and so import never needs a key."""
    global _client
    if _client is None:
        _client = anthropic.Anthropic()  # reads ANTHROPIC_API_KEY
    return _client


def complete(messages: list[dict]) -> dict:
    """One model turn.

    The caller (the task pane) owns the loop: when stop_reason is "tool_use" it runs
    the tools in Excel, appends the tool_result blocks, and calls again.
    """
    response = get_client().messages.create(
        model=MODEL,
        max_tokens=16000,
        system=SYSTEM_PROMPT,
        tools=TOOLS,
        messages=messages,
        thinking={"type": "adaptive"},
        output_config={"effort": EFFORT},
        cache_control={"type": "ephemeral"},  # cache the growing prefix between loop steps
    )
    # exclude_none keeps the blocks replayable as-is (thinking signatures included).
    return response.model_dump(mode="json", exclude_none=True)
```

- [ ] **Step 7: The app**

`backend/app/main.py`:

```python
from typing import Any

import anthropic
from fastapi import FastAPI, HTTPException
from pydantic import BaseModel, Field

from . import agent

app = FastAPI(title="Crunched KISS backend")


class ChatRequest(BaseModel):
    # An Anthropic Messages API history (text, tool_use and tool_result blocks).
    # The API validates the shape; we pass it through unchanged.
    messages: list[dict[str, Any]] = Field(min_length=1)


@app.get("/api/health")
def health() -> dict:
    return {"ok": True, "model": agent.MODEL}


@app.post("/api/chat")
def chat(req: ChatRequest) -> dict:
    try:
        return agent.complete(req.messages)
    except anthropic.RateLimitError as e:
        raise HTTPException(429, "Rate limited by Anthropic; retry in a moment.") from e
    except anthropic.APIStatusError as e:
        raise HTTPException(502, f"Anthropic API error {e.status_code}: {e.message}") from e
    except anthropic.APIConnectionError as e:
        raise HTTPException(502, "Could not reach the Anthropic API.") from e
```

- [ ] **Step 8: Run tests, confirm they pass**

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude/backend && pytest -q
```

Expected: `2 passed`.

- [ ] **Step 9: Real call with the real key**

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude/backend
uvicorn app.main:app --host 127.0.0.1 --port 8000 --reload --env-file .env
```

In another terminal:

```bash
curl -s http://127.0.0.1:8000/api/chat -H 'Content-Type: application/json' \
  -d '{"messages":[{"role":"user","content":"What is in this workbook?"}]}' | python3 -m json.tool | head -40
```

Expected: `"stop_reason": "tool_use"` and a `tool_use` block named `get_workbook_overview`. That proves the schemas and prompt make Claude reach for the tools.

- [ ] **Step 10: Commit**

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude
git add backend
git commit -m "feat: FastAPI /api/chat that wraps one Claude tool-use turn"
```

---

### Task 3: One origin — proxy `/api` through webpack (0:55–1:05)

**Files:**
- Modify: `addin/webpack.config.js` (the `devServer` block)

- [ ] **Step 1: Check the dev-server major version**

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude/addin && npm ls webpack-dev-server
```

- [ ] **Step 2: Add the proxy**

Inside the existing `devServer: { ... }` object add, for webpack-dev-server 5:

```js
proxy: [{ context: ["/api"], target: "http://127.0.0.1:8000", changeOrigin: true }],
```

or for webpack-dev-server 4:

```js
proxy: { "/api": { target: "http://127.0.0.1:8000", changeOrigin: true } },
```

- [ ] **Step 3: Verify from the pane's origin**

Restart `npm run dev-server`, then:

```bash
curl -sk https://localhost:3000/api/health
```

Expected: `{"ok":true,"model":"claude-opus-5"}`. In the pane's inspector console, `fetch("/api/health").then(r => r.json()).then(console.log)` prints the same, proving the WebView can reach the backend with no second certificate and no CORS.

- [ ] **Step 4: Commit**

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude
git add addin/webpack.config.js
git commit -m "feat: proxy /api to the backend so the pane has one HTTPS origin"
```

---

### Task 4: Excel layer (1:05–1:35)

**Files:**
- Create: `addin/vitest.config.ts`, `addin/src/taskpane/excel/a1.ts`, `a1.test.ts`, `format.ts`, `format.test.ts`, `client.ts`, `addin/src/taskpane/agent/tools.ts`
- Modify: `addin/package.json` (scripts, devDependencies)

- [ ] **Step 1: Test tooling**

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude/addin
npm i -D vitest @anthropic-ai/sdk
```

`@anthropic-ai/sdk` is imported with `import type` only, so it never ships to the pane; it gives exact types for message and block shapes instead of hand-written duplicates.

`addin/vitest.config.ts`:

```ts
import { defineConfig } from "vitest/config";

export default defineConfig({
  test: { include: ["src/**/*.test.ts"] },
});
```

Add to `package.json` scripts: `"test": "vitest run"`.

In `addin/tsconfig.json` add `"skipLibCheck": true` to `compilerOptions`. The scaffold targets ES5 with an old `lib`, and without this flag `tsc` reports errors inside the SDK's own `.d.ts` files (private fields, `AsyncIterable`) even though the plan's code is clean. The scaffold also sets `"pretty": true`, so judge `tsc` by its exit code, not by grepping for `error TS`.

- [ ] **Step 2: Failing tests for A1 maths**

`addin/src/taskpane/excel/a1.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import { colToLetter, letterToCol, topLeft } from "./a1";

describe("a1", () => {
  it("converts column numbers to letters and back", () => {
    expect(colToLetter(1)).toBe("A");
    expect(colToLetter(26)).toBe("Z");
    expect(colToLetter(27)).toBe("AA");
    expect(colToLetter(703)).toBe("AAA");
    expect(letterToCol("AA")).toBe(27);
  });

  it("finds the top-left cell of sheet-qualified and absolute addresses", () => {
    expect(topLeft("'My Sheet'!$B$3:D10")).toEqual({ row: 3, col: 2 });
    expect(topLeft("A1")).toEqual({ row: 1, col: 1 });
  });

  it("rejects garbage", () => {
    expect(() => topLeft("nope")).toThrow(/Cannot parse/);
  });
});
```

Run: `npm test` → expected: FAIL, module `./a1` not found.

- [ ] **Step 3: Implement `a1.ts`**

```ts
/** 1 -> A, 27 -> AA */
export function colToLetter(col: number): string {
  let s = "";
  while (col > 0) {
    const r = (col - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    col = Math.floor((col - 1) / 26);
  }
  return s;
}

export function letterToCol(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n;
}

/** Top-left cell of "Sheet!B3:D10", "'My Sheet'!$B$3", or "B3". */
export function topLeft(address: string): { row: number; col: number } {
  const bang = address.lastIndexOf("!");
  const local = bang >= 0 ? address.slice(bang + 1) : address;
  const first = local.split(":")[0];
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(first);
  if (!m) throw new Error(`Cannot parse address: ${address}`);
  return { row: Number(m[2]), col: letterToCol(m[1]) };
}
```

Run: `npm test` → expected: a1 tests PASS.

- [ ] **Step 4: Failing tests for the table format**

`addin/src/taskpane/excel/format.test.ts`:

```ts
import { describe, expect, it } from "vitest";
import { formatTable } from "./format";

describe("formatTable", () => {
  it("labels rows and columns from the address so the model can cite cells", () => {
    expect(formatTable("Budget!B3:C4", [["Revenue", 100], ["Cost", 40]])).toBe(
      "\tB\tC\n3\tRevenue\t100\n4\tCost\t40",
    );
  });

  it("appends formulas in braces when they differ from the value", () => {
    expect(formatTable("A1", [[140]], [["=B3+B4"]])).toBe("\tA\n1\t140 {=B3+B4}");
  });

  it("renders empty cells as empty strings", () => {
    expect(formatTable("A1:B1", [["", null]])).toBe("\tA\tB\n1\t\t");
  });
});
```

Run: `npm test` → expected: FAIL, module `./format` not found.

- [ ] **Step 5: Implement `format.ts`**

```ts
import { colToLetter, topLeft } from "./a1";

export type Cell = string | number | boolean | null;

/** Tab-separated table with a column-letter header and row-number gutter. */
export function formatTable(address: string, values: Cell[][], formulas?: Cell[][]): string {
  const { row: r0, col: c0 } = topLeft(address);
  const cols = values[0]?.length ?? 0;
  const header = ["", ...Array.from({ length: cols }, (_, i) => colToLetter(c0 + i))].join("\t");
  const lines = values.map((row, i) =>
    [String(r0 + i), ...row.map((v, j) => renderCell(v, formulas?.[i]?.[j]))].join("\t"),
  );
  return [header, ...lines].join("\n");
}

function renderCell(value: Cell, formula?: Cell): string {
  const text = value === null ? "" : String(value);
  const isFormula = typeof formula === "string" && formula.startsWith("=");
  return isFormula && formula !== text ? `${text} {${formula}}` : text;
}
```

Run: `npm test` → expected: all PASS.

- [ ] **Step 6: The four Office.js executors** (verified by hand in Task 6; no unit test)

`addin/src/taskpane/excel/client.ts`:

```ts
/* global Excel */
import { formatTable, type Cell } from "./format";

export const MAX_CELLS = 2000; // mirror of backend/app/tools.py
const MAX_FIND_RESULTS = 50;
const PREVIEW_COLS = 20;

export async function getWorkbookOverview(): Promise<string> {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    const active = ctx.workbook.worksheets.getActiveWorksheet();
    active.load("name");
    const selection = ctx.workbook.getSelectedRange();
    selection.load("address");
    await ctx.sync();

    // valuesOnly=true ignores formatted-but-empty cells, which otherwise inflate used ranges.
    const used = sheets.items.map((ws) => {
      const r = ws.getUsedRangeOrNullObject(true);
      r.load("address,rowCount,columnCount");
      return r;
    });
    await ctx.sync();

    const previews = used.map((r) => {
      if (r.isNullObject) return null;
      const head = r.getRow(0).getAbsoluteResizedRange(1, Math.min(r.columnCount, PREVIEW_COLS));
      head.load("values");
      return head;
    });
    await ctx.sync();

    const lines = sheets.items.map((ws, i) => {
      const r = used[i];
      if (r.isNullObject) return `- ${ws.name}: empty`;
      const headers = previews[i]!.values[0].map((v) => String(v ?? "")).join(" | ");
      return `- ${ws.name}: used ${r.address} (${r.rowCount} rows x ${r.columnCount} cols); first row: ${headers}`;
    });
    return [`Active sheet: ${active.name}`, `Selection: ${selection.address}`, "Sheets:", ...lines].join("\n");
  });
}

export async function readRange(sheet: string, address: string, includeFormulas: boolean): Promise<string> {
  return Excel.run(async (ctx) => {
    const range = ctx.workbook.worksheets.getItem(sheet).getRange(address);
    range.load("address,rowCount,columnCount,cellCount");
    await ctx.sync();
    if (range.cellCount > MAX_CELLS) {
      throw new Error(
        `${range.address} has ${range.cellCount} cells (${range.rowCount}x${range.columnCount}); ` +
          `the limit is ${MAX_CELLS}. Read it in smaller blocks.`,
      );
    }
    range.load(includeFormulas ? "values,formulas" : "values");
    await ctx.sync();
    const table = formatTable(range.address, range.values, includeFormulas ? range.formulas : undefined);
    return `${range.address} (${range.rowCount}x${range.columnCount})\n${table}`;
  });
}

export async function writeRange(sheet: string, address: string, values: Cell[][]): Promise<string> {
  const rows = values.length;
  const cols = values[0]?.length ?? 0;
  if (rows === 0 || cols === 0) throw new Error("values must be a non-empty 2-D array");
  if (values.some((r) => r.length !== cols)) throw new Error("every row in values must have the same length");
  if (rows * cols > MAX_CELLS) {
    throw new Error(`${rows * cols} cells exceeds the ${MAX_CELLS}-cell write limit; split into several writes`);
  }
  return Excel.run(async (ctx) => {
    let target = ctx.workbook.worksheets.getItem(sheet).getRange(address);
    target.load("rowCount,columnCount");
    await ctx.sync();
    if (target.rowCount === 1 && target.columnCount === 1 && (rows > 1 || cols > 1)) {
      target = target.getAbsoluteResizedRange(rows, cols);
    } else if (target.rowCount !== rows || target.columnCount !== cols) {
      throw new Error(`address is ${target.rowCount}x${target.columnCount} but values are ${rows}x${cols}`);
    }
    // `formulas` accepts plain values too; strings starting with "=" become formulas.
    target.formulas = values;
    target.load("address");
    await ctx.sync();
    return `Wrote ${rows}x${cols} cells to ${target.address}`;
  });
}

export async function find(
  query: string,
  sheet: string | undefined,
  matchCase: boolean,
  completeMatch: boolean,
): Promise<string> {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();
    const targets = sheet ? [sheets.getItem(sheet)] : sheets.items;
    const hits = targets.map((ws) => {
      const areas = ws.findAllOrNullObject(query, { completeMatch, matchCase });
      areas.load("cellCount,areas/items/address");
      return areas;
    });
    await ctx.sync();

    let total = 0;
    const addresses: string[] = [];
    for (const h of hits) {
      if (h.isNullObject) continue;
      total += h.cellCount;
      for (const a of h.areas.items) addresses.push(a.address);
    }
    if (total === 0) return `No cells contain "${query}".`;
    const shown = addresses.slice(0, MAX_FIND_RESULTS);
    const note = addresses.length > shown.length ? ` (showing the first ${shown.length} areas)` : "";
    return `${total} matching cell(s) for "${query}"${note}:\n${shown.join("\n")}`;
  });
}
```

- [ ] **Step 7: Dispatcher**

`addin/src/taskpane/agent/tools.ts`:

```ts
import * as excel from "../excel/client";

type Args = Record<string, any>;

/** Maps a tool_use name to its Office.js executor. Mirrors backend/app/tools.py. */
export async function executeTool(name: string, input: unknown): Promise<string> {
  const a = (input ?? {}) as Args;
  switch (name) {
    case "get_workbook_overview":
      return excel.getWorkbookOverview();
    case "read_range":
      return excel.readRange(a.sheet, a.address, Boolean(a.include_formulas));
    case "write_range":
      return excel.writeRange(a.sheet, a.address, a.values);
    case "find":
      return excel.find(a.query, a.sheet, Boolean(a.match_case), Boolean(a.complete_match));
    default:
      throw new Error(`Unknown tool: ${name}`);
  }
}
```

- [ ] **Step 8: Type-check and commit**

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude/addin && npx tsc --noEmit && npm test
cd .. && git add addin && git commit -m "feat: Office.js executors for overview, read, write, find with size caps"
```

---

### Task 5: The agent loop (1:35–1:55)

**Files:**
- Create: `addin/src/taskpane/api/chat.ts`, `addin/src/taskpane/agent/loop.ts`
- Test: `addin/src/taskpane/agent/loop.test.ts`

- [ ] **Step 1: The HTTP client**

`addin/src/taskpane/api/chat.ts`:

```ts
import type { ContentBlockParam, MessageParam, StopReason } from "@anthropic-ai/sdk/resources/messages";

/** What /api/chat returns: the assistant's blocks, typed as params because we send them straight back. */
export interface ChatResponse {
  content: ContentBlockParam[];
  stop_reason: StopReason | null;
}

export async function chat(messages: MessageParam[]): Promise<ChatResponse> {
  const res = await fetch("/api/chat", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ messages }),
  });
  if (!res.ok) {
    const body = await res.json().catch(() => ({}));
    throw new Error(body.detail ?? `Backend error ${res.status}`);
  }
  return res.json();
}
```

- [ ] **Step 2: Failing loop tests**

`addin/src/taskpane/agent/loop.test.ts`:

```ts
import { describe, expect, it, vi } from "vitest";
import type { ChatResponse } from "../api/chat";
import { MAX_STEPS, runTurn } from "./loop";

const askForTool: ChatResponse = {
  stop_reason: "tool_use",
  content: [
    { type: "text", text: "Let me look." },
    { type: "tool_use", id: "tu_1", name: "read_range", input: { sheet: "S", address: "A1" } },
  ],
};
const finalAnswer: ChatResponse = { stop_reason: "end_turn", content: [{ type: "text", text: "A1 is 42." }] };

describe("runTurn", () => {
  it("executes tool calls and feeds results back until the model stops", async () => {
    const chat = vi.fn().mockResolvedValueOnce(askForTool).mockResolvedValueOnce(finalAnswer);
    const execute = vi.fn().mockResolvedValue("A1\t42");

    const out = await runTurn([], "what is in A1?", { chat, execute });

    expect(execute).toHaveBeenCalledWith("read_range", { sheet: "S", address: "A1" });
    expect(out.map((m) => m.role)).toEqual(["user", "assistant", "user", "assistant"]);
    expect(out[2].content).toEqual([{ type: "tool_result", tool_use_id: "tu_1", content: "A1\t42" }]);
    expect(chat).toHaveBeenLastCalledWith(out.slice(0, 3));
  });

  it("reports executor failures as is_error results instead of aborting", async () => {
    const chat = vi.fn().mockResolvedValueOnce(askForTool).mockResolvedValueOnce(finalAnswer);
    const execute = vi.fn().mockRejectedValue(new Error("ItemNotFound"));

    const out = await runTurn([], "x", { chat, execute });

    expect(out[2].content).toEqual([
      { type: "tool_result", tool_use_id: "tu_1", content: "Error: ItemNotFound", is_error: true },
    ]);
  });

  it("gives up after MAX_STEPS tool rounds", async () => {
    const chat = vi.fn().mockResolvedValue(askForTool);
    const execute = vi.fn().mockResolvedValue("ok");

    await expect(runTurn([], "loop forever", { chat, execute })).rejects.toThrow(/tool steps/);
    expect(chat).toHaveBeenCalledTimes(MAX_STEPS);
  });
});
```

Run: `npm test` → expected: FAIL, module `./loop` not found.

- [ ] **Step 3: Implement the loop**

`addin/src/taskpane/agent/loop.ts`:

```ts
import type { MessageParam, ToolResultBlockParam, ToolUseBlockParam } from "@anthropic-ai/sdk/resources/messages";
import type { ChatResponse } from "../api/chat";

export type ChatFn = (messages: MessageParam[]) => Promise<ChatResponse>;
export type ExecuteFn = (name: string, input: unknown) => Promise<string>;

export const MAX_STEPS = 12;

interface Deps {
  chat: ChatFn;
  execute: ExecuteFn;
  /** Called after every change so the UI can render partial progress. */
  onUpdate?: (messages: MessageParam[]) => void;
}

/**
 * One user turn: call the model, run any tools it asks for in Excel, repeat until it
 * answers. Returns the full history to keep for the next turn.
 */
export async function runTurn(history: MessageParam[], userText: string, deps: Deps): Promise<MessageParam[]> {
  let messages: MessageParam[] = [...history, { role: "user", content: userText }];

  for (let step = 0; step < MAX_STEPS; step++) {
    const reply = await deps.chat(messages);
    // Keep every block, including thinking blocks, so the history replays verbatim.
    messages = [...messages, { role: "assistant", content: reply.content }];
    deps.onUpdate?.(messages);
    if (reply.stop_reason !== "tool_use") return messages;

    const calls = reply.content.filter((b): b is ToolUseBlockParam => b.type === "tool_use");
    const results: ToolResultBlockParam[] = [];
    for (const call of calls) {
      // Sequential on purpose: Office.js batches are simpler to reason about one at a time.
      try {
        results.push({ type: "tool_result", tool_use_id: call.id, content: await deps.execute(call.name, call.input) });
      } catch (err) {
        results.push({
          type: "tool_result",
          tool_use_id: call.id,
          content: `Error: ${(err as Error).message}`,
          is_error: true,
        });
      }
    }
    // All results for one assistant message go back in ONE user message.
    messages = [...messages, { role: "user", content: results }];
    deps.onUpdate?.(messages);
  }
  throw new Error(`Stopped after ${MAX_STEPS} tool steps without a final answer.`);
}
```

Run: `npm test` → expected: all PASS (a1, format, loop).

- [ ] **Step 4: Commit**

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude
git add addin && git commit -m "feat: frontend agent loop with tool execution and error results"
```

---

### Task 6: Minimal chat UI and the tracer bullet (1:55–2:25)

**Files:**
- Modify: `addin/src/taskpane/components/App.tsx`, `addin/src/taskpane/taskpane.css`
- Create: `addin/src/taskpane/components/MessageList.tsx`, `Composer.tsx`

- [ ] **Step 1: App state and wiring**

`addin/src/taskpane/components/App.tsx`:

```tsx
import * as React from "react";
import type { MessageParam } from "@anthropic-ai/sdk/resources/messages";
import { chat } from "../api/chat";
import { runTurn } from "../agent/loop";
import { executeTool } from "../agent/tools";
import { Composer } from "./Composer";
import { MessageList } from "./MessageList";

export default function App() {
  const [messages, setMessages] = React.useState<MessageParam[]>([]);
  const [busy, setBusy] = React.useState(false);
  const [error, setError] = React.useState<string | null>(null);

  async function send(text: string) {
    if (!text.trim() || busy) return;
    setBusy(true);
    setError(null);
    try {
      setMessages(await runTurn(messages, text, { chat, execute: executeTool, onUpdate: setMessages }));
    } catch (err) {
      setError((err as Error).message);
    } finally {
      setBusy(false);
    }
  }

  return (
    <div className="app">
      <header className="header">
        <strong>Crunched</strong>
        <button onClick={() => { setMessages([]); setError(null); }} disabled={busy}>New chat</button>
      </header>
      <MessageList messages={messages} busy={busy} />
      {error && <div className="error">{error}</div>}
      <Composer onSend={send} disabled={busy} />
    </div>
  );
}
```

- [ ] **Step 2: Message rendering**

`addin/src/taskpane/components/MessageList.tsx`:

```tsx
import * as React from "react";
import type { MessageParam } from "@anthropic-ai/sdk/resources/messages";

export function MessageList({ messages, busy }: { messages: MessageParam[]; busy: boolean }) {
  const endRef = React.useRef<HTMLDivElement>(null);
  React.useEffect(() => endRef.current?.scrollIntoView({ behavior: "smooth" }), [messages, busy]);
  return (
    <div className="messages">
      {messages.map((m, i) => <Message key={i} message={m} />)}
      {busy && <div className="bubble assistant muted">Working…</div>}
      <div ref={endRef} />
    </div>
  );
}

function Message({ message }: { message: MessageParam }) {
  if (typeof message.content === "string") return <div className="bubble user">{message.content}</div>;
  return (
    <>
      {message.content.map((block, i) => {
        switch (block.type) {
          case "text":
            return <div key={i} className="bubble assistant">{block.text}</div>;
          case "tool_use":
            return <div key={i} className="tool">▸ {block.name} {describe(block.input)}</div>;
          case "tool_result":
            return (
              <details key={i} className="tool-result">
                <summary>{block.is_error ? "⚠ tool error" : "result"}</summary>
                <pre>{typeof block.content === "string" ? block.content : JSON.stringify(block.content)}</pre>
              </details>
            );
          default:
            return null; // thinking blocks stay in history but are not shown
        }
      })}
    </>
  );
}

function describe(input: unknown): string {
  const { values, ...rest } = (input ?? {}) as Record<string, unknown>;
  const summary = JSON.stringify(rest);
  return Array.isArray(values) ? `${summary} (${values.length} rows)` : summary;
}
```

- [ ] **Step 3: Composer**

`addin/src/taskpane/components/Composer.tsx`:

```tsx
import * as React from "react";

export function Composer({ onSend, disabled }: { onSend: (text: string) => void; disabled: boolean }) {
  const [text, setText] = React.useState("");
  function submit() {
    const t = text.trim();
    if (!t) return;
    setText("");
    onSend(t);
  }
  return (
    <div className="composer">
      <textarea
        value={text}
        rows={2}
        placeholder="Ask me anything…"
        disabled={disabled}
        onChange={(e) => setText(e.target.value)}
        onKeyDown={(e) => {
          if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            submit();
          }
        }}
      />
      <button onClick={submit} disabled={disabled || !text.trim()}>Send</button>
    </div>
  );
}
```

- [ ] **Step 4: Styles** — replace `addin/src/taskpane/taskpane.css`:

```css
html, body, #container { height: 100%; margin: 0; font: 13px/1.4 -apple-system, "Segoe UI", sans-serif; }
.app { display: flex; flex-direction: column; height: 100%; }
.header { display: flex; justify-content: space-between; align-items: center; padding: 8px 12px; border-bottom: 1px solid #ddd; }
.messages { flex: 1; overflow-y: auto; padding: 12px; display: flex; flex-direction: column; gap: 8px; }
.bubble { max-width: 90%; padding: 8px 10px; border-radius: 10px; white-space: pre-wrap; }
.bubble.user { align-self: flex-end; background: #dbeafe; }
.bubble.assistant { align-self: flex-start; background: #f3f4f6; }
.bubble.muted { color: #6b7280; font-style: italic; }
.tool { font-family: ui-monospace, Menlo, monospace; font-size: 11px; color: #6b7280; }
.tool-result { font-size: 11px; color: #6b7280; }
.tool-result pre { max-height: 160px; overflow: auto; background: #fafafa; padding: 6px; border-radius: 6px; }
.error { margin: 0 12px 8px; padding: 8px; background: #fee2e2; color: #991b1b; border-radius: 6px; }
.composer { display: flex; gap: 8px; padding: 8px 12px; border-top: 1px solid #ddd; }
.composer textarea { flex: 1; resize: none; padding: 6px; border: 1px solid #ccc; border-radius: 6px; }
.suggestions { padding: 12px; display: flex; flex-direction: column; gap: 8px; }
.chip { text-align: left; padding: 8px 10px; border: 1px solid #ccc; border-radius: 999px; background: #fff; cursor: pointer; }
```

- [ ] **Step 5: Tracer bullet in Excel** (M2)

Both servers running (`npm run dev-server` in `addin/`, `uvicorn … --env-file .env` in `backend/`). Reload the pane. In a blank workbook type `Revenue`, `100`, `120` into A1:A3. Then, in the pane:

1. "Write hello in B1" → expect a `write_range` tool line, then B1 reads `hello`.
2. "What's in this sheet?" → expect `get_workbook_overview` then `read_range`, then a sentence describing A1:B3.
3. "Put the sum of A2:A3 in A4 as a formula" → A4 contains `=SUM(A2:A3)` = 220.

If a tool errors, the pane shows a "⚠ tool error" card and Claude's follow-up; open the inspector console for the Office.js message.

- [ ] **Step 6: Commit**

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude
git add addin && git commit -m "feat: chat UI wired to the agent loop; end-to-end read/write works"
```

---

### Task 7: Prove "any size" (2:25–2:50)

**Files:**
- Create: `scripts/make_big_workbook.py`

- [ ] **Step 1: Generator script**

```python
"""Build scripts/big.xlsx: a 1,000,000-cell Data sheet plus a small Budget sheet with
deliberate mistakes, to show the agent inspects big workbooks without reading them whole."""
from openpyxl import Workbook

ROWS, COLS = 5000, 200

wb = Workbook(write_only=True)
data = wb.create_sheet("Data")
data.append([f"Metric_{c}" for c in range(1, COLS + 1)])
for r in range(2, ROWS + 1):
    data.append([r * c for c in range(1, COLS + 1)])

budget = wb.create_sheet("Budget")
budget.append(["Line item", "2024", "2025", "2026"])
budget.append(["Revenue", 1000, 1200, 1500])
budget.append(["COGS", 400, 450, 500])
budget.append(["Gross profit", "=B2-B3", "=C2-C3", 1000])   # 2026 is hard-coded: a planted error
budget.append(["Margin", "=B4/B2", "=C4/C2", "=D4/D2"])
budget.append(["Per unit", "=B4/B9", "=C4/C9", "=D4/D9"])  # divides by empty cells: #DIV/0!

wb.save("scripts/big.xlsx")
print("wrote scripts/big.xlsx")
```

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude
source backend/.venv/bin/activate && python scripts/make_big_workbook.py
open scripts/big.xlsx
```

Expected: `wrote scripts/big.xlsx` in about 3 seconds; the file is roughly 6 MB. Excel takes a few seconds to open it.

- [ ] **Step 2: Exercise the caps in the pane** (M3)

With `big.xlsx` open and the pane loaded:

1. "How big is this workbook and what are the Data headers?" → one `get_workbook_overview` call, answer within seconds, no `read_range` of Data. Note the wall time.
2. "What is the value of Metric_150 on row 4000?" → expect `find` (or overview) then a small `read_range`; never the whole sheet. If Claude tries `Data!A1:GR5000`, the executor returns the 1,000,000-cell error and Claude narrows: that is the designed behaviour, keep it.
3. "Check the Budget sheet for errors" → expect `read_range` with `include_formulas: true`, and the answer to name the hard-coded `Budget!D4` and the `#DIV/0!` row.

- [ ] **Step 3: Fix what you found**

Typical fixes at this point: the overview is slow because `getRow(0)` on a huge used range loads too much (cap `PREVIEW_COLS` lower); `find` returns thousands of areas on Data (cap is already 50, confirm the message stays short); `formatTable` output of 2,000 numeric cells is ~15 KB, which is fine.

- [ ] **Step 4: Commit**

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude
git add scripts addin && git commit -m "feat: 1M-cell test workbook; tune caps for large sheets"
```

---

### Task 8: Polish (2:50–3:15)

**Files:**
- Create: `addin/src/taskpane/components/SuggestedPrompts.tsx`
- Modify: `addin/src/taskpane/components/App.tsx`

- [ ] **Step 1: Suggested prompts** (shown only on an empty conversation)

```tsx
import * as React from "react";

const PROMPTS = [
  "What's in this workbook?",
  "Error-check the formulas on the active sheet",
  "Add a Total row under the selected range",
];

export function SuggestedPrompts({ onPick }: { onPick: (text: string) => void }) {
  return (
    <div className="suggestions">
      <p>Hi, I'm Crunched, your AI analyst in Excel. I can read and edit this workbook. Try:</p>
      {PROMPTS.map((p) => (
        <button key={p} className="chip" onClick={() => onPick(p)}>{p}</button>
      ))}
    </div>
  );
}
```

In `App.tsx` add `import { SuggestedPrompts } from "./SuggestedPrompts";` next to the other component imports, then replace `<MessageList … />` with:

```tsx
{messages.length === 0 ? <SuggestedPrompts onPick={send} /> : <MessageList messages={messages} busy={busy} />}
```

- [ ] **Step 2: Hand-verify the states**

Empty state shows chips; clicking one sends it. Backend down (`Ctrl-C` uvicorn) → red error banner with "Backend error 504" or similar, composer re-enabled. Wrong sheet name in a prompt ("read Sheet99") → tool error card, Claude recovers. "New chat" clears everything.

- [ ] **Step 3: Type-check, test, commit**

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude/addin && npx tsc --noEmit && npm test
cd .. && git add addin && git commit -m "feat: suggested prompts and error states"
```

---

### Task 9: README and submission (3:15–3:45)

**Files:**
- Modify: `README.md` (keep the exercise text under a "The brief" heading at the bottom, or drop it; the reviewers wrote it)

- [ ] **Step 1: Write the README** with these sections, each a few sentences:

1. **What this is** — one paragraph, plus a screenshot or GIF of the pane if time allows.
2. **Run it on a Mac** — the exact commands from Tasks 1, 2, 3 (`uv venv`, `.env`, `uvicorn`, `npm run dev-server`, `npm start`), the keychain prompt, the wef fallback, `OfficeWebAddinDeveloperExtras`.
3. **Architecture** — the diagram and the one-turn sequence from §1 of this plan; why the loop is in the pane; why one origin.
4. **Handling any workbook size** — the four tools, the caps, the paging error, the 1M-cell script and the measured overview time.
5. **Testing** — what is unit-tested and why Office.js/UI are hand-verified; `pytest` and `npm test` commands.
6. **Trade-offs and what I'd do next** — streaming; write preview/confirm; `scan_errors` tool that walks `valueTypes` in chunks in-process; persisting conversations; Excel Online; refusal fallbacks; moving the loop server-side with WebSockets if tools ever need to run off-client.
7. **Time log** — honest per-task minutes.

- [ ] **Step 2: Final verification** (`@superpowers:verification-before-completion`)

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude/backend && pytest -q
cd ../addin && npx tsc --noEmit && npm test
```

Then run the three demo prompts from Task 7 Step 2 once more in Excel. Only after all pass:

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss-claude
git add README.md && git commit -m "docs: setup, architecture, trade-offs"
git push -u origin claude-macbook-plan
```

The other implementation lives on `feat/excel-taskpane-agent`. Decide which branch is the submission (or merge one into `main`) before sharing the link. Add `markusskagemo` and `larsgmu` as collaborators if the repo is private; email the link to recruiting@usecrunched.com.

---

## 5. Demo script for the 15-minute call

1. Open `big.xlsx`, show the Data sheet is 5,000 × 200. Ask "How big is this workbook?" and point at the single overview call in the pane.
2. Ask "Check the Budget sheet for errors." Show the `include_formulas` read and the two findings.
3. Ask "Fix the hard-coded 2026 gross profit with a formula." Show the `write_range` line and `Budget!D4` now containing `=D2-D3`.
4. Walk the code: `loop.ts` (30 lines, tested), `agent.py` (one call), `tools.py` (the contract). Explain why the loop lives in the pane and why there is one origin.
5. Name what is deliberately missing (§6).

## 6. Deliberately left out of the four hours

- **Streaming text.** Non-streaming keeps the loop a plain request/response; adding SSE per step is mechanical.
- **Write confirmation.** Writes apply immediately and are listed in the chat. A preview/approve step is the first thing to add for real users.
- **Refusal fallbacks.** Claude Opus 5 supports a server-side `fallbacks: "default"` parameter (beta header `server-side-fallback-2026-07-01`, via `client.beta.messages.create`). A spreadsheet assistant almost never trips a refusal, and an unknown beta header would take down the demo, so it is documented rather than enabled.
- **Strict tool schemas** (`strict: true`) and `disable_parallel_tool_use`. Inputs are validated by the executors; strictness is a one-line upgrade once the schemas are settled.
- **Conversation persistence, auth, Excel Online.** Out of scope for a local demo.
- **Reason / Agent Mode toggles** from the screenshot. `effort` is already an env var; a toggle would just set it per request.
