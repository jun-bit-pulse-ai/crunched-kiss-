# Crunched — 4-Hour Implementation Plan

## What We're Building
A simplified **Crunched**: an AI analyst that lives in Excel's task pane. User sends chat messages; AI reads/writes the spreadsheet. Must handle workbooks of any size.

> **Deployment scope:** This is a local-dev-first 4-hour build. We target **Excel Desktop** with local HTTPS (mkcert) as the primary path. Cloud deployment (Vercel, etc.) is explicitly out of scope for the 4-hour window — it adds manifest re-configuration, backend hosting, and Excel security policy complexity that eats the time budget.

**Stack:** React + TypeScript + Office.js (frontend) | Python + FastAPI (backend) | Claude/OpenAI (agent)

---

## Scope & Deployment

| In Scope (4 hours) | Out of Scope |
|---|---|
| Local HTTPS dev server with mkcert | Cloud deployment (Vercel, Netlify, etc.) |
| Excel Desktop sideload via local manifest | Excel Online (can work later, not now) |
| Single-machine frontend + backend | Multi-service hosting architecture |
| Feature-complete local MVP | CI/CD, auth, billing |

**Why no Vercel?** The Excel add-in manifest must be sideloaded from your local machine and points to `https://localhost:3000`. Moving to Vercel requires updating the manifest with production URLs, deploying the FastAPI backend separately, handling CORS, and wrestling with Excel Desktop's external URL security policies. That alone is 1–2 hours — time we don't have. Build locally first, deploy later if desired.

---

## Workflow & Agent Assignment

### Trunk-Based Development

All agents work on `main`. No long-lived feature branches.

```
main  ← everyone pushes here
├── frontend/   ← Frontend Agent owns
├── backend/    ← Backend Agent owns  
└── manifest.xml, certs/  ← shared, coordinate in Slack/chat
```

**Why trunk-based:** Frontend and backend live in separate directories → zero merge conflicts. No branch switching overhead. Single source of truth. Progress visible in real-time.

**Rule:** If two agents must touch the same file, the second agent waits or the first agent finishes the file entirely before the second starts.

### GitHub Issues as Task Board

Create 4 issues so agents can self-assign and parallelize:

| Issue | Title | Agent | Est. Duration |
|---|---|---|---|
| #1 | Hour 1 — Scaffold, certs, manifest | Infra Agent | 60 min |
| #2 | Hour 2 — Frontend chat + Excel service | Frontend Agent | 60 min |
| #3 | Hour 3 — FastAPI backend + Anthropic agent | Backend Agent | 60 min |
| #4 | Hour 4 — Integration, polish, E2E test | Integration Agent | 60 min |

**Issue template per task:**
```markdown
## Task: [Hour X — Title]

### Deliverables
- [ ] File A
- [ ] File B

### Depends on
- Issue #Y (if any)

### Acceptance Criteria
- [ ] Criteria 1
- [ ] Criteria 2
```

**Agent picks up issue → self-assigns → checks boxes → closes on done.**

### Parallelization Strategy

```
Time  Issue #1 (Infra)        Issue #2 (Frontend)     Issue #3 (Backend)
─────────────────────────────────────────────────────────────────────────
0:00  yo office + mkcert
0:10  webpack HTTPS config
0:20  manifest.xml
0:30  ──────────────────►  ChatPanel.tsx
0:40  ──────────────────►  InputBox.tsx
0:50  ──────────────────►  api.ts (stub)
1:00  DONE #1              ──────────────────►  excel.ts
1:10                       excel.ts cont.
1:20                       ──────────────────►  FastAPI scaffold
1:30                       ──────────────────►  Pydantic models
1:40  ◄──────────────────  api.ts (wire to real backend)
1:50                       ──────────────────►  agent.py
2:00  ◄──────────────────  E2E read test     Test backend
2:10                                               agent loop
2:20                                               agent loop cont.
2:30  ◄──────────────────  E2E write test    Test /chat endpoint
2:40  ◄──────────────────  Integration       Polish
2:50  ◄──────────────────  Error handling    Error handling
3:00  ◄──────────────────  Suggested chips   DONE #3
3:10  Large workbook cap
3:20  DONE #2              ◄─────────────────  Final integration
3:30  ◄────────────────────────────────────  E2E test
3:40  ◄────────────────────────────────────  Bug fixes
3:50  ◄────────────────────────────────────  Final polish
4:00  DONE #4 (all)
```

**Key insight:** Issues #1, #2, #3 can overlap. Issue #4 (Integration) starts once #2 and #3 have their core files ready.

---

## Architecture

```
┌─────────────────────────────────────────────────────────────┐
│  Excel Desktop / Excel Online                               │
│  ┌─────────────────────────────────────────────────────┐   │
│  │  Task Pane (React + Office.js)                      │   │
│  │  ┌──────────────┐  ┌──────────────────────────┐    │   │
│  │  │ Chat UI      │  │ Message Input            │    │   │
│  │  │ (messages)   │  │ [Ask me anything...]     │    │   │
│  │  └──────────────┘  └──────────────────────────┘    │   │
│  └─────────────────────────────────────────────────────┘   │
└─────────────────────────────────────────────────────────────┘
                              │
                              │ HTTPS (mkcert)
                              ▼
┌─────────────────────────────────────────────────────────────┐
│  FastAPI Backend                                            │
│  ┌─────────────┐    ┌─────────────┐    ┌────────────────┐  │
│  │ /chat       │───▶│ Agent Loop  │───▶│ Claude / OpenAI│  │
│  │ endpoint    │◄───│ (stateless) │◄───│ API            │  │
│  └─────────────┘    └─────────────┘    └────────────────┘  │
│         │                                                   │
│         │ Actions: read_range, write_range, get_metadata    │
│         ▼                                                   │
│  ┌────────────────────────────────────────────────────────┐│
│  │  Frontend executes via Office.js, returns results      ││
│  └────────────────────────────────────────────────────────┘│
└─────────────────────────────────────────────────────────────┘
```

**Key insight:** The agent doesn't touch Excel directly. It returns *actions* (e.g., "read range A1:D10"). The frontend executes them via Office.js and sends results back. This avoids CORS, sandboxing, and certificate hell.

---

## File Structure (Auto-Generated)

```
crunched/
├── manifest.xml              # Excel add-in manifest
├── certs/                    # mkcert generated (gitignored)
│
├── frontend/                 # React + Office.js task pane
│   ├── package.json
│   ├── tsconfig.json
│   ├── webpack.config.js     # HTTPS dev server
│   ├── public/
│   │   └── index.html
│   └── src/
│       ├── index.tsx         # Office.onReady() bootstrap
│       ├── App.tsx           # Main layout
│       ├── components/
│       │   ├── ChatPanel.tsx
│       │   ├── Message.tsx
│       │   └── InputBox.tsx
│       ├── services/
│       │   ├── api.ts        # Backend HTTP client
│       │   └── excel.ts      # Office.js wrappers
│       └── types/
│           └── index.ts
│
├── backend/                  # Python + FastAPI
│   ├── requirements.txt
│   ├── main.py               # FastAPI app
│   ├── agent.py              # AI agent loop
│   ├── models.py             # Pydantic schemas
│   └── tools.py              # Excel tool definitions
│
└── README.md
```

---

## 4-Hour Breakdown

### Hour 1: Scaffold & Certificates
- [ ] **0:00–0:10** — Clone/use Microsoft's React quickstart: `yo office --projectType taskpane --name crunched --host excel --js tsx`
- [ ] **0:10–0:20** — Clean up scaffold: remove sample code, keep Office.js boilerplate
- [ ] **0:20–0:35** — Generate HTTPS certs with mkcert:
  ```bash
  brew install mkcert
  mkcert -install
  mkcert localhost 127.0.0.1 ::1
  mv localhost+2.pem localhost+2-key.pem certs/
  ```
- [ ] **0:35–0:50** — Configure webpack dev server for HTTPS (point to certs)
- [ ] **0:50–1:00** — Test: `npm run dev` → sideload manifest → task pane opens in Excel
  - **Mac:** Copy manifest to WEF folder: `cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` then restart Excel

### Hour 2: Frontend Core
- [ ] **1:00–1:15** — Build `ChatPanel.tsx`: message list, scrollable, basic styling
- [ ] **1:15–1:30** — Build `InputBox.tsx`: textarea, send button, submit on Enter
- [ ] **1:30–1:45** — Build `excel.ts`: Office.js wrappers for:
  - `getActiveWorkbookInfo()` → sheet names, used ranges
  - `readRange(address)` → 2D array of values
  - `writeRange(address, values)` → write 2D array
- [ ] **1:45–2:00` — Build `api.ts`: HTTP client to FastAPI `/chat`, handle action responses

### Hour 3: Backend & Agent
- [ ] **2:00–2:15** — Scaffold FastAPI: `main.py` with `/chat` POST endpoint
- [ ] **2:15–2:30` — Define Pydantic models:
  ```python
  class ChatRequest(BaseModel):
      message: str
      context: dict | None = None  # Previous action results

  class Action(BaseModel):
      type: str  # "read_range" | "write_range" | "get_metadata" | "respond"
      params: dict

  class ChatResponse(BaseModel):
      actions: list[Action]
      message: str
  ```
- [ ] **2:30–2:50` — Build agent loop (`agent.py`):
  - System prompt with available tools (read_range, write_range, get_metadata)
  - User message + context → LLM → parse JSON action response
  - Use Claude 3.5 Sonnet via API key from README
- [ ] **2:50–3:00` — Test backend: `curl` the `/chat` endpoint, verify action generation

### Hour 4: Integration & Polish
- [ ] **3:00–3:20` — Wire frontend ↔ backend:
  - User sends message → POST to `/chat`
  - If action returned: execute via Office.js → send result back → get final response
  - Render AI message in chat
- [ ] **3:20–3:35` — Handle large workbooks:
  - Agent requests `get_metadata` first (sheet names, used ranges)
  - Agent reads specific ranges, not entire workbook
  - Frontend truncates very large ranges (>1000 cells) with warning
- [ ] **3:35–3:50` — Error handling + polish:
  - Loading states, error banners
  - Handle network errors gracefully
  - Add suggested prompt chips (from screenshot)
- [ ] **3:50–4:00` — End-to-end test: open a spreadsheet, chat with agent to read/write cells

---

## Critical Implementation Details

### 1. Agent System Prompt
```
You are Crunched, an AI analyst in Excel. You help users analyze spreadsheets,
build models, and find errors.

You can interact with the workbook through these actions:
1. get_metadata — Get sheet names and used ranges
   params: {}
2. read_range — Read values from a range
   params: {"sheet": "Sheet1", "range": "A1:D10"}
3. write_range — Write values to a range
   params: {"sheet": "Sheet1", "range": "A1", "values": [[1,2],[3,4]]}
4. respond — Send a message to the user (no action needed)
   params: {"text": "..."}

Rules:
- For large workbooks, request metadata first, then specific ranges
- Always respond with a valid JSON action
- If you need to read data to answer, use read_range
- If you need to modify data, use write_range
```

### 2. Office.js Excel Service
```typescript
// services/excel.ts
export async function getActiveWorkbookInfo() {
  return Excel.run(async (context) => {
    const workbook = context.workbook;
    const sheets = workbook.worksheets;
    sheets.load("items/name, items/usedRange");
    await context.sync();
    return sheets.items.map(s => ({
      name: s.name,
      usedRange: s.usedRange ? s.usedRange.address : null
    }));
  });
}

export async function readRange(sheet: string, range: string) {
  return Excel.run(async (context) => {
    const sheetObj = context.workbook.worksheets.getItem(sheet);
    const rangeObj = sheetObj.getRange(range);
    rangeObj.load("values, text");
    await context.sync();
    return { values: rangeObj.values, text: rangeObj.text };
  });
}

export async function writeRange(sheet: string, range: string, values: any[][]) {
  return Excel.run(async (context) => {
    const sheetObj = context.workbook.worksheets.getItem(sheet);
    const rangeObj = sheetObj.getRange(range);
    rangeObj.values = values;
    await context.sync();
  });
}
```

### 3. Backend Agent Loop
```python
# agent.py
import json
from anthropic import Anthropic

SYSTEM_PROMPT = """You are Crunched..."""  # (from above)

class Agent:
    def __init__(self, api_key: str):
        self.client = Anthropic(api_key=api_key)

    def process(self, user_message: str, context: dict = None) -> list[dict]:
        messages = [{"role": "user", "content": user_message}]
        if context:
            messages.append({"role": "user", "content": f"Context: {json.dumps(context)}"})

        response = self.client.messages.create(
            model="claude-3-5-sonnet-20241022",
            max_tokens=4096,
            system=SYSTEM_PROMPT,
            messages=messages
        )

        # Parse actions from response (expect JSON array of actions)
        content = response.content[0].text
        try:
            actions = json.loads(content)
            return actions if isinstance(actions, list) else [actions]
        except json.JSONDecodeError:
            return [{"type": "respond", "params": {"text": content}}]
```

### 4. FastAPI Endpoint
```python
# main.py
from fastapi import FastAPI
from fastapi.middleware.cors import CORSMiddleware
from models import ChatRequest, ChatResponse
from agent import Agent

app = FastAPI()
app.add_middleware(CORSMiddleware, allow_origins=["*"], allow_methods=["*"], allow_headers=["*"])

agent = Agent(api_key="...")

@app.post("/chat")
async def chat(request: ChatRequest) -> ChatResponse:
    actions = agent.process(request.message, request.context)
    return ChatResponse(actions=actions, message="")
```

---

## Handling Large Workbooks

| Problem | Solution |
|---------|----------|
| Can't load entire workbook into LLM context | Agent requests metadata first, then specific ranges |
| Used range is millions of cells | Frontend caps reads at 1000 cells; warns if truncated |
| LLM context window limits | Send only values, not formatting. Summarize in agent loop |
| Slow Office.js reads | Read in batches if needed; show loading spinner |

---

## macOS / MacBook Notes

### Sideloading the Add-in on Mac Excel

Excel on macOS has two ways to sideload add-ins:

**Option A: WEF Folder (Recommended — Auto-Loads)**
```bash
# Create the WEF directory if it doesn't exist
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/

# Copy your manifest there
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
```
- Excel will auto-detect the manifest on next launch
- No UI clicks needed after initial setup
- Remove the file to uninstall

**Option B: Insert → My Add-ins → Add from File**
1. Open Excel → Insert → Add-ins → My Add-ins
2. Click the `...` menu → Add from File
3. Select `manifest.xml`
4. The task pane should open automatically

### macOS-Specific Considerations

| Topic | macOS Behavior |
|---|---|
| **mkcert** | `mkcert -install` auto-adds to macOS Keychain. Excel's WebView trusts it. |
| **Excel version** | Requires Excel 2016+ (16.x) or Microsoft 365. Check: Excel → About Excel. |
| **File paths** | Use `~/Library/Containers/...` for WEF. Don't use Windows `%USERPROFILE%` paths. |
| **Office.js API** | `Excel.run()` works identically. `context.workbook` is fully supported. |
| **Task pane size** | Mac Excel task panes are resizable but default narrower. Design for ~320px min width. |
| **Keyboard shortcuts** | `Cmd+Enter` to send instead of `Ctrl+Enter` (handle both in InputBox). |

### Excel for Mac Limitations to Know

- **No shared runtime** on Mac — task pane is isolated. This is fine for our architecture.
- **Some advanced APIs** (like `Workbook.customXmlParts`) are not supported — we don't use them.
- **File dialogs** from Office.js behave slightly differently — we don't use them.

### Quick Verification Checklist (Mac)

```bash
# 1. Verify mkcert is trusted
security find-certificate -c "mkcert" ~/Library/Keychains/login.keychain-db

# 2. Verify WEF folder exists
ls ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/

# 3. Check Excel version (must be 16.x+)
# Excel → About Excel → look for "Version 16.XX"
```

---

## QA Plan

### Philosophy
QA is continuous, not a final gate. Every agent tests their own work before marking an issue done. The Integration Agent runs the full E2E suite before Hour 4 ends.

### Layer 1: Unit / Component Tests (Per-Agent, ~5 min each)

| Component | Test | How | Pass Criteria |
|---|---|---|---|
| **HTTPS certs** | Cert is trusted | `curl -v https://localhost:3000` | No TLS error |
| **Webpack dev server** | Serves over HTTPS | Open `https://localhost:3000` in browser | Loads without cert warning |
| **FastAPI** | `/health` responds | `curl http://localhost:8000/health` | Returns `{"status":"ok"}` |
| **Backend /chat** | Returns valid actions | `curl -X POST -d '{"message":"read A1"}' http://localhost:8000/chat` | Returns JSON with `actions` array |
| **Office.js excel.ts** | readRange works | Call `readRange("Sheet1","A1")` in browser console | Returns `{values:[["hello"]]}` |
| **Office.js excel.ts** | writeRange works | Call `writeRange("Sheet1","A1",[["test"]])` | Cell A1 shows "test" |

### Layer 2: Integration Tests (Hour 4, ~10 min)

| Test | Steps | Expected Result |
|---|---|---|
| **Chat → Read** | Type "What's in A1?" with value "Hello" in A1 | AI responds with "Hello" |
| **Chat → Write** | Type "Put 42 in B2" | Cell B2 shows 42, AI confirms |
| **Chat → Metadata** | Type "What sheets are there?" | AI lists sheet names |
| **Large workbook** | Open workbook with 10k rows, ask "What's in Sheet1?" | Agent reads metadata first, then specific range. No crash. |
| **Error handling** | Kill backend, send message | Frontend shows "Cannot connect to server" error |

### Layer 3: Edge Cases (If Time Permits)

| Scenario | Expected Behavior |
|---|---|
| User sends empty message | Nothing happens or gentle prompt |
| Agent returns invalid JSON | Frontend shows "Something went wrong" |
| Range read >1000 cells | Frontend truncates, shows warning toast |
| Write to protected sheet | Excel error surfaced in chat |
| Backend timeout (>30s) | Frontend shows timeout message, offers retry |

### QA Checklist (Before Calling Done)

```markdown
## QA Checklist — Issue #X

- [ ] All files compile/build without errors
- [ ] Manual test steps above pass
- [ ] No console errors in Excel task pane
- [ ] No console errors in backend logs
- [ ] macOS WEF sideload works (if touching manifest)
```

---

## Local Development Setup (macOS)

```bash
# 1. Certificates
brew install mkcert
mkcert -install
mkcert localhost 127.0.0.1 ::1

# 2. Frontend
cd frontend
npm install
npm run dev  # HTTPS webpack dev server on localhost:3000

# 3. Backend
cd backend
python -m venv venv && source venv/bin/activate
pip install -r requirements.txt
uvicorn main:app --reload --ssl-keyfile=../certs/localhost+2-key.pem --ssl-certfile=../certs/localhost+2.pem

# 4. Sideload add-in (macOS)
mkdir -p ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
cp manifest.xml ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/
# Then restart Excel — task pane auto-opens
```

**Windows alternative:** Insert → Add-ins → Upload My Add-in → select `manifest.xml`

---

## Success Criteria (4-Hour Done)

- [ ] Task pane opens in Excel with chat UI
- [ ] User can type a message and get AI response
- [ ] AI can read a range from the active workbook
- [ ] AI can write a range to the active workbook
- [ ] Works with a large workbook (agent doesn't crash on metadata read)
- [ ] HTTPS local dev works without browser warnings

---

## Suggested Prompts (Pre-Load Chips)
From the screenshot, include these quick-action chips below the input:
- "Error check Budget model"
- "Check assumptions in Revenue model"
- "Create comparison dashboard"

---

## Post-4-Hour Stretch Goals (If Time Permits)
1. Streaming responses for faster perceived latency
2. Conversation history persistence
3. Multiple tool calls in a single turn
4. Better error messages from Excel (formula errors, #REF!, etc.)
5. Support for Excel Online (may already work with same manifest)

---

## Handoff for Claude Code Agents

**Start here:** Hour 1 scaffold → `yo office` quickstart + mkcert certs
**Critical path:** Frontend chat UI → Excel service → Backend agent → Integration
**Known traps:** HTTPS certificates, Office.js async context, CORS on FastAPI, **macOS WEF folder path for sideloading**
**Test end-to-end early:** Don't wait until Hour 4 to try reading from Excel

**GO.**
