# Feature List

The assignment is the chat loop, five tools, and the 2k / 8-round policy. Everything else is week-one polish or an intentional cut — same split as the README.

| When | What |
|---|---|
| Hours 1–4 | Chat pane, five Office.js tools, FastAPI turn, 2k/8-round policy, one-origin proxy, million-cell demo |
| Week one | Tool cards, bounded history and requests, undo, write-confirm, persistence, CI, and pre-Office.js bounds checks |
| Optional polish | Formula explainer; useful, but outside the core loop and shipping argument |
| Cut | Clarifying buttons, follow-up chips, tour / ELI5, leftover plan docs |

## Chat & Interaction

| Feature | Status | Description |
|---|---|---|
| Natural language chat | ✅ Shipped | Type questions in Excel's task pane sidebar |
| Chat autoscroll | ✅ Shipped | Thread auto-scrolls to newest message |
| New chat button | ✅ Shipped | Clear history and start fresh |
| Composer auto-focus | ✅ Shipped | Textarea focuses on mount and after "New chat" |
| Demo prompt chips | ✅ Shipped | Empty-state suggestion buttons ("How big is this workbook?") |
| Conversation persistence | ✅ Shipped | `localStorage` keyed by workbook name so chats survive reloads |
| Markdown rendering | ✅ Shipped | Bold, bullets, headings, tables and inline code render as elements; the user's own text stays verbatim |

## AI & Tools

| Feature | Status | Description |
|---|---|---|
| Tool-use loop | ✅ Shipped | Claude requests data → pane executes → results returned |
| 5 Excel tools | ✅ Shipped | `list_workbook_meta`, `read_range`, `write_range`, `get_selection`, `find` |
| Tool visibility | ✅ Shipped | Every tool call appears as a card in the chat thread |
| Human tool names | ✅ Shipped | "Read range" instead of `read_range` in tool cards |
| 8-round cap | ✅ Shipped | Hard limit of 8 tool rounds per user message |
| 2,000-cell read cap | ✅ Shipped | Large reads are truncated; small reads preferred |
| Metadata-first policy | ✅ Shipped | Model sees sheet list + header preview before any data |
| Formula detection | ✅ Shipped | Tool cards show "· formulas" tag when values ≠ formulas |
| Large-sheet fixture | ✅ Shipped | 1M-cell test workbook (`scripts/big.xlsx`) |
| Undo stack | ✅ Shipped | Snapshot cells before `write_range`; Undo button in header |
| Formula explainer | ✅ Shipped | Select a cell → "Explain this formula" → plain-English breakdown |

## Excel Integration

| Feature | Status | Description |
|---|---|---|
| Office.js runtime | ✅ Shipped | All Excel I/O through `frontend/src/app/services/excel.ts` |
| Live selection | ✅ Shipped | Pane shows current Excel selection as a pill |
| Write validation | ✅ Shipped | `write_range` validated before `Excel.run` |
| A1 bounds before Office.js | ✅ Shipped | Million-cell addresses rejected in JS; `Excel.run` tests mock the host |
| Write confirm | ✅ Shipped | Apply / Don't write before any AI write lands |
| Undo stack | ✅ Shipped | Snapshot cells before `write_range`; Undo button |
| Sideload manifest | ✅ Shipped | `manifest.xml` for desktop Excel on macOS |
| HTTPS dev server | ✅ Shipped | Webpack on `https://localhost:3000` with trusted certs |
| One-origin proxy | ✅ Shipped | `/api/*` proxied to uvicorn — no CORS needed |

## UX & Polish

| Feature | Status | Description |
|---|---|---|
| Dark theme | ✅ Shipped | Charcoal background with copper accents |
| Tool card errors | ✅ Shipped | Red border + "Tool error:" aria-label for failed calls |
| Friendly errors | ✅ Shipped | Human-readable error messages in the banner |
| Status spinner | ✅ Shipped | Animated indicator while waiting for Claude |
| Selection pill | ✅ Shipped | Shows "Crunched can see Budget!B2:D5" |
| Narrow-pane layout | ✅ Shipped | Tool cards collapse to single column at ≤320px |
| Reduced motion | ✅ Shipped | Respects `prefers-reduced-motion` |

## Accessibility

| Feature | Status | Description |
|---|---|---|
| Aria-live region | ✅ Shipped | `aria-live="polite"` on chat thread |
| Tool card labels | ✅ Shipped | `aria-label` with name + error state |
| Composer label | ✅ Shipped | `aria-label="Message"` on textarea |
| Focus outlines | ✅ Shipped | Visible `:focus-visible` styles on all interactive elements |

## Tests

| Feature | Status | Description |
|---|---|---|
| Backend pytest | ✅ Shipped | HTTP contract, tool schemas, cell caps, agent loop, request limits |
| Frontend mocha | ✅ Shipped | Policy, storage, tool cards, demo chips, undo, write confirm, formula explainer, mocked `Excel.run` |

## Cut (on purpose)

| Feature | Why |
|---|---|
| Clarifying-question buttons | Model can ask in prose; extra parser UI |
| Follow-up suggestion chips | Empty-state demo chips already drive the walkthrough |
| Guided tour / ELI5 | Interview decoration |
| Draft plan docs | They read as a second product next to the shipped README |

## Future work

These were scoped, issued (#29–#31), and then closed as not planned for this take-home scope. They are solid next features for a production version.

| Feature | GitHub issue | Why it was cut | What would make it come back |
|---|---|---|---|
| Streaming responses | #30 | Office.js add-ins are small; full messages are fast enough | Large model outputs where perceived latency matters |
| Excel Online support | #31 | Desktop Excel has the full Office.js API; Online is a subset | User base that lives in browser-first Excel |
