# Feature List

## Chat & Interaction

| Feature | Status | Description |
|---|---|---|
| Natural language chat | ✅ Shipped | Type questions in Excel's task pane sidebar |
| Markdown rendering | ✅ Shipped | Bold, italic, code, lists, tables in chat bubbles |
| Chat autoscroll | ✅ Shipped | Thread auto-scrolls to newest message |
| New chat button | ✅ Shipped | Clear history and start fresh |
| Composer auto-focus | ✅ Shipped | Textarea focuses on mount and after "New chat" |
| Demo prompt chips | ✅ Shipped | Empty-state suggestion buttons ("How big is this workbook?") |
| Suggested follow-ups | ✅ Shipped | Context-aware chip buttons after each assistant reply |
| Clarifying questions | ✅ Shipped | Multiple-choice buttons when the model is uncertain |
| Conversation persistence | ✅ Shipped | `localStorage` keyed by workbook name so chats survive reloads |

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
| Backend pytest | ✅ Shipped | 24 tests: HTTP contract, tool schemas, cell/body caps, CORS, agent loop |
| Frontend mocha | ✅ Shipped | 14 test files: agent loop, excel policy, storage, undo, markdown, clarifying questions, suggestions, tool cards, formula explainer, tour, ELI5, API paths, demo prompts, conversation |

## Future work

These were scoped, issued (#29–#31), and then closed as not planned for this take-home scope. They are solid next features for a production version.

| Feature | GitHub issue | Why it was cut | What would make it come back |
|---|---|---|---|
| Write-confirm dialog | #29 | Adds UI friction for a demo where writes apply immediately | Production use on live financial models |
| Streaming responses | #30 | Office.js add-ins are small; full messages are fast enough | Large model outputs where perceived latency matters |
| Excel Online support | #31 | Desktop Excel has the full Office.js API; Online is a subset | User base that lives in browser-first Excel |

