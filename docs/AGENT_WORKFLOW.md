# Running multiple coding agents on one repo

How this project was actually built, what broke, and the playbook I would hand
someone starting the same way tomorrow.

This is a field report, not theory. Every failure below happened in this
repository during the sprint, and the fixes are the ones that worked.

---

## 1. Context

`crunched-kiss` is a four-hour take-home: an Excel task-pane agent that chats
with Claude and reads and writes the live workbook, including sheets with a
million cells. It was built by **three implementation agents working in parallel
on one trunk** (Claude Code, Cursor, Kimi), a continuous QA agent, and me as
integrator and final reviewer. The four-hour core was then hardened over a longer
"week one" tail.

---

## 2. The pipeline

Nine stages. The ordering matters more than any individual tool.

```
1 research  →  2 plan  →  3 review the plan  →  4 conflict-free implementation spec
                                                          │
                        ┌─────────────────┬───────────────┼───────────────┐
                     5 agent A         agent B         agent C      (issue-scoped)
                        └─────────────────┴───────────────┴───────────────┘
                                          │
   6 design pass  →  7 continuous QA (agents + me)  →  8 hardening: a11y, security
                                          │
                              9 cleanup: branches, PRs, issues
```

### Stage 1 — Research

Read the brief, the existing product, and the constraints that will actually bite
before writing any plan. For this project the binding constraints were: Office
add-ins run in a sandboxed WebView that refuses plain HTTP, and Office.js is the
only path to the workbook. Both shaped every later decision.

### Stage 2 — Plan

Write the architecture yourself: the loop, the state boundaries, the tool
contract. This is the part agents must not choose. Two competing plans were
drafted here, which turned out to be useful — comparing them exposed assumptions
neither made explicit.

### Stage 3 — Review the plan before building

The plan went to a reviewer agent with a deliberately narrow brief: find gaps
that would cause real problems during implementation, not stylistic preferences.
It caught four environment issues that would each have cost a scaffold or a
debugging session — the Node version the generator refuses, a `pytest` import
path, a type-check that fails inside a dependency's own type declarations, and a
colliding checkout.

**Reviewing the plan is cheaper than reviewing the code it produces.**

### Stage 4 — Draft the implementation spec to avoid conflicts

This is the stage that decides whether parallel agents help or hurt. The rule
that worked:

> Split by **file ownership**, not by task.

Each GitHub issue named the exact files its agent could touch, its acceptance
test, and what it must not go near. Agents sharing a goal collide; agents sharing
a *file* corrupt each other's work.

Two things had to land **before** the fan-out:

1. **The tool contract** — `backend/app/tools.py` (schemas) and
   `dispatchExcelTool` (executors). Both sides of a contract are the single most
   conflict-prone thing in a parallel build.
2. **The one-origin decision** — the dev server proxies `/api` to the backend, so
   there is one HTTPS origin, no CORS, and no second certificate.

### Stage 5 — Implement with parallel agents

Three issues, three agents, three worktrees, trunk-based with squash merges:

| Issue | Agent | Scope |
|---|---|---|
| #2 | Cursor | Proxy `/api` through the add-in origin |
| #3 | Claude Code | `find` tool + header preview |
| #4 | Kimi | 1M-cell workbook fixture |

Then integration on trunk, and a README walkthrough once the pieces existed.

### Stage 6 — Design pass

A separate agent audited the pane against the plan for UX gaps rather than bugs:
friendly error messages instead of raw HTTP status codes, a truncation indicator
so a partial read is visible, a live selection pill, and a status spinner so a
slow tool call does not read as frozen.

### Stage 7 — Continuous QA, including me

A QA agent ran on a fixed interval with a standing brief: pull, run both suites,
type-check, check contract drift, read recent diffs, hunt for real bugs, and open
PRs — but never commit to trunk. I reviewed every PR and drove the demo myself.

Both halves were necessary. The agent found things I would not have: a history
window that would make the API reject long conversations, a private key one
`git add -A` from being committed. I found things it would not: that the *demo*
was wrong even when the code was right.

### Stage 8 — Hardening: accessibility and security

Once the core worked, two dedicated passes:

- **Accessibility** (issue #24): ARIA labelling on tool cards and the thread,
  `:focus-visible` rings, `prefers-reduced-motion` on the status spinner, and a
  composer that stays usable while the agent is working.
- **Security**: request-body cap (1 MB) and message-count cap on the backend,
  sheet-name sanitising so a hostile sheet name cannot inject into the system
  prompt, a closed tool allowlist so an unknown tool name never reaches the pane,
  CORS removed entirely in favour of the single origin, and a `.gitignore` gap
  that had left an RSA private key committable.

### Stage 9 — Cleanup

Branches, PRs and issues get closed or documented. Feature ideas that were
considered and rejected (streaming, Excel Online) were closed **as issues with a
reason**, not silently dropped — that record is the evidence of scope control.

Two traps here.

**`git branch --merged` lies when you squash-merge.** A squash creates a new
commit, so the branch's own commits are never ancestors of `main` and every
merged branch still reports as unmerged. Key deletion off the **pull request
state**, not off git ancestry:

```bash
gh pr list --state merged --json headRefName -q '.[].headRefName'
```

**Cleanup is the stage most often skipped, including here.** At the time of
writing this repository still has nine remote branches whose work is fully merged
— readable, because every branch maps to a numbered PR, but not tidy. Worth
saying plainly rather than describing an aspiration as a practice. And per 4.4,
whoever cleans up should delete only branches they created or whose PR they can
see is merged.

---

## 3. What was achieved

Numbers from the repository, not estimates:

| | |
|---|---|
| Commits on `main` | 61 |
| Pull requests merged | 25 |
| Issues opened and resolved | 17 |
| Tests | 157 (29 backend, 128 frontend) |
| Type-check | clean, enforced in CI |

Shipped, verified running in desktop Excel against a 1,000,000-cell workbook:

- Chat pane with the agent loop **in the pane**, because tools can only execute
  inside Excel's WebView; the backend is one stateless Claude turn
- Five tools: workbook metadata with header preview, ranged read, write, current
  selection, and `find`
- Size discipline: metadata first, 2,000-cell read cap, 8 tool rounds per
  message, then a forced text reply
- Every tool call visible as a card in the thread — nothing happens invisibly
- Write-confirm: **Apply / Don't write** before any cell changes
- Undo that restores formulas rather than the numbers they displayed
- Conversation persistence keyed by workbook, surviving pane reload
- Markdown rendering, formula explainer, demo prompt chips
- CI running both suites on every push

The demo, end to end: *"How big is this workbook?"* answers from a single
metadata call with no read of the 5,000 × 200 sheet; *"Fix the hard-coded Gross
profit"* chains find → read-with-formulas → write, pauses for Apply, and changes
`Budget!D4` from `1000` to `=D2-D3` — confirmed by reading the saved `.xlsx`, not
by screenshot.

---

## 4. What actually went wrong

These are the real incidents, in the order they hurt.

### 4.1 Two agents ran `npm install` in the same folder

The dev server started failing with `Can't resolve
'html-webpack-plugin/lib/loader.js'`. The package looked corrupt. It was not:
one agent ran `npm ci && npm test`, a second launched `rm -rf node_modules && npm
install`, and I had started my own `npm ci` in between. Each deletes and rebuilds
`node_modules` underneath the others, so webpack was compiling against a
directory being emptied. Every package was equally broken; the error just named
the first one webpack happened to need.

Recovering needed all three stopped, then `rm -rf node_modules` by hand (an
`npm ci` alone failed with `ENOTEMPTY`), then a single clean install.

> **Rule: one owner for dependencies and long-running processes.** Installs, the
> dev server, and the backend belong to exactly one agent. Everyone else treats
> them as read-only infrastructure.

### 4.2 The Python venv was rebuilt with a different interpreter

`pytest` died with `No module named '_pytest'` while the package sat plainly on
disk. An agent had recreated `.venv` using its own **app-bundled Python 3.12**,
while the installed packages lived in the `python3.14` directory. Nothing could
import. It also silently coupled the backend to a path inside another
application's bundle.

> **Rule: pin the interpreter in the setup script**, and never let an agent
> recreate a shared environment with whatever Python it happens to be running.

### 4.3 Two agents shipped the same fix

I opened a PR fixing chat auto-scroll. While it sat in review, another agent
shipped the identical fix inside a larger feature PR. Mine became a duplicate and
I closed it. Later the same thing nearly happened again on a documentation gap.

> **Rule: claim the issue before writing code**, and re-check `main` immediately
> before opening the PR. Cheap, and it saves a whole branch.

### 4.4 Cleanup ate someone else's work

A branch and worktree were pruned by another agent's tidy-up while a commit still
lived only there. Separately, an agent's `git stash` was left pointing at a
parent object that was never written, so the stash was unrecoverable.

> **Rule: cleanup only touches worktrees you created.** Never prune branches you
> did not make.

### 4.5 The environment itself generated conflicts

The repo sat in `~/Documents`, which syncs to iCloud. With several agents writing
quickly, iCloud produced conflict copies: `FINAL_PLAN 2.md`, `agent 2.py`,
`frontend/dist 2/`. Some were stale copies of live source files sitting inside
the Python package directory. `git status` collapses untracked directories, so
`git add -A` would have committed build output and an old `agent.py` without
anyone noticing.

> **Rule: keep multi-agent repos out of synced folders**, and add a
> `* [0-9].*` / `* [0-9]/` ignore rule as a backstop.

### 4.6 A stale `index.lock` froze every agent

One agent crashed mid-write and left `.git/index.lock`. Every other agent's git
command failed until it was removed. Worth checking first whenever git behaves
strangely.

---

## 5. The playbook

### 5.1 Before any agent starts

1. **Write the architecture yourself.** The loop, the state boundaries, and the
   tool contract are the parts that decide whether the codebase stays coherent.
   Agents fill in shapes; they do not choose them.
2. **Turn the plan into issues that name files.** One issue per agent. Each lists
   its exclusive files, its acceptance test, and what it must not touch.
3. **Freeze the contract first.** Here it was `backend/app/tools.py` (schemas)
   and `dispatchExcelTool` (executors). Both sides of a contract are the single
   most conflict-prone thing in a parallel build; land them before fanning out.
4. **Set the constraint prompt.** Parallel agents are additive by default: they
   add files, features and abstractions rather than refine what exists. Counter
   it explicitly:

   > Reject any addition that increases file count or state complexity unless the
   > issue asks for it. Prefer refining an existing function to introducing a new
   > abstraction. If you believe a new module is required, say why first.

   Without this, the sprint grew a guided tour, an "explain like I'm five" mode,
   clarifying-question buttons and follow-up chips — all later deleted for scope.
   That deletion was the right call, but the work was wasted.

### 5.2 While they run

| Rule | Why |
|---|---|
| Trunk only, short-lived branches, squash merge | Long-lived branches guarantee conflicts with several agents |
| One worktree per agent, created by that agent | Isolates working trees; nobody stashes over anyone |
| Exclusive file lists per issue | The only reliable collision prevention |
| One owner for installs, dev server, backend | See 4.1 and 4.2 |
| Re-check `main` before opening a PR | See 4.3 |
| Never discard another agent's uncommitted work | If the tree is dirty, stop and report |

### 5.3 The QA agent is the highest-value seat

Run a continuous QA agent on a fixed interval with a standing brief: pull, run
both suites, type-check, check contract drift, read recent diffs, and hunt for
real bugs. Give it authority to open PRs but never to commit to trunk.

In this project that loop found, with live proof each time:

- **Undo destroyed formulas.** The snapshot read `.values` (computed results) and
  restored them, so undoing a write over `=D2-D3` left the number `1000`. Silent
  data loss, reachable from the documented demo.
- **Conversations leaked between workbooks.** Saved chats were keyed by sheet
  names, and every blank workbook is `Sheet1`.
- **The reopened thread contradicted the sheet.** Persistence only ran on one
  code path, so an undo note vanished and the restored thread still claimed a
  write had landed.
- **An RSA private key was one `git add -A` from being committed.**
- **A history-truncation bug** that would have made the API reject long
  conversations mid-demo.

None were visible from the diff alone. All were found by running the thing.

### 5.4 Verify by consequence, not by screenshot

The strongest verification is an **observable side effect outside the app**. For
a spreadsheet agent that means: drive the pane, then read the `.xlsx` from disk
and assert on the cell.

That is how the undo bug was proved rather than argued:

```
Budget!D4 held  =D2-D3
overwrite with 0, press Apply, press Undo
read the saved file  →  D4 = 1000        # formula gone
```

and how the fix was confirmed:

```
same sequence after the fix
read the saved file  →  D4 = "=D2-D3"    # formula restored
```

A screenshot would have shown a plausible number in both cases.

### 5.5 Close the loop programmatically

The remaining gap in this workflow is that a human still drives the pane. The
next step is a UI automation agent — Playwright or the Office add-in test
harness — so a card is not "done" until a script clicks **Apply** and asserts the
resulting cell. Today the unit tests cover pure logic and the Office.js layer is
hand-checked; that boundary is defensible for four hours but is the first thing I
would automate with more time.

### 5.6 Prefer tool schemas to prompt conventions

Anything the model must emit for the app to parse should be a **tool with a JSON
schema**, not a string convention in the system prompt. Agents drop conventions
under context pressure, and the failure is silent.

This bit us concretely: the clarifying-question feature told the model to reply
with lettered options inside a blockquote, and the parser required the letter at
the start of the line. The `> ` prefix defeated it, so the buttons never
appeared — and because the text still rendered, it looked like it worked. The
tests passed because every fixture used the unprefixed form.

---

## 6. Scope integrity

Parallel agents produce volume, and volume without a split reads as scope creep —
or worse, as dishonesty about effort. Keep an explicit table of **what shipped in
the timebox** versus **what came after**.

| | |
|---|---|
| Four-hour core | Chat pane, agent loop, five Excel tools, cell caps, one HTTPS origin |
| Week one | Tool cards, write-confirm, undo, persistence, request limits, accessibility, security caps, CI, Markdown rendering |
| Cut deliberately | Guided tour, ELI5, clarifying buttons, follow-up chips, streaming, Excel Online, write-preview diffing |

Three habits keep that table honest:

1. **Close rejected ideas as issues with a reason.** Streaming (#30) and Excel
   Online (#31) were opened, considered, and closed as out of scope. A closed
   issue with a rationale is evidence of judgement; a silently dropped idea is
   indistinguishable from having never thought of it.
2. **Delete what does not earn its place.** The sprint grew a guided tour, an
   ELI5 mode, clarifying-question buttons and follow-up chips. All four were
   removed in one commit, with their tests, and the README's cut list updated in
   the same change. Removing agent output is a normal part of the loop, not an
   admission of failure.
3. **Never let an agent widen its own scope.** Issues name files. An agent that
   wants a new module has to say why first.

---

## 7. Have an agent rate you as the interviewer

The most useful review of this project came from pointing an agent at it with a
different brief: *you are the hiring engineer, score this submission and be
specific about what would worry you.*

That produced findings a QA pass never would, because they were about **signal**
rather than correctness:

- *"The after-list is now longer than the core. Looks like a swarm unless you say
  so."* → became the scope table in section 6.
- *"`CRUNCHED_ASK` is still a prompt convention, not a tool schema."* → a real
  fragility; the same class of bug later bit the clarifying-question parser
  (5.6).
- *"The card exists in code. Untested UI is the new risk."* → drove
  verification-by-consequence (5.4), and remains the honest gap (5.5).

Those findings were then worked as issues like any other (#35, "fix
interviewer-audit findings"), not treated as commentary.

Two things make this work. Give the reviewer the **artefacts an interviewer would
actually see** — the repo, the README, the PR history — not your intentions. And
ask for a **score with reasons**, because a rubric forces prioritisation in a way
"any feedback?" does not.

---

## 8. Talking about it

| Question | Answer |
|---|---|
| How do you use AI in your process? | I own the architecture, the state boundaries and the tool contract in the first pass. Agents implement against issues that name their files. |
| How do you avoid an AI-generated mess? | Trunk-based development, one worktree per agent, exclusive file ownership, squash merges, and a continuous QA agent that can open PRs but never commit to trunk. |
| How do you know it works? | Verification by consequence. The agent drives Excel, then a script reads the saved workbook and asserts the cell. That is how the undo data-loss bug was found and how its fix was confirmed. |
| What did you cut? | A documented list, kept separate from what shipped in the timebox. Rejected ideas are closed issues with reasons, not silence. |
| How do you control scope with agents that keep adding? | Issues name files, not goals. A constraint prompt rejects new abstractions unless asked for. Four features were deleted in one commit when they stopped earning their place. |
| How do you know your own judgement is not the bottleneck? | I have an agent review the work as the hiring engineer, with a rubric, and work its findings as issues. |

The honest headline: parallel agents multiply throughput and they multiply
environment failures. Most of the operational cost in this project was not code —
it was three agents fighting over one `node_modules`, one virtual environment,
and one git index.
