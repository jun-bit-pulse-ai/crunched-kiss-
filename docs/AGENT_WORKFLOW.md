# Running multiple coding agents on one repo

How this project was actually built, what broke, and the playbook I would hand
someone starting the same way tomorrow.

This is a field report, not theory. Every failure below happened in this
repository during the sprint, and the fixes are the ones that worked.

---

## 1. Context

`crunched-kiss` is a four-hour take-home: an Excel task-pane agent that chats
with Claude and reads and writes the live workbook, including sheets with a
million cells. It was built by **three coding agents working in parallel on one
trunk**, plus a fourth continuous QA agent, and then hardened over a longer
"week one" tail.

The shape of the work:

```
research → plan → conflict-free implementation spec
                          │
        ┌─────────────────┼─────────────────┐
     agent A           agent B           agent C        (parallel, issue-scoped)
        └─────────────────┼─────────────────┘
                     integration
                          │
                    QA agent (loop)  ──→ bug PRs ──→ merge
```

The split that made it work was **file ownership, not task ownership**. Each
GitHub issue named the exact files its agent was allowed to touch. Agents that
share a goal collide; agents that share a *file* corrupt each other's work.

---

## 2. What actually went wrong

These are the real incidents, in the order they hurt.

### 2.1 Two agents ran `npm install` in the same folder

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

### 2.2 The Python venv was rebuilt with a different interpreter

`pytest` died with `No module named '_pytest'` while the package sat plainly on
disk. An agent had recreated `.venv` using its own **app-bundled Python 3.12**,
while the installed packages lived in the `python3.14` directory. Nothing could
import. It also silently coupled the backend to a path inside another
application's bundle.

> **Rule: pin the interpreter in the setup script**, and never let an agent
> recreate a shared environment with whatever Python it happens to be running.

### 2.3 Two agents shipped the same fix

I opened a PR fixing chat auto-scroll. While it sat in review, another agent
shipped the identical fix inside a larger feature PR. Mine became a duplicate and
I closed it. Later the same thing nearly happened again on a documentation gap.

> **Rule: claim the issue before writing code**, and re-check `main` immediately
> before opening the PR. Cheap, and it saves a whole branch.

### 2.4 Cleanup ate someone else's work

A branch and worktree were pruned by another agent's tidy-up while a commit still
lived only there. Separately, an agent's `git stash` was left pointing at a
parent object that was never written, so the stash was unrecoverable.

> **Rule: cleanup only touches worktrees you created.** Never prune branches you
> did not make.

### 2.5 The environment itself generated conflicts

The repo sat in `~/Documents`, which syncs to iCloud. With several agents writing
quickly, iCloud produced conflict copies: `FINAL_PLAN 2.md`, `agent 2.py`,
`frontend/dist 2/`. Some were stale copies of live source files sitting inside
the Python package directory. `git status` collapses untracked directories, so
`git add -A` would have committed build output and an old `agent.py` without
anyone noticing.

> **Rule: keep multi-agent repos out of synced folders**, and add a
> `* [0-9].*` / `* [0-9]/` ignore rule as a backstop.

### 2.6 A stale `index.lock` froze every agent

One agent crashed mid-write and left `.git/index.lock`. Every other agent's git
command failed until it was removed. Worth checking first whenever git behaves
strangely.

---

## 3. The playbook

### 3.1 Before any agent starts

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

### 3.2 While they run

| Rule | Why |
|---|---|
| Trunk only, short-lived branches, squash merge | Long-lived branches guarantee conflicts with several agents |
| One worktree per agent, created by that agent | Isolates working trees; nobody stashes over anyone |
| Exclusive file lists per issue | The only reliable collision prevention |
| One owner for installs, dev server, backend | See 2.1 and 2.2 |
| Re-check `main` before opening a PR | See 2.3 |
| Never discard another agent's uncommitted work | If the tree is dirty, stop and report |

### 3.3 The QA agent is the highest-value seat

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

### 3.4 Verify by consequence, not by screenshot

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

### 3.5 Close the loop programmatically

The remaining gap in this workflow is that a human still drives the pane. The
next step is a UI automation agent — Playwright or the Office add-in test
harness — so a card is not "done" until a script clicks **Apply** and asserts the
resulting cell. Today the unit tests cover pure logic and the Office.js layer is
hand-checked; that boundary is defensible for four hours but is the first thing I
would automate with more time.

### 3.6 Prefer tool schemas to prompt conventions

Anything the model must emit for the app to parse should be a **tool with a JSON
schema**, not a string convention in the system prompt. Agents drop conventions
under context pressure, and the failure is silent.

This bit us concretely: the clarifying-question feature told the model to reply
with lettered options inside a blockquote, and the parser required the letter at
the start of the line. The `> ` prefix defeated it, so the buttons never
appeared — and because the text still rendered, it looked like it worked. The
tests passed because every fixture used the unprefixed form.

---

## 4. Scope integrity

Keep an explicit table of **what shipped in the timebox** versus **what came
after**. Parallel agents produce volume, and volume without that split reads as
scope creep or as dishonesty about effort.

| | |
|---|---|
| Four-hour core | Chat pane, agent loop, the five Excel tools, cell caps, one HTTPS origin |
| Week one | Tool cards, write-confirm, undo, persistence, request limits, CI, Markdown rendering |
| Cut deliberately | Guided tour, ELI5, clarifying buttons, follow-up chips, streaming, Excel Online |

The cut row matters most. It shows the additive pressure in 3.1 was recognised
and reversed.

---

## 5. Talking about it

| Question | Answer |
|---|---|
| How do you use AI in your process? | I own the architecture, the state boundaries and the tool contract in the first pass. Agents implement against issues that name their files. |
| How do you avoid an AI-generated mess? | Trunk-based development, one worktree per agent, exclusive file ownership, squash merges, and a continuous QA agent that can open PRs but never commit to trunk. |
| How do you know it works? | Verification by consequence. The agent drives Excel, then a script reads the saved workbook and asserts the cell. That is how the undo data-loss bug was found and how its fix was confirmed. |
| What did you cut? | A documented list, kept separate from what shipped in the timebox. |

The honest headline: parallel agents multiply throughput and they multiply
environment failures. Most of the operational cost in this project was not code —
it was three agents fighting over one `node_modules`, one virtual environment,
and one git index.
