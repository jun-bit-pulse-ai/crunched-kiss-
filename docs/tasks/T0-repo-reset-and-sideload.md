# T0 — Reset the repo to the single-process layout and prove the pane opens in desktop Excel

**Who:** the HUMAN, in a terminal. Not an agent task (cert trust and Excel launch are interactive).
**Estimate:** 25 min (0:00-0:25). **Depends on:** nothing. **Unblocks:** T1.

## Why this shape

`main` currently carries a committed two-process draft: `frontend/` (a complete hand-rolled Office-Addin-TaskPane-React
layout: webpack-dev-server 5, office-addin-dev-certs, office-addin-debugging, manifest on https://localhost:3000, with
`node_modules` already installed), `backend/` (FastAPI), `certs/`, helper scripts, `FINAL_PLAN.md` plus three earlier plans,
and a README that replaced the exercise text (the original is at commit `2987b87`). `yo` is not installed and this machine
runs Node 24; the old README records that generator-office refused an odd Node version once. So: **do not run `yo office`.**
Promote `frontend/` to the repo root (it is the official quickstart layout minus sample code) and archive everything
two-process under `docs/history/` with `git mv`, so the grader finds one architecture and the process history in one place.
`yo office` stays the fallback only if step 5 fails for 15 minutes.

Keep `src/app/**` (the old chat UI) in place for now: it renders without a backend and is enough to prove the WebView loads
our origin. T1 archives it.

## Steps

### 1. Sync and archive the two-process draft (5 min)

```bash
cd /Users/junseki/Documents/GitHub/crunched-kiss
git pull --ff-only                                   # another session pushes to this repo; start from origin's tip
rm -rf backend/.venv .venv backend/app/__pycache__ backend/tests/__pycache__ backend/.pytest_cache
mkdir -p docs/history/two-process-draft
git mv README.md docs/history/two-process-draft/README-two-process.md
git show 2987b87:README.md > README.md               # restore the original exercise text; T8 keeps it at the bottom of the final README
git mv IMPLEMENTATION_PLAN.md ai-orchestrator-plan.md CLAUDE_MACBOOK_PLAN.md FINAL_PLAN.md docs/history/
git mv backend docs/history/two-process-draft/backend
git mv scripts docs/history/two-process-draft/scripts
git rm -r -q certs .env.example
```

### 2. Promote `frontend/` to the root (3 min)

```bash
mv frontend/node_modules ./node_modules 2>/dev/null || true       # reuse the install already done in frontend/
git mv frontend/manifest.xml frontend/package.json frontend/package-lock.json frontend/tsconfig.json frontend/webpack.config.js .
git mv frontend/assets assets
git mv frontend/src src
git mv frontend/test test
rmdir frontend || { ls -la frontend; echo "remove leftovers above, then: rm -rf frontend"; }
ls manifest.xml package.json webpack.config.js tsconfig.json src/taskpane/index.tsx assets/icon-32.png   # all must exist
cp docs/tasks/assets/CLAUDE.md CLAUDE.md
cp docs/tasks/assets/TIMELOG.md docs/TIMELOG.md
```

### 3. Install and configure (4 min)

```bash
npm install                                         # fast: node_modules was moved; this only reconciles the lockfile
KEY=$(grep -E '^ANTHROPIC_API_KEY=' .env 2>/dev/null | cut -d= -f2-)   # keep a key if .env already has one
cat > .env <<EOF
ANTHROPIC_API_KEY=${KEY:-sk-ant-PASTE_THE_COMPANY_KEY_HERE}
ANTHROPIC_MODEL=claude-sonnet-5
ANTHROPIC_REASON_MODEL=claude-opus-5
ANTHROPIC_EFFORT=medium
CRUNCHED_STREAMING=true
DIRECT_ANTHROPIC_KEY=
EOF
grep -c 'PASTE' .env && echo "^ if 1: paste the company key into .env now"
git check-ignore .env            # must print ".env" (the committed .gitignore already lists it)
defaults write com.microsoft.Excel OfficeWebAddinDeveloperExtras -bool true   # enables right-click > Inspect Element in the pane
```

### 4. Cert spike in Safari BEFORE Excel (3 min)

Safari shares the macOS trust store with Excel's WKWebView; Chrome proves nothing.

```bash
npx office-addin-dev-certs install     # one sudo/Keychain prompt; installs the localhost dev CA
npx office-addin-dev-certs verify      # must say the certificate is valid/installed
npm run dev-server                     # webpack serve on https://localhost:3000
```

Open **https://localhost:3000/taskpane.html** in Safari. Expected: no certificate warning; the old "Crunched" chat UI
renders (its send button will fail without a backend; irrelevant). If Safari warns:
`npx office-addin-dev-certs install --machine`, then reload. When green, stop the dev server with Ctrl-C (step 5 starts its own).

### 5. Sideload into desktop Excel (5 min)

```bash
osascript -e 'quit app "Microsoft Excel"' 2>/dev/null; sleep 2
npm start        # = office-addin-debugging start manifest.xml: starts the HTTPS dev server, copies manifest.xml into
                 #   ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/, launches Excel
```

In Excel: open a blank workbook, **Home** tab, group "Crunched", button "Crunched" -> the task pane opens with the old chat UI
over https://localhost:3000. If the button is missing: **Insert > Add-ins > My Add-ins > Developer Add-ins > Crunched**.
Right-click inside the pane -> **Inspect Element** must open Safari Web Inspector (console shows no cert errors).

Blank-pane playbook (in order):
1. `npx office-addin-dev-certs verify`; if not valid, `npx office-addin-dev-certs install --machine`.
2. Quit Excel; `rm -rf ~/Library/Containers/com.microsoft.Excel/Data/Library/Caches/*`; `npm run stop`; `npm start`.
3. `npx office-addin-manifest validate manifest.xml` (must pass).
4. Remove stale manifests: `ls ~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` and delete anything not named
   manifest.xml; retry.

Fallback if still blank after 15 minutes: `npx -y yo office` in the scratchpad (Task Pane / React / TypeScript / Excel / XML
manifest / name Crunched), rsync it into the root over the current files, `npm install`, `npm start` again. Everything downstream
only relies on `src/taskpane/index.tsx` mounting `#container` inside `Office.onReady`, which the generator also produces.

### 6. Commit and hand over (2 min)

```bash
git add -A
git status --short | grep -Ev '^(R|A|M|D) +(docs/|src/|assets/|test/|manifest.xml|package|tsconfig|webpack|CLAUDE.md|README.md|.gitignore)' ; echo "(anything printed above is unexpected)"
git commit -m "Reset to single-process layout: promote add-in scaffold to root, archive two-process draft"
git push                        # the repo is shared with another session; publish the new layout before agents branch from it
```

Leave `npm start` running in this terminal for the rest of the session. Agents never start or stop it.
Fill the first two rows of `docs/TIMELOG.md`. Start T1 (`docs/tasks/T1-contracts-and-stubs.md`).

## Acceptance

- Desktop Excel shows the "Crunched" task pane over https://localhost:3000 with no certificate warning, and right-click >
  Inspect Element opens Web Inspector.
- `git status` is clean after the commit; `ls` at the root shows no `frontend/`, `backend/`, `certs/`, `FINAL_PLAN.md`.
- `README.md` starts with the exercise text (`grep -c 'Time limit: 4 hours' README.md` prints 1).
- `git check-ignore .env` prints `.env`; `CLAUDE.md` and `docs/TIMELOG.md` exist.
- `npx office-addin-dev-certs verify` reports a valid installed certificate.
