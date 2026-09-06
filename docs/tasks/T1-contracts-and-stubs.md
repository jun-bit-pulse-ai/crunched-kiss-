# T1 — Contracts, dependencies, stubs, test config (the fan-out gate)

**Agent model:** claude-sonnet-5. **Estimate:** 15 min. **Depends on:** T0 (committed). **Unblocks:** T2, T3, T4, T5, T6.
**Runs in:** the main working tree `/Users/junseki/Documents/GitHub/crunched-kiss` on branch `main` (no worktree; you are the only agent).

## Context

You are working in `/Users/junseki/Documents/GitHub/crunched-kiss`, an Excel task-pane add-in
(React 18 + TypeScript + webpack 5 + Office.js). A dev server is running in another terminal; NEVER run
`npm start`, `npm run stop`, `npm run dev-server`, or anything that opens Excel or a browser. Read `CLAUDE.md`
first. You are the ONLY task allowed to run `npm install`/`npm uninstall` and to edit `package.json`,
`tsconfig.json`, `.gitignore`; after you finish, those files are frozen.

Your job is purely mechanical: lay down the shared contract, the config, and compilable stubs so five
other agents can work in parallel on disjoint files. Make no design decisions; everything is specified below.

## You own (create or replace exactly these)

- `package.json`, `package-lock.json` (dependency changes + scripts only)
- `tsconfig.json`, `vitest.config.ts`, `.gitignore`, `.env.example`
- `src/agent/contracts.ts` (verbatim copy), `src/agent/client.ts` (verbatim copy), `src/agent/contracts.test.ts`
- `src/agent/loop.ts` (stub), `src/excel/tools.ts` (stub), `src/excel/devConsole.ts` (final)
- `src/taskpane/index.tsx` (stub)
- Moving `src/app/**` and `test/**` into `docs/history/two-process-draft/frontend/` (the one allowed touch of `docs/`)

Do NOT touch: `manifest.xml`, `webpack.config.js`, `src/taskpane/taskpane.html`, `src/commands/**`, `assets/**`, `CLAUDE.md`, `docs/PLAN.md`, `docs/tasks/**`.

## Steps

1. Dependencies (once, for everyone):
   ```bash
   npm uninstall mocha ts-node @types/mocha
   npm i @anthropic-ai/sdk
   npm i -D vitest dotenv exceljs tsx
   ```
   In `package.json` `scripts`, set `"test": "vitest run"`, add `"typecheck": "tsc --noEmit -p tsconfig.json"`,
   `"smoke": "tsx scripts/agentSmoke.ts"`, `"fixture": "node scripts/make-fixture.mjs"`. Keep `build`,
   `dev-server`, `start`, `stop`, `validate` and the `config` block exactly as they are.

2. `tsconfig.json`: change `"include": ["src", "test"]` to `"include": ["src", "scripts"]`; delete the
   `"ts-node"` block; keep everything else. If `npx tsc` later complains about Node globals in
   `scripts/` or `process` in `src/agent/client.ts`, add `"types": ["node", "office-js", "react", "react-dom"]`
   to `compilerOptions` (all four packages are installed).

3. `vitest.config.ts` at the repo root:
   ```ts
   import { defineConfig } from 'vitest/config';
   export default defineConfig({ test: { include: ['src/**/*.test.ts'], environment: 'node' } });
   ```

4. Archive the old two-process UI so it stops compiling into the bundle:
   ```bash
   mkdir -p docs/history/two-process-draft/frontend
   git mv src/app docs/history/two-process-draft/frontend/app
   git mv test docs/history/two-process-draft/frontend/test
   ```

5. Contract files, VERBATIM:
   ```bash
   mkdir -p src/agent src/excel
   cp docs/tasks/assets/contracts.ts src/agent/contracts.ts
   cp docs/tasks/assets/client.ts src/agent/client.ts
   ```
   Both files have been type-checked against the current `@anthropic-ai/sdk`. If `npx tsc` nevertheless
   reports an error inside them (e.g. an SDK version without `strict` on `Anthropic.Tool`), make the
   smallest fix (a cast) and report the exact diff in your final message under `CONTRACT DEVIATION`.

6. Stubs with the final signatures (other tasks replace the whole file):
   `src/agent/loop.ts`
   ```ts
   import type { RunAgentTurn } from './contracts';
   /** Replaced by task T4. */
   export const runAgentTurn: RunAgentTurn = async () => {
     throw new Error('runAgentTurn not implemented yet (task T4)');
   };
   ```
   `src/excel/tools.ts`
   ```ts
   import type { ToolExecutor } from '../agent/contracts';
   /** Replaced by task T3. */
   export const executeTool: ToolExecutor = async (name) => ({ ok: false, error: `${name} not implemented yet (task T3)` });
   ```

7. `src/excel/devConsole.ts` (final; nobody else edits it):
   ```ts
   import type Anthropic from '@anthropic-ai/sdk';
   import { createClient } from '../agent/client';
   import type { ToolExecutor } from '../agent/contracts';
   import { executeTool } from './tools';

   declare global {
     interface Window { crunched?: { run: ToolExecutor; client: Anthropic } }
   }
   /** Exposes `crunched.run(name, input)` and `crunched.client` in the pane's Web Inspector for manual testing. */
   export function installDevConsole(): void {
     window.crunched = { run: executeTool, client: createClient() };
   }
   ```

8. `src/taskpane/index.tsx` (stub; T5 replaces it). Keep the `#container` root and `Office.onReady`:
   ```tsx
   import { createRoot } from 'react-dom/client';
   import { installDevConsole } from '../excel/devConsole';

   /* global document, Office */
   Office.onReady(() => {
     installDevConsole();
     const el = document.getElementById('container');
     if (!el) return;
     createRoot(el).render(
       <div style={{ padding: 16, fontFamily: 'system-ui, sans-serif', fontSize: 13 }}>
         Crunched — chat UI pending (task T5). Open Web Inspector and try{' '}
         <code>await crunched.run('get_workbook_overview', {'{}'})</code>.
       </div>,
     );
   });
   ```

9. `src/agent/contracts.test.ts` (vitest) with these cases:
   - `isToolName('read_range')` is true; `isToolName('nope')` is false; `TOOL_NAMES.length === 7`;
     `TOOL_DEFINITIONS.map(t => t.name)` equals `TOOL_NAMES` in the same order; every definition has
     `strict === true`, `input_schema.additionalProperties === false`, and `required` listing every key of `properties`.
   - `formatOverviewForPrompt` of a 2-sheet overview contains both sheet names; with `maxChars = 60` it ends with `[truncated]`.
   - `overviewSignature` changes when a sheet's row count changes and is stable otherwise.
   - `serializeToolData('read_range', page)` for a 2x2 page (`columns: ['A','B']`, `firstRow: 5`,
     `mode: 'formulas'`, `formulas` with `'=A5*2'` at [0][1]) starts with `Sheet1!A5:B6 | mode=formulas | page 1/1`,
     its second line is `\tA\tB`, and its third line contains `{=A5*2}`.
   - `serializeToolData('write_range', result)` does not contain the string `"previous"` but contains `previousCells`.
   - `capToolResultText('x'.repeat(100), 10)` has length < 100 and contains `[truncated`.

10. `.env.example` (exactly):
    ```
    ANTHROPIC_API_KEY=sk-ant-...
    ANTHROPIC_MODEL=claude-sonnet-5
    ANTHROPIC_REASON_MODEL=claude-opus-5
    ANTHROPIC_EFFORT=medium
    CRUNCHED_STREAMING=true
    DIRECT_ANTHROPIC_KEY=
    ```

11. `.gitignore`: replace with
    ```
    node_modules/
    dist/
    .env
    fixtures/*.xlsx
    .office-addin-dev-certs/
    .DS_Store
    *.log
    ```

12. Verify, then commit:
    ```bash
    npx tsc --noEmit -p tsconfig.json
    npm test
    npx webpack --mode development && ! grep -rl 'sk-ant' dist/ ; rm -rf dist
    git add -A && git commit -m "T1: contracts, client, stubs, vitest, deps"
    ```
    (`npx webpack` evaluates `webpack.config.js`, which reads office-addin-dev-certs installed in T0; it must not start a server.)

## Acceptance (the orchestrator checks these before fan-out)

- `npx tsc --noEmit -p tsconfig.json` and `npm test` pass; `npm test` shows the contracts test file with >= 8 passing cases.
- `diff docs/tasks/assets/contracts.ts src/agent/contracts.ts` and `diff docs/tasks/assets/client.ts src/agent/client.ts` print nothing (or the exact deviation is reported).
- `grep -c 'strict: true,' src/agent/contracts.ts` prints 7.
- `node -p "Object.keys(require('./package.json').devDependencies).join(' ')"` includes vitest, dotenv, exceljs, tsx and not mocha/ts-node; `dependencies` includes `@anthropic-ai/sdk`.
- `git show --stat HEAD` touches only the owned files plus the two `git mv` moves.

## When done, report

Final message sections: `Files` (created/changed/moved), `Verification` (the commands and their results),
`CONTRACT DEVIATION` (only if step 5 needed a fix, with the exact diff), `Notes for the human`
(anything odd about the scaffold, e.g. tsconfig `types` added).
