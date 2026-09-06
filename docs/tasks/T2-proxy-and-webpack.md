# T2 — Key-holding proxy inside webpack-dev-server (+ env injection into the bundle)

**Agent model:** claude-sonnet-5. **Estimate:** 25 min. **Depends on:** T1. **Runs in parallel with:** T3, T4, T5, T6.
**Worktree:** `/Users/junseki/Documents/GitHub/ck-T2` on branch `task/T2` (created by the orchestrator; `node_modules` is a symlink to the main tree's).

## Context

You are working in an Excel task-pane add-in (React + TypeScript, webpack 5, webpack-dev-server 5.x which
bundles http-proxy-middleware 2.x). Read `CLAUDE.md` first. Read `src/agent/client.ts` (frozen): the browser
SDK is constructed with `apiKey: 'injected-by-proxy'` and `baseURL: window.location.origin + '/api/anthropic'`,
and it reads a global `__CRUNCHED_ENV__` object that YOUR DefinePlugin must provide. Never run `npm start` /
`npm run stop` / `npm run dev-server`; the human restarts the server. Do not run `npm install` (dotenv and
webpack are installed).

GOAL: the Anthropic API key lives only in Node (webpack-dev-server, read from `.env`) and never in the
browser bundle. The task pane calls the SDK against the same-origin path `/api/anthropic`, which the dev server
proxies to `https://api.anthropic.com`, injecting `x-api-key`. Same origin means no CORS, no second cert, no mkcert.

## You own

- `webpack.config.js` (only file you edit)

Do NOT touch anything else. In particular not `src/**`, `package.json`, `manifest.xml`, `.env`, `docs/**`.

## Steps

1. At the top of `webpack.config.js` add:
   ```js
   const webpack = require("webpack");
   require("dotenv").config({ path: path.resolve(__dirname, ".env") });
   const ANTHROPIC_API_KEY = process.env.ANTHROPIC_API_KEY || "";
   if (!ANTHROPIC_API_KEY) console.warn("[crunched] ANTHROPIC_API_KEY is not set in .env; /api/anthropic will return 401");
   ```
   (`path` is already required.) Delete the `mkcertHttpsOptions` function and its call: the `certs/` directory no
   longer exists. `getHttpsOptions()` must return the office-addin-dev-certs options only. Keep everything else the
   scaffold does (entries, HtmlWebpackPlugin x2, CopyWebpackPlugin, urlDev/urlProd replacement, `hot`, `headers`, port 3000).

2. In `plugins`, add (NEVER put the API key in here):
   ```js
   new webpack.DefinePlugin({
     __CRUNCHED_ENV__: JSON.stringify({
       ANTHROPIC_MODEL: process.env.ANTHROPIC_MODEL || "claude-sonnet-5",
       ANTHROPIC_REASON_MODEL: process.env.ANTHROPIC_REASON_MODEL || "claude-opus-5",
       ANTHROPIC_EFFORT: process.env.ANTHROPIC_EFFORT || "medium",
       CRUNCHED_STREAMING: process.env.CRUNCHED_STREAMING || "true",
       DIRECT_ANTHROPIC_KEY: process.env.DIRECT_ANTHROPIC_KEY || "",
     }),
   }),
   ```
   `DIRECT_ANTHROPIC_KEY` is a documented demo-only escape hatch (empty by default): when the human sets it, the
   pane bypasses the proxy and the key is in the bundle. Leave the `|| ""` default.

3. In `devServer`, add `compress: false` (SSE must not be buffered) and the proxy. Detect the installed versions first:
   ```bash
   node -p "require('webpack-dev-server/package.json').version"
   node -p "require('http-proxy-middleware/package.json').version"
   ```
   For webpack-dev-server 5 (array form) with http-proxy-middleware 2.x:
   ```js
   proxy: [
     {
       context: ["/api/anthropic"],
       target: "https://api.anthropic.com",
       changeOrigin: true,
       secure: true,
       pathRewrite: { "^/api/anthropic": "" },
       onProxyReq(proxyReq) {
         proxyReq.setHeader("x-api-key", ANTHROPIC_API_KEY);   // overrides the SDK's placeholder key
         proxyReq.removeHeader("origin");                     // make it look like a server-side request
         proxyReq.removeHeader("anthropic-dangerous-direct-browser-access");
       },
     },
   ],
   ```
   If http-proxy-middleware is 3.x, the hook is `on: { proxyReq(proxyReq) { ... } }` instead of `onProxyReq`.
   If webpack-dev-server is 4.x, use the object form `proxy: { "/api/anthropic": { ...same options... } }`.
   Also keep the `headers: { "x-api-key": ANTHROPIC_API_KEY }` option OUT (the hook is the single mechanism; two
   mechanisms make debugging ambiguous).

4. Verify without Excel:
   ```bash
   node -e "require('./webpack.config.js')({}, { mode: 'development' }).then(c => { console.log(JSON.stringify(c.devServer.proxy, (k, v) => typeof v === 'function' ? '[fn]' : v)); console.log('compress', c.devServer.compress) })"
   npx tsc --noEmit -p tsconfig.json
   npx webpack --mode development
   grep -rl 'sk-ant' dist/ ; echo "(nothing above = key not in bundle)"
   grep -c 'claude-sonnet-5' dist/taskpane.js   # >= 1: __CRUNCHED_ENV__ was inlined
   rm -rf dist
   git add webpack.config.js && git commit -m "T2: same-origin key-holding proxy + env injection"
   ```

## Acceptance (run by the human in the main tree after merging your branch and restarting the server with `npm run stop; npm start`)

```bash
# 1. models list through the proxy (no key sent by curl => proves injection)
curl -sk -H 'anthropic-version: 2023-06-01' https://localhost:3000/api/anthropic/v1/models | head -c 300
# 2. no key in the served bundle
curl -sk https://localhost:3000/taskpane.js | grep -c 'sk-ant'     # prints 0
# 3. SSE passes through unbuffered (first "event:" lines must appear within ~2 s)
curl -skN -X POST https://localhost:3000/api/anthropic/v1/messages \
  -H 'anthropic-version: 2023-06-01' -H 'content-type: application/json' \
  -d '{"model":"claude-sonnet-5","max_tokens":50,"stream":true,"messages":[{"role":"user","content":"say hi"}]}' | head -n 6
```
4. In the pane's Web Inspector (right-click the pane > Inspect Element), the SDK itself works through the proxy from inside WKWebView:
```js
await crunched.client.messages.create({ model: 'claude-sonnet-5', max_tokens: 20, messages: [{ role: 'user', content: 'say hi' }] })
```
returns a Message object (`stop_reason: 'end_turn'`), and the Network tab shows the request going to
`https://localhost:3000/api/anthropic/v1/messages`, never to api.anthropic.com.

If 4 fails with a 401: the hook did not run; check the middleware version and hook name. If it fails with a CORS
or network error: `baseURL` is wrong (client.ts is frozen; report it). If streaming (check 3) buffers: confirm
`compress: false` landed and that no `onProxyRes` was added.

## When done, report

Final message: the exact `devServer` block you added (proxy + compress), the webpack-dev-server and
http-proxy-middleware versions detected, the output of the four verification commands, and the reminder that the
human must run `npm run stop; npm start` before the acceptance curls. Include `CONTRACT CHANGE REQUEST` only if
`src/agent/client.ts` is incompatible with what you built.
