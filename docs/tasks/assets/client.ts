// src/agent/client.ts
// Model configuration + SDK client factories. Written verbatim in T1; FROZEN afterwards.
// Browser: the real API key never enters the bundle. The SDK is pointed at the same-origin path
// /api/anthropic, which webpack-dev-server proxies to https://api.anthropic.com while injecting
// x-api-key from .env (see webpack.config.js, task T2). Node (scripts/agentSmoke.ts): the SDK reads
// ANTHROPIC_API_KEY from the environment directly.
import Anthropic from '@anthropic-ai/sdk';
import { DEFAULT_EFFORT, DEFAULT_MODEL, LIMITS, REASON_EFFORT, REASON_MODEL, type Effort } from './contracts';

interface CrunchedEnv {
  ANTHROPIC_MODEL?: string;
  ANTHROPIC_REASON_MODEL?: string;
  ANTHROPIC_EFFORT?: string;
  CRUNCHED_STREAMING?: string;
  /** Demo-only escape hatch: when set the pane calls api.anthropic.com directly and the key IS in the bundle. */
  DIRECT_ANTHROPIC_KEY?: string;
}

// Injected by webpack.DefinePlugin as a JSON object literal (T2). Undeclared in Node/vitest, hence the typeof guard.
declare const __CRUNCHED_ENV__: CrunchedEnv | undefined;

function readEnv(): CrunchedEnv {
  if (typeof __CRUNCHED_ENV__ !== 'undefined' && __CRUNCHED_ENV__) return __CRUNCHED_ENV__;
  if (typeof process !== 'undefined' && process.env) {
    const e = process.env;
    return {
      ANTHROPIC_MODEL: e.ANTHROPIC_MODEL,
      ANTHROPIC_REASON_MODEL: e.ANTHROPIC_REASON_MODEL,
      ANTHROPIC_EFFORT: e.ANTHROPIC_EFFORT,
      CRUNCHED_STREAMING: e.CRUNCHED_STREAMING,
      DIRECT_ANTHROPIC_KEY: e.DIRECT_ANTHROPIC_KEY,
    };
  }
  return {};
}

const env = readEnv();

export const MODEL_CONFIG = {
  defaultModel: env.ANTHROPIC_MODEL || DEFAULT_MODEL,
  reasonModel: env.ANTHROPIC_REASON_MODEL || REASON_MODEL,
  defaultEffort: (env.ANTHROPIC_EFFORT || DEFAULT_EFFORT) as Effort,
  reasonEffort: REASON_EFFORT,
  streaming: (env.CRUNCHED_STREAMING || 'true') !== 'false',
  directKey: env.DIRECT_ANTHROPIC_KEY || '',
} as const;

/** Browser client for the task pane. Same-origin proxy holds the key unless DIRECT_ANTHROPIC_KEY is set. */
export function createClient(): Anthropic {
  if (MODEL_CONFIG.directKey) {
    return new Anthropic({
      apiKey: MODEL_CONFIG.directKey,
      dangerouslyAllowBrowser: true, // throwaway take-home key only; see README "Security notes"
      maxRetries: 2,
      timeout: LIMITS.REQUEST_TIMEOUT_MS,
    });
  }
  return new Anthropic({
    apiKey: 'injected-by-proxy',
    baseURL: `${window.location.origin}/api/anthropic`,
    dangerouslyAllowBrowser: true, // safe here: no secret in the browser; the proxy holds the key
    maxRetries: 2,
    timeout: LIMITS.REQUEST_TIMEOUT_MS,
  });
}

/** Node client for scripts (reads ANTHROPIC_API_KEY from the environment). */
export function createNodeClient(): Anthropic {
  return new Anthropic({ maxRetries: 2, timeout: LIMITS.REQUEST_TIMEOUT_MS });
}
