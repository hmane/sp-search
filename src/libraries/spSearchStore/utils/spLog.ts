/**
 * T5.D6 — central logging shim with PII redaction and production gate.
 *
 * Two reasons SP Search can't keep using `console.log(message, payload)`
 * directly:
 *
 *   1. **PII leak risk** — any object handed to `console.*` that
 *      contains `queryText` / `userEmail` / `userId` etc. ends up
 *      visible in DevTools forever. Tenant admins can't audit what
 *      every web part writes; the safe default is "redact at the
 *      source". Audit acceptance signal: "DevTools console shows 0
 *      occurrences of literal user query".
 *
 *   2. **Production noise** — `console.log` calls scattered across
 *      `SearchOrchestrator` / `SearchManagerService` / etc. fire on
 *      every keystroke + search. In a real tenant the F12 console
 *      should be quiet unless `?debug=1` is in the URL. Audit
 *      acceptance signal: "with `?debug=1` queryText reads
 *      `[redacted]`" — i.e. debug mode unlocks the verbose stream,
 *      but the redactor is unconditional.
 *
 * The shim is a thin wrapper around `console.*` — no dependency on
 * SPContext.logger or DebugCollector, so it works in tests, in
 * non-SPFx contexts, and at module-init time before any SPFx
 * context is wired.
 */

export const REDACTED_PLACEHOLDER = '[redacted]';

/**
 * Keys whose values are always stripped before logging. The list is
 * intentionally short and explicit — adding wildcard matching is a
 * recipe for false positives (e.g. `keyExpression` is not a PII key).
 *
 * Convention: canonical camelCase. Down-stream code should use these
 * exact names when serializing payloads so the redactor catches them.
 */
export const PII_KEYS: ReadonlyArray<string> = [
  'queryText',
  'searchTerms',
  'kql',           // SearchOrchestrator's DebugCollector.setLastQuery uses `kql` for the rendered query string
  'query',         // generic fallback when the orchestrator labels the field `query: ...`
  'userId',
  'userEmail',
  'userDisplayName',
  'email',
  'loginName',
  'upn',
];

/**
 * Deep-walks the value, replacing any PII-keyed leaf with the
 * `REDACTED_PLACEHOLDER` string. Returns a new object — never
 * mutates the input. Non-object primitives, null, and undefined
 * pass through.
 */
export function redactPII(value: unknown): unknown {
  if (value === null || value === undefined) {
    return value;
  }
  if (Array.isArray(value)) {
    return value.map(redactPII);
  }
  if (typeof value === 'object') {
    const source = value as Record<string, unknown>;
    const result: Record<string, unknown> = {};
    const keys = Object.keys(source);
    for (let i: number = 0; i < keys.length; i++) {
      const key = keys[i];
      if (PII_KEYS.indexOf(key) >= 0) {
        result[key] = REDACTED_PLACEHOLDER;
      } else {
        result[key] = redactPII(source[key]);
      }
    }
    return result;
  }
  return value;
}

export type LogLevel = 'debug' | 'info' | 'warn' | 'error';

export interface ILogEnvironment {
  /** True when running in a real tenant (not workbench / localhost). */
  isProduction: boolean;
  /** True when `?debug=1` / `?isDebug=1` / session-storage debug flag is active. */
  isDebug: boolean;
}

/**
 * Production gate. warn + error always log; info + debug are silenced
 * in production unless the admin has opted in via `?debug=1`. In
 * workbench / localhost everything logs.
 *
 * Exported so unit tests can drive the gate with a synthetic
 * environment instead of mucking with `window.location`.
 */
export function isLogLevelEnabled(level: LogLevel, env: ILogEnvironment): boolean {
  if (level === 'warn' || level === 'error') {
    return true;
  }
  if (!env.isProduction) {
    return true;
  }
  return env.isDebug;
}

// ─── Runtime environment detection ──────────────────────────

function detectEnvironment(): ILogEnvironment {
  if (typeof window === 'undefined' || !window.location) {
    return { isProduction: false, isDebug: false };
  }
  const host = window.location.hostname || '';
  // SPFx workbench is `localhost`; tenant URLs end in `sharepoint.com` /
  // `.sharepoint.us` / `.sharepoint.cn` etc. Treat anything not `localhost`
  // and not `127.0.0.1` as production for logging purposes.
  const isProduction = host !== 'localhost' && host !== '127.0.0.1';
  const search = window.location.search || '';
  const urlDebug = /[?&](isDebug|debug)=(1|true)\b/.test(search);
  let sessionDebug = false;
  try {
    sessionDebug = window.sessionStorage && window.sessionStorage.getItem('sp-search-debug') === '1';
  } catch {
    sessionDebug = false;
  }
  return { isProduction, isDebug: urlDebug || sessionDebug };
}

// ─── Public logger ──────────────────────────────────────────

/**
 * Source tag the shim prefixes onto every console line. Stable so
 * admins can grep for `[SP Search]` in F12 to find every log this
 * solution writes.
 */
const PREFIX = '[SP Search]';

function emit(level: LogLevel, message: string, payload: unknown): void {
  const env = detectEnvironment();
  if (!isLogLevelEnabled(level, env)) {
    return;
  }
  const redacted = payload === undefined ? undefined : redactPII(payload);
  // Map debug → console.log (DevTools' Verbose channel) so this works
  // in browsers without console.debug (rare in 2026 but cheap).
  const sink = level === 'debug' ? console.log
    : level === 'info' ? console.info
    : level === 'warn' ? console.warn
    : console.error;
  if (redacted === undefined) {
    sink(PREFIX, message);
  } else {
    sink(PREFIX, message, redacted);
  }
}

export const spLog = {
  debug(message: string, payload?: unknown): void { emit('debug', message, payload); },
  info(message: string, payload?: unknown): void { emit('info', message, payload); },
  warn(message: string, payload?: unknown): void { emit('warn', message, payload); },
  error(message: string, payload?: unknown): void { emit('error', message, payload); },
};
