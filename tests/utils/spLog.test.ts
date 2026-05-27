import {
  redactPII,
  isLogLevelEnabled,
  PII_KEYS,
  REDACTED_PLACEHOLDER,
} from '../../src/libraries/spSearchStore/utils/spLog';

/**
 * T5.D6 — logging discipline. Two pure behaviours under test:
 *
 *   1. `redactPII(payload)` — strips any PII-shaped key from the
 *      payload (queryText / userId / etc.) before it reaches the
 *      console. Returns a new object; never mutates the input.
 *
 *   2. `isLogLevelEnabled(level)` — production-gate predicate:
 *      info/debug are silenced in production unless `?debug=1` (or
 *      the session-storage debug flag) is set. warn/error always
 *      log regardless of environment.
 */

describe('redactPII', () => {
  it('redacts the canonical leak — queryText at the top level', () => {
    const result = redactPII({ queryText: 'budget report' });
    expect(result).toEqual({ queryText: REDACTED_PLACEHOLDER });
  });

  it('redacts nested queryText inside an arbitrary container', () => {
    const result = redactPII({
      action: 'SEARCH',
      payload: { queryText: 'budget', duration: 42 },
    }) as Record<string, unknown>;
    const payload = result.payload as Record<string, unknown>;
    expect(payload.queryText).toBe(REDACTED_PLACEHOLDER);
    expect(payload.duration).toBe(42);
  });

  it('redacts every key in the PII_KEYS list', () => {
    const input: Record<string, unknown> = {};
    for (let i: number = 0; i < PII_KEYS.length; i++) {
      input[PII_KEYS[i]] = 'sensitive-value-' + String(i);
    }
    const result = redactPII(input) as Record<string, unknown>;
    for (let i: number = 0; i < PII_KEYS.length; i++) {
      expect(result[PII_KEYS[i]]).toBe(REDACTED_PLACEHOLDER);
    }
  });

  it('passes non-PII keys through untouched', () => {
    const result = redactPII({
      duration: 42,
      providerId: 'sp-search',
      totalCount: 17,
      page: 1,
    });
    expect(result).toEqual({
      duration: 42,
      providerId: 'sp-search',
      totalCount: 17,
      page: 1,
    });
  });

  it('handles arrays of objects — each element redacted independently', () => {
    const result = redactPII({
      history: [
        { queryText: 'first', count: 5 },
        { queryText: 'second', count: 7 },
      ],
    }) as Record<string, unknown>;
    const history = result.history as Array<Record<string, unknown>>;
    expect(history[0].queryText).toBe(REDACTED_PLACEHOLDER);
    expect(history[0].count).toBe(5);
    expect(history[1].queryText).toBe(REDACTED_PLACEHOLDER);
    expect(history[1].count).toBe(7);
  });

  it('redacts a top-level array', () => {
    const result = redactPII([
      { queryText: 'a' },
      { queryText: 'b' },
    ]) as Array<Record<string, unknown>>;
    expect(result[0].queryText).toBe(REDACTED_PLACEHOLDER);
    expect(result[1].queryText).toBe(REDACTED_PLACEHOLDER);
  });

  it('passes primitives through', () => {
    expect(redactPII('a string')).toBe('a string');
    expect(redactPII(42)).toBe(42);
    expect(redactPII(true)).toBe(true);
    expect(redactPII(null)).toBe(null);
    expect(redactPII(undefined)).toBe(undefined);
  });

  it('does not mutate the input', () => {
    const input = { queryText: 'budget', count: 7 };
    const copy = { ...input };
    redactPII(input);
    expect(input).toEqual(copy);
  });

  it('is case-sensitive on key names (queryText redacted, querytext not)', () => {
    // PII_KEYS list uses canonical camelCase; defence-in-depth is the
    // SearchOrchestrator's choice of consistent field names, not a
    // case-folded match here.
    const result = redactPII({ querytext: 'still raw', queryText: 'budget' }) as Record<string, unknown>;
    expect(result.queryText).toBe(REDACTED_PLACEHOLDER);
    expect(result.querytext).toBe('still raw');
  });
});

describe('isLogLevelEnabled — production gate', () => {
  it('warn always logs (production gate exemption)', () => {
    expect(isLogLevelEnabled('warn', { isProduction: true, isDebug: false })).toBe(true);
    expect(isLogLevelEnabled('warn', { isProduction: false, isDebug: false })).toBe(true);
  });

  it('error always logs (production gate exemption)', () => {
    expect(isLogLevelEnabled('error', { isProduction: true, isDebug: false })).toBe(true);
    expect(isLogLevelEnabled('error', { isProduction: false, isDebug: false })).toBe(true);
  });

  it('info is silenced in production by default', () => {
    expect(isLogLevelEnabled('info', { isProduction: true, isDebug: false })).toBe(false);
  });

  it('debug is silenced in production by default', () => {
    expect(isLogLevelEnabled('debug', { isProduction: true, isDebug: false })).toBe(false);
  });

  it('info is enabled in production when ?debug=1 (or session flag) is set', () => {
    expect(isLogLevelEnabled('info', { isProduction: true, isDebug: true })).toBe(true);
  });

  it('debug is enabled in production when ?debug=1 is set', () => {
    expect(isLogLevelEnabled('debug', { isProduction: true, isDebug: true })).toBe(true);
  });

  it('info + debug always log outside production (workbench, localhost)', () => {
    expect(isLogLevelEnabled('info', { isProduction: false, isDebug: false })).toBe(true);
    expect(isLogLevelEnabled('debug', { isProduction: false, isDebug: false })).toBe(true);
  });
});
