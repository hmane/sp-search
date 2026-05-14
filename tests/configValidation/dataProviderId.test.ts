/**
 * T3.D7 — canonical dataProviderId validator.
 *
 * Each vertical row in `verticalsCollection` carries a `dataProviderId`
 * (free-text). Empty string is allowed (admin opts out — orchestrator
 * falls back to the default registered provider). Non-empty values must
 * match a registered provider id; otherwise the orchestrator's silent
 * fallback masks the misconfiguration. Reference: audit Roadmap T3.D7.
 */

import {
  validateDataProviderId,
  validateVerticalDataProviderIds,
} from '../../src/libraries/spSearchStore/configValidation/dataProviderId';

describe('validateDataProviderId', () => {
  const REGISTERED = ['sharepoint-search', 'graph-search', 'graph-people'];

  it('passes for empty string (opt-out)', () => {
    expect(validateDataProviderId('', REGISTERED)).toBe('');
  });

  it('passes for undefined', () => {
    expect(validateDataProviderId(undefined, REGISTERED)).toBe('');
  });

  it('passes for an exact-match registered id', () => {
    expect(validateDataProviderId('graph-search', REGISTERED)).toBe('');
  });

  it('passes for a case-insensitive match', () => {
    expect(validateDataProviderId('Graph-Search', REGISTERED)).toBe('');
  });

  it('flags an unknown id with Did-You-Mean when a near match exists', () => {
    const msg = validateDataProviderId('graph', REGISTERED);
    expect(msg).not.toBe('');
    expect(msg).toMatch(/Did you mean/i);
    expect(msg).toMatch(/graph-search|graph-people/);
  });

  it('flags an unknown id without Did-You-Mean when no near match', () => {
    const msg = validateDataProviderId('totally-not-a-provider', REGISTERED);
    expect(msg).not.toBe('');
    expect(msg).toMatch(/not a registered/i);
  });

  it('passes silently when registered list is empty (cold registry)', () => {
    expect(validateDataProviderId('anything', [])).toBe('');
  });
});

describe('validateVerticalDataProviderIds', () => {
  const REGISTERED = ['sharepoint-search', 'graph-search'];

  it('returns empty array when all verticals are valid', () => {
    const verticals = [
      { key: 'all', label: 'All', dataProviderId: '' },
      { key: 'docs', label: 'Documents', dataProviderId: 'sharepoint-search' },
    ];
    expect(validateVerticalDataProviderIds(verticals, REGISTERED)).toEqual([]);
  });

  it('flags each invalid row with rowIndex + verticalKey + dataProviderId', () => {
    const verticals = [
      { key: 'all', label: 'All', dataProviderId: '' },
      { key: 'people', label: 'People', dataProviderId: 'graph' },
    ];
    const issues = validateVerticalDataProviderIds(verticals, REGISTERED);
    expect(issues).toHaveLength(1);
    expect(issues[0].rowIndex).toBe(1);
    expect(issues[0].message).toMatch(/people/);
    expect(issues[0].message).toMatch(/graph/);
  });

  it('returns empty array when registered list is empty', () => {
    const verticals = [{ key: 'a', label: 'A', dataProviderId: 'bogus' }];
    expect(validateVerticalDataProviderIds(verticals, [])).toEqual([]);
  });
});
