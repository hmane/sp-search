import {
  setWebPartContextId,
  unregisterWebPartContextId,
  getRegisteredContextIds,
  subscribeContextIdChanges,
  getPeerContextIds,
  _resetContextIdRegistryForTesting,
} from '../../src/libraries/spSearchStore/utils/contextIdRegistry';

/**
 * T3.D2 — context-id registry tests. The registry is the shared source of
 * truth the mismatch banner reads to detect when two web parts on the same
 * page point at different `searchContextId` values.
 */

describe('contextIdRegistry — T3.D2', () => {
  beforeEach(() => {
    _resetContextIdRegistryForTesting();
  });

  it('starts empty', () => {
    expect(getRegisteredContextIds().size).toBe(0);
  });

  it('registers a web part with its context id', () => {
    setWebPartContextId('wp-1', 'default');
    expect(getRegisteredContextIds().get('wp-1')).toBe('default');
  });

  it('updates the context id on a re-register', () => {
    setWebPartContextId('wp-1', 'default');
    setWebPartContextId('wp-1', 'hr-search');
    expect(getRegisteredContextIds().get('wp-1')).toBe('hr-search');
  });

  it('unregister removes the web part', () => {
    setWebPartContextId('wp-1', 'default');
    unregisterWebPartContextId('wp-1');
    expect(getRegisteredContextIds().has('wp-1')).toBe(false);
  });

  it('getPeerContextIds returns peers whose id differs from this one', () => {
    setWebPartContextId('wp-box', 'hr-search');
    setWebPartContextId('wp-results', 'policy-search'); // mismatched
    setWebPartContextId('wp-filters', 'hr-search'); // matches box
    expect(getPeerContextIds('wp-box', 'hr-search')).toEqual(['policy-search']);
  });

  it('getPeerContextIds returns empty when all peers match this id', () => {
    setWebPartContextId('wp-box', 'hr-search');
    setWebPartContextId('wp-results', 'hr-search');
    expect(getPeerContextIds('wp-box', 'hr-search')).toEqual([]);
  });

  it('getPeerContextIds returns empty when there are no peers', () => {
    setWebPartContextId('wp-box', 'hr-search');
    expect(getPeerContextIds('wp-box', 'hr-search')).toEqual([]);
  });

  it('getPeerContextIds deduplicates peer ids (two peers, same value)', () => {
    setWebPartContextId('wp-box', 'hr-search');
    setWebPartContextId('wp-results', 'policy-search');
    setWebPartContextId('wp-filters', 'policy-search');
    const peers = getPeerContextIds('wp-box', 'hr-search');
    expect(peers).toHaveLength(1);
    expect(peers).toEqual(['policy-search']);
  });

  it('subscribeContextIdChanges fires on set + unregister; the returned function unsubscribes', () => {
    let calls = 0;
    const unsubscribe = subscribeContextIdChanges(() => { calls++; });
    setWebPartContextId('wp-1', 'a');
    setWebPartContextId('wp-1', 'b');
    unregisterWebPartContextId('wp-1');
    expect(calls).toBe(3);
    unsubscribe();
    setWebPartContextId('wp-2', 'c');
    expect(calls).toBe(3);
  });

  it('subscribeContextIdChanges does NOT fire when set is a no-op (same value)', () => {
    setWebPartContextId('wp-1', 'a');
    let calls = 0;
    subscribeContextIdChanges(() => { calls++; });
    setWebPartContextId('wp-1', 'a'); // no change
    expect(calls).toBe(0);
  });
});
