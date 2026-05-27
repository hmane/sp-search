/**
 * T5.D4 — extensible DebugPanel tab registry tests. Verifies the
 * cross-bundle registration contract that T3.D8 (Multi-Context audit
 * panel) consumes.
 */

import * as React from 'react';
import {
  registerDebugTab,
  getRegisteredDebugTabs,
  unregisterDebugTab,
  _resetDebugTabRegistryForTesting,
} from '../../src/libraries/spSearchStore/debug/debugTabRegistry';

const renderEmpty = (): React.ReactElement => React.createElement('div');

beforeEach((): void => {
  _resetDebugTabRegistryForTesting();
});

describe('registerDebugTab', () => {
  it('exposes a tab via getRegisteredDebugTabs', () => {
    registerDebugTab('multi-context', 'Multi-Context', renderEmpty);
    const all = getRegisteredDebugTabs();
    expect(all).toHaveLength(1);
    expect(all[0].id).toBe('multi-context');
    expect(all[0].label).toBe('Multi-Context');
    expect(typeof all[0].render).toBe('function');
  });

  it('overwrites a previous registration for the same id', () => {
    registerDebugTab('foo', 'Foo v1', renderEmpty);
    registerDebugTab('foo', 'Foo v2', renderEmpty);
    const all = getRegisteredDebugTabs();
    expect(all).toHaveLength(1);
    expect(all[0].label).toBe('Foo v2');
  });

  it('preserves registration order when no sortOrder specified', () => {
    registerDebugTab('first', 'First', renderEmpty);
    registerDebugTab('second', 'Second', renderEmpty);
    registerDebugTab('third', 'Third', renderEmpty);
    const ids = getRegisteredDebugTabs().map((t) => t.id);
    expect(ids).toEqual(['first', 'second', 'third']);
  });

  it('respects explicit sortOrder over registration order', () => {
    registerDebugTab('z', 'Z', renderEmpty, { sortOrder: 100 });
    registerDebugTab('a', 'A', renderEmpty, { sortOrder: 1 });
    registerDebugTab('m', 'M', renderEmpty, { sortOrder: 50 });
    const ids = getRegisteredDebugTabs().map((t) => t.id);
    expect(ids).toEqual(['a', 'm', 'z']);
  });
});

describe('unregisterDebugTab', () => {
  it('removes a tab from the registry', () => {
    registerDebugTab('temp', 'Temp', renderEmpty);
    unregisterDebugTab('temp');
    expect(getRegisteredDebugTabs()).toHaveLength(0);
  });

  it('is a no-op for an unknown id', () => {
    expect(() => unregisterDebugTab('never')).not.toThrow();
    expect(getRegisteredDebugTabs()).toHaveLength(0);
  });
});

describe('getRegisteredDebugTabs', () => {
  it('returns an empty array on a fresh registry', () => {
    expect(getRegisteredDebugTabs()).toEqual([]);
  });
});
