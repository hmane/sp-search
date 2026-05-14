/**
 * T4.D12 — preset-suggestion registry tests.
 *
 * The registry is window-backed (like contextIdRegistry) so each web part
 * bundle's separate webpack entry can read/write the same Map. Verifies
 * the basic record/consume/clear/subscribe contract.
 */

import {
  recordPresetSuggestion,
  consumePresetSuggestion,
  clearPresetSuggestion,
  subscribePresetSuggestionChanges,
  _resetPresetSuggestionRegistryForTesting,
  type IPresetSuggestion,
} from '../../src/libraries/spSearchStore/utils/presetSuggestionRegistry';

const SAMPLE: IPresetSuggestion = {
  id: 'documents',
  label: 'Documents',
  filterSuggestions: [
    { managedProperty: 'FileType', label: 'File type', urlAlias: 'ft', filterType: 'checkbox' },
  ],
  recordedAt: 0,
};

beforeEach(() => {
  _resetPresetSuggestionRegistryForTesting();
});

describe('recordPresetSuggestion', () => {
  it('stores a suggestion under the given context ID', () => {
    recordPresetSuggestion('ctx-a', SAMPLE);
    const got = consumePresetSuggestion('ctx-a');
    expect(got).toBeDefined();
    expect(got!.id).toBe('documents');
    expect(got!.filterSuggestions).toHaveLength(1);
  });

  it('overwrites a previous suggestion for the same context ID', () => {
    recordPresetSuggestion('ctx-a', SAMPLE);
    recordPresetSuggestion('ctx-a', { ...SAMPLE, id: 'people', label: 'People' });
    expect(consumePresetSuggestion('ctx-a')!.id).toBe('people');
  });

  it('keeps suggestions per-context isolated', () => {
    recordPresetSuggestion('ctx-a', { ...SAMPLE, id: 'documents' });
    recordPresetSuggestion('ctx-b', { ...SAMPLE, id: 'people' });
    expect(consumePresetSuggestion('ctx-a')!.id).toBe('documents');
    expect(consumePresetSuggestion('ctx-b')!.id).toBe('people');
  });

  it('stamps recordedAt on write', () => {
    recordPresetSuggestion('ctx-a', { ...SAMPLE, recordedAt: 0 });
    const got = consumePresetSuggestion('ctx-a')!;
    expect(got.recordedAt).toBeGreaterThan(0);
  });
});

describe('consumePresetSuggestion', () => {
  it('returns undefined for an unknown context', () => {
    expect(consumePresetSuggestion('missing-ctx')).toBeUndefined();
  });

  it('does NOT remove the suggestion from the registry (peek semantics)', () => {
    recordPresetSuggestion('ctx-a', SAMPLE);
    consumePresetSuggestion('ctx-a');
    expect(consumePresetSuggestion('ctx-a')).toBeDefined();
  });
});

describe('clearPresetSuggestion', () => {
  it('removes the suggestion for a context', () => {
    recordPresetSuggestion('ctx-a', SAMPLE);
    clearPresetSuggestion('ctx-a');
    expect(consumePresetSuggestion('ctx-a')).toBeUndefined();
  });

  it('no-ops for a context with no suggestion', () => {
    expect(() => clearPresetSuggestion('no-such-ctx')).not.toThrow();
  });
});

describe('subscribePresetSuggestionChanges', () => {
  it('fires listeners on record', () => {
    const listener = jest.fn();
    const unsub = subscribePresetSuggestionChanges(listener);
    recordPresetSuggestion('ctx-a', SAMPLE);
    expect(listener).toHaveBeenCalled();
    unsub();
  });

  it('fires listeners on clear', () => {
    recordPresetSuggestion('ctx-a', SAMPLE);
    const listener = jest.fn();
    const unsub = subscribePresetSuggestionChanges(listener);
    clearPresetSuggestion('ctx-a');
    expect(listener).toHaveBeenCalled();
    unsub();
  });

  it('stops firing after unsubscribe', () => {
    const listener = jest.fn();
    const unsub = subscribePresetSuggestionChanges(listener);
    unsub();
    recordPresetSuggestion('ctx-a', SAMPLE);
    expect(listener).not.toHaveBeenCalled();
  });
});
