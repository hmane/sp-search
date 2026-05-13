import {
  validateSearchState,
  ISearchStateValidationResult,
  IValidatedSearchState,
} from '../../src/libraries/spSearchStore/utils/searchStateSchema';

/**
 * T2.D3 — saved-search JSON schema validation on restore (closes SEC-004).
 *
 * The validator parses a raw `searchState` string into an
 * `IValidatedSearchState` (the shape SavedSearchList / SearchHistory feed
 * into `store.setState`). Malformed input — including parse failures, wrong
 * top-level types, and per-field type mismatches — short-circuits with a
 * tagged-union error result so callers can show a MessageBar and skip the
 * apply rather than corrupting the store.
 */

function expectValid(result: ISearchStateValidationResult): IValidatedSearchState {
  if (!result.ok) {
    // Cast: ts-jest's non-strict mode doesn't narrow the discriminated union on `!result.ok`.
    const errors = (result as { ok: false; errors: string[] }).errors;
    throw new Error('Expected valid result, got errors: ' + errors.join('; '));
  }
  return (result as { ok: true; state: IValidatedSearchState }).state;
}

function expectInvalid(result: ISearchStateValidationResult): string[] {
  if (result.ok) {
    const state = (result as { ok: true; state: IValidatedSearchState }).state;
    throw new Error('Expected invalid result, got: ' + JSON.stringify(state));
  }
  return (result as { ok: false; errors: string[] }).errors;
}

describe('validateSearchState — happy paths', () => {
  it('accepts an empty object', () => {
    const result = validateSearchState('{}');
    const state = expectValid(result);
    expect(state).toEqual({});
  });

  it('accepts a full valid state', () => {
    const raw = JSON.stringify({
      queryText: 'budget',
      currentVerticalKey: 'documents',
      activeFilters: [
        { filterName: 'FileType', value: 'docx', displayValue: 'Word', operator: 'OR' },
      ],
      sort: { property: 'LastModifiedTime', direction: 'Descending' },
    });
    const state = expectValid(validateSearchState(raw));
    expect(state.queryText).toBe('budget');
    expect(state.currentVerticalKey).toBe('documents');
    expect(state.activeFilters).toHaveLength(1);
    expect(state.sort?.direction).toBe('Descending');
  });

  it('accepts queryText alone', () => {
    const state = expectValid(validateSearchState(JSON.stringify({ queryText: 'hello' })));
    expect(state.queryText).toBe('hello');
    expect(state.activeFilters).toBeUndefined();
  });

  it('accepts empty activeFilters array', () => {
    const state = expectValid(validateSearchState(JSON.stringify({ activeFilters: [] })));
    expect(state.activeFilters).toEqual([]);
  });

  it('treats empty string as empty object', () => {
    const state = expectValid(validateSearchState(''));
    expect(state).toEqual({});
  });

  it('treats undefined as empty object', () => {
    const state = expectValid(validateSearchState(undefined));
    expect(state).toEqual({});
  });
});

describe('validateSearchState — malformation cases (6 required for T2.D3)', () => {
  it('1. rejects non-JSON input', () => {
    const errors = expectInvalid(validateSearchState('not-json-at-all'));
    expect(errors.join(' ')).toMatch(/parse|invalid json/i);
  });

  it('2. rejects JSON null (no top-level object)', () => {
    expectInvalid(validateSearchState('null'));
  });

  it('3. rejects non-object top-level value', () => {
    const errors = expectInvalid(validateSearchState('42'));
    expect(errors.join(' ')).toMatch(/object/i);
  });

  it('4. rejects array top-level value', () => {
    expectInvalid(validateSearchState('[]'));
  });

  it('5. rejects non-string queryText', () => {
    const errors = expectInvalid(validateSearchState(JSON.stringify({ queryText: 123 })));
    expect(errors.join(' ')).toMatch(/queryText/);
  });

  it('6. rejects non-array activeFilters', () => {
    const errors = expectInvalid(validateSearchState(JSON.stringify({ activeFilters: 'not-an-array' })));
    expect(errors.join(' ')).toMatch(/activeFilters/);
  });

  it('7. rejects activeFilters element missing filterName', () => {
    const errors = expectInvalid(validateSearchState(JSON.stringify({
      activeFilters: [{ value: 'docx', operator: 'OR' }],
    })));
    expect(errors.join(' ')).toMatch(/filterName/);
  });

  it('8. rejects activeFilters element with bad operator', () => {
    const errors = expectInvalid(validateSearchState(JSON.stringify({
      activeFilters: [{ filterName: 'F', value: 'v', operator: 'XOR' }],
    })));
    expect(errors.join(' ')).toMatch(/operator/);
  });

  it('9. rejects sort.direction not Ascending|Descending', () => {
    const errors = expectInvalid(validateSearchState(JSON.stringify({
      sort: { property: 'X', direction: 'sideways' },
    })));
    expect(errors.join(' ')).toMatch(/direction/i);
  });

  it('10. rejects sort missing property', () => {
    expectInvalid(validateSearchState(JSON.stringify({
      sort: { direction: 'Ascending' },
    })));
  });

  it('11. rejects non-string currentVerticalKey', () => {
    expectInvalid(validateSearchState(JSON.stringify({ currentVerticalKey: 42 })));
  });
});

describe('validateSearchState — defence-in-depth', () => {
  it('accumulates multiple errors, not just the first', () => {
    const errors = expectInvalid(validateSearchState(JSON.stringify({
      queryText: 123,
      activeFilters: 'wrong',
      sort: { property: 'X', direction: 'sideways' },
    })));
    expect(errors.length).toBeGreaterThanOrEqual(3);
  });

  it('strips unknown top-level keys (does not propagate untrusted fields)', () => {
    const state = expectValid(validateSearchState(JSON.stringify({
      queryText: 'safe',
      __proto__: { polluted: true },
      isAdmin: true,
    })));
    expect(state).toEqual({ queryText: 'safe' });
    expect((state as Record<string, unknown>).isAdmin).toBeUndefined();
  });
});
