/**
 * T3.D10 — initialization-order diagnostic tests.
 */

import {
  recordWebPartInit,
  recordFirstSearch,
  getInitOrderDiagnostic,
  hasInitOrderIssue,
  clearInitOrderDiagnostic,
  _resetInitOrderDiagnosticForTesting,
} from '../../src/libraries/spSearchStore/utils/initOrderDiagnostic';

beforeEach((): void => {
  _resetInitOrderDiagnosticForTesting();
});

describe('initOrderDiagnostic', () => {
  it('returns undefined for an unknown context', () => {
    expect(getInitOrderDiagnostic('never')).toBeUndefined();
    expect(hasInitOrderIssue('never')).toBe(false);
  });

  it('records web parts that registered against a context', () => {
    recordWebPartInit('ctx-a', 'SpSearchResultsWebPart');
    recordWebPartInit('ctx-a', 'SpSearchFiltersWebPart');
    const d = getInitOrderDiagnostic('ctx-a')!;
    expect(d.registeredWebParts.has('SpSearchResultsWebPart')).toBe(true);
    expect(d.registeredWebParts.has('SpSearchFiltersWebPart')).toBe(true);
  });

  it('marks firstSearchFired + filterConfigAtFirstSearch on recordFirstSearch', () => {
    recordWebPartInit('ctx-a', 'SpSearchResultsWebPart');
    recordFirstSearch('ctx-a', 0);
    const d = getInitOrderDiagnostic('ctx-a')!;
    expect(d.firstSearchFired).toBe(true);
    expect(d.filterConfigAtFirstSearch).toBe(0);
  });

  it('subsequent recordFirstSearch calls do not overwrite the original measurement', () => {
    recordFirstSearch('ctx-a', 0);
    recordFirstSearch('ctx-a', 5); // simulated second search
    expect(getInitOrderDiagnostic('ctx-a')!.filterConfigAtFirstSearch).toBe(0);
  });

  it('hasInitOrderIssue is true only when Filters registered AFTER first-search ran AND filterConfig was empty', () => {
    recordWebPartInit('ctx-a', 'SpSearchResultsWebPart');
    recordFirstSearch('ctx-a', 0);
    recordWebPartInit('ctx-a', 'SpSearchFiltersWebPart');
    expect(hasInitOrderIssue('ctx-a')).toBe(true);
  });

  it('hasInitOrderIssue is false when Filters registered BEFORE first search', () => {
    recordWebPartInit('ctx-a', 'SpSearchResultsWebPart');
    recordWebPartInit('ctx-a', 'SpSearchFiltersWebPart');
    recordFirstSearch('ctx-a', 3);
    expect(hasInitOrderIssue('ctx-a')).toBe(false);
  });

  it('hasInitOrderIssue is false when Filters registered late but filterConfig was non-empty', () => {
    recordWebPartInit('ctx-a', 'SpSearchResultsWebPart');
    recordFirstSearch('ctx-a', 4); // filterConfig already non-empty (e.g. seeded from URL)
    recordWebPartInit('ctx-a', 'SpSearchFiltersWebPart');
    expect(hasInitOrderIssue('ctx-a')).toBe(false);
  });

  it('keeps diagnostics per-context isolated', () => {
    recordWebPartInit('ctx-a', 'SpSearchResultsWebPart');
    recordFirstSearch('ctx-a', 0);
    recordWebPartInit('ctx-a', 'SpSearchFiltersWebPart');
    recordWebPartInit('ctx-b', 'SpSearchResultsWebPart');
    expect(hasInitOrderIssue('ctx-a')).toBe(true);
    expect(hasInitOrderIssue('ctx-b')).toBe(false);
  });

  it('clearInitOrderDiagnostic resets the state for a context', () => {
    recordWebPartInit('ctx-a', 'SpSearchResultsWebPart');
    recordFirstSearch('ctx-a', 0);
    recordWebPartInit('ctx-a', 'SpSearchFiltersWebPart');
    expect(hasInitOrderIssue('ctx-a')).toBe(true);
    clearInitOrderDiagnostic('ctx-a');
    expect(hasInitOrderIssue('ctx-a')).toBe(false);
  });
});
