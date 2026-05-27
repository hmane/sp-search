/**
 * T2.D3 — saved-search JSON schema validation (closes SEC-004).
 *
 * Saved searches and history entries persist their state as a JSON string on
 * the `SearchState` column. The restore paths in `SavedSearchList.tsx` and
 * `SearchHistory.tsx` previously `JSON.parse`'d the column and cast straight
 * to a TypeScript type — a parseable but malformed payload (wrong field
 * types, missing required fields, prototype-pollution attempts) would
 * silently flow into `store.setState` and corrupt the search context.
 *
 * This module exposes a single `validateSearchState(raw)` function that
 * returns a tagged-union result. Callers should:
 *   1. Call `validateSearchState(savedSearch.searchState)`.
 *   2. On `{ ok: false, errors }`, show a MessageBar and skip the apply.
 *   3. On `{ ok: true, state }`, pass `state` to `store.setState`.
 *
 * Unknown top-level keys are silently stripped — defence-in-depth against
 * future fields that haven't been schema-versioned yet.
 */

/** Operator on an active-filter row. */
type FilterOperator = 'AND' | 'OR';

/** Sort direction on the stored sort field. */
type SortDirection = 'Ascending' | 'Descending';

interface IValidatedActiveFilter {
  filterName: string;
  value: string;
  displayValue?: string;
  operator: FilterOperator;
}

interface IValidatedSort {
  property: string;
  direction: SortDirection;
}

export interface IValidatedSearchState {
  queryText?: string;
  currentVerticalKey?: string;
  activeFilters?: IValidatedActiveFilter[];
  sort?: IValidatedSort;
}

export type ISearchStateValidationResult =
  | { ok: true; state: IValidatedSearchState }
  | { ok: false; errors: string[] };

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

function validateActiveFilter(value: unknown, index: number, errors: string[]): IValidatedActiveFilter | undefined {
  if (!isPlainObject(value)) {
    errors.push('activeFilters[' + String(index) + '] must be an object');
    return undefined;
  }
  let ok = true;
  if (typeof value.filterName !== 'string' || !value.filterName) {
    errors.push('activeFilters[' + String(index) + '].filterName must be a non-empty string');
    ok = false;
  }
  if (typeof value.value !== 'string') {
    errors.push('activeFilters[' + String(index) + '].value must be a string');
    ok = false;
  }
  if (value.operator !== 'AND' && value.operator !== 'OR') {
    errors.push('activeFilters[' + String(index) + '].operator must be "AND" or "OR"');
    ok = false;
  }
  if (value.displayValue !== undefined && typeof value.displayValue !== 'string') {
    errors.push('activeFilters[' + String(index) + '].displayValue, when present, must be a string');
    ok = false;
  }
  if (!ok) {
    return undefined;
  }
  return {
    filterName: value.filterName as string,
    value: value.value as string,
    operator: value.operator as FilterOperator,
    displayValue: value.displayValue as string | undefined,
  };
}

function validateSort(value: unknown, errors: string[]): IValidatedSort | undefined {
  if (!isPlainObject(value)) {
    errors.push('sort must be an object');
    return undefined;
  }
  let ok = true;
  if (typeof value.property !== 'string' || !value.property) {
    errors.push('sort.property must be a non-empty string');
    ok = false;
  }
  if (value.direction !== 'Ascending' && value.direction !== 'Descending') {
    errors.push('sort.direction must be "Ascending" or "Descending"');
    ok = false;
  }
  if (!ok) {
    return undefined;
  }
  return {
    property: value.property as string,
    direction: value.direction as SortDirection,
  };
}

export function validateSearchState(raw: string | undefined): ISearchStateValidationResult {
  if (raw === undefined || raw === '' || raw === '{}') {
    return { ok: true, state: {} };
  }

  let parsed: unknown;
  try {
    parsed = JSON.parse(raw);
  } catch (err) {
    return { ok: false, errors: ['Could not parse searchState as JSON: ' + (err as Error).message] };
  }

  if (!isPlainObject(parsed)) {
    return { ok: false, errors: ['searchState must be a JSON object at the top level'] };
  }

  const errors: string[] = [];
  const state: IValidatedSearchState = {};

  if ('queryText' in parsed) {
    if (typeof parsed.queryText !== 'string') {
      errors.push('queryText must be a string');
    } else {
      state.queryText = parsed.queryText;
    }
  }

  if ('currentVerticalKey' in parsed) {
    if (typeof parsed.currentVerticalKey !== 'string') {
      errors.push('currentVerticalKey must be a string');
    } else {
      state.currentVerticalKey = parsed.currentVerticalKey;
    }
  }

  if ('activeFilters' in parsed) {
    const af = parsed.activeFilters;
    if (!Array.isArray(af)) {
      errors.push('activeFilters must be an array');
    } else {
      const valid: IValidatedActiveFilter[] = [];
      for (let i: number = 0; i < af.length; i++) {
        const v = validateActiveFilter(af[i], i, errors);
        if (v) {
          valid.push(v);
        }
      }
      // Only set the field if every element validated — partial-trust silently
      // dropping a filter is worse than failing the whole apply.
      if (errors.length === 0) {
        state.activeFilters = valid;
      }
    }
  }

  if ('sort' in parsed) {
    const s = validateSort(parsed.sort, errors);
    if (s) {
      state.sort = s;
    }
  }

  if (errors.length > 0) {
    return { ok: false, errors };
  }
  return { ok: true, state };
}
