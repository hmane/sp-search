import type { ISearchHistoryEntry } from '@interfaces/index';
import { validateSearchState, type IValidatedSearchState } from '@store/utils/searchStateSchema';

export interface ISearchHistoryDisplay {
  title: string;
  metaParts: string[];
}

function isDefaultVertical(value: string | undefined): boolean {
  const normalized = (value || '').trim().toLowerCase();
  return normalized === '' || normalized === 'all';
}

function cleanFilterValue(value: string): string {
  return value.replace(/^"|"$/g, '').trim();
}

function formatFilterLabel(filterName: string, value: string): string {
  const cleanedValue = cleanFilterValue(value);
  if (!filterName) {
    return cleanedValue;
  }
  if (filterName.toLowerCase() === cleanedValue.toLowerCase()) {
    return cleanedValue;
  }
  return filterName + ': ' + cleanedValue;
}

function getValidatedState(entry: ISearchHistoryEntry): IValidatedSearchState | undefined {
  const validation = validateSearchState(entry.searchState);
  return validation.ok ? validation.state : undefined;
}

function trimDefaultVerticalPrefix(title: string): string {
  return title.replace(/^all\s*[•-]\s*/i, '').trim();
}

export function getHistoryDisplay(entry: ISearchHistoryEntry): ISearchHistoryDisplay {
  const state = getValidatedState(entry);
  const queryText = (state?.queryText || '').trim();
  const vertical = (state?.currentVerticalKey || entry.vertical || '').trim();
  const filters = state?.activeFilters || [];
  const filterLabels: string[] = [];

  for (let i = 0; i < filters.length && filterLabels.length < 3; i++) {
    const filter = filters[i];
    const value = filter.displayValue || filter.value;
    if (value) {
      filterLabels.push(formatFilterLabel(filter.filterName, value));
    }
  }

  let title = '';
  if (queryText && filterLabels.length > 0) {
    title = queryText + ' • ' + filterLabels.join(' • ');
  } else if (queryText) {
    title = queryText;
  } else if (filterLabels.length > 0 && !isDefaultVertical(vertical)) {
    title = vertical + ' • ' + filterLabels.join(' • ');
  } else if (filterLabels.length > 0) {
    title = filterLabels.join(' • ');
  } else if (!isDefaultVertical(vertical)) {
    title = 'Browse ' + vertical;
  } else {
    title = trimDefaultVerticalPrefix(entry.queryText) || 'Browse all results';
  }

  const metaParts: string[] = [];
  if (!isDefaultVertical(vertical)) {
    metaParts.push('Scope: ' + vertical);
  }
  if (entry.useCount > 1) {
    metaParts.push('Used ' + String(entry.useCount) + ' times');
  }
  metaParts.push(String(entry.resultCount) + ' results');

  return {
    title,
    metaParts,
  };
}
