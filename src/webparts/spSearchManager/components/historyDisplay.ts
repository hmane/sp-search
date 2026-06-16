import type { IFilterConfig, ISearchHistoryEntry } from '@interfaces/index';
import { validateSearchState, type IValidatedSearchState } from '@store/utils/searchStateSchema';
import { formatRefinerValueForDisplay } from '@store/utils/refinerDisplay';

export interface ISearchHistoryDisplay {
  title: string;
  metaParts: string[];
}

/**
 * Build a lookup from a refiner's managed-property name (lowercased) to its
 * admin-configured display name (alias), so history shows "Status: Yes" rather
 * than "RefinableString06: Yes".
 */
export function buildFilterAliasMap(filterConfig: IFilterConfig[] | undefined): Record<string, string> {
  const map: Record<string, string> = {};
  if (!filterConfig) {
    return map;
  }
  for (const cfg of filterConfig) {
    const key = (cfg.managedProperty || '').trim().toLowerCase();
    const alias = (cfg.displayName || '').trim();
    if (key && alias) {
      map[key] = alias;
    }
  }
  return map;
}

function isDefaultVertical(value: string | undefined): boolean {
  const normalized = (value || '').trim().toLowerCase();
  return normalized === '' || normalized === 'all';
}

function formatFilterLabel(filterName: string, value: string, filterAliases?: Record<string, string>): string {
  // Decode FQL/hex/taxonomy tokens so history shows the label, not the raw id.
  const cleanedValue = formatRefinerValueForDisplay(value);
  // Prefer the admin alias for the refiner; fall back to the managed-property name.
  const label = (filterAliases && filterAliases[(filterName || '').toLowerCase()]) || filterName;
  if (!label) {
    return cleanedValue;
  }
  if (label.toLowerCase() === cleanedValue.toLowerCase()) {
    return cleanedValue;
  }
  return label + ': ' + cleanedValue;
}

function getValidatedState(entry: ISearchHistoryEntry): IValidatedSearchState | undefined {
  const validation = validateSearchState(entry.searchState);
  return validation.ok ? validation.state : undefined;
}

function trimDefaultVerticalPrefix(title: string): string {
  return title.replace(/^all\s*[•-]\s*/i, '').trim();
}

export function getHistoryDisplay(
  entry: ISearchHistoryEntry,
  filterAliases?: Record<string, string>
): ISearchHistoryDisplay {
  const state = getValidatedState(entry);
  const queryText = (state?.queryText || '').trim();
  const vertical = (state?.currentVerticalKey || entry.vertical || '').trim();
  const filters = state?.activeFilters || [];
  const filterLabels: string[] = [];

  for (let i = 0; i < filters.length && filterLabels.length < 3; i++) {
    const filter = filters[i];
    const value = filter.displayValue || filter.value;
    if (value) {
      filterLabels.push(formatFilterLabel(filter.filterName, value, filterAliases));
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
