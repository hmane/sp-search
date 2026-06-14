import { StateCreator } from 'zustand';
import { ISearchStore, IFilterSlice, IActiveFilter, IRefiner, IRefinerValue } from '@interfaces/index';
import { areFiltersEquivalent } from '@store/utils/filterValueMatching';
import { DebugCollector } from '../../debug';

/**
 * Merge new refiners from a search response with existing display refiners.
 * Uses the newest search response as the visible option set. The only values
 * preserved from previous responses are active selections that SharePoint
 * omitted from the narrowed response, so users can still remove them without
 * seeing every stale zero-count option.
 */
function isActiveRefinerValue(
  filterName: string,
  value: IRefinerValue,
  activeFilters: IActiveFilter[]
): boolean {
  const candidate: IActiveFilter = {
    filterName: filterName,
    value: value.value,
    displayValue: value.name || undefined,
    operator: 'OR',
  };

  for (let i = 0; i < activeFilters.length; i++) {
    if (areFiltersEquivalent(activeFilters[i], candidate)) {
      return true;
    }
  }

  return false;
}

function appendMissingActiveValues(
  filterName: string,
  values: IRefinerValue[],
  activeFilters: IActiveFilter[]
): IRefinerValue[] {
  const nextValues = values.slice();

  for (let i = 0; i < activeFilters.length; i++) {
    const active = activeFilters[i];
    if (active.filterName !== filterName) {
      continue;
    }

    let alreadyPresent = false;
    for (let j = 0; j < nextValues.length; j++) {
      if (isActiveRefinerValue(filterName, nextValues[j], [active])) {
        alreadyPresent = true;
        break;
      }
    }

    if (!alreadyPresent) {
      nextValues.push({
        name: active.displayValue || active.value,
        value: active.value,
        count: 0,
        isSelected: true,
      });
    }
  }

  return nextValues;
}

function mergeRefiners(existing: IRefiner[], incoming: IRefiner[], activeFilters: IActiveFilter[]): IRefiner[] {
  if (existing.length === 0) {
    return incoming.map(function (refiner: IRefiner): IRefiner {
      return {
        filterName: refiner.filterName,
        values: appendMissingActiveValues(refiner.filterName, refiner.values, activeFilters),
      };
    });
  }

  // Build a lookup of incoming refiners by name
  const incomingMap = new Map<string, IRefiner>();
  for (let i = 0; i < incoming.length; i++) {
    incomingMap.set(incoming[i].filterName, incoming[i]);
  }

  const merged: IRefiner[] = [];

  // Process each existing refiner — merge with incoming
  for (let i = 0; i < existing.length; i++) {
    const prev = existing[i];
    const next = incomingMap.get(prev.filterName);

    if (!next) {
      const activeValues = appendMissingActiveValues(
        prev.filterName,
        prev.values.filter(function (prevVal: IRefinerValue): boolean {
          return isActiveRefinerValue(prev.filterName, prevVal, activeFilters);
        }).map(function (prevVal: IRefinerValue): IRefinerValue {
          return {
            name: prevVal.name,
            value: prevVal.value,
            count: 0,
            isSelected: prevVal.isSelected,
          };
        }),
        activeFilters
      );
      if (activeValues.length > 0) {
        merged.push({
          filterName: prev.filterName,
          values: activeValues,
        });
      }
    } else {
      // Build value lookup from new results
      const nextValueMap = new Map<string, IRefinerValue>();
      for (let j = 0; j < next.values.length; j++) {
        nextValueMap.set(next.values[j].value, next.values[j]);
      }

      const seenValues = new Set<string>();
      const mergedValues: IRefinerValue[] = [];

      for (let j = 0; j < next.values.length; j++) {
        seenValues.add(next.values[j].value);
        mergedValues.push(next.values[j]);
      }

      for (let j = 0; j < prev.values.length; j++) {
        const prevVal = prev.values[j];
        if (!nextValueMap.has(prevVal.value) && isActiveRefinerValue(prev.filterName, prevVal, activeFilters)) {
          mergedValues.push({
            name: prevVal.name,
            value: prevVal.value,
            count: 0,
            isSelected: prevVal.isSelected,
          });
          seenValues.add(prevVal.value);
        }
      }

      merged.push({
        filterName: prev.filterName,
        values: appendMissingActiveValues(prev.filterName, mergedValues, activeFilters).filter(function (value: IRefinerValue, index: number, all: IRefinerValue[]): boolean {
          return all.findIndex(function (candidate: IRefinerValue): boolean {
            return candidate.value === value.value;
          }) === index && (seenValues.has(value.value) || isActiveRefinerValue(prev.filterName, value, activeFilters));
        }),
      });
      incomingMap.delete(prev.filterName);
    }
  }

  // Add any entirely new refiners not in the existing set
  incomingMap.forEach((refiner) => {
    merged.push({
      filterName: refiner.filterName,
      values: appendMissingActiveValues(refiner.filterName, refiner.values, activeFilters),
    });
  });

  return merged;
}

export const createFilterSlice: StateCreator<ISearchStore, [], [], IFilterSlice> = (set, get) => ({
  activeFilters: [],
  availableRefiners: [],
  displayRefiners: [],
  filterConfig: [],
  isRefining: false,
  operatorBetweenFilters: 'AND',

  setRefiner: (filter: IActiveFilter): void => {
    const current = get().activeFilters;

    // Check if this is a date range filter (uses FQL range() syntax)
    const isDateRangeFilter = filter.value.indexOf('range(') === 0 ||
      filter.value.indexOf('range(datetime') >= 0;

    if (isDateRangeFilter) {
      // Date range filters are mutually exclusive per property name
      // First check for existing filter for this property, then toggle/replace
      const existingIndex = current.findIndex(
        (f) => f.filterName === filter.filterName
      );

      if (existingIndex >= 0 && areFiltersEquivalent(current[existingIndex], filter)) {
        // Same value clicked again — toggle off
        const updated = [...current];
        updated.splice(existingIndex, 1);
        set({ activeFilters: updated, currentPage: 1 });
      } else if (existingIndex >= 0) {
        // Different value for same property — replace
        const updated = [...current];
        updated[existingIndex] = filter;
        set({ activeFilters: updated, currentPage: 1 });
      } else {
        // New date filter — add
        set({ activeFilters: [...current, filter], currentPage: 1 });
      }
    } else {
      // Standard toggle behavior for non-date filters
      const existing = current.findIndex(
        (f) => areFiltersEquivalent(f, filter)
      );
      if (existing >= 0) {
        // Already exists — remove it (toggle off)
        const updated = [...current];
        updated.splice(existing, 1);
        set({ activeFilters: updated, currentPage: 1 });
      } else {
        set({ activeFilters: [...current, filter], currentPage: 1 });
      }
    }
    DebugCollector.logEvent('FILTER', { action: 'set', filterName: filter.filterName, value: filter.value });
  },

  removeRefiner: (filterKey: string, value?: string): void => {
    const current = get().activeFilters;
    const updated = value
      ? current.filter((f) => !(f.filterName === filterKey && f.value === value))
      : current.filter((f) => f.filterName !== filterKey);
    set({ activeFilters: updated, currentPage: 1 });
    DebugCollector.logEvent('FILTER', { action: 'remove', filterName: filterKey, value: value || '*' });
  },

  clearAllFilters: (): void => {
    set({ activeFilters: [], currentPage: 1, displayRefiners: [] });
    DebugCollector.logEvent('FILTER', { action: 'clearAll' });
  },

  setOperatorBetweenFilters: (operator: 'AND' | 'OR'): void => {
    set({ operatorBetweenFilters: operator });
  },

  setAvailableRefiners: (refiners: IRefiner[]): void => {
    const prev = get().displayRefiners;
    const activeFilters = get().activeFilters;
    // If no active filters, show raw refiners (no merge needed — fresh results)
    if (activeFilters.length === 0) {
      set({ availableRefiners: refiners, displayRefiners: refiners });
    } else {
      // Merge with previous only to preserve active selections omitted by
      // SharePoint's narrowed refiner response.
      const merged = mergeRefiners(prev, refiners, activeFilters);
      set({ availableRefiners: refiners, displayRefiners: merged });
    }
  },
});
