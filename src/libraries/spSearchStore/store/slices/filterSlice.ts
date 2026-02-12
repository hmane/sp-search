import { StateCreator } from 'zustand';
import { ISearchStore, IFilterSlice, IActiveFilter, IRefiner, IRefinerValue } from '@interfaces/index';

/**
 * Merge new refiners from a search response with existing display refiners.
 * Keeps all previously-seen values visible so filter options don't disappear
 * when a filter narrows results. Values not in the new response keep their
 * previous counts but are preserved for user selection/deselection.
 */
function mergeRefiners(existing: IRefiner[], incoming: IRefiner[]): IRefiner[] {
  if (existing.length === 0) {
    return incoming;
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
      // Refiner no longer returned — keep previous values with 0 counts
      const zeroed: IRefinerValue[] = prev.values.map((v) => ({
        name: v.name,
        value: v.value,
        count: 0,
        isSelected: v.isSelected,
      }));
      merged.push({ filterName: prev.filterName, values: zeroed });
    } else {
      // Build value lookup from new results
      const nextValueMap = new Map<string, IRefinerValue>();
      for (let j = 0; j < next.values.length; j++) {
        nextValueMap.set(next.values[j].value, next.values[j]);
      }

      // Start with all previous values — update counts from new results
      const seenValues = new Set<string>();
      const mergedValues: IRefinerValue[] = [];

      for (let j = 0; j < prev.values.length; j++) {
        const prevVal = prev.values[j];
        const nextVal = nextValueMap.get(prevVal.value);
        seenValues.add(prevVal.value);

        if (nextVal) {
          // Value exists in new results — use new count
          mergedValues.push(nextVal);
        } else {
          // Value no longer in results — keep with 0 count
          mergedValues.push({
            name: prevVal.name,
            value: prevVal.value,
            count: 0,
            isSelected: prevVal.isSelected,
          });
        }
      }

      // Add any NEW values from the incoming results that weren't in previous
      for (let j = 0; j < next.values.length; j++) {
        if (!seenValues.has(next.values[j].value)) {
          mergedValues.push(next.values[j]);
        }
      }

      merged.push({ filterName: prev.filterName, values: mergedValues });
      incomingMap.delete(prev.filterName);
    }
  }

  // Add any entirely new refiners not in the existing set
  incomingMap.forEach((refiner) => {
    merged.push(refiner);
  });

  return merged;
}

export const createFilterSlice: StateCreator<ISearchStore, [], [], IFilterSlice> = (set, get) => ({
  activeFilters: [],
  availableRefiners: [],
  displayRefiners: [],
  filterConfig: [],
  isRefining: false,

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

      if (existingIndex >= 0 && current[existingIndex].value === filter.value) {
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
        (f) => f.filterName === filter.filterName && f.value === filter.value
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
  },

  removeRefiner: (filterKey: string, value?: string): void => {
    const current = get().activeFilters;
    const updated = value
      ? current.filter((f) => !(f.filterName === filterKey && f.value === value))
      : current.filter((f) => f.filterName !== filterKey);
    set({ activeFilters: updated, currentPage: 1 });
  },

  clearAllFilters: (): void => {
    set({ activeFilters: [], currentPage: 1, displayRefiners: [] });
  },

  setAvailableRefiners: (refiners: IRefiner[]): void => {
    const prev = get().displayRefiners;
    const activeFilters = get().activeFilters;
    // If no active filters, show raw refiners (no merge needed — fresh results)
    if (activeFilters.length === 0) {
      set({ availableRefiners: refiners, displayRefiners: refiners });
    } else {
      // Merge with previous to preserve all known values
      const merged = mergeRefiners(prev, refiners);
      set({ availableRefiners: refiners, displayRefiners: merged });
    }
  },
});
