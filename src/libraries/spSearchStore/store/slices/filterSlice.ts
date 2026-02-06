import { StateCreator } from 'zustand';
import { ISearchStore, IFilterSlice, IActiveFilter, IRefiner } from '@interfaces/index';

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
        set({ activeFilters: updated });
      } else if (existingIndex >= 0) {
        // Different value for same property — replace
        const updated = [...current];
        updated[existingIndex] = filter;
        set({ activeFilters: updated });
      } else {
        // New date filter — add
        set({ activeFilters: [...current, filter] });
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
        set({ activeFilters: updated });
      } else {
        set({ activeFilters: [...current, filter] });
      }
    }
  },

  removeRefiner: (filterKey: string, value?: string): void => {
    const current = get().activeFilters;
    const updated = value
      ? current.filter((f) => !(f.filterName === filterKey && f.value === value))
      : current.filter((f) => f.filterName !== filterKey);
    set({ activeFilters: updated });
  },

  clearAllFilters: (): void => {
    set({ activeFilters: [] });
  },

  setAvailableRefiners: (refiners: IRefiner[]): void => {
    set({ availableRefiners: refiners, displayRefiners: refiners });
  },
});
