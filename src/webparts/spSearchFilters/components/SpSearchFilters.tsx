import * as React from 'react';
import styles from './SpSearchFilters.module.scss';
import type { ISpSearchFiltersProps } from './ISpSearchFiltersProps';
import FilterGroup from './FilterGroup';
import FilterPillBar from './FilterPillBar';
import type {
  IRefiner,
  IActiveFilter,
  IFilterConfig,
  ISearchStore
} from '@interfaces/index';

/**
 * Finds the IFilterConfig that matches a given managed property name.
 */
function findFilterConfig(filterName: string, configs: IFilterConfig[]): IFilterConfig | undefined {
  for (let i: number = 0; i < configs.length; i++) {
    if (configs[i].managedProperty === filterName) {
      return configs[i];
    }
  }
  return undefined;
}

/**
 * Custom hook to subscribe to Zustand store state outside of React context.
 * Uses the vanilla store API with subscribe + getState.
 */
function useStoreState<T>(
  store: ISpSearchFiltersProps['store'],
  selector: (state: ISearchStore) => T
): T | undefined {
  const [state, setState] = React.useState<T | undefined>(function (): T | undefined {
    return store ? selector(store.getState()) : undefined;
  });

  React.useEffect(function (): (() => void) | undefined {
    if (!store) {
      return undefined;
    }
    // Set initial state
    setState(selector(store.getState()));

    // Subscribe to changes
    const unsubscribe: () => void = store.subscribe(function (newState: ISearchStore): void {
      setState(selector(newState));
    });

    return unsubscribe;
  }, [store, selector]);

  return state;
}

const SpSearchFilters: React.FC<ISpSearchFiltersProps> = (props: ISpSearchFiltersProps): React.ReactElement => {
  const { store, applyMode, showClearAll } = props;

  // Stable selectors to avoid re-subscriptions
  const selectRefiners = React.useCallback(function (s: ISearchStore): IRefiner[] {
    return s.availableRefiners;
  }, []);

  const selectActiveFilters = React.useCallback(function (s: ISearchStore): IActiveFilter[] {
    return s.activeFilters;
  }, []);

  const selectFilterConfig = React.useCallback(function (s: ISearchStore): IFilterConfig[] {
    return s.filterConfig;
  }, []);

  const availableRefiners: IRefiner[] | undefined = useStoreState(store, selectRefiners);
  const activeFilters: IActiveFilter[] | undefined = useStoreState(store, selectActiveFilters);
  const filterConfig: IFilterConfig[] | undefined = useStoreState(store, selectFilterConfig);

  // Use safe defaults
  const refiners: IRefiner[] = availableRefiners || [];
  const filters: IActiveFilter[] = activeFilters || [];
  const configs: IFilterConfig[] = filterConfig || [];

  // Pending filters for manual mode
  const [pendingFilters, setPendingFilters] = React.useState<IActiveFilter[]>([]);
  const [hasPendingChanges, setHasPendingChanges] = React.useState<boolean>(false);

  // Determine which filters to display: pending (manual mode) or live (instant mode)
  const displayFilters: IActiveFilter[] = applyMode === 'manual' && hasPendingChanges
    ? pendingFilters
    : filters;

  // Sync pending filters when store filters change (manual mode)
  React.useEffect(function (): void {
    if (applyMode === 'manual' && !hasPendingChanges) {
      setPendingFilters(filters);
    }
  }, [filters, applyMode, hasPendingChanges]);

  /** Handle toggling a refiner checkbox. */
  function handleToggleRefiner(filter: IActiveFilter): void {
    if (!store) {
      return;
    }

    if (applyMode === 'instant') {
      store.getState().setRefiner(filter);
    } else {
      // Manual mode: update pending filters locally
      const current: IActiveFilter[] = hasPendingChanges ? pendingFilters : filters;
      let existingIndex: number = -1;
      for (let i: number = 0; i < current.length; i++) {
        if (current[i].filterName === filter.filterName && current[i].value === filter.value) {
          existingIndex = i;
          break;
        }
      }

      let updated: IActiveFilter[];
      if (existingIndex >= 0) {
        updated = current.slice();
        updated.splice(existingIndex, 1);
      } else {
        updated = current.concat([filter]);
      }
      setPendingFilters(updated);
      setHasPendingChanges(true);
    }
  }

  /** Handle removing a specific filter from the pill bar. */
  function handleRemoveFilter(filterName: string, value: string): void {
    if (!store) {
      return;
    }

    if (applyMode === 'instant') {
      store.getState().removeRefiner(filterName, value);
    } else {
      const current: IActiveFilter[] = hasPendingChanges ? pendingFilters : filters;
      const updated: IActiveFilter[] = [];
      for (let i: number = 0; i < current.length; i++) {
        if (!(current[i].filterName === filterName && current[i].value === value)) {
          updated.push(current[i]);
        }
      }
      setPendingFilters(updated);
      setHasPendingChanges(true);
    }
  }

  /** Handle clearing all filters. */
  function handleClearAll(): void {
    if (!store) {
      return;
    }

    if (applyMode === 'instant') {
      store.getState().clearAllFilters();
    } else {
      setPendingFilters([]);
      setHasPendingChanges(true);
    }
  }

  /** Apply pending filters in manual mode. */
  function handleApply(): void {
    if (!store) {
      return;
    }

    const storeState: ISearchStore = store.getState();

    // Clear current filters first
    storeState.clearAllFilters();

    // Apply each pending filter
    for (let i: number = 0; i < pendingFilters.length; i++) {
      storeState.setRefiner(pendingFilters[i]);
    }

    setHasPendingChanges(false);
  }

  if (!store) {
    return (
      <div className={styles.spSearchFilters}>
        <div className={styles.emptyState}>
          No search context configured. Please set a Search Context ID in the web part properties.
        </div>
      </div>
    );
  }

  if (refiners.length === 0) {
    return (
      <div className={styles.spSearchFilters}>
        <div className={styles.emptyState}>
          No filters available. Perform a search to see available filters.
        </div>
      </div>
    );
  }

  return (
    <div className={styles.spSearchFilters}>
      {/* Active filter pill bar */}
      <FilterPillBar
        activeFilters={displayFilters}
        filterConfig={configs}
        onRemoveFilter={handleRemoveFilter}
        onClearAll={handleClearAll}
        showClearAll={showClearAll}
      />

      {/* Filter groups */}
      {refiners.map(function (refiner: IRefiner): React.ReactElement {
        const config: IFilterConfig | undefined = findFilterConfig(refiner.filterName, configs);
        return (
          <FilterGroup
            key={refiner.filterName}
            refiner={refiner}
            config={config}
            activeFilters={displayFilters}
            onToggleRefiner={handleToggleRefiner}
          />
        );
      })}

      {/* Apply button for manual mode */}
      {applyMode === 'manual' && hasPendingChanges && (
        <div className={styles.applyBar}>
          <button
            type="button"
            className={styles.applyButton}
            onClick={handleApply}
            aria-label="Apply filters"
          >
            Apply filters
          </button>
        </div>
      )}
    </div>
  );
};

export default SpSearchFilters;
