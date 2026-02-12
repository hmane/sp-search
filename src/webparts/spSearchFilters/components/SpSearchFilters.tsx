import * as React from 'react';
import { IconButton } from '@fluentui/react/lib/Button';
import { Shimmer, ShimmerElementType } from '@fluentui/react/lib/Shimmer';
import { createLazyComponent } from 'spfx-toolkit/lib/utilities/lazyLoader';
import styles from './SpSearchFilters.module.scss';
import type { ISpSearchFiltersProps } from './ISpSearchFiltersProps';
import FilterGroup from './FilterGroup';
import type {
  IRefiner,
  IActiveFilter,
  IFilterConfig,
  ISearchStore
} from '@interfaces/index';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const VisualFilterBuilder: any = createLazyComponent(
  () => import('./VisualFilterBuilder') as any,
  { errorMessage: 'Failed to load visual filter builder' }
);

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

    // Subscribe to changes — only update if selector returns a new reference
    const unsubscribe: () => void = store.subscribe(function (newState: ISearchStore): void {
      const nextValue: T = selector(newState);
      setState(function (prev: T | undefined): T | undefined {
        return prev === nextValue ? prev : nextValue;
      });
    });

    return unsubscribe;
  }, [store, selector]);

  return state;
}

const SpSearchFilters: React.FC<ISpSearchFiltersProps> = (props: ISpSearchFiltersProps): React.ReactElement => {
  const { store, applyMode, enableVisualFilterBuilder } = props;

  // Stable selectors to avoid re-subscriptions
  const selectRefiners = React.useCallback(function (s: ISearchStore): IRefiner[] {
    return s.displayRefiners;
  }, []);

  const selectActiveFilters = React.useCallback(function (s: ISearchStore): IActiveFilter[] {
    return s.activeFilters;
  }, []);

  const selectFilterConfig = React.useCallback(function (s: ISearchStore): IFilterConfig[] {
    return s.filterConfig;
  }, []);

  const selectIsLoading = React.useCallback(function (s: ISearchStore): boolean {
    return s.isLoading;
  }, []);

  const availableRefiners: IRefiner[] | undefined = useStoreState(store, selectRefiners);
  const activeFilters: IActiveFilter[] | undefined = useStoreState(store, selectActiveFilters);
  const filterConfig: IFilterConfig[] | undefined = useStoreState(store, selectFilterConfig);
  const isLoading: boolean | undefined = useStoreState(store, selectIsLoading);

  // Use safe defaults
  const refiners: IRefiner[] = availableRefiners || [];
  const filters: IActiveFilter[] = activeFilters || [];
  const configs: IFilterConfig[] = filterConfig || [];

  // Pending filters for manual mode
  const [pendingFilters, setPendingFilters] = React.useState<IActiveFilter[]>([]);
  const [hasPendingChanges, setHasPendingChanges] = React.useState<boolean>(false);

  // Visual Filter Builder toggle
  const [isBuilderOpen, setIsBuilderOpen] = React.useState<boolean>(false);

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

  /** Handle applying filters from the visual filter builder. */
  function handleBuilderApply(builderFilters: IActiveFilter[]): void {
    if (!store) {
      return;
    }
    const storeState: ISearchStore = store.getState();
    storeState.clearAllFilters();
    for (let i = 0; i < builderFilters.length; i++) {
      storeState.setRefiner(builderFilters[i]);
    }
    setIsBuilderOpen(false);
    setHasPendingChanges(false);
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
        <div className={styles.emptyState} role="status">
          No search context configured. Please set a Search Context ID in the web part properties.
        </div>
      </div>
    );
  }

  if (refiners.length === 0 && !isLoading) {
    return (
      <div className={styles.spSearchFilters}>
        <div className={styles.emptyState} role="status">
          No filters available. Perform a search to see available filters.
        </div>
      </div>
    );
  }

  if (refiners.length === 0 && isLoading) {
    return (
      <div className={styles.spSearchFilters}>
        {[0, 1, 2].map(function (i: number): React.ReactElement {
          return (
            <div key={i} className={styles.shimmerGroup}>
              <Shimmer
                shimmerElements={[
                  { type: ShimmerElementType.line, height: 14, width: '45%' }
                ]}
                width="100%"
              />
              <Shimmer
                shimmerElements={[
                  { type: ShimmerElementType.line, height: 12, width: '80%' }
                ]}
                width="100%"
                style={{ marginTop: 8 }}
              />
              <Shimmer
                shimmerElements={[
                  { type: ShimmerElementType.line, height: 12, width: '65%' }
                ]}
                width="100%"
                style={{ marginTop: 6 }}
              />
              <Shimmer
                shimmerElements={[
                  { type: ShimmerElementType.line, height: 12, width: '55%' }
                ]}
                width="100%"
                style={{ marginTop: 6 }}
              />
            </div>
          );
        })}
      </div>
    );
  }

  return (
    <div className={styles.spSearchFilters}>
      {/* Visual Filter Builder toggle */}
      {enableVisualFilterBuilder && (
        <div className={styles.visualFilterBuilderToggle}>
          <IconButton
            iconProps={{ iconName: isBuilderOpen ? 'Cancel' : 'Filter' }}
            title={isBuilderOpen ? 'Close visual filter builder' : 'Open visual filter builder'}
            ariaLabel={isBuilderOpen ? 'Close visual filter builder' : 'Open visual filter builder'}
            onClick={function (): void { setIsBuilderOpen(!isBuilderOpen); }}
            checked={isBuilderOpen}
          />
        </div>
      )}

      {/* Visual Filter Builder panel */}
      {enableVisualFilterBuilder && isBuilderOpen && (
        <VisualFilterBuilder
          refiners={refiners}
          filterConfig={configs}
          activeFilters={displayFilters}
          onApplyFilters={handleBuilderApply}
          onCancel={function (): void { setIsBuilderOpen(false); }}
        />
      )}

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
