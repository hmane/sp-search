import * as React from 'react';
import * as strings from 'SpSearchFiltersWebPartStrings';
import { IconButton, DefaultButton, PrimaryButton } from '@fluentui/react/lib/Button';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { Shimmer, ShimmerElementType } from '@fluentui/react/lib/Shimmer';
import { useMediaQuery } from 'spfx-toolkit/lib/hooks/useViewport';
import { lazyBridge } from '../../../utilities/lazyBridge';
import styles from './SpSearchFilters.module.scss';
import type { ISpSearchFiltersProps } from './ISpSearchFiltersProps';
import FilterGroup from './FilterGroup';
import type {
  IRefiner,
  IActiveFilter,
  IFilterConfig,
  IReplaceRefinerValuesPayload,
  ISearchStore
} from '@interfaces/index';
import { areFiltersEquivalent } from '@store/utils/filterValueMatching';
import { isInAudience, fetchManagedProperties, getCachedSchema } from '@services/index';
// T4.D5 — edit-mode validator surfaces a Did-You-Mean MessageBar for any
// filter row whose `managedProperty` doesn't match a known managed property
// in the cached schema. Validator passes silently when schema cache is cold.
import { validateManagedPropertyCollection } from '@store/configValidation/sharedValidators';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import type { IManagedProperty } from '@store/interfaces/ISearchDataProvider';
// T4.D12 — cross-web-part preset propagation. When Results applies a
// preset, the suggestion lands in the registry; this component subscribes
// and renders an edit-mode MessageBar offering to apply the suggested filters.
import {
  consumePresetSuggestion,
  clearPresetSuggestion,
  subscribePresetSuggestionChanges,
  type IPresetSuggestion,
} from '@store/utils/presetSuggestionRegistry';
// T5.D1 — cross-bundle singleton DebugFab + Panel host.
import { DebugFabHost } from '../../../utilities/DebugFabHost';
import { ShortcutHelpModalHost } from '../../../utilities/ShortcutHelpModal';

const VisualFilterBuilder = lazyBridge(
  () => import(/* webpackChunkName: 'VisualFilterBuilder' */ './VisualFilterBuilder') as unknown as Promise<{ default: React.ComponentType<Record<string, unknown>> }>,
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

function buildDisplayRefiners(refiners: IRefiner[], configs: IFilterConfig[]): IRefiner[] {
  // Admin-configured `configs` order is canonical — drive the loop from configs
  // and look up the matching server-returned refiner bucket. Refiners that the
  // server returned but the admin didn't configure get appended at the end so
  // they're still discoverable.
  const merged: IRefiner[] = [];
  const used = new Set<string>();
  const refinerByName = new Map<string, IRefiner>();
  for (let i: number = 0; i < refiners.length; i++) {
    refinerByName.set(refiners[i].filterName, refiners[i]);
  }

  for (let i: number = 0; i < configs.length; i++) {
    const config = configs[i];
    const bucket = refinerByName.get(config.managedProperty);

    if (bucket) {
      merged.push(bucket);
      used.add(config.managedProperty);
      continue;
    }

    const canRenderWithoutBuckets =
      config.filterType === 'people' ||
      config.filterType === 'daterange' ||
      config.filterType === 'toggle' ||
      config.filterType === 'text';

    if (canRenderWithoutBuckets) {
      merged.push({
        filterName: config.managedProperty,
        values: []
      });
      used.add(config.managedProperty);
    }
  }

  for (let i: number = 0; i < refiners.length; i++) {
    if (!used.has(refiners[i].filterName)) {
      merged.push(refiners[i]);
    }
  }

  return merged;
}

function isSingleValueFilter(config: IFilterConfig | undefined): boolean {
  if (!config) {
    return false;
  }

  if (config.multiValues === false) {
    return true;
  }

  return config.filterType === 'daterange' ||
    config.filterType === 'slider' ||
    config.filterType === 'toggle' ||
    config.filterType === 'text';
}

function hasActiveSelection(filterName: string, filters: IActiveFilter[]): boolean {
  for (let i = 0; i < filters.length; i++) {
    if (filters[i].filterName === filterName) {
      return true;
    }
  }
  return false;
}

function clearDependentFilters(
  current: IActiveFilter[],
  changedFilterName: string,
  configs: IFilterConfig[]
): IActiveFilter[] {
  const dependentNames = new Set<string>();
  const queue: string[] = [changedFilterName];

  while (queue.length > 0) {
    const parent = queue.shift() as string;
    for (let i = 0; i < configs.length; i++) {
      const config = configs[i];
      if (config.dependsOn === parent && config.resetWhenParentChanges === true && !dependentNames.has(config.managedProperty)) {
        dependentNames.add(config.managedProperty);
        queue.push(config.managedProperty);
      }
    }
  }

  if (dependentNames.size === 0) {
    return current;
  }

  return current.filter(function (filter: IActiveFilter): boolean {
    return !dependentNames.has(filter.filterName);
  });
}

function buildRenderableRefiner(
  refiner: IRefiner,
  config: IFilterConfig | undefined,
  activeFilters: IActiveFilter[]
): IRefiner | undefined {
  if (!config) {
    return refiner;
  }

  if (config.dependsOn && config.showWhenParentHasValue && !hasActiveSelection(config.dependsOn, activeFilters)) {
    return undefined;
  }

  if (config.hideZeroCountValues !== true) {
    return refiner;
  }

  const visibleValues = refiner.values.filter(function (value): boolean {
    if (value.count > 0) {
      return true;
    }

    const candidate: IActiveFilter = {
      filterName: refiner.filterName,
      value: value.value,
      displayValue: value.name || undefined,
      operator: config.operator
    };

    for (let i = 0; i < activeFilters.length; i++) {
      if (areFiltersEquivalent(activeFilters[i], candidate)) {
        return true;
      }
    }

    return false;
  });

  return {
    filterName: refiner.filterName,
    values: visibleValues
  };
}

function areFiltersEqual(a: IActiveFilter[], b: IActiveFilter[]): boolean {
  if (a.length !== b.length) {
    return false;
  }

  for (let i: number = 0; i < a.length; i++) {
    if (
      a[i].filterName !== b[i].filterName ||
      a[i].value !== b[i].value ||
      a[i].displayValue !== b[i].displayValue ||
      a[i].operator !== b[i].operator
    ) {
      return false;
    }
  }

  return true;
}

function buildNextFilters(
  current: IActiveFilter[],
  nextFilter: IActiveFilter,
  config: IFilterConfig | undefined
): IActiveFilter[] {
  const sameValueIndex = current.findIndex(function (filter: IActiveFilter): boolean {
    return areFiltersEquivalent(filter, nextFilter);
  });

  if (sameValueIndex >= 0) {
    const updated = current.slice();
    updated.splice(sameValueIndex, 1);
    return updated;
  }

  if (isSingleValueFilter(config)) {
    const withoutSameName = current.filter(function (filter: IActiveFilter): boolean {
      return filter.filterName !== nextFilter.filterName;
    });
    return withoutSameName.concat([nextFilter]);
  }

  return current.concat([nextFilter]);
}

/**
 * Pure helper used by the multi-value batched callback. Removes every active
 * filter matching `filterName`, then appends the supplied values (filtering
 * out any that name a different filter). Always returns a new array reference
 * so the Zustand store + orchestrator subscribe trigger reliably.
 */
export function applyReplaceRefinerValues(
  current: IActiveFilter[],
  filterName: string,
  values: IActiveFilter[]
): IActiveFilter[] {
  const kept = current.filter(function (f: IActiveFilter): boolean {
    return f.filterName !== filterName;
  });
  const accepted = values.filter(function (v: IActiveFilter): boolean {
    return v.filterName === filterName;
  });
  return kept.concat(accepted);
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
  const { store, applyMode, showClearAll, enableVisualFilterBuilder, isEditMode, searchContextId, onApplyPresetFilters } = props;

  // T4.D5 — fetch schema on mount in edit mode so the synchronous
  // validator below has something to compare against. Fire-and-forget;
  // the validator handles missing schema gracefully.
  const [schema, setSchema] = React.useState<IManagedProperty[] | undefined>(() => getCachedSchema());
  React.useEffect((): void => {
    if (!isEditMode) { return; }
    if (schema && schema.length > 0) { return; }
    fetchManagedProperties()
      .then((result): void => {
        if (result.status === 'success' && result.properties.length > 0) {
          setSchema(result.properties);
        }
      })
      .catch((): void => { /* silent — validator passes when schema missing */ });
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [isEditMode]);

  // T4.D12 — cross-web-part preset suggestion. Subscribe to registry
  // changes in edit mode and re-render when the Results web part publishes
  // a preset application. View mode never subscribes — the MessageBar is
  // an admin-only affordance.
  const [presetSuggestion, setPresetSuggestion] = React.useState<IPresetSuggestion | undefined>(
    () => isEditMode && searchContextId ? consumePresetSuggestion(searchContextId) : undefined
  );
  React.useEffect((): (() => void) | undefined => {
    if (!isEditMode || !searchContextId) { return undefined; }
    setPresetSuggestion(consumePresetSuggestion(searchContextId));
    const unsub = subscribePresetSuggestionChanges((): void => {
      setPresetSuggestion(consumePresetSuggestion(searchContextId));
    });
    return unsub;
  }, [isEditMode, searchContextId]);

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

  // Stream D / #5 — audience-targeted refiners. `currentUserGroups` is
  // populated by storeRegistry's fire-and-forget call to `resolveUserGroupIds()`
  // on context init; until it resolves (or if the Graph call fails) all
  // audience-targeted refiners stay hidden (fail-closed).
  const selectCurrentUserGroups = React.useCallback(function (s: ISearchStore): string[] {
    return s.currentUserGroups;
  }, []);

  const selectIsLoading = React.useCallback(function (s: ISearchStore): boolean {
    return s.isLoading;
  }, []);

  const selectOperatorBetweenFilters = React.useCallback(function (s: ISearchStore): 'AND' | 'OR' {
    return s.operatorBetweenFilters;
  }, []);

  const availableRefiners: IRefiner[] | undefined = useStoreState(store, selectRefiners);
  const activeFilters: IActiveFilter[] | undefined = useStoreState(store, selectActiveFilters);
  const filterConfig: IFilterConfig[] | undefined = useStoreState(store, selectFilterConfig);
  const isLoading: boolean | undefined = useStoreState(store, selectIsLoading);
  const operatorBetweenFilters: 'AND' | 'OR' | undefined = useStoreState(store, selectOperatorBetweenFilters);
  const currentUserGroups: string[] | undefined = useStoreState(store, selectCurrentUserGroups);

  // Use safe defaults
  const refiners: IRefiner[] = availableRefiners || [];
  const filters: IActiveFilter[] = activeFilters || [];
  const userGroups: string[] = currentUserGroups || [];
  const configs: IFilterConfig[] = (filterConfig || []).filter((config: IFilterConfig): boolean => {
    // Stream D / #5 — hide refiners whose audienceGroups exclude the current user.
    if (!config.audienceGroups || config.audienceGroups.length === 0) {
      return true;
    }
    return isInAudience(config.audienceGroups, userGroups);
  });

  // Pending filters for manual mode
  const [pendingFilters, setPendingFilters] = React.useState<IActiveFilter[]>([]);
  const [hasPendingChanges, setHasPendingChanges] = React.useState<boolean>(false);

  // Visual Filter Builder toggle
  const [isBuilderOpen, setIsBuilderOpen] = React.useState<boolean>(false);

  // T1.D1 — phone-width drawer state. Below 640px the filter body collapses
  // behind a "Show filters" toggle that opens a Fluent Panel (off-canvas
  // surface with built-in FocusTrapZone + Escape-to-close + light dismiss
  // backdrop). Desktop ignores both pieces of state.
  const isMobile: boolean = useMediaQuery('(max-width: 639px)', false);
  const [isDrawerOpen, setIsDrawerOpen] = React.useState<boolean>(false);
  const openDrawer = React.useCallback((): void => { setIsDrawerOpen(true); }, []);
  const closeDrawer = React.useCallback((): void => { setIsDrawerOpen(false); }, []);

  // When the viewport grows past 640px while the drawer is open, drop the
  // drawer flag — the body is about to render inline so leaving the modal
  // mounted would briefly trap focus inside an invisible surface.
  React.useEffect((): void => {
    if (!isMobile && isDrawerOpen) {
      setIsDrawerOpen(false);
    }
  }, [isMobile, isDrawerOpen]);

  // Determine which filters to display: pending (manual mode) or live (instant mode)
  const displayFilters: IActiveFilter[] = applyMode === 'manual' && hasPendingChanges
    ? pendingFilters
    : filters;
  const displayRefiners: IRefiner[] = React.useMemo(
    function (): IRefiner[] {
      return buildDisplayRefiners(refiners, configs);
    },
    [refiners, configs]
  );
  const renderableRefiners: IRefiner[] = React.useMemo(
    function (): IRefiner[] {
      const next: IRefiner[] = [];
      for (let i = 0; i < displayRefiners.length; i++) {
        const refiner = displayRefiners[i];
        const config = findFilterConfig(refiner.filterName, configs);
        const renderable = buildRenderableRefiner(refiner, config, displayFilters);
        if (renderable) {
          next.push(renderable);
        }
      }
      return next;
    },
    [configs, displayFilters, displayRefiners]
  );

  // Sync pending filters when store filters change (manual mode)
  React.useEffect(function (): void {
    if (applyMode === 'manual' && !hasPendingChanges) {
      setPendingFilters(filters);
    }
  }, [filters, applyMode, hasPendingChanges]);

  // Detect when activeFilters changes externally to match pending state (manual mode)
  React.useEffect(function (): void {
    if (applyMode === 'manual' && hasPendingChanges) {
      const pendingMatchesStore = areFiltersEqual(pendingFilters, filters);
      if (pendingMatchesStore) {
        setHasPendingChanges(false);
      }
    }
  }, [filters]); // eslint-disable-line react-hooks/exhaustive-deps

  /** Handle toggling a refiner checkbox. */
  function handleToggleRefiner(filter: IActiveFilter): void {
    if (!store) {
      return;
    }

    const config = findFilterConfig(filter.filterName, configs);

    if (applyMode === 'instant') {
      const nextFilters = clearDependentFilters(
        buildNextFilters(filters, filter, config),
        filter.filterName,
        configs
      );
      store.setState({
        activeFilters: nextFilters,
        currentPage: 1
      });
    } else {
      const current: IActiveFilter[] = hasPendingChanges ? pendingFilters : filters;
      const updated = clearDependentFilters(
        buildNextFilters(current, filter, config),
        filter.filterName,
        configs
      );
      setPendingFilters(updated);
      setHasPendingChanges(!areFiltersEqual(updated, filters));
    }
  }

  /** Multi-value batched: replace all values for a single filterName in one call. */
  function handleReplaceRefinerValues(payload: IReplaceRefinerValuesPayload): void {
    if (!store) {
      return;
    }

    if (applyMode === 'instant') {
      const replaced = applyReplaceRefinerValues(filters, payload.filterName, payload.values);
      const nextFilters = clearDependentFilters(replaced, payload.filterName, configs);
      store.setState({ activeFilters: nextFilters, currentPage: 1 });
    } else {
      const current: IActiveFilter[] = hasPendingChanges ? pendingFilters : filters;
      const replaced = applyReplaceRefinerValues(current, payload.filterName, payload.values);
      const updated = clearDependentFilters(replaced, payload.filterName, configs);
      setPendingFilters(updated);
      setHasPendingChanges(!areFiltersEqual(updated, filters));
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
    // T1.D1 — closing the drawer on Apply matches the iOS/Android filter
    // pattern: the user explicitly committed; the surface gets out of the way.
    setIsDrawerOpen(false);
  }

  /** Clear all active filters. Respects apply mode: instant dispatches to store, manual clears pending. */
  function handleClearAll(): void {
    if (!store) {
      return;
    }

    if (applyMode === 'instant') {
      store.getState().clearAllFilters();
    } else {
      setPendingFilters([]);
      setHasPendingChanges(!areFiltersEqual([], filters));
    }
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

  if (renderableRefiners.length === 0 && !isLoading) {
    return (
      <div className={styles.spSearchFilters}>
        <div className={styles.emptyState} role="status">
          No filters available. Perform a search to see available filters.
        </div>
      </div>
    );
  }

  if (renderableRefiners.length === 0 && isLoading) {
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

  // T4.D5 — edit-mode-only managed-property validation MessageBar block.
  // Runs against the cached schema; admin sees Did-You-Mean copy for any
  // filterConfig row whose `managedProperty` doesn't match a known property.
  const validationBlock: React.ReactNode = isEditMode ? ((): React.ReactNode => {
    const issues = validateManagedPropertyCollection(
      (filterConfig || []).map((c) => ({ managedProperty: c.managedProperty, property: c.managedProperty })),
      schema
    );
    if (issues.length === 0) { return null; }
    return (
      <>
        {issues.map((issue) => (
          <MessageBar
            key={issue.id}
            messageBarType={issue.severity === 'error' ? MessageBarType.error : MessageBarType.warning}
            isMultiline={true}
            styles={{ root: { marginBottom: 4 } }}
          >
            {issue.message}
          </MessageBar>
        ))}
      </>
    );
  })() : null;

  // T4.D12 — preset suggestion MessageBar. Edit-mode only. Shows the
  // preset's `filterSuggestions` count, offers an Apply button that calls
  // back to the web part to write `filtersCollection`, and a Dismiss that
  // clears the suggestion. Existing managed-property rows are not
  // duplicated — the web part filters them out in `onApplyPresetFilters`.
  const presetBlock: React.ReactNode = isEditMode && presetSuggestion && presetSuggestion.filterSuggestions.length > 0 ? (
    <MessageBar
      messageBarType={MessageBarType.info}
      isMultiline={true}
      styles={{ root: { marginBottom: 4 } }}
      actions={(
        <div>
          <PrimaryButton
            text="Apply"
            onClick={(): void => {
              if (onApplyPresetFilters && presetSuggestion) {
                onApplyPresetFilters(presetSuggestion.filterSuggestions);
              }
              if (searchContextId) {
                clearPresetSuggestion(searchContextId);
              }
            }}
          />
          <DefaultButton
            text="Dismiss"
            onClick={(): void => {
              if (searchContextId) { clearPresetSuggestion(searchContextId); }
            }}
            style={{ marginLeft: 8 }}
          />
        </div>
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      ) as any}
    >
      The <strong>{presetSuggestion.label}</strong> preset suggests {presetSuggestion.filterSuggestions.length} filter{presetSuggestion.filterSuggestions.length === 1 ? '' : 's'} for this web part — apply now?
    </MessageBar>
  ) : null;

  // T1.D1 — body content is rendered identically inline (desktop) and inside
  // the mobile drawer Panel. Extracted so we only describe the filter UI once.
  const filterBody: React.ReactElement = (
    <>
      {presetBlock}
      {validationBlock}
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
          refiners={displayRefiners}
          filterConfig={configs}
          activeFilters={displayFilters}
          operatorBetweenFilters={operatorBetweenFilters || 'AND'}
          onApplyFilters={handleBuilderApply}
          onCancel={function (): void { setIsBuilderOpen(false); }}
        />
      )}

      {/* Clear All button */}
      {showClearAll && displayFilters.length > 0 && (
        <div className={styles.clearAllBar}>
          <button
            type="button"
            className={styles.clearAllButton}
            onClick={handleClearAll}
            aria-label={strings.ClearAllFiltersLabel}
          >
            {strings.ClearAllFiltersLabel}
          </button>
        </div>
      )}

      {/* Filter groups */}
      {renderableRefiners.map(function (refiner: IRefiner): React.ReactElement {
        const config: IFilterConfig | undefined = findFilterConfig(refiner.filterName, configs);
        return (
          <FilterGroup
            key={refiner.filterName}
            refiner={refiner}
            config={config}
            activeFilters={displayFilters}
            onToggleRefiner={handleToggleRefiner}
            onReplaceRefinerValues={handleReplaceRefinerValues}
          />
        );
      })}

      {/* Apply button for manual mode */}
      {applyMode === 'manual' && hasPendingChanges && ((): React.ReactNode => {
        // T1.D8 — pending-change count. Counts the symmetric difference
        // between pendingFilters and store filters (an "applied" filter
        // counts once whether it's an addition or a removal).
        const pendingKeys = new Set(pendingFilters.map((f) => f.filterName + '|' + f.value));
        const liveKeys = new Set(filters.map((f) => f.filterName + '|' + f.value));
        let changeCount = 0;
        pendingKeys.forEach((k) => { if (!liveKeys.has(k)) { changeCount++; } });
        liveKeys.forEach((k) => { if (!pendingKeys.has(k)) { changeCount++; } });
        const buttonText = changeCount === 1 ? 'Apply 1 change' : 'Apply ' + changeCount + ' changes';
        return (
          <div className={styles.applyBar}>
            <button
              type="button"
              className={styles.applyButton}
              onClick={handleApply}
              aria-label={buttonText}
            >
              {buttonText}
            </button>
          </div>
        );
      })()}
    </>
  );

  // T1.D1 — mobile (≤639px): collapse the body behind a "Show filters" toggle.
  // The active-filter count is part of the label so the user sees what's
  // applied without having to open the drawer. Fluent Panel handles focus
  // trap, Escape-to-close, light dismiss, aria-modal, and motion-reduction.
  if (isMobile) {
    const activeCount: number = displayFilters.length;
    const toggleLabel: string = activeCount > 0
      ? `${strings.ShowFiltersLabel || 'Show filters'} (${activeCount})`
      : (strings.ShowFiltersLabel || 'Show filters');

    return (
      <div className={styles.spSearchFilters}>
        <div className={styles.drawerToggleBar}>
          <button
            type="button"
            className={styles.drawerToggleButton}
            onClick={openDrawer}
            aria-haspopup="dialog"
            aria-expanded={isDrawerOpen}
            aria-label={activeCount > 0
              ? `${strings.ShowFiltersLabel || 'Show filters'}, ${activeCount} active`
              : (strings.ShowFiltersLabel || 'Show filters')}
          >
            {toggleLabel}
          </button>
        </div>
        <Panel
          isOpen={isDrawerOpen}
          onDismiss={closeDrawer}
          type={PanelType.smallFluid}
          headerText={strings.FiltersPanelHeaderLabel || 'Filters'}
          isLightDismiss={true}
          closeButtonAriaLabel={strings.CloseFiltersLabel || 'Close filters'}
        >
          <div className={styles.drawerContent}>
            {filterBody}
          </div>
        </Panel>
      </div>
    );
  }

  return (
    <div className={styles.spSearchFilters}>
      {filterBody}
      {/* T5.D1 — singleton DebugFab host. */}
      {store && <DebugFabHost store={store} />}
      {/* T2.D9 — singleton shortcut help modal host (cross-bundle owner claim). */}
      <ShortcutHelpModalHost />
    </div>
  );
};

export default SpSearchFilters;
