import * as React from 'react';
import { Shimmer, ShimmerElementType } from '@fluentui/react/lib/Shimmer';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';
import { ErrorBoundary } from 'spfx-toolkit/lib/components/ErrorBoundary';
import { sanitizeHtml } from 'spfx-toolkit/lib/utilities/htmlUtils/sanitizeHtml';
import { lazyBridge } from '../../../utilities/lazyBridge';
import type { ISpSearchResultsProps } from './ISpSearchResultsProps';
import {
  ISearchResult,
  IPromotedResultItem,
  ISortField,
  ISortableProperty,
  IActiveFilter,
  IFilterConfig,
  ISearchScope
} from '@interfaces/index';
import ResultToolbar from './ResultToolbar';
import ActiveFilterPillBar from './ActiveFilterPillBar';
import Pagination from './Pagination';
import { spLog } from '@store/utils/spLog';

// ─── Lazy-loaded layouts (code-split per layout) ─────────
const ListLayout = lazyBridge(
  () => import(/* webpackChunkName: 'ListLayout' */ './ListLayout') as unknown as Promise<{ default: React.ComponentType<Record<string, unknown>> }>,
  { errorMessage: 'Failed to load list layout' }
);
const CompactLayout = lazyBridge(
  () => import(/* webpackChunkName: 'CompactLayout' */ './CompactLayout') as unknown as Promise<{ default: React.ComponentType<Record<string, unknown>> }>,
  { errorMessage: 'Failed to load compact layout' }
);
const DataGridLayout = lazyBridge(
  () => import(/* webpackChunkName: 'DataGridLayout' */ './DataGridLayout') as unknown as Promise<{ default: React.ComponentType<Record<string, unknown>> }>,
  { errorMessage: 'Failed to load data grid layout' }
);
import styles from './SpSearchResults.module.scss';
import { validateWebPartConfig, IConfigWarning, ConfigWarningLevel } from './configValidation';

// ─── Lazy layout preloaders ──────────────────────────────
// Called on hover of the corresponding toolbar button to warm the webpack
// chunk before the user actually clicks. No-ops on subsequent calls because
// webpack deduplicates dynamic imports after the first load.
const LAYOUT_PRELOADERS: Record<string, () => void> = {
  list:    (): void => { import(/* webpackChunkName: 'ListLayout' */    './ListLayout')    .catch((): void => { /* ignore preload error */ }); },
  compact: (): void => { import(/* webpackChunkName: 'CompactLayout' */ './CompactLayout') .catch((): void => { /* ignore preload error */ }); },
  card:    (): void => { import(/* webpackChunkName: 'CardLayout' */    './CardLayout')    .catch((): void => { /* ignore preload error */ }); },
  people:  (): void => { import(/* webpackChunkName: 'PeopleLayout' */  './PeopleLayout')  .catch((): void => { /* ignore preload error */ }); },
  grid:    (): void => { import(/* webpackChunkName: 'DataGridLayout' */'./DataGridLayout') .catch((): void => { /* ignore preload error */ }); },
  gallery: (): void => { import(/* webpackChunkName: 'GalleryLayout' */ './GalleryLayout') .catch((): void => { /* ignore preload error */ }); },
};

// ─── Lazy-loaded layouts and panels ──────────────────────
const CardLayout = lazyBridge(
  () => import(/* webpackChunkName: 'CardLayout' */ './CardLayout') as unknown as Promise<{ default: React.ComponentType<Record<string, unknown>> }>,
  { errorMessage: 'Failed to load card layout' }
);
const PeopleLayout = lazyBridge(
  () => import(/* webpackChunkName: 'PeopleLayout' */ './PeopleLayout') as unknown as Promise<{ default: React.ComponentType<Record<string, unknown>> }>,
  { errorMessage: 'Failed to load people layout' }
);
const GalleryLayout = lazyBridge(
  () => import(/* webpackChunkName: 'GalleryLayout' */ './GalleryLayout') as unknown as Promise<{ default: React.ComponentType<Record<string, unknown>> }>,
  { errorMessage: 'Failed to load gallery layout' }
);
const ResultDetailPanel = lazyBridge(
  () => import(/* webpackChunkName: 'ResultDetailPanel' */ './ResultDetailPanel') as unknown as Promise<{ default: React.ComponentType<Record<string, unknown>> }>,
  { errorMessage: 'Failed to load detail panel' }
);

import { safeNavigate } from '@store/utils/safeNavigate';
// T3.D10 — init-order diagnostic helpers.
import { hasInitOrderIssue, clearInitOrderDiagnostic } from '@store/utils/initOrderDiagnostic';
// T2.D11 — layout-agnostic export.
import { exportItemsAsCsv, exportItemsAsXlsx } from './exportShared';

// T5.D1 — singleton DebugFab + Panel imported from the cross-bundle host.
import { DebugFabHost } from '../../../utilities/DebugFabHost';
import { ShortcutHelpModalHost } from '../../../utilities/ShortcutHelpModal';

/**
 * Custom hook that subscribes to the Zustand vanilla store and
 * returns selected state slices. Uses store.subscribe for efficient
 * re-renders by comparing relevant fields.
 */
function useStoreState(
  props: ISpSearchResultsProps
): {
  items: ISearchResult[];
  totalCount: number;
  currentPage: number;
  pageSize: number;
  isLoading: boolean;
  hasSearched: boolean;
  error: string | undefined;
  queryText: string;
  activeLayoutKey: string;
  availableLayouts: string[];
  promotedResults: IPromotedResultItem[];
  sort: ISortField | undefined;
  sortableProperties: ISortableProperty[];
  previewPanel: { isOpen: boolean; item: ISearchResult | undefined };
  activeFilters: IActiveFilter[];
  filterConfig: IFilterConfig[];
  querySuggestion: string | undefined;
  showPaging: boolean;
  pageRange: number;
  scope: ISearchScope;
} {
  const { store } = props;

  const emptyState = React.useMemo(() => ({
    items: [] as ISearchResult[],
    totalCount: 0,
    currentPage: 1,
    pageSize: 25,
    isLoading: true,
    hasSearched: false,
    error: undefined as string | undefined,
    queryText: '',
    activeLayoutKey: 'list',
    availableLayouts: ['list', 'compact', 'grid'],
    promotedResults: [] as IPromotedResultItem[],
    sort: undefined as ISortField | undefined,
    sortableProperties: [] as ISortableProperty[],
    previewPanel: { isOpen: false, item: undefined as ISearchResult | undefined },
    activeFilters: [] as IActiveFilter[],
    filterConfig: [] as IFilterConfig[],
    querySuggestion: undefined as string | undefined,
    showPaging: true,
    pageRange: 5,
    scope: { id: 'all', label: 'All SharePoint' } as ISearchScope,
  }), []);

  const getSnapshot = React.useCallback((): {
    items: ISearchResult[];
    totalCount: number;
    currentPage: number;
    pageSize: number;
    isLoading: boolean;
    hasSearched: boolean;
    error: string | undefined;
    queryText: string;
    activeLayoutKey: string;
    availableLayouts: string[];
    promotedResults: IPromotedResultItem[];
    sort: ISortField | undefined;
    sortableProperties: ISortableProperty[];
    previewPanel: { isOpen: boolean; item: ISearchResult | undefined };
    activeFilters: IActiveFilter[];
    filterConfig: IFilterConfig[];
    querySuggestion: string | undefined;
    showPaging: boolean;
    pageRange: number;
    scope: ISearchScope;
  } => {
    if (!store) {
      return emptyState;
    }
    const state = store.getState();
    return {
      items: state.items,
      totalCount: state.totalCount,
      currentPage: state.currentPage,
      pageSize: state.pageSize,
      isLoading: state.isLoading,
      hasSearched: state.hasSearched,
      error: state.error,
      queryText: state.queryText,
      activeLayoutKey: state.activeLayoutKey,
      availableLayouts: state.availableLayouts,
      promotedResults: state.promotedResults,
      sort: state.sort,
      sortableProperties: state.sortableProperties,
      previewPanel: state.previewPanel,
      activeFilters: state.activeFilters,
      filterConfig: state.filterConfig,
      querySuggestion: state.querySuggestion,
      showPaging: state.showPaging,
      pageRange: state.pageRange,
      scope: state.scope,
    };
  }, [store, emptyState]);

  const [storeState, setStoreState] = React.useState(getSnapshot);

  React.useEffect((): (() => void) => {
    if (!store) {
      return (): void => { /* noop */ };
    }
    const unsubscribe = store.subscribe((): void => {
      const next = getSnapshot();
      setStoreState((prev) => {
        // Shallow comparison of relevant fields to avoid unnecessary re-renders
        if (
          prev.items === next.items &&
          prev.totalCount === next.totalCount &&
          prev.currentPage === next.currentPage &&
          prev.pageSize === next.pageSize &&
          prev.isLoading === next.isLoading &&
          prev.hasSearched === next.hasSearched &&
          prev.error === next.error &&
          prev.queryText === next.queryText &&
          prev.activeLayoutKey === next.activeLayoutKey &&
          prev.availableLayouts === next.availableLayouts &&
          prev.promotedResults === next.promotedResults &&
          prev.sort === next.sort &&
          prev.sortableProperties === next.sortableProperties &&
          prev.previewPanel === next.previewPanel &&
          prev.activeFilters === next.activeFilters &&
          prev.filterConfig === next.filterConfig &&
          prev.querySuggestion === next.querySuggestion &&
          prev.showPaging === next.showPaging &&
          prev.pageRange === next.pageRange &&
          prev.scope === next.scope
        ) {
          return prev;
        }
        return next;
      });
    });
    return unsubscribe;
  }, [store, getSnapshot]);

  return storeState;
}

/**
 * Renders the loading shimmer placeholders that match the list layout shape.
 */
const LoadingShimmer: React.FC<{ count: number }> = (shimmerProps) => {
  const titleWidths = ['54%', '48%', '58%', '44%'];
  const urlWidths = ['38%', '46%', '34%', '41%'];
  const summaryWidths = [
    ['92%', '78%'],
    ['88%', '72%'],
    ['94%', '75%'],
    ['86%', '69%']
  ];
  const metaWidths = [
    ['92', '68', '54'],
    ['104', '74', '48'],
    ['88', '70', '56'],
    ['96', '64', '52']
  ];

  const rows: React.ReactElement[] = [];
  for (let i: number = 0; i < shimmerProps.count; i++) {
    const titleWidth = titleWidths[i % titleWidths.length];
    const urlWidth = urlWidths[i % urlWidths.length];
    const summaryWidthSet = summaryWidths[i % summaryWidths.length];
    const metaWidthSet = metaWidths[i % metaWidths.length];

    rows.push(
      <li key={i} className={`${styles.resultCard} ${styles.shimmerResultCard}`} role="presentation">
        <div className={styles.resultIcon}>
          <Shimmer
            shimmerElements={[
              { type: ShimmerElementType.line, height: 28, width: 28 }
            ]}
            width={28}
          />
        </div>

        <div className={styles.resultBody}>
          <div className={styles.shimmerTitleRow}>
            <Shimmer
              shimmerElements={[
                { type: ShimmerElementType.line, height: 18, width: titleWidth },
                { type: ShimmerElementType.gap, width: 10 },
                { type: ShimmerElementType.line, height: 16, width: 42 }
              ]}
              width="100%"
            />
          </div>

          <div className={styles.shimmerUrlRow}>
            <Shimmer
              shimmerElements={[
                { type: ShimmerElementType.line, height: 10, width: urlWidth }
              ]}
              width="100%"
            />
          </div>

          <div className={styles.shimmerSummaryRow}>
            <Shimmer
              shimmerElements={[
                { type: ShimmerElementType.line, height: 12, width: summaryWidthSet[0] }
              ]}
              width="100%"
            />
          </div>

          <div className={styles.shimmerSummaryRow}>
            <Shimmer
              shimmerElements={[
                { type: ShimmerElementType.line, height: 12, width: summaryWidthSet[1] }
              ]}
              width="100%"
            />
          </div>

          <div className={styles.shimmerMetaRow}>
            <Shimmer
              shimmerElements={[
                { type: ShimmerElementType.circle, height: 24 },
                { type: ShimmerElementType.gap, width: 8 },
                { type: ShimmerElementType.line, height: 12, width: parseInt(metaWidthSet[0], 10) },
                { type: ShimmerElementType.gap, width: 16 },
                { type: ShimmerElementType.line, height: 12, width: parseInt(metaWidthSet[1], 10) },
                { type: ShimmerElementType.gap, width: 16 },
                { type: ShimmerElementType.line, height: 12, width: parseInt(metaWidthSet[2], 10) }
              ]}
              width="100%"
            />
          </div>
        </div>
      </li>
    );
  }
  return <ul className={`${styles.resultList} ${styles.shimmerContainer}`} aria-hidden="true">{rows}</ul>;
};

/**
 * Renders the promoted results section — highlighted "Recommended" cards above
 * the main results. Supports session-only dismissal.
 */
const PromotedResultsSection: React.FC<{ items: IPromotedResultItem[] }> = function PromotedResultsSection(sectionProps) {
  const [dismissedUrls, setDismissedUrls] = React.useState<Set<string>>(new Set());

  if (!sectionProps.items || sectionProps.items.length === 0) {
    // eslint-disable-next-line @rushstack/no-new-null
    return null;
  }

  const visibleItems = sectionProps.items.filter(function (item: IPromotedResultItem): boolean {
    return !dismissedUrls.has(item.url);
  });

  if (visibleItems.length === 0) {
    // eslint-disable-next-line @rushstack/no-new-null
    return null;
  }

  function handleDismiss(url: string): void {
    setDismissedUrls(function (prev: Set<string>): Set<string> {
      const next = new Set(prev);
      next.add(url);
      return next;
    });
  }

  return (
    <div className={styles.promotedResults} role="region" aria-label="Recommended results">
      <div className={styles.promotedHeader}>
        <Icon iconName="FavoriteStar" />
        <span>Recommended</span>
      </div>
      {visibleItems.map(function (item: IPromotedResultItem, index: number): React.ReactElement {
        return (
          <div key={item.url + '-' + String(index)} className={styles.promotedCard}>
            <div className={styles.promotedIcon}>
              {item.iconUrl ? (
                <img src={item.iconUrl} alt="" style={{ width: 20, height: 20 }} />
              ) : (
                <Icon iconName="FavoriteStar" />
              )}
            </div>
            <div className={styles.promotedContent}>
              <h3 className={styles.promotedTitle}>
                <a href={item.url} target="_blank" rel="noopener noreferrer">
                  {item.title}
                </a>
              </h3>
              {item.description && (
                <p className={styles.promotedDescription}>{item.description}</p>
              )}
            </div>
            <button
              className={styles.promotedDismiss}
              onClick={function (): void { handleDismiss(item.url); }}
              aria-label={'Dismiss promoted result: ' + item.title}
              title="Dismiss"
              type="button"
            >
              <Icon iconName="Cancel" />
            </button>
          </div>
        );
      })}
    </div>
  );
};

interface IEmptyStateProps {
  queryText: string;
  hasActiveFilters: boolean;
  /** True once a search has executed — gates the custom-HTML message off the initial page state. */
  hasSearched: boolean;
  /** Admin-supplied HTML rendered when set + a search returned zero results. Sanitized via sanitizeHtml. */
  customMessage: string;
  onClearFilters: () => void;
  onReset: () => void;
}

/**
 * Context-aware zero-results recovery panel.
 * Surfaces the most likely recovery action first based on what is active:
 * - Active filters → offer to clear them (most common cause of zero results)
 * - No filters     → suggest broader keywords
 * Always offers a hard reset to wipe all state and start fresh.
 */
const EmptyState: React.FC<IEmptyStateProps> = (emptyProps) => {
  const { queryText, hasActiveFilters, hasSearched, customMessage, onClearFilters, onReset } = emptyProps;

  // Admin-supplied HTML wins when a search returned zero results — the recovery
  // buttons stay so users still get "clear filters" / "start over".
  if (hasSearched && customMessage) {
    return (
      <div className={styles.emptyState} role="status">
        <div
          className={styles.emptyCustom}
          dangerouslySetInnerHTML={{ __html: sanitizeHtml(customMessage) }}
        />
        {hasActiveFilters && (
          <div className={styles.emptyRecovery}>
            <button className={styles.emptyRecoveryButton} onClick={onClearFilters} type="button">
              Clear all filters
            </button>
          </div>
        )}
        {(queryText || hasActiveFilters) && (
          <button className={styles.emptyResetLink} onClick={onReset} type="button">
            Start over
          </button>
        )}
      </div>
    );
  }

  // Context-aware empty state messaging
  let title: React.ReactNode;
  let description: string;

  if (queryText && hasActiveFilters) {
    title = <>No results for <span className={styles.emptyQuery}>&#x201C;{queryText}&#x201D;</span></>;
    description = 'No results match your search and filters.';
  } else if (queryText && !hasActiveFilters) {
    title = <>No results for <span className={styles.emptyQuery}>&#x201C;{queryText}&#x201D;</span></>;
    description = 'No results found. Check your spelling or try broader search terms.';
  } else if (!queryText && hasActiveFilters) {
    title = 'No results found';
    description = 'Your filters might be too specific.';
  } else {
    title = 'Search';
    description = 'Enter a search term to get started.';
  }

  // T1.D5 — neutral icon. `SearchIssue` reads as a warning/error (warning
  // triangle over a magnifying glass); using it for the "no results" /
  // "enter a query" state miscommunicates that something went wrong.
  // `Search` is the iconographic match for both idle ("start typing") and
  // no-results ("we looked, nothing found") states.
  const emptyIconName: string = hasSearched ? 'SearchAndApps' : 'Search';

  return (
    <div className={styles.emptyState} role="status">
      <div className={styles.emptyIcon}>
        <Icon iconName={emptyIconName} />
      </div>
      <h3 className={styles.emptyTitle}>{title}</h3>
      <p className={styles.emptyDescription}>{description}</p>
      {hasActiveFilters && (
        <div className={styles.emptyRecovery}>
          <button className={styles.emptyRecoveryButton} onClick={onClearFilters} type="button">
            Clear all filters
          </button>
        </div>
      )}
      {(queryText || hasActiveFilters) && (
        <button className={styles.emptyResetLink} onClick={onReset} type="button">
          Start over
        </button>
      )}
    </div>
  );
};

/** Delay (ms) before showing the loading overlay. Sub-threshold searches complete
 *  without ever displaying the spinner, avoiding a distracting flash. */
const LOADING_OVERLAY_DELAY_MS = 300;

/**
 * SPSearchResults — main container component.
 * Subscribes to the shared Zustand store and orchestrates the display of
 * promoted results, toolbar, layout, and pagination.
 */
const SpSearchResults: React.FC<ISpSearchResultsProps> = (props) => {
  const {
    store,
    orchestrator,
    searchContextId,
    showResultCount,
    showSortDropdown,
    showDeleteConfirmation,
    enablePreviewPanel,
    hideWebPartWhenNoResults,
    emptyResultsMessage,
    titleDisplayMode,
    isEditMode,
    defaultLayout,
    selectedPropertyColumns,
    gridPropertyColumns,
    compactPropertyColumns,
    showColumnChooser,
    queryTemplate,
    graphOrgService,
    linkConfig
  } = props;

  const {
    items,
    totalCount,
    currentPage,
    pageSize,
    isLoading,
    hasSearched,
    error,
    queryText,
    activeLayoutKey,
    availableLayouts,
    promotedResults,
    sort,
    sortableProperties,
    previewPanel,
    activeFilters,
    filterConfig,
    querySuggestion,
    showPaging,
    pageRange,
    scope
  } = useStoreState(props);

  // T5.D1 — DebugFab/Panel now mounted via shared DebugFabHost.

  const effectiveDefaultLayout = React.useMemo((): string => {
    const configured = defaultLayout || 'list';
    return availableLayouts.indexOf(configured) >= 0 ? configured : (availableLayouts[0] || 'list');
  }, [availableLayouts, defaultLayout]);

  // Sync store when URL-provided layout is not in available layouts
  React.useEffect((): void => {
    if (activeLayoutKey && availableLayouts.length > 0 && availableLayouts.indexOf(activeLayoutKey) < 0) {
      store.getState().setLayout(effectiveDefaultLayout);
    }
  }, [activeLayoutKey, availableLayouts, effectiveDefaultLayout, store]);

  // ─── Store action callbacks ─────────────────────────────────
  // T1.D11 — preserve scroll position across layout swaps. Capture
  // `window.scrollY` before the layout change fires, then restore it on
  // the next animation frame after the new layout has painted. All
  // layouts render in the same parent so the same scroll offset lands
  // the user at roughly the same row in the new view.
  const handleLayoutChange = React.useCallback((key: string): void => {
    const savedScrollY = typeof window !== 'undefined' ? window.scrollY : 0;
    store.getState().setLayout(key);
    if (typeof window !== 'undefined' && savedScrollY > 0) {
      // Two RAFs — first lets React commit the new layout DOM, second
      // lets the new layout's first paint complete so heights are real.
      window.requestAnimationFrame((): void => {
        window.requestAnimationFrame((): void => {
          window.scrollTo({ top: savedScrollY, behavior: 'auto' });
        });
      });
    }
  }, [store]);

  const handlePreloadLayout = React.useCallback((key: string): void => {
    const preload = LAYOUT_PRELOADERS[key];
    if (preload) { preload(); }
  }, []);

  const handleSortChange = React.useCallback((newSort: ISortField): void => {
    store.getState().setSort(newSort);
  }, [store]);

  const resultsContainerRef = React.useRef<HTMLDivElement>(null);

  const handlePageChange = React.useCallback((page: number): void => {
    store.getState().setPage(page);
    // Scroll to top of results container for better UX
    if (resultsContainerRef.current) {
      resultsContainerRef.current.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }
  }, [store]);

  // ─── Error dismissal ───────────────────────────────────────
  const handleDismissError = React.useCallback((): void => {
    store.getState().setError(undefined);
  }, [store]);

  // ─── Detail panel handlers ─────────────────────────────────
  const handlePreviewItem = React.useCallback((item: ISearchResult): void => {
    store.getState().setPreviewItem(item);
  }, [store]);

  // ─── Click tracking ───────────────────────────────────────
  const handleItemClick = React.useCallback((item: ISearchResult, position: number): void => {
    if (orchestrator) {
      orchestrator.logClickedItem(item.url, item.title, position);
    }
  }, [orchestrator]);

  // Page-boundary Next-from-detail-panel handoff. When the user clicks
  // Next on the last item of a page, we trigger a page load and set
  // this ref; an effect below watches for the new page's items and
  // jumps the panel forward to items[0] so the click feels responsive.
  const pendingNextPageItemRef = React.useRef<boolean>(false);
  const lastSeenPageRef = React.useRef<number>(currentPage);
  React.useEffect((): void => {
    if (lastSeenPageRef.current !== currentPage) {
      lastSeenPageRef.current = currentPage;
      if (pendingNextPageItemRef.current && items.length > 0) {
        store.getState().setPreviewItem(items[0]);
      }
      pendingNextPageItemRef.current = false;
    }
  }, [currentPage, items, store]);

  const handleDismissPreviewPanel = React.useCallback((): void => {
    store.getState().setPreviewItem(undefined);
    // Clear the page-boundary handoff so dismissing during the load
    // doesn't re-open the panel when the new page arrives.
    pendingNextPageItemRef.current = false;
  }, [store]);

  // ─── Filter pill bar handlers ──────────────────────────────
  const handleRemoveFilter = React.useCallback(function (filterName: string): void {
    store.getState().removeRefiner(filterName);
  }, [store]);

  const handleClearAllFilters = React.useCallback(function (): void {
    store.getState().clearAllFilters();
  }, [store]);

  // ─── Reset handler — navigate to base page without params ─
  const handleReset = React.useCallback(function (): void {
    // safe: read-only URL capture for serialization (Found.D4 exempt).
    const url = new URL(window.location.href);
    url.search = '';
    safeNavigate(url.toString());
  }, []);

  // ─── Loading overlay (delayed to avoid flash for fast searches) ───────────
  const [showOverlay, setShowOverlay] = React.useState(false);
  React.useEffect((): (() => void) => {
    if (!isLoading || items.length === 0) {
      setShowOverlay(false);
      return (): void => { /* noop */ };
    }
    // Only reveal the overlay after LOADING_OVERLAY_DELAY_MS — sub-threshold
    // searches complete before it ever appears, so the user sees no flash.
    const timer = setTimeout((): void => { setShowOverlay(true); }, LOADING_OVERLAY_DELAY_MS);
    return (): void => { clearTimeout(timer); };
  }, [isLoading, items.length]);

  // ─── "Did you mean" handler ────────────────────────────────
  const handleQuerySuggestionClick = React.useCallback(function (): void {
    if (querySuggestion) {
      store.getState().setQueryText(querySuggestion);
    }
  }, [store, querySuggestion]);

  // ─── Determine which layout to render ──────────────────────
  const renderLayout = (): React.ReactElement | undefined => {
    // T1.D4 — fresh page (no search triggered yet) renders the idle
    // EmptyState ("Enter a search term to get started."), not a shimmer.
    // Shimmer is reserved for actually-loading states. If a search IS in
    // flight on first render (e.g. auto-search from URL state), the next
    // branch catches it.
    if (!hasSearched && !isLoading) {
      return (
        <EmptyState
          queryText={queryText}
          hasActiveFilters={activeFilters.length > 0}
          hasSearched={false}
          customMessage={emptyResultsMessage}
          onClearFilters={handleClearAllFilters}
          onReset={handleReset}
        />
      );
    }

    // First search in flight (or refresh with no previous results to retain)
    // — show the skeleton rather than flickering to empty state.
    if (isLoading && items.length === 0) {
      return <LoadingShimmer count={5} />;
    }

    // Search completed with no results.
    if (items.length === 0) {
      return (
        <EmptyState
          queryText={queryText}
          hasActiveFilters={activeFilters.length > 0}
          hasSearched={hasSearched}
          customMessage={emptyResultsMessage}
          onClearFilters={handleClearAllFilters}
          onReset={handleReset}
        />
      );
    }

    // Build the layout content for the current items.
    let layoutContent: React.ReactElement;
    switch (activeLayoutKey) {
      case 'compact':
        layoutContent = (
          <CompactLayout
            items={items}
            searchContextId={searchContextId}
            compactPropertyColumns={compactPropertyColumns}
            titleDisplayMode={titleDisplayMode}
            onItemClick={handleItemClick}
            linkConfig={linkConfig}
            onOpenInSidePanel={handlePreviewItem}
          />
        );
        break;

      case 'card':
        layoutContent = (
          <CardLayout
            items={items}
            searchContextId={searchContextId}
            titleDisplayMode={titleDisplayMode}
            onPreviewItem={enablePreviewPanel ? handlePreviewItem : undefined}
            onItemClick={handleItemClick}
            linkConfig={linkConfig}
            onOpenInSidePanel={handlePreviewItem}
          />
        );
        break;

      case 'people':
        layoutContent = (
          <PeopleLayout
            items={items}
            onPreviewItem={enablePreviewPanel ? handlePreviewItem : undefined}
            onItemClick={handleItemClick}
            graphOrgService={graphOrgService}
          />
        );
        break;

      case 'grid':
        layoutContent = (
          <DataGridLayout
            items={items}
            searchContextId={searchContextId}
            gridPropertyColumns={gridPropertyColumns}
            titleDisplayMode={titleDisplayMode}
            totalCount={totalCount}
            pageSize={pageSize}
            currentPage={currentPage}
            showPaging={showPaging}
            pageRange={pageRange}
            showDeleteConfirmation={showDeleteConfirmation}
            showColumnChooser={showColumnChooser}
            sort={sort}
            sortableProperties={sortableProperties}
            onPreviewItem={enablePreviewPanel ? handlePreviewItem : undefined}
            onItemClick={handleItemClick}
            onPageChange={handlePageChange}
            onSortChange={handleSortChange}
            onFallback={(): void => handleLayoutChange('list')}
            linkConfig={linkConfig}
            onOpenInSidePanel={handlePreviewItem}
          />
        );
        break;

      case 'gallery':
        layoutContent = (
          <GalleryLayout
            items={items}
            titleDisplayMode={titleDisplayMode}
            onPreviewItem={enablePreviewPanel ? handlePreviewItem : undefined}
            onItemClick={handleItemClick}
            linkConfig={linkConfig}
            onOpenInSidePanel={handlePreviewItem}
          />
        );
        break;

      default:
        // 'list' is the default layout
        layoutContent = (
          <ListLayout
            items={items}
            searchContextId={searchContextId}
            scope={scope}
            titleDisplayMode={titleDisplayMode}
            onItemClick={handleItemClick}
            linkConfig={linkConfig}
            onOpenInSidePanel={handlePreviewItem}
          />
        );
    }

    return layoutContent;
  };

  if (hideWebPartWhenNoResults && !isEditMode && hasSearched && !isLoading && !error && items.length === 0) {
    return null;
  }

  return (
    <ErrorBoundary enableRetry={true} maxRetries={3}>
      <div ref={resultsContainerRef} className={styles.spSearchResults}>
        {/* Admin diagnostic notices — edit mode only */}
        {isEditMode && searchContextId === 'default' && (
          <MessageBar messageBarType={MessageBarType.info} isMultiline={true} styles={{ root: { marginBottom: 8 } }}>
            Using the <strong>default</strong> search context. Set a unique Search Context ID in the property pane
            when using multiple independent search experiences on the same page.
          </MessageBar>
        )}
        {/* T3.D10 — init-order diagnostic. Renders when Filters web part
            registered AFTER the first search ran with empty filterConfig
            (URL-deep-linked filter values silently failed to apply). The
            Retry button re-runs the search now that filterConfig is
            populated. View mode hides this entirely. */}
        {isEditMode && hasInitOrderIssue(searchContextId) && (
          <MessageBar
            messageBarType={MessageBarType.warning}
            isMultiline={true}
            styles={{ root: { marginBottom: 8 } }}
            actions={(
              <div>
                <button
                  type="button"
                  onClick={(): void => {
                    clearInitOrderDiagnostic(searchContextId);
                    // Re-fire the orchestrator's search with the now-loaded filterConfig.
                    if (orchestrator) {
                      orchestrator.triggerSearch().catch(function noop(): void { /* handled in orchestrator */ });
                    } else {
                      store.getState().setError('Search is not ready. Reload the page and try again.');
                    }
                  }}
                  style={{ padding: '4px 10px', cursor: 'pointer' }}
                >
                  Retry
                </button>
              </div>
            // eslint-disable-next-line @typescript-eslint/no-explicit-any
            ) as any}
          >
            First search ran before the Filters web part loaded — URL filters may not have been applied.
            Click Retry to re-run the search now that the filter configuration is loaded, or reload the page.
          </MessageBar>
        )}
        {isEditMode && ((): React.ReactNode => {
          const configWarnings: IConfigWarning[] = validateWebPartConfig({
            defaultLayout: effectiveDefaultLayout,
            availableLayouts,
            selectedPropertyColumns: selectedPropertyColumns || [],
            gridPropertyColumns: gridPropertyColumns || [],
            queryTemplate: queryTemplate || '{searchTerms}',
          });
          if (configWarnings.length === 0) { return null; }
          const levelToMessageBarType = (level: ConfigWarningLevel): MessageBarType => {
            if (level === 'error')   { return MessageBarType.error; }
            if (level === 'warning') { return MessageBarType.warning; }
            return MessageBarType.info;
          };
          return (
            <>
              {configWarnings.map((w) => (
                <MessageBar
                  key={w.id}
                  messageBarType={levelToMessageBarType(w.level)}
                  isMultiline={true}
                  styles={{ root: { marginBottom: 4 } }}
                >
                  {w.message}
                </MessageBar>
              ))}
            </>
          );
        })()}

        {/* Error message bar */}
        {error && (
          <div className={styles.errorContainer} role="alert">
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={false}
              onDismiss={handleDismissError}
              dismissButtonAriaLabel="Dismiss"
            >
              {error}
            </MessageBar>
          </div>
        )}

        {/* "Did you mean" suggestion */}
        {querySuggestion && !isLoading && (
          <div className={styles.querySuggestion} role="status">
            <Icon iconName="Lightbulb" className={styles.querySuggestionIcon} />
            <span>Did you mean: </span>
            <button
              className={styles.querySuggestionLink}
              onClick={handleQuerySuggestionClick}
              type="button"
            >
              {querySuggestion}
            </button>
          </div>
        )}

        {/* Promoted results */}
        <PromotedResultsSection items={promotedResults} />

        {/* Result toolbar — count, sort, layout toggle */}
        {(items.length > 0 || isLoading) && (
          <ResultToolbar
            totalCount={totalCount}
            activeLayoutKey={activeLayoutKey}
            availableLayouts={availableLayouts}
            sort={sort}
            sortableProperties={sortableProperties}
            showResultCount={showResultCount}
            showSortDropdown={showSortDropdown && ['list', 'compact', 'grid'].indexOf(activeLayoutKey) >= 0}
            onLayoutChange={handleLayoutChange}
            onSortChange={handleSortChange}
            onPreloadLayout={handlePreloadLayout}
            onExportCsv={(): void => {
              // Bulk selection retired — export the full current page.
              exportItemsAsCsv(items, {
                configuredColumns: selectedPropertyColumns,
              });
            }}
            onExportXlsx={(): void => {
              exportItemsAsXlsx(items, {
                configuredColumns: selectedPropertyColumns,
              }).catch((err): void => { spLog.error('XLSX export failed', { error: err }); });
            }}
          />
        )}

        {/* Active filter pills */}
        <ActiveFilterPillBar
          activeFilters={activeFilters}
          filterConfig={filterConfig}
          onRemoveFilter={handleRemoveFilter}
          onClearAll={handleClearAllFilters}
        />

        {/* Active layout — wrapped to provide overlay positioning context */}
        <div className={styles.resultsWrapper}>
          {renderLayout()}
          {showOverlay && (
            <div className={styles.loadingOverlay} role="status" aria-busy="true" aria-label="Loading results">
              <Spinner size={SpinnerSize.medium} label="Loading..." ariaLive="assertive" />
            </div>
          )}
        </div>

        {/* Pagination */}
        {items.length > 0 && !isLoading && (
          <Pagination
            currentPage={currentPage}
            totalCount={totalCount}
            pageSize={pageSize}
            showPaging={showPaging}
            pageRange={pageRange}
            onPageChange={handlePageChange}
          />
        )}

        {/* T2.D7 — Detail panel with next/previous navigation. Items list +
            current index are passed so the panel can render arrow buttons
            and listen for Alt+Left/Right. When the user advances past the
            last on-page item, the parent triggers a page load via
            `handleNextPage` (existing pager). */}
        {enablePreviewPanel && previewPanel.isOpen && ((): React.ReactNode => {
          const currentIndex = previewPanel.item
            ? items.findIndex((it) => it.key === previewPanel.item!.key)
            : -1;
          const handleNavigate = (delta: number): void => {
            if (currentIndex < 0) { return; }
            const nextIndex = currentIndex + delta;
            if (nextIndex >= 0 && nextIndex < items.length) {
              store.getState().setPreviewItem(items[nextIndex]);
              return;
            }
            // Forward past the last item — if a next page exists, page +1
            // and arm the page-boundary handoff so the panel auto-advances
            // to items[0] of the new page once it loads. (Pager bounds
            // check upstream.)
            if (delta > 0 && currentPage * pageSize < totalCount) {
              pendingNextPageItemRef.current = true;
              handlePageChange(currentPage + 1);
            }
          };
          return (
            <ResultDetailPanel
              isOpen={previewPanel.isOpen}
              item={previewPanel.item}
              onDismiss={handleDismissPreviewPanel}
              currentIndex={currentIndex}
              totalOnPage={items.length}
              hasNextPage={currentPage * pageSize < totalCount}
              onNavigate={handleNavigate}
            />
          );
        })()}
      </div>
        {/* T5.D1 — singleton DebugFab + Panel. Host gates internally on
            DebugCollector.isActive() AND the cross-bundle owner claim so
            multi-web-part pages render exactly one FAB. */}
        <DebugFabHost store={store} />
        {/* T2.D9 — singleton shortcut help modal host (cross-bundle owner claim). */}
        <ShortcutHelpModalHost />
    </ErrorBoundary>
  );
};

export default SpSearchResults;
