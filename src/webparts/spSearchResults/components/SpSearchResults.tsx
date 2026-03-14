import * as React from 'react';
import { Shimmer, ShimmerElementType } from '@fluentui/react/lib/Shimmer';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';
import { ErrorBoundary } from 'spfx-toolkit/lib/components/ErrorBoundary';
import { createLazyComponent } from 'spfx-toolkit/lib/utilities/lazyLoader';
import type { ISpSearchResultsProps } from './ISpSearchResultsProps';
import {
  ISearchResult,
  IPromotedResultItem,
  ISortField,
  ISortableProperty,
  IActiveFilter,
  IFilterConfig
} from '@interfaces/index';
import ResultToolbar from './ResultToolbar';
import ActiveFilterPillBar from './ActiveFilterPillBar';
import ListLayout from './ListLayout';
import CompactLayout from './CompactLayout';
import DataGridLayout from './DataGridLayout';
import Pagination from './Pagination';
import styles from './SpSearchResults.module.scss';
import { validateWebPartConfig, IConfigWarning, ConfigWarningLevel } from './configValidation';

// ─── Lazy layout preloaders ──────────────────────────────
// Called on hover of the corresponding toolbar button to warm the webpack
// chunk before the user actually clicks. No-ops on subsequent calls because
// webpack deduplicates dynamic imports after the first load.
const LAYOUT_PRELOADERS: Record<string, () => void> = {
  card:    (): void => { import(/* webpackChunkName: 'CardLayout' */    './CardLayout')    .catch((): void => { /* ignore preload error */ }); },
  people:  (): void => { import(/* webpackChunkName: 'PeopleLayout' */  './PeopleLayout')  .catch((): void => { /* ignore preload error */ }); },
  grid:    (): void => { /* DataGridLayout is bundled eagerly to avoid runtime chunk-load failures. */ },
  gallery: (): void => { import(/* webpackChunkName: 'GalleryLayout' */ './GalleryLayout') .catch((): void => { /* ignore preload error */ }); },
};

// ─── Lazy-loaded layouts and panels ──────────────────────
// Type assertions needed due to @types/react mismatch between sp-search and spfx-toolkit
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const CardLayout: any = createLazyComponent(
  () => import(/* webpackChunkName: 'CardLayout' */ './CardLayout') as any,
  { errorMessage: 'Failed to load card layout' }
);
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const PeopleLayout: any = createLazyComponent(
  () => import(/* webpackChunkName: 'PeopleLayout' */ './PeopleLayout') as any,
  { errorMessage: 'Failed to load people layout' }
);
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const GalleryLayout: any = createLazyComponent(
  () => import(/* webpackChunkName: 'GalleryLayout' */ './GalleryLayout') as any,
  { errorMessage: 'Failed to load gallery layout' }
);
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const ResultDetailPanel: any = createLazyComponent(
  () => import(/* webpackChunkName: 'ResultDetailPanel' */ './ResultDetailPanel') as any,
  { errorMessage: 'Failed to load detail panel' }
);
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
    pageRange: 5
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
      pageRange: state.pageRange
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
          prev.pageRange === next.pageRange
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
  const { queryText, hasActiveFilters, onClearFilters, onReset } = emptyProps;
  return (
    <div className={styles.emptyState} role="status">
      <div className={styles.emptyIcon}>
        <Icon iconName="SearchIssue" />
      </div>
      <h3 className={styles.emptyTitle}>
        {queryText
          ? <>No results for <span className={styles.emptyQuery}>&#x201C;{queryText}&#x201D;</span></>
          : 'No results found'}
      </h3>
      {hasActiveFilters ? (
        <div className={styles.emptyRecovery}>
          <p className={styles.emptyDescription}>
            Your active filters may be narrowing results too much.
          </p>
          <button className={styles.emptyRecoveryButton} onClick={onClearFilters} type="button">
            Clear all filters
          </button>
        </div>
      ) : (
        <p className={styles.emptyDescription}>
          Check your spelling or try broader search terms.
        </p>
      )}
      <button className={styles.emptyResetLink} onClick={onReset} type="button">
        Start over
      </button>
    </div>
  );
};

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
    titleDisplayMode,
    isEditMode,
    defaultLayout,
    selectedPropertyColumns,
    gridPropertyColumns,
    compactPropertyColumns,
    queryTemplate,
    graphOrgService
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
    pageRange
  } = useStoreState(props);

  const effectiveDefaultLayout = React.useMemo((): string => {
    const configured = defaultLayout || 'list';
    return availableLayouts.indexOf(configured) >= 0 ? configured : (availableLayouts[0] || 'list');
  }, [availableLayouts, defaultLayout]);

  // ─── Store action callbacks ─────────────────────────────────
  const handleLayoutChange = React.useCallback((key: string): void => {
    store.getState().setLayout(key);
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

  const handleDismissPreviewPanel = React.useCallback((): void => {
    store.getState().setPreviewItem(undefined);
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
    const url = new URL(window.location.href);
    url.search = '';
    window.location.href = url.toString();
  }, []);

  // ─── Loading overlay (delayed to avoid flash for fast searches) ───────────
  const [showOverlay, setShowOverlay] = React.useState(false);
  React.useEffect((): (() => void) => {
    if (!isLoading || items.length === 0) {
      setShowOverlay(false);
      return (): void => { /* noop */ };
    }
    // Only reveal the overlay after 300ms — sub-300ms searches complete before
    // it ever appears, so the user sees no flash.
    const timer = setTimeout((): void => { setShowOverlay(true); }, 300);
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
    // No search has completed yet — show skeleton, never "No results found".
    // This covers both the very first page load and the case where the store
    // initializes before the Results web part has registered a data provider.
    if (!hasSearched) {
      return <LoadingShimmer count={5} />;
    }

    // A refresh search is in progress with no previous results to retain.
    // Show the skeleton again rather than flickering to empty state.
    if (isLoading && items.length === 0) {
      return <LoadingShimmer count={5} />;
    }

    // Search completed with no results.
    if (items.length === 0) {
      return (
        <EmptyState
          queryText={queryText}
          hasActiveFilters={activeFilters.length > 0}
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
            sort={sort}
            sortableProperties={sortableProperties}
            onPreviewItem={enablePreviewPanel ? handlePreviewItem : undefined}
            onItemClick={handleItemClick}
            onPageChange={handlePageChange}
            onSortChange={handleSortChange}
            onFallback={(): void => handleLayoutChange('list')}
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
          />
        );
        break;

      default:
        // 'list' is the default layout
        layoutContent = (
          <ListLayout
            items={items}
            searchContextId={searchContextId}
            titleDisplayMode={titleDisplayMode}
            onItemClick={handleItemClick}
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
            showSortDropdown={showSortDropdown}
            onLayoutChange={handleLayoutChange}
            onSortChange={handleSortChange}
            onPreloadLayout={handlePreloadLayout}
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

        {/* Detail panel — lazy-loaded, only renders when preview panel is open */}
        {enablePreviewPanel && previewPanel.isOpen && (
          <ResultDetailPanel
            isOpen={previewPanel.isOpen}
            item={previewPanel.item}
            onDismiss={handleDismissPreviewPanel}
          />
        )}
      </div>
    </ErrorBoundary>
  );
};

export default SpSearchResults;
