import * as React from 'react';
import { Shimmer, ShimmerElementType } from '@fluentui/react/lib/Shimmer';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { ErrorBoundary } from 'spfx-toolkit/lib/components/ErrorBoundary';
import { createLazyComponent } from 'spfx-toolkit/lib/utilities/lazyLoader';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import type { ISpSearchResultsProps } from './ISpSearchResultsProps';
import {
  ISearchResult,
  IPromotedResultItem,
  ISortField,
  IActiveFilter,
  IFilterConfig,
  ISearchContext,
  IActionProvider
} from '@interfaces/index';
import { getManagerService } from '@store/store';
import ResultToolbar from './ResultToolbar';
import BulkActionsToolbar from './BulkActionsToolbar';
import ActiveFilterPillBar from './ActiveFilterPillBar';
import ListLayout from './ListLayout';
import CompactLayout from './CompactLayout';
import Pagination from './Pagination';
import styles from './SpSearchResults.module.scss';

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
const DataGridLayout: any = createLazyComponent(
  () => import(/* webpackChunkName: 'DataGridLayout' */ './DataGridLayout') as any,
  { errorMessage: 'Failed to load data grid layout' }
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
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const SpSearchManager: any = createLazyComponent(
  () => import(/* webpackChunkName: 'SearchManager' */ '@webparts/spSearchManager/components/SpSearchManager') as any,
  { errorMessage: 'Failed to load search manager' }
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
  error: string | undefined;
  activeLayoutKey: string;
  promotedResults: IPromotedResultItem[];
  sort: ISortField | undefined;
  bulkSelection: string[];
  previewPanel: { isOpen: boolean; item: ISearchResult | undefined };
  isSearchManagerOpen: boolean;
  activeFilters: IActiveFilter[];
  filterConfig: IFilterConfig[];
} {
  const { store } = props;

  const getSnapshot = React.useCallback((): {
    items: ISearchResult[];
    totalCount: number;
    currentPage: number;
    pageSize: number;
    isLoading: boolean;
    error: string | undefined;
    activeLayoutKey: string;
    promotedResults: IPromotedResultItem[];
    sort: ISortField | undefined;
    bulkSelection: string[];
    previewPanel: { isOpen: boolean; item: ISearchResult | undefined };
    isSearchManagerOpen: boolean;
    activeFilters: IActiveFilter[];
    filterConfig: IFilterConfig[];
  } => {
    const state = store.getState();
    return {
      items: state.items,
      totalCount: state.totalCount,
      currentPage: state.currentPage,
      pageSize: state.pageSize,
      isLoading: state.isLoading,
      error: state.error,
      activeLayoutKey: state.activeLayoutKey,
      promotedResults: state.promotedResults,
      sort: state.sort,
      bulkSelection: state.bulkSelection,
      previewPanel: state.previewPanel,
      isSearchManagerOpen: state.isSearchManagerOpen,
      activeFilters: state.activeFilters,
      filterConfig: state.filterConfig
    };
  }, [store]);

  const [storeState, setStoreState] = React.useState(getSnapshot);

  React.useEffect((): (() => void) => {
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
          prev.error === next.error &&
          prev.activeLayoutKey === next.activeLayoutKey &&
          prev.promotedResults === next.promotedResults &&
          prev.sort === next.sort &&
          prev.bulkSelection === next.bulkSelection &&
          prev.previewPanel === next.previewPanel &&
          prev.isSearchManagerOpen === next.isSearchManagerOpen &&
          prev.activeFilters === next.activeFilters &&
          prev.filterConfig === next.filterConfig
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
  const rows: React.ReactElement[] = [];
  for (let i: number = 0; i < shimmerProps.count; i++) {
    rows.push(
      <div key={i} className={styles.shimmerRow}>
        <Shimmer
          shimmerElements={[
            { type: ShimmerElementType.circle, height: 32 },
            { type: ShimmerElementType.gap, width: 12 },
            { type: ShimmerElementType.line, height: 16, width: '40%' }
          ]}
          width="100%"
        />
      </div>
    );
  }
  return <div className={styles.shimmerContainer}>{rows}</div>;
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

/**
 * Renders the empty state when no results are found.
 */
const EmptyState: React.FC = () => (
  <div className={styles.emptyState} role="status">
    <div className={styles.emptyIcon}>
      <Icon iconName="SearchIssue" />
    </div>
    <h3 className={styles.emptyTitle}>No results found</h3>
    <p className={styles.emptyDescription}>
      Try adjusting your search query or filters to find what you are looking for.
    </p>
  </div>
);

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
    theme,
    showResultCount,
    showSortDropdown,
    enableSelection
  } = props;

  const {
    items,
    totalCount,
    currentPage,
    pageSize,
    isLoading,
    error,
    activeLayoutKey,
    promotedResults,
    sort,
    bulkSelection,
    previewPanel,
    isSearchManagerOpen,
    activeFilters,
    filterConfig
  } = useStoreState(props);

  // Get the manager service for the Search Manager panel
  const managerService = getManagerService(searchContextId);

  // ─── Store action callbacks ─────────────────────────────────
  const handleLayoutChange = React.useCallback((key: string): void => {
    store.getState().setLayout(key);
  }, [store]);

  const handleSortChange = React.useCallback((newSort: ISortField): void => {
    store.getState().setSort(newSort);
  }, [store]);

  const handlePageChange = React.useCallback((page: number): void => {
    store.getState().setPage(page);
  }, [store]);

  const handleToggleSelection = React.useCallback((key: string, multiSelect: boolean): void => {
    store.getState().toggleSelection(key, multiSelect);
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

  const handleClearSelection = React.useCallback(function (): void {
    store.getState().clearSelection();
  }, [store]);

  // ─── Search Manager panel dismiss ─────────────────────────
  const handleDismissSearchManager = React.useCallback((): void => {
    store.getState().toggleSearchManager();
  }, [store]);

  const selectedItems = React.useMemo((): ISearchResult[] => {
    if (!bulkSelection || bulkSelection.length === 0) {
      return [];
    }
    const selected = new Set(bulkSelection);
    return items.filter((item) => selected.has(item.key));
  }, [bulkSelection, items]);

  const bulkActions = React.useMemo((): IActionProvider[] => {
    return store.getState().registries.actions.getAll();
  }, [store]);

  const actionContext = React.useMemo((): ISearchContext => ({
    searchContextId,
    siteUrl: SPContext.webAbsoluteUrl || '',
    scope: store.getState().scope,
  }), [searchContextId, store]);

  // ─── Determine which layout to render ──────────────────────
  const renderLayout = (): React.ReactElement | undefined => {
    if (isLoading) {
      return <LoadingShimmer count={5} />;
    }

    if (items.length === 0) {
      return <EmptyState />;
    }

    switch (activeLayoutKey) {
      case 'compact':
        return <CompactLayout items={items} onItemClick={handleItemClick} />;

      case 'card':
        return (
          <CardLayout items={items} onPreviewItem={handlePreviewItem} onItemClick={handleItemClick} />
        );

      case 'people':
        return (
          <PeopleLayout items={items} onPreviewItem={handlePreviewItem} onItemClick={handleItemClick} />
        );

      case 'datagrid':
        return (
          <DataGridLayout
            items={items}
            enableSelection={enableSelection}
            selectedKeys={bulkSelection}
            onToggleSelection={handleToggleSelection}
            onPreviewItem={handlePreviewItem}
            onItemClick={handleItemClick}
          />
        );

      case 'gallery':
        return (
          <GalleryLayout items={items} onPreviewItem={handlePreviewItem} onItemClick={handleItemClick} />
        );

      default:
        // 'list' layout is the default
        return (
          <ListLayout
            items={items}
            enableSelection={enableSelection}
            selectedKeys={bulkSelection}
            onToggleSelection={handleToggleSelection}
            onItemClick={handleItemClick}
          />
        );
    }
  };

  return (
    <ErrorBoundary enableRetry={true} maxRetries={3}>
      <div className={styles.spSearchResults}>
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

        {/* Promoted results */}
        <PromotedResultsSection items={promotedResults} />

        {/* Result toolbar — count, sort, layout toggle */}
        {(items.length > 0 || isLoading) && (
          <ResultToolbar
            totalCount={totalCount}
            activeLayoutKey={activeLayoutKey}
            sort={sort}
            showResultCount={showResultCount}
            showSortDropdown={showSortDropdown}
            onLayoutChange={handleLayoutChange}
            onSortChange={handleSortChange}
          />
        )}

        {enableSelection && bulkSelection.length > 0 && (
          <BulkActionsToolbar
            selectedItems={selectedItems}
            actions={bulkActions}
            context={actionContext}
            onClearSelection={handleClearSelection}
          />
        )}

        {/* Active filter pills */}
        <ActiveFilterPillBar
          activeFilters={activeFilters}
          filterConfig={filterConfig}
          onRemoveFilter={handleRemoveFilter}
          onClearAll={handleClearAllFilters}
        />

        {/* Active layout */}
        {renderLayout()}

        {/* Pagination */}
        {items.length > 0 && !isLoading && (
          <Pagination
            currentPage={currentPage}
            totalCount={totalCount}
            pageSize={pageSize}
            onPageChange={handlePageChange}
          />
        )}

        {/* Detail panel — lazy-loaded, only renders when preview panel is open */}
        {previewPanel.isOpen && (
          <ResultDetailPanel
            isOpen={previewPanel.isOpen}
            item={previewPanel.item}
            onDismiss={handleDismissPreviewPanel}
          />
        )}

        {/* Search Manager panel — lazy-loaded, only renders when manager panel is open */}
        {isSearchManagerOpen && managerService && (
          <Panel
            isOpen={isSearchManagerOpen}
            onDismiss={handleDismissSearchManager}
            type={PanelType.medium}
            headerText="Search Manager"
            closeButtonAriaLabel="Close"
            isLightDismiss={true}
          >
            <SpSearchManager
              store={store}
              service={managerService}
              theme={theme}
              mode="panel"
            />
          </Panel>
        )}
      </div>
    </ErrorBoundary>
  );
};

export default SpSearchResults;
