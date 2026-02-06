import * as React from 'react';
import { Shimmer, ShimmerElementType } from '@fluentui/react/lib/Shimmer';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { ErrorBoundary } from 'spfx-toolkit/lib/components/ErrorBoundary';
import type { ISpSearchResultsProps } from './ISpSearchResultsProps';
import {
  ISearchResult,
  IPromotedResultItem,
  ISortField
} from '@interfaces/index';
import { getManagerService } from '@store/store';
import ResultToolbar from './ResultToolbar';
import ListLayout from './ListLayout';
import CompactLayout from './CompactLayout';
import Pagination from './Pagination';
import styles from './SpSearchResults.module.scss';

// ─── Lazy-loaded layouts and panels ──────────────────────
const CardLayout = React.lazy(() => import(/* webpackChunkName: 'CardLayout' */ './CardLayout'));
const PeopleLayout = React.lazy(() => import(/* webpackChunkName: 'PeopleLayout' */ './PeopleLayout'));
const DataGridLayout = React.lazy(() => import(/* webpackChunkName: 'DataGridLayout' */ './DataGridLayout'));
const GalleryLayout = React.lazy(() => import(/* webpackChunkName: 'GalleryLayout' */ './GalleryLayout'));
const ResultDetailPanel = React.lazy(() => import(/* webpackChunkName: 'ResultDetailPanel' */ './ResultDetailPanel'));
const SpSearchManager = React.lazy(() => import(/* webpackChunkName: 'SearchManager' */ '@webparts/spSearchManager/components/SpSearchManager'));

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
      isSearchManagerOpen: state.isSearchManagerOpen
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
          prev.isSearchManagerOpen === next.isSearchManagerOpen
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
 * Suspense fallback shimmer for lazy-loaded layouts.
 */
const LayoutFallbackShimmer: React.FC = () => (
  <div className={styles.shimmerContainer}>
    <Shimmer
      shimmerElements={[
        { type: ShimmerElementType.line, height: 120, width: '100%' }
      ]}
      width="100%"
    />
    <Shimmer
      shimmerElements={[
        { type: ShimmerElementType.line, height: 120, width: '100%' }
      ]}
      width="100%"
      style={{ marginTop: 12 }}
    />
    <Shimmer
      shimmerElements={[
        { type: ShimmerElementType.line, height: 120, width: '100%' }
      ]}
      width="100%"
      style={{ marginTop: 12 }}
    />
  </div>
);

/**
 * Renders the promoted results section — highlighted cards above the main results.
 */
const PromotedResultsSection: React.FC<{ items: IPromotedResultItem[] }> = (sectionProps) => {
  if (!sectionProps.items || sectionProps.items.length === 0) {
    // eslint-disable-next-line @rushstack/no-new-null
    return null;
  }

  return (
    <div className={styles.promotedResults}>
      {sectionProps.items.map((item: IPromotedResultItem, index: number) => (
        <div key={item.url + '-' + String(index)} className={styles.promotedCard}>
          <div className={styles.promotedIcon}>
            {item.iconUrl ? (
              <img src={item.iconUrl} alt="" style={{ width: 20, height: 20 }} />
            ) : (
              <Icon iconName="FavoriteStar" />
            )}
          </div>
          <div className={styles.promotedContent}>
            <div className={styles.promotedBadge}>Promoted</div>
            <h3 className={styles.promotedTitle}>
              <a href={item.url} target="_blank" rel="noopener noreferrer">
                {item.title}
              </a>
            </h3>
            {item.description && (
              <p className={styles.promotedDescription}>{item.description}</p>
            )}
          </div>
        </div>
      ))}
    </div>
  );
};

/**
 * Renders the empty state when no results are found.
 */
const EmptyState: React.FC = () => (
  <div className={styles.emptyState}>
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
    isSearchManagerOpen
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

  // ─── Search Manager panel dismiss ─────────────────────────
  const handleDismissSearchManager = React.useCallback((): void => {
    store.getState().toggleSearchManager();
  }, [store]);

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
          <React.Suspense fallback={<LayoutFallbackShimmer />}>
            <CardLayout items={items} onPreviewItem={handlePreviewItem} onItemClick={handleItemClick} />
          </React.Suspense>
        );

      case 'people':
        return (
          <React.Suspense fallback={<LayoutFallbackShimmer />}>
            <PeopleLayout items={items} onPreviewItem={handlePreviewItem} onItemClick={handleItemClick} />
          </React.Suspense>
        );

      case 'datagrid':
        return (
          <React.Suspense fallback={<LayoutFallbackShimmer />}>
            <DataGridLayout
              items={items}
              enableSelection={enableSelection}
              selectedKeys={bulkSelection}
              onToggleSelection={handleToggleSelection}
              onPreviewItem={handlePreviewItem}
              onItemClick={handleItemClick}
            />
          </React.Suspense>
        );

      case 'gallery':
        return (
          <React.Suspense fallback={<LayoutFallbackShimmer />}>
            <GalleryLayout items={items} onPreviewItem={handlePreviewItem} onItemClick={handleItemClick} />
          </React.Suspense>
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
          <div className={styles.errorContainer}>
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
          <React.Suspense fallback={<div />}>
            <ResultDetailPanel
              isOpen={previewPanel.isOpen}
              item={previewPanel.item}
              onDismiss={handleDismissPreviewPanel}
            />
          </React.Suspense>
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
            <React.Suspense fallback={<div style={{ padding: 20 }}>Loading...</div>}>
              <SpSearchManager
                store={store}
                service={managerService}
                theme={theme}
                mode="panel"
              />
            </React.Suspense>
          </Panel>
        )}
      </div>
    </ErrorBoundary>
  );
};

export default SpSearchResults;
