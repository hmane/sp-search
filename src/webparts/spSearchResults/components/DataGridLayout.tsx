import * as React from 'react';
import { ISearchResult, ISortField, ISortableProperty } from '@interfaces/index';
import DataGridContent from './DataGridContent';
import { ISelectedPropertyColumn } from './ISpSearchResultsProps';
import { TitleDisplayMode } from './documentTitleUtils';
import styles from './SpSearchResults.module.scss';

export interface IDataGridLayoutProps {
  items: ISearchResult[];
  searchContextId: string;
  gridPropertyColumns: ISelectedPropertyColumn[];
  titleDisplayMode: TitleDisplayMode;
  totalCount: number;
  pageSize: number;
  currentPage: number;
  showPaging: boolean;
  pageRange: number;
  showDeleteConfirmation: boolean;
  sort: ISortField | undefined;
  sortableProperties: ISortableProperty[];
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
  /** Called when the pager changes page. */
  onPageChange: (page: number) => void;
  onSortChange: (sort: ISortField) => void;
  /**
   * Called when a runtime render error is caught inside the grid.
   * Use this to switch to a safe fallback layout (e.g. list).
   */
  onFallback?: () => void;
}

interface IDataGridRenderErrorState {
  hasError: boolean;
}

/**
 * Error boundary that isolates DevExtreme DataGrid render errors from the
 * outer createLazyComponent chunk-load error boundary in SpSearchResults.
 *
 * When a runtime error is caught (e.g. a DevExtreme prop mismatch or cell
 * renderer throw), this boundary:
 *   1. Logs the real error message and stack — not the generic chunk-load message.
 *   2. Renders a brief status message while onFallback switches the layout.
 *   3. Calls onFallback (deferred) to trigger the layout switch in the store.
 *
 * Chunk-load failures (network errors loading the DataGridLayout bundle) are
 * still handled by the outer createLazyComponent boundary.
 */
class DataGridRenderErrorBoundary extends React.Component<
  { children: React.ReactNode; onFallback?: () => void },
  IDataGridRenderErrorState
> {
  constructor(props: { children: React.ReactNode; onFallback?: () => void }) {
    super(props);
    this.state = { hasError: false };
  }

  static getDerivedStateFromError(): IDataGridRenderErrorState {
    return { hasError: true };
  }

  componentDidCatch(error: Error, info: React.ErrorInfo): void {
    // Log the actual error — distinct from the chunk-load "Failed to load data grid layout" message
    console.error(
      '[SP Search] DataGrid render error (runtime, not a chunk-load failure)' +
      '\nMessage:', error.message,
      '\nStack:', error.stack,
      '\nComponent stack:', info.componentStack
    );
  }

  componentDidUpdate(_prevProps: unknown, prevState: IDataGridRenderErrorState): void {
    if (!prevState.hasError && this.state.hasError && this.props.onFallback) {
      console.warn('[SP Search] DataGrid render error boundary: triggering fallback to list layout.');
      // Defer to avoid triggering a Zustand state update during React's error recovery cycle
      setTimeout(this.props.onFallback, 0);
    }
  }

  render(): React.ReactNode {
    if (this.state.hasError) {
      return (
        <div className={styles.dataGridError} role="status">
          Grid view encountered an error. Switching to list view.
        </div>
      );
    }
    return this.props.children;
  }
}

const DataGridLayout: React.FC<IDataGridLayoutProps> = (props) => {
  return (
    <div className={styles.dataGridContainer}>
      <DataGridRenderErrorBoundary onFallback={props.onFallback}>
        <DataGridContent
          items={props.items}
          searchContextId={props.searchContextId}
          selectedPropertyColumns={props.gridPropertyColumns}
          titleDisplayMode={props.titleDisplayMode}
          totalCount={props.totalCount}
          pageSize={props.pageSize}
          currentPage={props.currentPage}
          showPaging={props.showPaging}
          pageRange={props.pageRange}
          showDeleteConfirmation={props.showDeleteConfirmation}
          sort={props.sort}
          sortableProperties={props.sortableProperties}
          onPreviewItem={props.onPreviewItem}
          onItemClick={props.onItemClick}
          onPageChange={props.onPageChange}
          onSortChange={props.onSortChange}
        />
      </DataGridRenderErrorBoundary>
    </div>
  );
};

export default DataGridLayout;
