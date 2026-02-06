import * as React from 'react';
import { ISearchResult } from '@interfaces/index';
import styles from './SpSearchResults.module.scss';

export interface IDataGridLayoutProps {
  items: ISearchResult[];
  enableSelection: boolean;
  selectedKeys: string[];
  onToggleSelection: (key: string, multiSelect: boolean) => void;
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
}

// ─── Lazy-loaded DevExtreme DataGrid ─────────────────────
const DataGrid = React.lazy(() => import(/* webpackChunkName: 'DevExtremeDataGrid' */ 'devextreme-react/data-grid'));

/**
 * Formats an ISO date string into a short locale-friendly format.
 */
function formatDate(isoDate: string): string {
  if (!isoDate) {
    return '';
  }
  try {
    const d: Date = new Date(isoDate);
    if (isNaN(d.getTime())) {
      return '';
    }
    const month: number = d.getMonth() + 1;
    const day: number = d.getDate();
    const year: number = d.getFullYear();
    return month + '/' + day + '/' + year;
  } catch {
    return '';
  }
}

/**
 * Formats a file size in bytes into a human-readable string.
 */
function formatFileSize(bytes: number): string {
  if (!bytes || bytes <= 0) {
    return '';
  }
  if (bytes < 1024) {
    return bytes + ' B';
  }
  if (bytes < 1048576) {
    return Math.round(bytes / 1024) + ' KB';
  }
  return (bytes / 1048576).toFixed(1) + ' MB';
}

/**
 * Transform search results into a flat data structure for the grid.
 */
function transformForGrid(items: ISearchResult[]): Array<{
  key: string;
  title: string;
  url: string;
  author: string;
  modified: string;
  fileType: string;
  fileSize: string;
  siteName: string;
}> {
  return items.map((item: ISearchResult) => ({
    key: item.key,
    title: item.title,
    url: item.url,
    author: item.author?.displayText || '',
    modified: formatDate(item.modified),
    fileType: (item.fileType || '').toUpperCase(),
    fileSize: formatFileSize(item.fileSize),
    siteName: item.siteName || ''
  }));
}

/**
 * DataGridLayout — renders search results as a DevExtreme DataGrid.
 * Supports column sorting, filtering, selection, and row click for preview.
 *
 * DevExtreme DataGrid is lazy-loaded to minimize initial bundle size.
 */
const DataGridLayout: React.FC<IDataGridLayoutProps> = (props) => {
  const { items, enableSelection, selectedKeys, onToggleSelection, onPreviewItem, onItemClick } = props;

  const gridData = React.useMemo(() => transformForGrid(items), [items]);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const handleRowClick = React.useCallback((e: any): void => {
    const rowKey: string = e.data?.key;
    if (!rowKey) {
      return;
    }
    const matchingItem = items.find((item: ISearchResult) => item.key === rowKey);

    // Log the click
    if (onItemClick && matchingItem) {
      onItemClick(matchingItem, (e.rowIndex || 0) + 1);
    }

    // Open preview panel if available
    if (onPreviewItem && matchingItem) {
      onPreviewItem(matchingItem);
    }
  }, [items, onItemClick, onPreviewItem]);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const handleSelectionChanged = React.useCallback((e: any): void => {
    if (!enableSelection || !e.selectedRowKeys || e.selectedRowKeys.length === 0) {
      return;
    }
    // For now, handle single selections; multi-select can be added later
    const selectedKey = String(e.selectedRowKeys[0]);
    if (selectedKey) {
      onToggleSelection(selectedKey, false);
    }
  }, [enableSelection, onToggleSelection]);

  // Custom cell render for title column with link
  const titleCellRender = React.useCallback((cellData: { value: string; data: { key: string; url: string }; rowIndex?: number }): React.ReactElement => {
    const handleLinkClick = (e: React.MouseEvent): void => {
      e.stopPropagation(); // Prevent row click from also firing

      // Log the click for analytics parity with other layouts
      if (onItemClick) {
        const matchingItem = items.find((item: ISearchResult) => item.key === cellData.data.key);
        if (matchingItem) {
          const position = (cellData.rowIndex !== undefined ? cellData.rowIndex : 0) + 1;
          onItemClick(matchingItem, position);
        }
      }
    };

    return (
      <a
        href={cellData.data.url}
        target="_blank"
        rel="noopener noreferrer"
        className={styles.gridTitleLink}
        onClick={handleLinkClick}
      >
        {cellData.value}
      </a>
    );
  }, [items, onItemClick]);

  return (
    <div className={styles.dataGridContainer}>
      <React.Suspense fallback={<div style={{ padding: 20 }}>Loading grid...</div>}>
        <DataGrid
          dataSource={gridData}
          keyExpr="key"
          showBorders={true}
          columnAutoWidth={true}
          rowAlternationEnabled={true}
          hoverStateEnabled={true}
          onRowClick={handleRowClick}
          onSelectionChanged={handleSelectionChanged}
          selection={enableSelection ? { mode: 'multiple', showCheckBoxesMode: 'always' } : undefined}
          selectedRowKeys={enableSelection ? selectedKeys : undefined}
          height="auto"
        >
          {/* Title column with custom render */}
          <div data-options="dxTemplate" data-name="titleTemplate">
            {titleCellRender}
          </div>
        </DataGrid>
      </React.Suspense>
    </div>
  );
};

export default DataGridLayout;
