import * as React from 'react';
// eslint-disable-next-line @typescript-eslint/no-explicit-any
const DxDataGrid: any = React.lazy(() => import('devextreme-react/data-grid') as any);
import { ISearchResult } from '@interfaces/index';
import { formatShortDate, formatFileSize } from './documentTitleUtils';
import DocumentTitleHoverCard from './DocumentTitleHoverCard';
import styles from './SpSearchResults.module.scss';

export interface IDataGridLayoutProps {
  items: ISearchResult[];
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
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
    modified: formatShortDate(item.modified),
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
  const { items, onPreviewItem, onItemClick } = props;

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

  // Custom cell render for title column with hover card + link
  const titleCellRender = React.useCallback((cellData: { value: string; data: { key: string; url: string }; rowIndex?: number }): React.ReactElement => {
    const matchingItem = items.find((item: ISearchResult) => item.key === cellData.data.key);
    const position = (cellData.rowIndex !== undefined ? cellData.rowIndex : 0) + 1;

    if (!matchingItem) {
      return (
        <a href={cellData.data.url} target="_blank" rel="noopener noreferrer" className={styles.gridTitleLink}>
          {cellData.value}
        </a>
      );
    }

    return (
      <DocumentTitleHoverCard item={matchingItem} position={position} onItemClick={onItemClick} hostDisplay="block">
        {(handleClick): React.ReactNode => (
          <a
            href={cellData.data.url}
            target="_blank"
            rel="noopener noreferrer"
            className={styles.gridTitleLink}
            onClick={(e: React.MouseEvent): void => {
              e.stopPropagation();
              handleClick(e);
            }}
          >
            {cellData.value}
          </a>
        )}
      </DocumentTitleHoverCard>
    );
  }, [items, onItemClick]);

  const columns = React.useMemo(() => [
    { dataField: 'title', caption: 'Name', cellRender: titleCellRender },
    { dataField: 'author', caption: 'Author', width: 160 },
    { dataField: 'modified', caption: 'Modified', width: 110 },
    { dataField: 'fileType', caption: 'Type', width: 70 },
    { dataField: 'fileSize', caption: 'Size', width: 80 },
    { dataField: 'siteName', caption: 'Site', width: 150 }
  ], [titleCellRender]);

  return (
    <div className={styles.dataGridContainer}>
      <React.Suspense fallback={<div className={styles.dataGridLoading}>Loading data grid...</div>}>
        <DxDataGrid
          dataSource={gridData}
          keyExpr="key"
          columns={columns}
          showBorders={false}
          showColumnLines={false}
          showRowLines={true}
          columnAutoWidth={false}
          rowAlternationEnabled={true}
          hoverStateEnabled={true}
          onRowClick={handleRowClick}
          height="auto"
          wordWrapEnabled={false}
        />
      </React.Suspense>
    </div>
  );
};

export default DataGridLayout;
