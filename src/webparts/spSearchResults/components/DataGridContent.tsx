import * as React from 'react';
import DataGrid, { Column } from 'devextreme-react/data-grid';
import { ISearchResult } from '@interfaces/index';
import { formatRelativeDate, formatDateTime, formatFileSize } from './documentTitleUtils';
import { FileTypeIcon, IconType, ImageSize } from '@pnp/spfx-controls-react/lib/FileTypeIcon';
import DocumentTitleHoverCard from './DocumentTitleHoverCard';
import styles from './SpSearchResults.module.scss';

export interface IDataGridContentProps {
  items: ISearchResult[];
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
}

interface IGridRow {
  key: string;
  title: string;
  url: string;
  author: string;
  authorInitials: string;
  modified: string;
  modifiedFull: string;
  fileType: string;
  fileSize: string;
  siteName: string;
}

function getInitials(name: string): string {
  if (!name) {
    return '?';
  }
  const parts: string[] = name.trim().split(/\s+/);
  if (parts.length >= 2) {
    return (parts[0][0] + parts[parts.length - 1][0]).toUpperCase();
  }
  return name.substring(0, 2).toUpperCase();
}

function getInitialsColor(name: string): string {
  const colors: string[] = [
    '#0078d4', '#498205', '#8764b8', '#ca5010',
    '#038387', '#986f0b', '#4f6bed', '#c239b3',
    '#e3008c', '#5c2d91', '#0099bc', '#8e562e'
  ];
  let hash: number = 0;
  for (let i: number = 0; i < name.length; i++) {
    hash = ((hash << 5) - hash) + name.charCodeAt(i);
    hash = hash & hash;
  }
  return colors[Math.abs(hash) % colors.length];
}

function transformForGrid(items: ISearchResult[]): IGridRow[] {
  return items.map((item: ISearchResult) => ({
    key: item.key,
    title: item.title,
    url: item.url,
    author: item.author?.displayText || '',
    authorInitials: getInitials(item.author?.displayText || ''),
    modified: formatRelativeDate(item.modified),
    modifiedFull: formatDateTime(item.modified),
    fileType: (item.fileType || '').toUpperCase(),
    fileSize: formatFileSize(item.fileSize),
    siteName: item.siteName || ''
  }));
}

const DataGridContent: React.FC<IDataGridContentProps> = (props) => {
  const { items, onPreviewItem, onItemClick } = props;

  const gridData = React.useMemo(() => transformForGrid(items), [items]);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const handleRowClick = React.useCallback((e: any): void => {
    const rowKey: string = e.data?.key;
    if (!rowKey) {
      return;
    }
    const matchingItem = items.find((item: ISearchResult) => item.key === rowKey);

    if (onItemClick && matchingItem) {
      onItemClick(matchingItem, (e.rowIndex || 0) + 1);
    }

    if (onPreviewItem && matchingItem) {
      onPreviewItem(matchingItem);
    }
  }, [items, onItemClick, onPreviewItem]);

  // ── Cell Renderers ──────────────────────────────────────────

  const titleCellRender = React.useCallback((cellData: { value: string; data: IGridRow; rowIndex?: number }): React.ReactElement => {
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
          <div className={styles.gridTitleCell}>
            <span className={styles.gridTitleIcon}>
              <FileTypeIcon type={IconType.image} path={matchingItem.url} size={ImageSize.small} />
            </span>
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
          </div>
        )}
      </DocumentTitleHoverCard>
    );
  }, [items, onItemClick]);

  const authorCellRender = React.useCallback((cellData: { data: IGridRow }): React.ReactElement => {
    const name: string = cellData.data.author;
    const initials: string = cellData.data.authorInitials;

    if (!name) {
      return <span className={styles.gridCellMuted}>{'\u2014'}</span>;
    }

    const bgColor: string = getInitialsColor(name);

    return (
      <div className={styles.gridAuthorCell}>
        <span
          className={styles.gridAuthorAvatar}
          style={{ backgroundColor: bgColor }}
          title={name}
        >
          {initials}
        </span>
        <span className={styles.gridAuthorName}>{name}</span>
      </div>
    );
  }, []);

  const modifiedCellRender = React.useCallback((cellData: { value: string; data: IGridRow }): React.ReactElement => {
    const relativeDate: string = cellData.value;
    const fullDate: string = cellData.data.modifiedFull;

    if (!relativeDate) {
      return <span className={styles.gridCellMuted}>{'\u2014'}</span>;
    }

    return (
      <span className={styles.gridDateCell} title={fullDate}>
        {relativeDate}
      </span>
    );
  }, []);

  const typeCellRender = React.useCallback((cellData: { data: IGridRow }): React.ReactElement => {
    const label: string = cellData.data.fileType;

    if (!label) {
      return <span className={styles.gridCellMuted}>{'\u2014'}</span>;
    }

    return (
      <span className={styles.gridTypeBadge}>
        {label}
      </span>
    );
  }, []);

  return (
    <DataGrid
      dataSource={gridData}
      keyExpr="key"
      showBorders={false}
      showColumnLines={false}
      showRowLines={true}
      columnAutoWidth={false}
      rowAlternationEnabled={true}
      hoverStateEnabled={true}
      onRowClick={handleRowClick}
      height="auto"
      wordWrapEnabled={false}
    >
      <Column dataField="title" caption="Name" cellRender={titleCellRender} minWidth={200} />
      <Column dataField="author" caption="Author" width={180} cellRender={authorCellRender} />
      <Column dataField="modified" caption="Modified" width={120} cellRender={modifiedCellRender} />
      <Column dataField="fileType" caption="Type" width={70} cellRender={typeCellRender} alignment="center" />
      <Column dataField="fileSize" caption="Size" width={80} alignment="right" />
      <Column dataField="siteName" caption="Site" width={150} />
    </DataGrid>
  );
};

export default DataGridContent;
