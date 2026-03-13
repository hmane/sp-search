import * as React from 'react';
import DataGrid, { Column, Scrolling, Sorting, StateStoring, Toolbar, Item, Paging, Pager } from 'devextreme-react/data-grid';
import { IconButton, DefaultButton } from '@fluentui/react/lib/Button';
import { IContextualMenuItem } from '@fluentui/react/lib/ContextualMenu';
import { confirm } from 'spfx-toolkit/lib/utilities/dialogService';
import { ISearchResult, ISortField, ISortableProperty } from '@interfaces/index';
import { PermissionKind } from '@pnp/sp/security';
import { hasPermissions } from '@pnp/sp/security/funcs';
import { buildDownloadUrl, copyTextToClipboard } from '@providers/actions/actionUtils';
import { formatRelativeDate, formatDateTime, formatFileSize, getResultAnchorProps, buildFormUrl, buildBrowserOpenUrl, formatTitleText, TitleDisplayMode } from './documentTitleUtils';
import { FileTypeIcon, IconType, ImageSize } from '@pnp/spfx-controls-react/lib/FileTypeIcon';
import DocumentTitleHoverCard from './DocumentTitleHoverCard';
import { ISelectedPropertyColumn } from './ISpSearchResultsProps';
import Pagination from './Pagination';
import styles from './SpSearchResults.module.scss';

export interface IDataGridContentProps {
  items: ISearchResult[];
  selectedPropertyColumns: ISelectedPropertyColumn[];
  titleDisplayMode: TitleDisplayMode;
  /** Total result count from the store — used to decide whether to activate virtual scrolling. */
  totalCount: number;
  /** Active page size from the store — virtual scrolling only activates when pageSize >= 25 and totalCount > pageSize. */
  pageSize: number;
  currentPage: number;
  showPaging: boolean;
  pageRange: number;
  /** Used to scope the localStorage key for column preferences — one entry per context per page. */
  searchContextId: string;
  showDeleteConfirmation: boolean;
  sort: ISortField | undefined;
  sortableProperties: ISortableProperty[];
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
  onPageChange: (page: number) => void;
  onSortChange: (sort: ISortField) => void;
}

type ColumnKind = 'title' | 'author' | 'date' | 'fileType' | 'fileSize' | 'url' | 'text';

interface IGridColumnConfig {
  property: string;
  caption: string;
  kind: ColumnKind;
  width?: number;
  minWidth?: number;
  alignment?: 'left' | 'center' | 'right';
}

interface IGridRow {
  key: string;
  __item: ISearchResult;
  [key: string]: unknown;
}

interface IItemPermissionState {
  checked: boolean;
  canEdit: boolean;
  canDelete: boolean;
}

const DEFAULT_COLUMNS: ISelectedPropertyColumn[] = [
  { property: 'Author', alias: 'Author' },
  { property: 'LastModifiedTime', alias: 'Modified' },
  { property: 'FileType', alias: 'Type' },
  { property: 'Size', alias: 'Size' },
  { property: 'SiteTitle', alias: 'Site' }
];

const GRID_HEADER_HEIGHT = 44;
const GRID_TOOLBAR_HEIGHT = 44;
const GRID_ROW_HEIGHT = 44;
const GRID_FRAME_HEIGHT = 16;
const GRID_MIN_HEIGHT = 220;
const GRID_MAX_HEIGHT = 500;

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

function getConfiguredColumns(selectedPropertyColumns: ISelectedPropertyColumn[]): ISelectedPropertyColumn[] {
  const source = [{ property: 'Title', alias: 'Name' }, ...(selectedPropertyColumns.length > 0 ? selectedPropertyColumns : DEFAULT_COLUMNS)];
  const seen = new Set<string>();
  const unique: ISelectedPropertyColumn[] = [];

  for (let i: number = 0; i < source.length; i++) {
    const property = (source[i].property || '').trim();
    if (!property) {
      continue;
    }
    const lookupKey = property.toLowerCase();
    if (seen.has(lookupKey)) {
      continue;
    }
    seen.add(lookupKey);
    unique.push({
      property,
      alias: (source[i].alias || '').trim()
    });
  }

  return unique;
}

function getColumnConfig(property: string, alias: string): IGridColumnConfig {
  const normalized = property.toLowerCase();

  if (normalized === 'title' || normalized === 'filename') {
    return { property, caption: alias || 'Name', kind: 'title', minWidth: 220 };
  }
  if (normalized === 'author' || normalized === 'authorowsuser' || normalized === 'displayauthor') {
    return { property, caption: alias || 'Author', kind: 'author', width: 180 };
  }
  if (normalized === 'lastmodifiedtime' || normalized === 'modified' || normalized === 'created') {
    return { property, caption: alias || property, kind: 'date', width: 140 };
  }
  if (normalized === 'filetype' || normalized === 'fileextension') {
    return { property, caption: alias || 'Type', kind: 'fileType', width: 80, alignment: 'center' };
  }
  if (normalized === 'size' || normalized === 'filesize') {
    return { property, caption: alias || 'Size', kind: 'fileSize', width: 90, alignment: 'right' };
  }
  if (normalized === 'path' || normalized.indexOf('url') >= 0 || normalized.indexOf('link') >= 0) {
    return { property, caption: alias || property, kind: 'url', minWidth: 220 };
  }
  return { property, caption: alias || property, kind: 'text', minWidth: 140 };
}

function resolvePropertyValue(item: ISearchResult, property: string): unknown {
  const normalized = property.toLowerCase();

  switch (normalized) {
    case 'title':
      return item.title;
    case 'filename':
      return typeof item.properties[property] === 'string' ? item.properties[property] : item.title;
    case 'path':
      return item.url;
    case 'author':
    case 'authorowsuser':
    case 'displayauthor':
      return item.author?.displayText || item.properties[property] || '';
    case 'created':
      return item.created || item.properties[property] || '';
    case 'lastmodifiedtime':
    case 'modified':
      return item.modified || item.properties[property] || '';
    case 'filetype':
    case 'fileextension':
      return item.fileType || item.properties[property] || '';
    case 'size':
    case 'filesize':
      return item.fileSize || item.properties[property] || 0;
    case 'sitetitle':
    case 'sitename':
      return item.siteName || item.properties[property] || '';
    default:
      if (Object.prototype.hasOwnProperty.call(item.properties, property)) {
        return item.properties[property];
      }
      return '';
  }
}

function formatTextValue(value: unknown): string {
  if (value === undefined || value === null || value === '') {
    return '--';
  }
  if (typeof value === 'string') {
    return value;
  }
  if (typeof value === 'number') {
    return value.toLocaleString();
  }
  if (typeof value === 'boolean') {
    return value ? 'Yes' : 'No';
  }
  if (Array.isArray(value)) {
    return value.map((entry: unknown) => formatTextValue(entry)).join(', ');
  }
  if (typeof value === 'object') {
    if ('displayText' in value && typeof value.displayText === 'string') {
      return value.displayText;
    }
    return JSON.stringify(value);
  }
  return String(value);
}

function getDxSortOrder(sort: ISortField | undefined, property: string): 'asc' | 'desc' | undefined {
  if (!sort || sort.property === 'Rank' || sort.property.toLowerCase() !== property.toLowerCase()) {
    return undefined;
  }
  return sort.direction === 'Descending' ? 'desc' : 'asc';
}

function getColumnSortCandidates(property: string): string[] {
  const normalized = property.toLowerCase();
  switch (normalized) {
    case 'author':
    case 'authorowsuser':
    case 'displayauthor':
      return ['DisplayAuthor', 'Author', 'AuthorOWSUSER'];
    case 'lastmodifiedtime':
    case 'modified':
      return ['LastModifiedTime', 'Modified'];
    case 'created':
      return ['Created'];
    case 'title':
    case 'filename':
      return ['Title', 'Filename'];
    case 'filetype':
    case 'fileextension':
      return ['FileType', 'FileExtension'];
    case 'size':
    case 'filesize':
      return ['Size', 'FileSize'];
    case 'sitetitle':
    case 'sitename':
      return ['SiteTitle', 'SiteName'];
    default:
      return [property];
  }
}

function resolveColumnSortProperty(property: string, sortableProperties: ISortableProperty[]): string | undefined {
  const candidates = getColumnSortCandidates(property);
  const configured = sortableProperties.length > 0
    ? sortableProperties.map((item) => item.property)
    : ['LastModifiedTime', 'Title', 'DisplayAuthor'];

  for (let i: number = 0; i < candidates.length; i++) {
    const candidate = candidates[i];
    for (let j: number = 0; j < configured.length; j++) {
      if (configured[j].toLowerCase() === candidate.toLowerCase()) {
        return configured[j];
      }
    }
  }

  return undefined;
}

function mapDxSortToStore(property: string, sortOrder: 'asc' | 'desc' | undefined): ISortField {
  if (!sortOrder) {
    return { property: 'Rank', direction: 'Ascending' };
  }

  return {
    property,
    direction: sortOrder === 'desc' ? 'Descending' : 'Ascending'
  };
}

/**
 * Returns the localStorage key used to store grid column preferences for a given
 * search context. Scoped by context ID so multi-context pages don't share state.
 */
function getGridPrefsKey(searchContextId: string): string {
  return 'sp-search-grid-cols-' + searchContextId;
}

function buildColumnSignature(columns: string[]): string {
  return [...columns].sort().join('|');
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function loadGridState(searchContextId: string, currentColumnKeys: string[]): any {
  try {
    const raw = localStorage.getItem(getGridPrefsKey(searchContextId));
    if (!raw) {
      return {};
    }
    const parsed = JSON.parse(raw);
    if (!parsed || !Array.isArray(parsed.columns)) {
      return {};
    }

    const savedSignature = typeof parsed.columnSignature === 'string' ? parsed.columnSignature : '';
    const currentSignature = buildColumnSignature(currentColumnKeys);
    const allowSavedOrder = savedSignature !== '' && savedSignature === currentSignature;

    return {
      columns: parsed.columns.map((column: {
        dataField?: string;
        visible?: boolean;
        width?: number | string;
        visibleIndex?: number;
      }) => ({
        dataField: column.dataField,
        visible: column.visible,
        width: column.width,
        ...(allowSavedOrder ? { visibleIndex: column.visibleIndex } : {})
      }))
    };
  } catch {
    return {};
  }
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
function saveGridState(searchContextId: string, state: any): void {
  if (!state) {
    return;
  }
  try {
    const columnKeys = Array.isArray(state.columns)
      ? state.columns
        .map((column: { dataField?: string }) => column.dataField)
        .filter((value: string | undefined): value is string => !!value)
      : [];

    const toStore = {
      columnSignature: buildColumnSignature(columnKeys),
      columns: Array.isArray(state.columns)
        ? state.columns.map((column: {
          dataField?: string;
          visible?: boolean;
          width?: number | string;
          visibleIndex?: number;
        }) => ({
          dataField: column.dataField,
          visible: column.visible,
          width: column.width,
          visibleIndex: column.visibleIndex
        }))
        : []
    };
    localStorage.setItem(getGridPrefsKey(searchContextId), JSON.stringify(toStore));
  } catch {
    // localStorage unavailable (private browsing, quota exceeded) — fail silently
  }
}

function transformForGrid(items: ISearchResult[], columns: IGridColumnConfig[]): IGridRow[] {
  return items.map((item: ISearchResult) => {
    const row: IGridRow = {
      key: item.key,
      __item: item
    };

    for (let i: number = 0; i < columns.length; i++) {
      row[columns[i].property] = resolvePropertyValue(item, columns[i].property);
    }

    return row;
  });
}

function getDeleteInfo(item: ISearchResult): { siteUrl: string; listId: string; itemId: number } | undefined {
  const siteUrl = String(item.properties.SPSiteURL || item.siteUrl || '').trim();
  const listId = String(item.properties.ListId || '').trim();
  const itemId = parseInt(String(item.properties.ListItemID || '0'), 10) || 0;

  if (!siteUrl || !listId || !itemId) {
    return undefined;
  }

  return { siteUrl: siteUrl.replace(/\/$/, ''), listId, itemId };
}

async function getFormDigestValue(siteUrl: string): Promise<string> {
  const response = await fetch(siteUrl + '/_api/contextinfo', {
    method: 'POST',
    headers: {
      Accept: 'application/json;odata=nometadata'
    }
  });

  if (!response.ok) {
    throw new Error('Failed to acquire form digest');
  }

  const json = await response.json() as { FormDigestValue?: string };
  return json.FormDigestValue || '';
}

async function recycleSearchResult(item: ISearchResult): Promise<void> {
  const deleteInfo = getDeleteInfo(item);
  if (!deleteInfo) {
    throw new Error('Delete metadata is unavailable for this item');
  }

  const digest = await getFormDigestValue(deleteInfo.siteUrl);
  const response = await fetch(
    deleteInfo.siteUrl + '/_api/web/lists(guid\'' + deleteInfo.listId + '\')/items(' + String(deleteInfo.itemId) + ')/recycle()',
    {
      method: 'POST',
      headers: {
        Accept: 'application/json;odata=nometadata',
        'X-RequestDigest': digest,
        'IF-MATCH': '*'
      }
    }
  );

  if (!response.ok) {
    throw new Error('Delete request failed');
  }
}

async function loadItemPermissions(item: ISearchResult): Promise<IItemPermissionState> {
  const deleteInfo = getDeleteInfo(item);
  if (!deleteInfo) {
    return { checked: true, canEdit: false, canDelete: false };
  }

  const response = await fetch(
    deleteInfo.siteUrl + '/_api/web/lists(guid\'' + deleteInfo.listId + '\')/items(' + String(deleteInfo.itemId) + ')?$select=EffectiveBasePermissions',
    {
      method: 'GET',
      headers: {
        Accept: 'application/json;odata=nometadata'
      }
    }
  );

  if (!response.ok) {
    throw new Error('Permission lookup failed');
  }

  const json = await response.json() as {
    EffectiveBasePermissions?: { High?: number | string; Low?: number | string };
  };

  const perms = {
    High: Number(json.EffectiveBasePermissions?.High || 0),
    Low: Number(json.EffectiveBasePermissions?.Low || 0)
  };

  return {
    checked: true,
    canEdit: hasPermissions(perms, PermissionKind.EditListItems),
    canDelete: hasPermissions(perms, PermissionKind.DeleteListItems)
  };
}

const DataGridContent: React.FC<IDataGridContentProps> = (props) => {
  const {
    items,
    selectedPropertyColumns,
    titleDisplayMode,
    totalCount,
    pageSize,
    currentPage,
    showPaging,
    pageRange,
    searchContextId,
    showDeleteConfirmation,
    sort,
    sortableProperties,
    onItemClick,
    onPageChange,
    onSortChange
  } = props;

  // Activate virtual scrolling only for larger page sizes where DOM virtualization
  // provides a measurable benefit. For common sizes like 10/25, let the grid grow
  // naturally and avoid an internal vertical scrollbar.
  const useVirtualScrolling = pageSize >= 50 && totalCount > pageSize;
  const gridRef = React.useRef<DataGrid | null>(null);
  const hostRef = React.useRef<HTMLDivElement | null>(null);
  const sortSyncTimeoutRef = React.useRef<number | undefined>(undefined);
  const [isFullscreen, setIsFullscreen] = React.useState<boolean>(false);
  const [permissionCache, setPermissionCache] = React.useState<Record<string, IItemPermissionState>>({});
  const [columnVisibility, setColumnVisibility] = React.useState<Record<string, boolean>>({});

  React.useEffect((): (() => void) => {
    if (!isFullscreen) {
      return (): void => { /* noop */ };
    }

    const previousOverflow = document.body.style.overflow;
    document.body.style.overflow = 'hidden';

    const handleEscape = (event: KeyboardEvent): void => {
      if (event.key === 'Escape') {
        setIsFullscreen(false);
      }
    };

    window.addEventListener('keydown', handleEscape);

    return (): void => {
      document.body.style.overflow = previousOverflow;
      window.removeEventListener('keydown', handleEscape);
    };
  }, [isFullscreen]);

  const columnConfigs = React.useMemo(
    (): IGridColumnConfig[] => getConfiguredColumns(selectedPropertyColumns).map((column) => {
      return getColumnConfig(column.property, column.alias);
    }),
    [selectedPropertyColumns]
  );

  const gridData = React.useMemo(
    (): IGridRow[] => transformForGrid(items, columnConfigs),
    [items, columnConfigs]
  );

  const estimatedGridHeight = React.useMemo((): number | 'auto' | '100%' => {
    if (isFullscreen) {
      return '100%';
    }

    const rowCount = gridData.length;
    if (rowCount === 0) {
      return GRID_MIN_HEIGHT;
    }

    if (!useVirtualScrolling) {
      return 'auto';
    }

    // Approximate the visible grid height from the currently loaded rows.
    // Even when the page size is large, keep the normal viewport compact and
    // let the grid scroll internally rather than dominating the whole page.
    const visibleRows = Math.min(Math.max(rowCount, 6), 8);

    const estimated = GRID_HEADER_HEIGHT +
      GRID_TOOLBAR_HEIGHT +
      GRID_FRAME_HEIGHT +
      (visibleRows * GRID_ROW_HEIGHT);

    return Math.max(GRID_MIN_HEIGHT, Math.min(GRID_MAX_HEIGHT, estimated));
  }, [gridData.length, isFullscreen, useVirtualScrolling]);

  React.useEffect((): void => {
    const permissionItems = items.filter((item) => !!getDeleteInfo(item) && !permissionCache[item.key]);
    if (permissionItems.length === 0) {
      return;
    }

    permissionItems.forEach((item) => {
      loadItemPermissions(item)
        .then((state): void => {
          setPermissionCache((prev) => ({ ...prev, [item.key]: state }));
        })
        .catch((): void => {
          setPermissionCache((prev) => ({
            ...prev,
            [item.key]: { checked: true, canEdit: false, canDelete: false }
          }));
        });
    });
  }, [items, permissionCache]);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const handleRowClick = React.useCallback((e: any): void => {
    const matchingItem = e.data?.__item as ISearchResult | undefined;
    if (!matchingItem) {
      return;
    }

    if (onItemClick) {
      onItemClick(matchingItem, (e.rowIndex || 0) + 1);
    }

  }, [onItemClick]);

  const syncVisibleColumnsFromGrid = React.useCallback((): void => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const instance = (gridRef.current as any)?.instance;
    if (!instance?.columnOption) {
      return;
    }
    const next: Record<string, boolean> = {};
    for (let i: number = 0; i < columnConfigs.length; i++) {
      const property = columnConfigs[i].property;
      next[property] = instance.columnOption(property, 'visible') !== false;
    }
    setColumnVisibility((prev) => {
      const prevKeys = Object.keys(prev);
      const nextKeys = Object.keys(next);
      if (prevKeys.length === nextKeys.length) {
        let changed = false;
        for (let i: number = 0; i < nextKeys.length; i++) {
          const key = nextKeys[i];
          if (prev[key] !== next[key]) {
            changed = true;
            break;
          }
        }
        if (!changed) {
          return prev;
        }
      }
      return next;
    });
  }, [columnConfigs]);

  const handleToggleColumnVisibility = React.useCallback((property: string): void => {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const instance = (gridRef.current as any)?.instance;
    if (!instance?.columnOption) {
      return;
    }
    const current = instance.columnOption(property, 'visible') !== false;
    instance.columnOption(property, 'visible', !current);
    syncVisibleColumnsFromGrid();
  }, [syncVisibleColumnsFromGrid]);

  // ─── Column preference persistence ─────────────────────────
  // StateStoring fires customLoad once on mount and customSave on every state change.
  // We filter to the columns slice only so sort/filter/paging driven by the Zustand
  // store are never overridden by cached grid state.
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const handleLoadGridState = React.useCallback((): any => {
    return loadGridState(searchContextId, columnConfigs.map((column) => column.property));
  }, [columnConfigs, searchContextId]);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const handleSaveGridState = React.useCallback((state: any): void => {
    saveGridState(searchContextId, state);
  }, [searchContextId]);

  React.useEffect((): (() => void) => {
    return (): void => {
      if (sortSyncTimeoutRef.current !== undefined) {
        window.clearTimeout(sortSyncTimeoutRef.current);
      }
    };
  }, []);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const handleGridOptionChanged = React.useCallback((e: any): void => {
    const fullName = String(e?.fullName || '');
    if (/^columns\[\d+\]\.visible$/.test(fullName)) {
      syncVisibleColumnsFromGrid();
      return;
    }
    if (!/^columns\[\d+\]\.sortOrder$/.test(fullName)) {
      return;
    }

    if (sortSyncTimeoutRef.current !== undefined) {
      window.clearTimeout(sortSyncTimeoutRef.current);
    }

    sortSyncTimeoutRef.current = window.setTimeout((): void => {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const instance = (gridRef.current as any)?.instance;
      let nextSort: ISortField = { property: 'Rank', direction: 'Ascending' };

      if (instance?.columnOption) {
        for (let i: number = 0; i < columnConfigs.length; i++) {
          const sortOrder = instance.columnOption(i, 'sortOrder') as 'asc' | 'desc' | undefined;
          const sortProperty = resolveColumnSortProperty(columnConfigs[i].property, sortableProperties);
          if (sortOrder && sortProperty) {
            nextSort = mapDxSortToStore(sortProperty, sortOrder);
            break;
          }
        }
      }

      const currentProperty = sort?.property || 'Rank';
      const currentDirection = sort?.direction || 'Ascending';
      if (currentProperty === nextSort.property && currentDirection === nextSort.direction) {
        return;
      }

      onSortChange(nextSort);
    }, 0);
  }, [columnConfigs, onSortChange, sort, sortableProperties, syncVisibleColumnsFromGrid]);

  const columnMenuItems = React.useMemo((): IContextualMenuItem[] => {
    return columnConfigs
      .filter((column) => column.kind !== 'title')
      .map((column) => ({
        key: column.property,
        text: column.caption,
        canCheck: true,
        checked: columnVisibility[column.property] !== false,
        onClick: (): void => {
          handleToggleColumnVisibility(column.property);
        }
      }));
  }, [columnConfigs, columnVisibility, handleToggleColumnVisibility]);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const titleCellRender = React.useCallback((cellData: { value: unknown; data: IGridRow; rowIndex?: number }): React.ReactElement => {
    const matchingItem = cellData.data.__item;
    const position = (cellData.rowIndex !== undefined ? cellData.rowIndex : 0) + 1;
    const title = formatTitleText(formatTextValue(cellData.value), titleDisplayMode);
    const linkProps = getResultAnchorProps(matchingItem);
    const permissionState = permissionCache[matchingItem.key];
    const viewUrl = buildFormUrl(matchingItem, 4);
    const editUrl = buildFormUrl(matchingItem, 6);
    const menuItems: IContextualMenuItem[] = [
      {
        key: 'view',
        text: 'View item',
        iconProps: { iconName: 'View' },
        disabled: !viewUrl,
        onClick: (): void => {
          if (viewUrl) {
            window.open(viewUrl, '_blank', 'noopener,noreferrer');
          }
        }
      },
      ...(permissionState?.canEdit && editUrl ? [{
        key: 'edit',
        text: 'Edit item',
        iconProps: { iconName: 'Edit' },
        onClick: (): void => {
          window.open(editUrl, '_blank', 'noopener,noreferrer');
        }
      }] : []),
      {
        key: 'divider-1',
        itemType: 1
      },
      {
        key: 'open',
        text: 'Open document in new tab',
        iconProps: { iconName: 'OpenInNewTab' },
        onClick: (): void => {
          if (onItemClick) {
            onItemClick(matchingItem, position);
          }
          window.open(buildBrowserOpenUrl(matchingItem), '_blank', 'noopener,noreferrer');
        }
      },
      {
        key: 'download',
        text: 'Download',
        iconProps: { iconName: 'Download' },
        onClick: (): void => {
          const downloadUrl = buildDownloadUrl(matchingItem.url);
          if (downloadUrl) {
            window.open(downloadUrl, '_blank', 'noopener,noreferrer');
          }
        }
      },
      {
        key: 'copyLink',
        text: 'Copy link',
        iconProps: { iconName: 'Link' },
        onClick: (): void => {
          copyTextToClipboard(matchingItem.url).catch((): void => {
            // Silently fail
          });
        }
      },
      ...(permissionState?.canDelete ? [{
        key: 'divider-2',
        itemType: 1
      }, {
        key: 'delete',
        text: 'Delete',
        iconProps: { iconName: 'Delete' },
        onClick: (): void => {
          (async (): Promise<void> => {
            if (showDeleteConfirmation) {
              const confirmed = await confirm(
                'Move "' + matchingItem.title + '" to the recycle bin?',
                { title: 'Delete item' }
              );
              if (!confirmed) {
                return;
              }
            }

            recycleSearchResult(matchingItem)
              .then((): void => {
                window.setTimeout((): void => {
                  window.location.reload();
                }, 150);
              })
              .catch((error: Error): void => {
                console.error('[SP Search] Failed to delete grid item.', error);
                window.alert('Unable to delete this item.');
              });
          })().catch((error: Error): void => {
            console.error('[SP Search] Failed to open delete confirmation.', error);
          });
        }
      }] : [])
    ];

    return (
      <DocumentTitleHoverCard item={matchingItem} position={position} onItemClick={onItemClick} hostDisplay="block">
        {(handleClick): React.ReactNode => (
          <div className={styles.gridTitleCell}>
            <span className={styles.gridTitleIcon}>
              <FileTypeIcon type={IconType.image} path={matchingItem.url} size={ImageSize.small} />
            </span>
            <div className={styles.gridTitleMain}>
              <a
                href={linkProps.href}
                target={linkProps.target}
                rel={linkProps.rel}
                className={titleDisplayMode === 'wrap' ? styles.gridTitleLinkWrap : styles.gridTitleLink}
                onClick={(e: React.MouseEvent): void => {
                  e.stopPropagation();
                  handleClick(e);
                }}
              >
                {title}
              </a>
              <IconButton
                iconProps={{ iconName: 'MoreVertical' }}
                title="More actions"
                ariaLabel="More actions"
                menuProps={{ items: menuItems }}
                className={styles.gridTitleActionButton}
                onClick={(event): void => {
                  event.stopPropagation();
                }}
              />
            </div>
          </div>
        )}
      </DocumentTitleHoverCard>
    );
  }, [onItemClick, permissionCache, showDeleteConfirmation, titleDisplayMode]);

  const authorCellRender = React.useCallback((cellData: { value: unknown }): React.ReactElement => {
    const name = formatTextValue(cellData.value);

    if (!cellData.value) {
      return <span className={styles.gridCellMuted}>--</span>;
    }

    const bgColor: string = getInitialsColor(name);

    return (
      <div className={styles.gridAuthorCell}>
        <span
          className={styles.gridAuthorAvatar}
          style={{ backgroundColor: bgColor }}
          title={name}
        >
          {getInitials(name)}
        </span>
        <span className={styles.gridAuthorName}>{name}</span>
      </div>
    );
  }, []);

  const dateCellRender = React.useCallback((cellData: { value: unknown }): React.ReactElement => {
    const rawValue = typeof cellData.value === 'string' ? cellData.value : '';
    if (!rawValue) {
      return <span className={styles.gridCellMuted}>--</span>;
    }

    return (
      <span className={styles.gridDateCell} title={formatDateTime(rawValue)}>
        {formatRelativeDate(rawValue)}
      </span>
    );
  }, []);

  const typeCellRender = React.useCallback((cellData: { value: unknown }): React.ReactElement => {
    const label = formatTextValue(cellData.value).toUpperCase();
    if (!cellData.value) {
      return <span className={styles.gridCellMuted}>--</span>;
    }

    return (
      <span className={styles.gridTypeBadge}>
        {label}
      </span>
    );
  }, []);

  const fileSizeCellRender = React.useCallback((cellData: { value: unknown }): React.ReactElement => {
    const size = typeof cellData.value === 'number'
      ? cellData.value
      : parseInt(String(cellData.value || '0'), 10) || 0;

    if (!size) {
      return <span className={styles.gridCellMuted}>--</span>;
    }

    return (
      <span title={String(size) + ' bytes'}>
        {formatFileSize(size)}
      </span>
    );
  }, []);

  const urlCellRender = React.useCallback((cellData: { value: unknown; data: IGridRow }): React.ReactElement => {
    const href = typeof cellData.value === 'string' && cellData.value ? cellData.value : cellData.data.__item.url;
    if (!href) {
      return <span className={styles.gridCellMuted}>--</span>;
    }

    return (
      <a href={href} target="_blank" rel="noopener noreferrer" className={styles.gridTitleLink}>
        {href}
      </a>
    );
  }, []);

  const textCellRender = React.useCallback((cellData: { value: unknown }): React.ReactElement => {
    const text = formatTextValue(cellData.value);
    if (text === '--') {
      return <span className={styles.gridCellMuted}>--</span>;
    }

    return <span title={text}>{text}</span>;
  }, []);

  const renderColumn = React.useCallback((column: IGridColumnConfig, index: number): React.ReactElement => {
    const sortProperty = resolveColumnSortProperty(column.property, sortableProperties);
    let cellRender:
      | ((cellData: { value: unknown; data: IGridRow; rowIndex?: number }) => React.ReactElement)
      | undefined;

    switch (column.kind) {
      case 'title':
        cellRender = titleCellRender;
        break;
      case 'author':
        cellRender = authorCellRender;
        break;
      case 'date':
        cellRender = dateCellRender;
        break;
      case 'fileType':
        cellRender = typeCellRender;
        break;
      case 'fileSize':
        cellRender = fileSizeCellRender;
        break;
      case 'url':
        cellRender = urlCellRender;
        break;
      default:
        cellRender = textCellRender;
        break;
    }

    return (
      <Column
        key={column.property}
        dataField={column.property}
        caption={column.caption}
        cellRender={cellRender}
        width={column.width}
        minWidth={column.minWidth}
        alignment={column.alignment}
        visibleIndex={index}
        allowHiding={column.kind !== 'title'}
        showInColumnChooser={column.kind !== 'title'}
        allowSorting={!!sortProperty}
        sortOrder={sortProperty ? getDxSortOrder(sort, sortProperty) : undefined}
        sortIndex={sortProperty && getDxSortOrder(sort, sortProperty) ? 0 : undefined}
      />
    );
  }, [authorCellRender, dateCellRender, fileSizeCellRender, sort, sortableProperties, textCellRender, titleCellRender, typeCellRender, urlCellRender]);

  return (
    <div
      ref={hostRef}
      className={isFullscreen ? styles.dataGridHostFullscreen : styles.dataGridHost}
    >
      <div className={styles.dataGridSurface}>
        <DataGrid
        ref={gridRef}
        dataSource={gridData}
        keyExpr="key"
        width="100%"
        showBorders={false}
        showColumnLines={false}
        showRowLines={true}
        columnAutoWidth={true}
        columnMinWidth={80}
        allowColumnResizing={true}
        allowColumnReordering={true}
        columnResizingMode="nextColumn"
        remoteOperations={{ sorting: true }}
        rowAlternationEnabled={true}
        hoverStateEnabled={true}
        onRowClick={handleRowClick}
        onOptionChanged={handleGridOptionChanged}
        onContentReady={syncVisibleColumnsFromGrid}
        height={estimatedGridHeight}
        wordWrapEnabled={false}
      >
      <Paging enabled={false} />
      <Pager visible={false} />
      <Sorting mode="single" />
      {/* Column preferences — persists visibility, widths, and order per search context */}
      <StateStoring
        enabled={true}
        type="custom"
        customLoad={handleLoadGridState}
        customSave={handleSaveGridState}
        savingTimeout={500}
      />
      {/* Scrolling — virtual mode when there are multiple pages of results */}
      <Scrolling mode={useVirtualScrolling ? 'virtual' : 'standard'} />

      {/* Toolbar — fullscreen toggle and column chooser */}
      <Toolbar>
        <Item
          location="before"
          render={(): React.ReactElement => (
            <button
              className={styles.gridToolbarButton}
              onClick={(): void => setIsFullscreen((prev) => !prev)}
              title={isFullscreen ? 'Exit full view' : 'Open grid in full view'}
              type="button"
            >
              <span className={styles.gridToolbarIcon} aria-hidden="true">{isFullscreen ? '⤢' : '⤢'}</span>
              {isFullscreen ? 'Exit full view' : 'Full view'}
            </button>
          )}
        />
        <Item
          location="before"
          render={(): React.ReactElement => (
            <DefaultButton
              iconProps={{ iconName: 'BulletedList' }}
              text="Columns"
              title="Choose visible columns"
              ariaLabel="Choose visible columns"
              className={styles.gridToolbarButton}
              menuProps={{ items: columnMenuItems }}
            />
          )}
        />
      </Toolbar>

      {columnConfigs.map(renderColumn)}
        </DataGrid>
      </div>
      {isFullscreen && showPaging && totalCount > pageSize && (
        <Pagination
          currentPage={currentPage}
          totalCount={totalCount}
          pageSize={pageSize}
          showPaging={showPaging}
          pageRange={pageRange}
          onPageChange={onPageChange}
        />
      )}
    </div>
  );
};

export default DataGridContent;
