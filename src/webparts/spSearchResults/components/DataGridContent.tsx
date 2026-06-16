import * as React from 'react';
import * as ReactDom from 'react-dom';
import DataGrid, { Column, Scrolling, Sorting, StateStoring, Toolbar, Item, Paging, Pager, LoadPanel } from 'devextreme-react/data-grid';
import { IconButton, DefaultButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { IContextualMenuItem } from '@fluentui/react/lib/ContextualMenu';
import { confirm } from 'spfx-toolkit/lib/utilities/dialogService';
import { ISearchResult, ISortField, ISortableProperty } from '@interfaces/index';
import { PermissionKind } from '@pnp/sp/security';
import { hasPermissions } from '@pnp/sp/security/funcs';
import { buildDownloadUrl, copyTextToClipboard } from '@providers/actions/actionUtils';
import { spLog } from '@store/utils/spLog';
import { buildFormUrl, buildBrowserOpenUrl, formatTitleText, TitleDisplayMode } from './documentTitleUtils';
import { resolveResultLink, type IResultLinkConfig } from './resultLink';
import { getFileTypeIconProps } from '@fluentui/react-file-type-icons';
import DocumentTitleHoverCard from './DocumentTitleHoverCard';
import {
  IColumnConfigItem,
  ColumnRenderer,
  normalizeColumnConfigItem,
} from './ColumnConfigField/columnConfig';
import {
  renderText,
  renderRichText,
  renderNumber,
  renderFileSize,
  renderBoolean,
  renderTags,
  renderPersona,
  renderDate,
  renderUrl,
  renderFileType,
  cleanSearchResultDisplayText,
} from './renderCell';
import Pagination from './Pagination';
import AddToCollectionButton from './AddToCollectionButton';
import { buildAddToCollectionMenuItem } from './buildRowActionMenu';
import styles from './SpSearchResults.module.scss';

export interface IDataGridContentProps {
  items: ISearchResult[];
  /** Stream B / Phase 1 — full IColumnConfigItem[] (alias / width / renderer / etc.). */
  columns: IColumnConfigItem[];
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
  /** Stream B / Phase 3 — when false, the "Columns" toolbar button is hidden. */
  showColumnChooser: boolean;
  sort: ISortField | undefined;
  sortableProperties: ISortableProperty[];
  onPreviewItem?: (item: ISearchResult) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
  onPageChange: (page: number) => void;
  onSortChange: (sort: ISortField) => void;
  // Stream C / #7
  linkConfig: IResultLinkConfig;
  onOpenInSidePanel?: (item: ISearchResult) => void;
}

type ColumnKind =
  | 'title'
  | 'author'      // auto-detected from property name (Author / DisplayAuthor) — uses the in-component avatar render
  | 'date'
  | 'fileType'
  | 'fileSize'
  | 'url'
  | 'text'
  // Stream B / Phase 2 — admin-picked renderer types dispatched via renderCell.tsx
  | 'persona'
  | 'richText'
  | 'number'
  | 'tags'
  | 'boolean';

interface IGridColumnConfig {
  property: string;
  caption: string;
  kind: ColumnKind;
  /** Carries the source column-config item through to the renderer dispatch. */
  column: IColumnConfigItem;
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

const DEFAULT_COLUMNS: IColumnConfigItem[] = [
  normalizeColumnConfigItem({ uniqueId: 'dx-default-0', property: 'Author', alias: 'Author' }),
  normalizeColumnConfigItem({ uniqueId: 'dx-default-1', property: 'LastModifiedTime', alias: 'Modified' }),
  normalizeColumnConfigItem({ uniqueId: 'dx-default-2', property: 'FileType', alias: 'Type' }),
  normalizeColumnConfigItem({ uniqueId: 'dx-default-3', property: 'Size', alias: 'Size' }),
  normalizeColumnConfigItem({ uniqueId: 'dx-default-4', property: 'SiteTitle', alias: 'Site' }),
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

const TITLE_COLUMN: IColumnConfigItem = normalizeColumnConfigItem({
  uniqueId: 'dx-title',
  property: 'Title',
  alias: 'Name',
  visibility: 'always',
});

function getConfiguredColumns(columns: IColumnConfigItem[]): IColumnConfigItem[] {
  const source = [TITLE_COLUMN, ...(columns.length > 0 ? columns : DEFAULT_COLUMNS)];
  const seen = new Set<string>();
  const unique: IColumnConfigItem[] = [];

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
    unique.push(source[i]);
  }

  return unique;
}

/**
 * Auto-detect a column's kind by property name. This is the pre-Phase-1
 * behaviour, reachable when `IColumnConfigItem.renderer === ''` (the migration
 * sentinel) and for the Title column.
 */
function autoDetectColumnKind(property: string): { kind: ColumnKind; width?: number; minWidth?: number; alignment?: 'left' | 'center' | 'right' } {
  const normalized = property.toLowerCase();

  if (normalized === 'title' || normalized === 'filename') {
    return { kind: 'title', minWidth: 220 };
  }
  if (normalized === 'author' || normalized === 'authorowsuser' || normalized === 'displayauthor') {
    return { kind: 'author', width: 180 };
  }
  if (normalized === 'lastmodifiedtime' || normalized === 'modified' || normalized === 'created') {
    return { kind: 'date', width: 140 };
  }
  if (normalized === 'filetype' || normalized === 'fileextension') {
    return { kind: 'fileType', width: 80, alignment: 'center' };
  }
  if (normalized === 'size' || normalized === 'filesize') {
    return { kind: 'fileSize', width: 90, alignment: 'right' };
  }
  if (normalized === 'path' || normalized.indexOf('url') >= 0 || normalized.indexOf('link') >= 0) {
    return { kind: 'url', minWidth: 220 };
  }
  return { kind: 'text', minWidth: 140 };
}

/**
 * Map an explicit `IColumnConfigItem.renderer` choice onto a cell renderer
 * kind. `''` defers to the auto-detect path so migrated items render
 * identically to today. The Phase-2 renderer types (`persona`, `richText`,
 * `number`, `tags`, `boolean`) each have their own kind dispatched via the
 * pure renderers in `renderCell.tsx`.
 */
function kindFromRenderer(renderer: ColumnRenderer, property: string): { kind: ColumnKind; width?: number; minWidth?: number; alignment?: 'left' | 'center' | 'right' } {
  switch (renderer) {
    case 'text':     return { kind: 'text', minWidth: 140 };
    case 'date':     return { kind: 'date', width: 140 };
    case 'fileType': return { kind: 'fileType', width: 80, alignment: 'center' };
    case 'fileSize': return { kind: 'fileSize', width: 90, alignment: 'right' };
    case 'url':      return { kind: 'url', minWidth: 220 };
    case 'persona':  return { kind: 'persona', width: 200 };
    case 'richText': return { kind: 'richText', minWidth: 200 };
    case 'number':   return { kind: 'number', width: 100, alignment: 'right' };
    case 'tags':     return { kind: 'tags', minWidth: 160 };
    case 'boolean':  return { kind: 'boolean', width: 70, alignment: 'center' };
    case '':
    default:
      return autoDetectColumnKind(property);
  }
}

function getColumnConfig(column: IColumnConfigItem): IGridColumnConfig {
  const property = column.property;
  const isTitle = property.toLowerCase() === 'title' || property.toLowerCase() === 'filename';
  // Title always renders via the title cell renderer regardless of admin choice.
  const dispatch = isTitle ? autoDetectColumnKind(property) : kindFromRenderer(column.renderer, property);
  const alias = (column.alias || '').trim();
  const caption = alias || property;
  const adminWidth = column.width;

  return {
    property,
    caption,
    kind: dispatch.kind,
    column,
    width: adminWidth !== undefined ? adminWidth : dispatch.width,
    minWidth: dispatch.minWidth,
    alignment: dispatch.alignment,
  };
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
    return cleanSearchResultDisplayText(value);
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
      return cleanSearchResultDisplayText(value.displayText);
    }
    return JSON.stringify(value);
  }
  return String(value);
}

// Author / people managed properties can carry several values joined by ';'
// (co-authors, plus service/app accounts that created or modified the file).
// Split on ';' ONLY — never comma — because a display name may itself contain a
// comma ("Lastname, Firstname").
function splitPeopleValue(value: unknown): string[] {
  if (Array.isArray(value)) {
    return value
      .map((entry: unknown) => formatTextValue(entry).trim())
      .filter((s: string) => s.length > 0 && s !== '--');
  }
  const text = formatTextValue(value).trim();
  if (!text || text === '--') {
    return [];
  }
  return text.split(/\s*;\s*/).map((s: string) => s.trim()).filter(Boolean);
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
  return 'sp-search-grid-v2-cols-' + searchContextId;
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

    const allowed = new Set(currentColumnKeys.map((key) => key.toLowerCase()));
    const seen = new Set<string>();
    const sanitizedColumns = parsed.columns
      .filter((column: { dataField?: string }) => {
        const dataField = String(column.dataField || '').trim();
        if (!dataField) {
          return false;
        }
        const normalized = dataField.toLowerCase();
        if (!allowed.has(normalized) || seen.has(normalized)) {
          return false;
        }
        seen.add(normalized);
        return true;
      })
      .map((column: {
        dataField?: string;
        visible?: boolean;
        width?: number | string;
        visibleIndex?: number;
      }) => ({
        dataField: column.dataField,
        visible: column.visible,
        width: column.width,
        ...(allowSavedOrder && typeof column.visibleIndex === 'number' ? { visibleIndex: column.visibleIndex } : {})
      }));

    return {
      columns: sanitizedColumns
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
    columns,
    titleDisplayMode,
    totalCount,
    pageSize,
    currentPage,
    showPaging,
    pageRange,
    searchContextId,
    showDeleteConfirmation,
    showColumnChooser,
    sort,
    sortableProperties,
    onItemClick,
    onPageChange,
    onSortChange,
    linkConfig,
    onOpenInSidePanel
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
  const [copiedUrlToast, setCopiedUrlToast] = React.useState<string | undefined>(undefined);

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
    (): IGridColumnConfig[] => getConfiguredColumns(columns).map(getColumnConfig),
    [columns]
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

  React.useEffect((): (() => void) => {
    const permissionItems = items.filter((item) => !!getDeleteInfo(item) && !permissionCache[item.key]);
    if (permissionItems.length === 0) {
      return (): void => { /* noop */ };
    }

    let isActive = true;
    Promise.all(permissionItems.map(async (item) => {
      try {
        const state = await loadItemPermissions(item);
        return [item.key, state] as const;
      } catch {
        return [item.key, { checked: true, canEdit: false, canDelete: false } satisfies IItemPermissionState] as const;
      }
    }))
      .then((entries): void => {
        if (!isActive || entries.length === 0) {
          return;
        }
        setPermissionCache((prev) => {
          const next = { ...prev };
          for (let i: number = 0; i < entries.length; i++) {
            next[entries[i][0]] = entries[i][1];
          }
          return next;
        });
      })
      .catch((): void => {
        // Swallow permission prefetch failures — per-item actions degrade safely.
      });

    return (): void => {
      isActive = false;
    };
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

  React.useEffect((): (() => void) | void => {
    if (!copiedUrlToast) {
      return;
    }

    const timeoutId = window.setTimeout((): void => {
      setCopiedUrlToast(undefined);
    }, 3200);

    return (): void => {
      window.clearTimeout(timeoutId);
    };
  }, [copiedUrlToast]);

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
    // Stream B / Phase 3 — visibility-aware filter + initial-check derivation.
    // `always` columns never appear in the menu; `defaultOff` columns start
    // unchecked unless the user has saved their state otherwise.
    return columnConfigs
      .filter((column) => column.kind !== 'title' && column.column.visibility !== 'always')
      .map((column) => {
        const saved = columnVisibility[column.property];
        const isOnByDefault = column.column.visibility !== 'defaultOff';
        const checked = saved === undefined ? isOnByDefault : saved !== false;
        return {
          key: column.property,
          text: column.caption,
          canCheck: true,
          checked,
          onClick: (): void => {
            handleToggleColumnVisibility(column.property);
          }
        };
      });
  }, [columnConfigs, columnVisibility, handleToggleColumnVisibility]);

  const handleExportXlsx = React.useCallback((): void => {
    import(/* webpackChunkName: 'xlsxExport' */ './exportXlsx')
      .then((m): void => {
        m.triggerXlsxDownload(gridData, columnConfigs, 'search-results.xlsx');
      })
      .catch((): void => {
        spLog.error('Failed to load XLSX export module');
      });
  }, [gridData, columnConfigs]);

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  const titleCellRender = React.useCallback((cellData: { value: unknown; data: IGridRow; rowIndex?: number }): React.ReactElement => {
    const matchingItem = cellData.data.__item;
    const position = (cellData.rowIndex !== undefined ? cellData.rowIndex : 0) + 1;
    const title = formatTitleText(formatTextValue(cellData.value), titleDisplayMode);
    const linkProps = resolveResultLink(matchingItem, linkConfig);
    const permissionState = permissionCache[matchingItem.key];
    const viewUrl = buildFormUrl(matchingItem, 4);
    const editUrl = buildFormUrl(matchingItem, 6);
    const buildMenuItems = (
      openAddToCollection: (event?: { preventDefault?: () => void; stopPropagation?: () => void }) => void
    ): IContextualMenuItem[] => [
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
      buildAddToCollectionMenuItem(openAddToCollection),
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
          copyTextToClipboard(matchingItem.url)
            .then((): void => {
              setCopiedUrlToast(matchingItem.url);
            })
            .catch((): void => {
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
                spLog.error('Failed to delete grid item', { error });
                window.alert('Unable to delete this item.');
              });
          })().catch((error: Error): void => {
            spLog.error('Failed to open delete confirmation', { error });
          });
        }
      }] : [])
    ];

    return (
      <DocumentTitleHoverCard
        item={matchingItem}
        position={position}
        onItemClick={onItemClick}
        hostDisplay="block"
        clickTarget={linkConfig.clickTarget}
        onOpenInSidePanel={onOpenInSidePanel}
      >
        {(handleClick): React.ReactNode => (
          <div className={styles.gridTitleCell}>
            <span className={styles.gridTitleIcon}>
              <Icon {...getFileTypeIconProps({ extension: matchingItem.fileType || '', size: 16 })} />
            </span>
            <div className={styles.gridTitleMain}>
              <a
                href={linkProps.href}
                target={linkProps.target}
                rel={linkProps.rel}
                // SharePoint Modern's SPA router intercepts <a> clicks in capture
                // phase and navigates the current tab itself — bypassing target=_blank
                // AND our React onClick's preventDefault. data-interception="off"
                // opts this anchor out so the browser handles the click natively
                // and our handleClick can Modal-ize the previewable.
                data-interception="off"
                className={titleDisplayMode === 'wrap' ? styles.gridTitleLinkWrap : styles.gridTitleLink}
                onClick={(e: React.MouseEvent): void => {
                  e.stopPropagation();
                  handleClick(e);
                }}
                >
                {title}
              </a>
              <div className={styles.gridTitleActions}>
                <AddToCollectionButton
                  item={matchingItem}
                  searchContextId={searchContextId}
                  triggerRenderer={(openAddToCollection): React.ReactNode => (
                    <IconButton
                      iconProps={{ iconName: 'MoreVertical' }}
                      title="More actions"
                      ariaLabel="More actions"
                      menuProps={{ items: buildMenuItems(openAddToCollection) }}
                      className={styles.gridTitleActionButton}
                      onClick={(event): void => {
                        event.stopPropagation();
                      }}
                    />
                  )}
                />
              </div>
            </div>
          </div>
        )}
      </DocumentTitleHoverCard>
    );
  }, [linkConfig, onItemClick, onOpenInSidePanel, permissionCache, searchContextId, showDeleteConfirmation, titleDisplayMode]);

  const authorCellRender = React.useCallback((cellData: { value: unknown }): React.ReactElement => {
    const people = splitPeopleValue(cellData.value);

    if (people.length === 0) {
      return <span className={styles.gridCellMuted}>--</span>;
    }

    // One row per person — each gets its own avatar with that person's correct
    // initials/color (a multi-value Author field shows several stacked rows).
    return (
      <div className={styles.gridAuthorList}>
        {people.map((person: string, idx: number) => (
          <div key={person + '-' + String(idx)} className={styles.gridAuthorCell}>
            <span
              className={styles.gridAuthorAvatar}
              style={{ backgroundColor: getInitialsColor(person) }}
              title={person}
            >
              {getInitials(person)}
            </span>
            <span className={styles.gridAuthorName}>{person}</span>
          </div>
        ))}
      </div>
    );
  }, []);

  // Date / fileType / fileSize / url / text renderers have moved to the pure
  // `renderCell.tsx` module as part of Stream B / Phase 2 — they're dispatched
  // directly inside `renderColumn`. The title and author cell renders stay
  // here because they close over local state (hover-card, permissions, etc.).

  const renderColumn = React.useCallback((column: IGridColumnConfig, index: number): React.ReactElement => {
    const sortProperty = resolveColumnSortProperty(column.property, sortableProperties);
    const cfg = column.column;
    let cellRender:
      | ((cellData: { value: unknown; data: IGridRow; rowIndex?: number }) => React.ReactElement)
      | undefined;

    switch (column.kind) {
      case 'title':
        cellRender = titleCellRender;
        break;
      case 'author':
        // Auto-detected author column — keep today's simple initials avatar.
        cellRender = authorCellRender;
        break;
      case 'date':
        cellRender = (cellData): React.ReactElement => renderDate(cellData.value, cfg);
        break;
      case 'fileType':
        cellRender = (cellData): React.ReactElement => renderFileType(cellData.value, cfg);
        break;
      case 'fileSize':
        cellRender = (cellData): React.ReactElement => renderFileSize(cellData.value, cfg);
        break;
      case 'url':
        cellRender = (cellData): React.ReactElement => {
          const value = typeof cellData.value === 'string' && cellData.value ? cellData.value : cellData.data.__item.url;
          return renderUrl(value, cfg);
        };
        break;
      // Stream B / Phase 2 — admin-picked explicit renderers
      case 'persona':
        cellRender = (cellData): React.ReactElement => renderPersona(cellData.value, cfg);
        break;
      case 'richText':
        cellRender = (cellData): React.ReactElement => renderRichText(cellData.value, cfg);
        break;
      case 'number':
        cellRender = (cellData): React.ReactElement => renderNumber(cellData.value, cfg);
        break;
      case 'tags':
        cellRender = (cellData): React.ReactElement => renderTags(cellData.value, cfg);
        break;
      case 'boolean':
        cellRender = (cellData): React.ReactElement => renderBoolean(cellData.value, cfg);
        break;
      default:
        cellRender = (cellData): React.ReactElement => renderText(cellData.value, cfg);
        break;
    }

    const isTitle = column.kind === 'title';
    const isAlways = cfg.visibility === 'always';
    const allowHiding = !isTitle && !isAlways;
    // Initial visibility: derived from the column-config item's `visibility`
    // (defaultOff starts hidden, defaultOn/always start visible). Saved state
    // from localStorage takes precedence via StateStoring once it loads.
    const savedVisible = columnVisibility[column.property];
    const initialVisible = savedVisible === undefined
      ? cfg.visibility !== 'defaultOff'
      : savedVisible !== false;

    return (
      <Column
        key={column.property}
        dataField={column.property}
        caption={column.caption}
        cellRender={cellRender}
        width={column.width}
        minWidth={column.minWidth}
        alignment={column.alignment}
        allowHiding={allowHiding}
        showInColumnChooser={allowHiding}
        visible={initialVisible}
        allowSorting={!!sortProperty}
        sortOrder={sortProperty ? getDxSortOrder(sort, sortProperty) : undefined}
        sortIndex={sortProperty && getDxSortOrder(sort, sortProperty) ? 0 : undefined}
      />
    );
  }, [authorCellRender, columnVisibility, sort, sortableProperties, titleCellRender]);

  return (
    <div
      ref={hostRef}
      className={isFullscreen ? styles.dataGridHostFullscreen : styles.dataGridHost}
    >
      {copiedUrlToast && typeof document !== 'undefined' && ReactDom.createPortal(
        <div className={styles.gridCopyToastViewport} role="status" aria-live="polite">
          <div className={styles.gridCopyToast}>
            <div className={styles.gridCopyToastHeader}>
              <Icon iconName="StatusCircleCheckmark" className={styles.gridCopyToastIcon} />
              <span className={styles.gridCopyToastTitle}>Link copied</span>
            </div>
            <div className={styles.gridCopyToastUrl} title={copiedUrlToast}>
              {copiedUrlToast}
            </div>
          </div>
        </div>,
        document.body
      )}
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
      <LoadPanel enabled={false} showPane={false} />
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

      {/* Toolbar — fullscreen toggle, column chooser, and XLSX export */}
      <Toolbar>
        {showColumnChooser && (
          <Item
            location="after"
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
        )}
        <Item
          location="after"
          render={(): React.ReactElement => (
            <TooltipHost content="Export to Excel">
              <IconButton
                iconProps={{ iconName: 'ExcelDocument' }}
                ariaLabel="Export to Excel"
                className={styles.gridToolbarIconBtn}
                onClick={handleExportXlsx}
              />
            </TooltipHost>
          )}
        />
        <Item
          location="after"
          render={(): React.ReactElement => (
            <TooltipHost content={isFullscreen ? 'Exit full view' : 'Expand to full view'}>
              <IconButton
                iconProps={{ iconName: isFullscreen ? 'BackToWindow' : 'FullScreen' }}
                ariaLabel={isFullscreen ? 'Exit full view' : 'Expand to full view'}
                onClick={(): void => setIsFullscreen((prev) => !prev)}
                className={styles.gridToolbarIconBtn}
              />
            </TooltipHost>
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
