/**
 * T2.D11 — layout-agnostic CSV/XLSX export.
 *
 * Today CSV/XLSX export is DataGrid-only (DevExtreme's built-in
 * Toolbar). The other five layouts (List, Compact, Card, People,
 * Gallery) have no export. This module wraps the existing
 * `triggerXlsxDownload` helper + adds a CSV serialiser so the
 * ResultToolbar Export menu works regardless of which layout is
 * active.
 *
 * "Export selection" semantics consume the T2.D2 bulk-selection
 * state — when the caller passes the selected `itemKey` set, only
 * those rows export. Otherwise all currently-rendered items export.
 */

import type { ISearchResult } from '@interfaces/index';
import type { IXlsxColumn, IXlsxRow } from './exportXlsx';

/**
 * Minimal column shape exporters consume. Wider than
 * `IColumnConfigItem` so callers can pass either the grid's full
 * column config or the simpler `selectedPropertyColumns` shape.
 */
export interface IExportableColumn {
  property: string;
  alias?: string;
  renderer?: string;
}

/** Default column set when no admin columns are configured. */
const DEFAULT_COLUMNS: IXlsxColumn[] = [
  { property: 'title', caption: 'Title', kind: 'text' },
  { property: 'url', caption: 'URL', kind: 'url' },
  { property: 'authorDisplay', caption: 'Author', kind: 'text' },
  { property: 'modified', caption: 'Modified', kind: 'date' },
  { property: 'fileType', caption: 'File type', kind: 'text' },
  { property: 'fileSize', caption: 'Size', kind: 'fileSize' },
  { property: 'siteName', caption: 'Site', kind: 'text' },
];

function columnFromConfig(c: IExportableColumn): IXlsxColumn {
  const kind: string = c.renderer || 'text';
  return {
    property: c.property,
    caption: c.alias || c.property,
    kind,
  };
}

function searchResultToRow(item: ISearchResult, _columns: IXlsxColumn[]): IXlsxRow {
  // Build a row keyed by every property a column might want. Adapters
  // flatten the ISearchResult shape (nested author, top-level
  // siteName / modified, properties bag) into a flat key-value map so
  // CSV/XLSX serialisation can read `row[col.property]` uniformly.
  const row: IXlsxRow = {
    key: item.key,
  };
  row.title = item.title;
  row.url = item.url;
  row.authorDisplay = item.author && item.author.displayText ? item.author.displayText : '';
  row.authorEmail = item.author && item.author.email ? item.author.email : '';
  row.modified = item.modified ? item.modified : '';
  row.created = item.created ? item.created : '';
  row.fileType = item.fileType || '';
  row.fileSize = typeof item.fileSize === 'number' ? item.fileSize : 0;
  row.siteName = item.siteName || '';
  // Flatten the admin-property bag so column.property === 'CustomMP'
  // resolves to item.properties.CustomMP.
  if (item.properties) {
    for (const k of Object.keys(item.properties)) {
      if (row[k] === undefined) {
        row[k] = item.properties[k];
      }
    }
  }
  return row;
}

export interface IExportItemsOptions {
  /** When provided, only items whose `key` is in this set export. */
  selectedItemKeys?: string[];
  /** Default-property column set when the layout has no admin-configured columns. */
  configuredColumns?: IExportableColumn[];
  /** File name without extension. Defaults to "search-results-<timestamp>". */
  baseFileName?: string;
}

function deriveColumns(opts: IExportItemsOptions): IXlsxColumn[] {
  if (opts.configuredColumns && opts.configuredColumns.length > 0) {
    return opts.configuredColumns.map(columnFromConfig);
  }
  return DEFAULT_COLUMNS;
}

function filterItems(items: ISearchResult[], opts: IExportItemsOptions): ISearchResult[] {
  if (!opts.selectedItemKeys || opts.selectedItemKeys.length === 0) {
    return items;
  }
  const set = new Set(opts.selectedItemKeys);
  return items.filter((i) => set.has(i.key));
}

function buildFileName(opts: IExportItemsOptions, extension: 'csv' | 'xlsx'): string {
  const ts = new Date().toISOString().substring(0, 19).replace(/[:T]/g, '-');
  return (opts.baseFileName || 'search-results-' + ts) + '.' + extension;
}

/**
 * Trigger an XLSX download for the current set of items. Lazy-loads
 * the `xlsx` chunk only on first use.
 */
export function exportItemsAsXlsx(items: ISearchResult[], opts: IExportItemsOptions = {}): Promise<void> {
  const columns = deriveColumns(opts);
  const filtered = filterItems(items, opts);
  const rows: IXlsxRow[] = filtered.map((it) => searchResultToRow(it, columns));
  const filename = buildFileName(opts, 'xlsx');
  return import(/* webpackChunkName: 'xlsxExport' */ './exportXlsx').then((mod) => {
    mod.triggerXlsxDownload(rows, columns, filename);
  });
}

/**
 * Trigger a CSV download for the current set of items. No async chunk
 * load — CSV serialisation is in-line. Matches the column rendering
 * the XLSX path uses, modulo the date/fileSize Excel-native types
 * (CSV emits the formatted strings).
 */
export function exportItemsAsCsv(items: ISearchResult[], opts: IExportItemsOptions = {}): void {
  const columns = deriveColumns(opts);
  const filtered = filterItems(items, opts);
  const rows: IXlsxRow[] = filtered.map((it) => searchResultToRow(it, columns));

  const headerLine = columns.map((c) => csvEscape(c.caption)).join(',');
  const dataLines = rows.map((row) =>
    columns.map((col) => csvEscape(formatCsvCell(row[col.property], col.kind))).join(',')
  );
  const csv = [headerLine, ...dataLines].join('\r\n');

  // BOM ensures Excel auto-detects UTF-8 on Windows.
  const blob = new Blob(['﻿' + csv], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = buildFileName(opts, 'csv');
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

function csvEscape(value: unknown): string {
  if (value === undefined || value === null) { return ''; }
  const str = String(value);
  // Wrap in quotes when the value contains comma, quote, newline, or
  // leading/trailing whitespace.
  if (/[",\r\n]/.test(str) || str !== str.trim()) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

function formatCsvCell(value: unknown, kind: string): string {
  if (value === undefined || value === null || value === '') { return ''; }
  if (kind === 'date') {
    const d = typeof value === 'string' ? new Date(value) : value as Date;
    if (d instanceof Date && !isNaN(d.getTime())) {
      return d.toISOString();
    }
    return String(value);
  }
  if (kind === 'fileSize') {
    return String(value);
  }
  if (Array.isArray(value)) {
    return (value as unknown[]).map((v) => String(v)).join('; ');
  }
  if (typeof value === 'object' && value !== null) {
    if ('displayText' in value && typeof (value as { displayText: unknown }).displayText === 'string') {
      return (value as { displayText: string }).displayText;
    }
    return JSON.stringify(value);
  }
  if (typeof value === 'boolean') {
    return value ? 'Yes' : 'No';
  }
  return String(value);
}
