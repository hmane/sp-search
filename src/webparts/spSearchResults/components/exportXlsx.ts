/**
 * XLSX export — lazily imported so the xlsx (SheetJS) library is only
 * bundled into a separate chunk and loaded on first use.
 *
 * This module is the ONLY file that imports from 'xlsx'. Any caller
 * that needs XLSX export should use:
 *
 *   import(/* webpackChunkName: 'xlsxExport' * / './exportXlsx')
 *     .then(m => m.triggerXlsxDownload(rows, columns, filename))
 *     .catch(() => { ... });
 */

import * as XLSX from 'xlsx';

export interface IXlsxRow {
  [property: string]: unknown;
  key: string;
}

export interface IXlsxColumn {
  property: string;
  caption: string;
  kind: string;
}

/**
 * Builds an XLSX workbook from the provided rows and column configs,
 * then triggers a browser download.
 *
 * Each column caption becomes a header in row 1.
 * Values are formatted the same way the CSV export formats them,
 * so the two exports are consistent.
 */
export function triggerXlsxDownload(
  rows: IXlsxRow[],
  columns: IXlsxColumn[],
  filename: string
): void {
  // Build a 2-D array: [header row, ...data rows]
  const headerRow: string[] = columns.map((col) => col.caption);
  const dataRows: unknown[][] = rows.map((row) =>
    columns.map((col) => formatXlsxCell(row[col.property], col.kind))
  );

  const worksheetData: unknown[][] = [headerRow, ...dataRows];

  const worksheet = XLSX.utils.aoa_to_sheet(worksheetData);

  // Auto-width: set each column width to the length of its longest cell (capped at 60 chars)
  const colWidths = columns.map((col, colIdx) => {
    let max = col.caption.length;
    for (const dataRow of dataRows) {
      const cell = dataRow[colIdx];
      const len = cell !== null && cell !== undefined ? String(cell).length : 0;
      if (len > max) { max = len; }
    }
    return { wch: Math.min(max + 2, 60) };
  });
  worksheet['!cols'] = colWidths;

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Search Results');

  XLSX.writeFile(workbook, filename);
}

function formatXlsxCell(value: unknown, kind: string): unknown {
  if (value === undefined || value === null || value === '') {
    return '';
  }

  switch (kind) {
    case 'date': {
      // Return a real Date object so Excel formats it as a date cell
      const raw = typeof value === 'string' ? value : '';
      if (!raw) { return ''; }
      const d = new Date(raw);
      return isNaN(d.getTime()) ? raw : d;
    }
    case 'fileSize': {
      // Return raw number so Excel can format/sort numerically
      const n = typeof value === 'number' ? value : parseInt(String(value || '0'), 10);
      return isNaN(n) ? 0 : n;
    }
    case 'url':
      return typeof value === 'string' ? value : String(value);
    default:
      return formatDisplayValue(value);
  }
}

function formatDisplayValue(value: unknown): unknown {
  if (typeof value === 'string') { return value; }
  if (typeof value === 'number') { return value; }
  if (typeof value === 'boolean') { return value ? 'Yes' : 'No'; }
  if (Array.isArray(value)) {
    return value.map((v) => formatDisplayValue(v)).join(', ');
  }
  if (typeof value === 'object' && value !== null) {
    if ('displayText' in value && typeof (value as { displayText: unknown }).displayText === 'string') {
      return (value as { displayText: string }).displayText;
    }
    return JSON.stringify(value);
  }
  return String(value);
}
