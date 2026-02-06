import type { IActionProvider, ISearchContext, ISearchResult } from '@interfaces/index';
import { normalizeUrl } from './actionUtils';

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

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

function buildCompareHtml(items: ISearchResult[]): string {
  const fields = [
    { label: 'Title', getValue: function (item: ISearchResult): string { return item.title || ''; } },
    { label: 'URL', getValue: function (item: ISearchResult): string { return normalizeUrl(item.url); } },
    { label: 'Author', getValue: function (item: ISearchResult): string { return item.author?.displayText || ''; } },
    { label: 'Modified', getValue: function (item: ISearchResult): string { return item.modified || ''; } },
    { label: 'Created', getValue: function (item: ISearchResult): string { return item.created || ''; } },
    { label: 'File Type', getValue: function (item: ISearchResult): string { return item.fileType || ''; } },
    { label: 'File Size', getValue: function (item: ISearchResult): string { return formatFileSize(item.fileSize); } },
    { label: 'Site', getValue: function (item: ISearchResult): string { return item.siteName || ''; } },
    { label: 'Site URL', getValue: function (item: ISearchResult): string { return normalizeUrl(item.siteUrl || ''); } },
  ];

  const headerCells = items.map(function (item): string {
    return '<th>' + escapeHtml(item.title || 'Untitled') + '</th>';
  }).join('');

  const rows = fields.map(function (field): string {
    const cells = items.map(function (item): string {
      const value = field.getValue(item) || '';
      const text = escapeHtml(value);
      if (field.label === 'URL' || field.label === 'Site URL') {
        return '<td><a href="' + escapeHtml(value) + '" target="_blank" rel="noopener noreferrer">' + text + '</a></td>';
      }
      return '<td>' + text + '</td>';
    }).join('');
    return '<tr><th>' + escapeHtml(field.label) + '</th>' + cells + '</tr>';
  }).join('');

  return (
    '<!doctype html>' +
    '<html><head><meta charset="utf-8" />' +
    '<title>Compare Metadata</title>' +
    '<style>' +
    'body{font-family:Segoe UI,Arial,sans-serif;margin:24px;color:#1b1a19;}' +
    'h1{font-size:20px;margin:0 0 16px 0;}' +
    'table{border-collapse:collapse;width:100%;font-size:12px;}' +
    'th,td{border:1px solid #e1dfdd;padding:8px;vertical-align:top;}' +
    'th{background:#f3f2f1;text-align:left;font-weight:600;}' +
    'a{color:#0078d4;text-decoration:none;}' +
    'a:hover{text-decoration:underline;}' +
    '</style>' +
    '</head><body>' +
    '<h1>Metadata Comparison</h1>' +
    '<table><thead><tr><th>Field</th>' + headerCells + '</tr></thead>' +
    '<tbody>' + rows + '</tbody></table>' +
    '</body></html>'
  );
}

export class CompareAction implements IActionProvider {
  public readonly id: string = 'compare';
  public readonly label: string = 'Compare';
  public readonly iconName: string = 'ColumnOptions';
  public readonly position: 'toolbar' | 'contextMenu' | 'both' = 'toolbar';
  public readonly isBulkEnabled: boolean = true;

  public isApplicable(_item: ISearchResult): boolean {
    return true;
  }

  public async execute(items: ISearchResult[], _context: ISearchContext): Promise<void> {
    if (!items || items.length < 2 || items.length > 3) {
      throw new Error('Select 2 or 3 items to compare.');
    }

    const win = window.open('', '_blank', 'noopener,noreferrer');
    if (!win) {
      throw new Error('Popup blocked. Allow pop-ups to compare items.');
    }
    win.document.write(buildCompareHtml(items));
    win.document.close();
  }
}
