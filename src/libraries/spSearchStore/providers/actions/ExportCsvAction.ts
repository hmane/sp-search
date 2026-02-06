import type { IActionProvider, ISearchContext, ISearchResult } from '@interfaces/index';
import { normalizeUrl } from './actionUtils';

function escapeCsv(value: string): string {
  if (value === undefined || value === null) {
    return '';
  }
  const text = String(value).replace(/\r?\n/g, ' ');
  if (text.indexOf('"') >= 0 || text.indexOf(',') >= 0) {
    return '"' + text.replace(/"/g, '""') + '"';
  }
  return text;
}

function buildCsv(items: ISearchResult[]): string {
  const headers = [
    'Title',
    'URL',
    'Author',
    'Modified',
    'Created',
    'File Type',
    'File Size',
    'Site',
    'Site URL'
  ];

  const rows = items.map(function (item): string {
    const values = [
      item.title || '',
      normalizeUrl(item.url),
      item.author?.displayText || '',
      item.modified || '',
      item.created || '',
      item.fileType || '',
      item.fileSize ? String(item.fileSize) : '',
      item.siteName || '',
      normalizeUrl(item.siteUrl || '')
    ];
    return values.map(escapeCsv).join(',');
  });

  return headers.map(escapeCsv).join(',') + '\r\n' + rows.join('\r\n');
}

function triggerDownload(csv: string): void {
  const blob = new Blob([csv], { type: 'text/csv;charset=utf-8' });
  const url = URL.createObjectURL(blob);
  const link = document.createElement('a');
  link.href = url;
  link.download = 'search-results-' + new Date().toISOString().slice(0, 10) + '.csv';
  document.body.appendChild(link);
  link.click();
  document.body.removeChild(link);
  URL.revokeObjectURL(url);
}

export class ExportCsvAction implements IActionProvider {
  public readonly id: string = 'exportCsv';
  public readonly label: string = 'Export CSV';
  public readonly iconName: string = 'ExcelDocument';
  public readonly position: 'toolbar' | 'contextMenu' | 'both' = 'toolbar';
  public readonly isBulkEnabled: boolean = true;

  public isApplicable(item: ISearchResult): boolean {
    return !!item.url;
  }

  public async execute(items: ISearchResult[], _context: ISearchContext): Promise<void> {
    if (!items || items.length === 0) {
      return;
    }
    const csv = buildCsv(items);
    triggerDownload(csv);
  }
}
