import type { IActionProvider, ISearchContext, ISearchResult } from '@interfaces/index';
import { buildDownloadUrl } from './actionUtils';

export class DownloadAction implements IActionProvider {
  public readonly id: string = 'download';
  public readonly label: string = 'Download';
  public readonly iconName: string = 'Download';
  public readonly position: 'toolbar' | 'contextMenu' | 'both' = 'toolbar';
  public readonly isBulkEnabled: boolean = true;

  public isApplicable(item: ISearchResult): boolean {
    return !!item.url;
  }

  public async execute(items: ISearchResult[], _context: ISearchContext): Promise<void> {
    for (let i = 0; i < items.length; i++) {
      const downloadUrl = buildDownloadUrl(items[i].url);
      if (downloadUrl) {
        window.open(downloadUrl, '_blank', 'noopener,noreferrer');
      }
    }
  }
}
