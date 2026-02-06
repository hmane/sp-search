import type { IActionProvider, ISearchContext, ISearchResult } from '@interfaces/index';
import { normalizeUrl } from './actionUtils';

export class OpenAction implements IActionProvider {
  public readonly id: string = 'open';
  public readonly label: string = 'Open';
  public readonly iconName: string = 'OpenInNewTab';
  public readonly position: 'toolbar' | 'contextMenu' | 'both' = 'both';
  public readonly isBulkEnabled: boolean = true;

  public isApplicable(item: ISearchResult): boolean {
    return !!item.url;
  }

  public async execute(items: ISearchResult[], _context: ISearchContext): Promise<void> {
    for (let i = 0; i < items.length; i++) {
      const url: string = normalizeUrl(items[i].url);
      if (url) {
        window.open(url, '_blank', 'noopener,noreferrer');
      }
    }
  }
}
