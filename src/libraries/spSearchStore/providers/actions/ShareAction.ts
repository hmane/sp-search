import type { IActionProvider, ISearchContext, ISearchResult } from '@interfaces/index';
import { buildShareLines, copyTextToClipboard } from './actionUtils';

export class ShareAction implements IActionProvider {
  public readonly id: string = 'share';
  public readonly label: string = 'Share';
  public readonly iconName: string = 'Share';
  public readonly position: 'toolbar' | 'contextMenu' | 'both' = 'toolbar';
  public readonly isBulkEnabled: boolean = true;

  public isApplicable(item: ISearchResult): boolean {
    return !!item.url;
  }

  public async execute(items: ISearchResult[], _context: ISearchContext): Promise<void> {
    if (!items || items.length === 0) {
      return;
    }
    const lines = buildShareLines(items);
    await copyTextToClipboard(lines.join('\n'));
  }
}
