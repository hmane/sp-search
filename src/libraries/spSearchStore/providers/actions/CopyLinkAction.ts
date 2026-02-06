import type { IActionProvider, ISearchContext, ISearchResult } from '@interfaces/index';
import { buildShareLines, copyTextToClipboard } from './actionUtils';

export class CopyLinkAction implements IActionProvider {
  public readonly id: string = 'copyLink';
  public readonly label: string = 'Copy link';
  public readonly iconName: string = 'Link';
  public readonly position: 'toolbar' | 'contextMenu' | 'both' = 'both';
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
