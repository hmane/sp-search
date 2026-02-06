import type { IActionProvider, ISearchContext, ISearchResult } from '@interfaces/index';
import { getStore } from '@store/store';

export class PreviewAction implements IActionProvider {
  public readonly id: string = 'preview';
  public readonly label: string = 'Preview';
  public readonly iconName: string = 'View';
  public readonly position: 'toolbar' | 'contextMenu' | 'both' = 'contextMenu';
  public readonly isBulkEnabled: boolean = false;

  public isApplicable(item: ISearchResult): boolean {
    return !!item.url;
  }

  public async execute(items: ISearchResult[], context: ISearchContext): Promise<void> {
    if (!items || items.length === 0) {
      return;
    }
    const store = getStore(context.searchContextId);
    store.getState().setPreviewItem(items[0]);
  }
}
