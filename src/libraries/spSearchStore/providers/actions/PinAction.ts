import type { IActionProvider, ISearchContext, ISearchResult } from '@interfaces/index';
import { getManagerService } from '@store/store';

export class PinAction implements IActionProvider {
  public readonly id: string = 'pin';
  public readonly label: string = 'Pin';
  public readonly iconName: string = 'Pin';
  public readonly position: 'toolbar' | 'contextMenu' | 'both' = 'toolbar';
  public readonly isBulkEnabled: boolean = true;

  public isApplicable(item: ISearchResult): boolean {
    return !!item.url;
  }

  public async execute(items: ISearchResult[], context: ISearchContext): Promise<void> {
    if (!items || items.length === 0) {
      return;
    }

    const collectionName = window.prompt('Pin to collection', 'Favorites');
    if (!collectionName || !collectionName.trim()) {
      return;
    }

    const service = getManagerService(context.searchContextId);
    if (!service) {
      throw new Error('Search Manager is not initialized.');
    }

    for (let i = 0; i < items.length; i++) {
      const item = items[i];
      await service.pinToCollection(
        collectionName.trim(),
        item.url,
        item.title || item.url,
        item.properties || {}
      );
    }
  }
}
