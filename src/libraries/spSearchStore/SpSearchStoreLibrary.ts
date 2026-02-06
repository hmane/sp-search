import { StoreApi } from 'zustand/vanilla';
import { ISearchStore } from '@interfaces/index';
import { getStore, disposeStore, hasStore } from '@store/store';

// Re-export everything consumers need
export { ISearchStore } from '@interfaces/index';
export { getStore, disposeStore, hasStore } from '@store/store';

/**
 * SPFx Library Component entry point.
 * Web parts consume this library to access shared Zustand stores.
 */
export class SpSearchStoreLibrary {
  public name(): string {
    return 'SpSearchStoreLibrary';
  }

  /**
   * Get or create a store for the given search context ID.
   */
  public getStore(searchContextId: string): StoreApi<ISearchStore> {
    return getStore(searchContextId);
  }

  /**
   * Dispose a store when all web parts in the context are unmounted.
   */
  public disposeStore(searchContextId: string): void {
    disposeStore(searchContextId);
  }

  /**
   * Check if a store exists for the given search context ID.
   */
  public hasStore(searchContextId: string): boolean {
    return hasStore(searchContextId);
  }
}
