import { ISearchDataProvider } from './ISearchDataProvider';
import { ISuggestionProvider } from './ISuggestionProvider';
import { IActionProvider } from './IActionProvider';
import { ILayoutDefinition } from './ILayoutDefinition';
import { IFilterTypeDefinition } from './IFilterTypeDefinition';

/**
 * Generic typed registry. All registries freeze after
 * first search execution to prevent mid-session mutations.
 */
export interface IRegistry<T extends { id: string }> {
  /** Register a provider. Duplicate IDs warn, first wins. Use force=true to override. */
  register: (provider: T, force?: boolean) => void;
  /** Get provider by ID */
  get: (id: string) => T | undefined;
  /** Get all registered providers */
  getAll: () => T[];
  /** Lock registry — prevents further mutations */
  freeze: () => void;
  /** Check if registry is frozen */
  isFrozen: () => boolean;
}

/**
 * Container for all per-store registries.
 * Each store instance has its own registry container.
 */
export interface IRegistryContainer {
  dataProviders: IRegistry<ISearchDataProvider>;
  suggestions: IRegistry<ISuggestionProvider>;
  actions: IRegistry<IActionProvider>;
  layouts: IRegistry<ILayoutDefinition>;
  filterTypes: IRegistry<IFilterTypeDefinition>;
}
