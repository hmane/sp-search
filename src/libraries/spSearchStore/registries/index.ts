export { Registry } from './Registry';

// Typed registry aliases for readability
import { Registry } from './Registry';
import {
  ISearchDataProvider,
  ISuggestionProvider,
  IActionProvider,
  ILayoutDefinition,
  IFilterTypeDefinition,
  IRegistryContainer
} from '@interfaces/index';

export type DataProviderRegistry = Registry<ISearchDataProvider>;
export type SuggestionProviderRegistry = Registry<ISuggestionProvider>;
export type ActionProviderRegistry = Registry<IActionProvider>;
export type LayoutRegistry = Registry<ILayoutDefinition>;
export type FilterTypeRegistry = Registry<IFilterTypeDefinition>;

/**
 * Create a fresh registry container with empty registries.
 */
export function createRegistryContainer(): IRegistryContainer {
  return {
    dataProviders: new Registry<ISearchDataProvider>('DataProvider'),
    suggestions: new Registry<ISuggestionProvider>('SuggestionProvider'),
    actions: new Registry<IActionProvider>('ActionProvider'),
    layouts: new Registry<ILayoutDefinition>('Layout'),
    filterTypes: new Registry<IFilterTypeDefinition>('FilterType'),
  };
}
