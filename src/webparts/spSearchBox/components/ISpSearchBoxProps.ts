import { StoreApi } from 'zustand/vanilla';
import { ISearchStore, ISearchScope } from '@interfaces/index';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface ISpSearchBoxProps {
  store: StoreApi<ISearchStore>;
  placeholder: string;
  debounceMs: number;
  searchBehavior: 'onEnter' | 'onButton' | 'both';
  enableScopeSelector: boolean;
  searchScopes: ISearchScope[];
  enableSuggestions: boolean;
  enableSearchManager: boolean;
  theme: IReadonlyTheme | undefined;
}
