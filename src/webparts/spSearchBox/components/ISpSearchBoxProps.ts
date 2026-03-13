import { StoreApi } from 'zustand/vanilla';
import { ISearchStore, ISearchScope } from '@interfaces/index';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export interface ISpSearchBoxProps {
  store: StoreApi<ISearchStore>;
  searchContextId: string;
  siteUrl: string;
  placeholder: string;
  debounceMs: number;
  searchBehavior: 'onEnter' | 'onButton' | 'both';
  resetSearchOnClear: boolean;
  enableScopeSelector: boolean;
  searchScopes: ISearchScope[];
  enableSuggestions: boolean;
  enableSharePointSuggestions: boolean;
  enableRecentSuggestions: boolean;
  enablePopularSuggestions: boolean;
  enableQuickResults: boolean;
  enablePropertySuggestions: boolean;
  suggestionsPerGroup: number;
  enableQueryBuilder: boolean;
  enableKqlMode: boolean;
  enableSearchManager: boolean;
  searchInNewPage: boolean;
  newPageUrl: string;
  newPageOpenBehavior: 'sameTab' | 'newTab';
  newPageParameterLocation: 'queryString' | 'hash';
  newPageQueryParameter: string;
  theme: IReadonlyTheme | undefined;
}
