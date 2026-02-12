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
  enableQueryBuilder: boolean;
  enableKqlMode: boolean;
  enableSearchManager: boolean;
  searchInNewPage: boolean;
  newPageUrl: string;
  queryInputTransformation: string;
  theme: IReadonlyTheme | undefined;
}
