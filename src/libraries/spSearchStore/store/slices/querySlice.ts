import { StateCreator } from 'zustand';
import { ISearchStore, IQuerySlice, ISearchScope, ISuggestion } from '@interfaces/index';

const defaultScope: ISearchScope = {
  id: 'all',
  label: 'All SharePoint',
};

export const createQuerySlice: StateCreator<ISearchStore, [], [], IQuerySlice> = (set, get) => ({
  queryText: '',
  queryTemplate: '{searchTerms}',
  scope: defaultScope,
  suggestions: [],
  isSearching: false,
  abortController: undefined,

  setQueryText: (text: string): void => {
    set({ queryText: text });
  },

  setScope: (scope: ISearchScope): void => {
    set({ scope });
  },

  setSuggestions: (suggestions: ISuggestion[]): void => {
    set({ suggestions });
  },

  cancelSearch: (): void => {
    const controller = get().abortController;
    if (controller) {
      controller.abort();
    }
    set({ abortController: undefined, isSearching: false });
  },
});
