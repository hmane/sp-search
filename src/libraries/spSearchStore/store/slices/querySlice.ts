import { StateCreator } from 'zustand';
import { ISearchStore, IQuerySlice, ISearchScope, ISuggestion } from '@interfaces/index';

const defaultScope: ISearchScope = {
  id: 'all',
  label: 'All SharePoint',
};

export const createQuerySlice: StateCreator<ISearchStore, [], [], IQuerySlice> = (set) => ({
  queryText: '',
  queryTemplate: '{searchTerms}',
  queryInputTransformation: '{searchTerms}',
  scope: defaultScope,
  suggestions: [],
  isSearching: false,

  setQueryText: (text: string): void => {
    set({ queryText: text });
  },

  setScope: (scope: ISearchScope): void => {
    set({ scope });
  },

  setSuggestions: (suggestions: ISuggestion[]): void => {
    set({ suggestions });
  },

  setQueryInputTransformation: (transformation: string): void => {
    set({ queryInputTransformation: transformation });
  },
});
