import { StateCreator } from 'zustand';
import { ISearchStore, IUISlice, ISearchResult } from '@interfaces/index';

// Default visible layouts — List, Compact, Grid.
// Card, People, and Gallery are opt-in: they require explicit admin enablement
// via the Results web part property pane (showCardLayout, showPeopleLayout, showGalleryLayout).
const DEFAULT_LAYOUTS: string[] = ['list', 'compact', 'grid'];

export const createUISlice: StateCreator<ISearchStore, [], [], IUISlice> = (set, get) => ({
  activeLayoutKey: 'list',
  availableLayouts: DEFAULT_LAYOUTS,
  isSearchManagerOpen: false,
  previewPanel: {
    isOpen: false,
    item: undefined,
  },
  currentUserGroups: [],

  setLayout: (key: string): void => {
    set({ activeLayoutKey: key });
  },

  setAvailableLayouts: (layouts: string[]): void => {
    set({ availableLayouts: layouts });
  },

  toggleSearchManager: (isOpen?: boolean): void => {
    const current = get().isSearchManagerOpen;
    set({ isSearchManagerOpen: isOpen !== undefined ? isOpen : !current });
  },

  setPreviewItem: (item: ISearchResult | undefined): void => {
    set({
      previewPanel: {
        isOpen: item !== undefined,
        item,
      },
    });
  },

  setCurrentUserGroups: (groups: string[]): void => {
    set({ currentUserGroups: groups });
  },
});
