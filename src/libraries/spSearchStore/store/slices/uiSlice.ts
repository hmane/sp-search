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
  bulkSelection: [],
  currentUserGroups: [],

  setLayout: (key: string): void => {
    // Clear bulk selection when switching layouts — selections are layout-specific
    // and invisible on layouts that don't render checkboxes.
    set({ activeLayoutKey: key, bulkSelection: [] });
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

  toggleSelection: (itemKey: string, multiSelect: boolean): void => {
    const current = get().bulkSelection;
    const index = current.indexOf(itemKey);

    if (index >= 0) {
      // Deselect
      const updated = [...current];
      updated.splice(index, 1);
      set({ bulkSelection: updated });
    } else if (multiSelect) {
      // Add to selection
      set({ bulkSelection: [...current, itemKey] });
    } else {
      // Replace selection
      set({ bulkSelection: [itemKey] });
    }
  },

  clearSelection: (): void => {
    set({ bulkSelection: [] });
  },

  setCurrentUserGroups: (groups: string[]): void => {
    set({ currentUserGroups: groups });
  },
});
