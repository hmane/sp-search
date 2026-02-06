import { StateCreator } from 'zustand';
import { ISearchStore, IUISlice, ISearchResult } from '@interfaces/index';

export const createUISlice: StateCreator<ISearchStore, [], [], IUISlice> = (set, get) => ({
  activeLayoutKey: 'list',
  isSearchManagerOpen: false,
  previewPanel: {
    isOpen: false,
    item: undefined,
  },
  bulkSelection: [],

  setLayout: (key: string): void => {
    set({ activeLayoutKey: key });
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
});
