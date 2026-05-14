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
    // T2.D2 — selection persists across layout switch (audit acceptance signal).
    // The 3 selection-aware layouts (list / compact / grid) all render their
    // selection from the same `bulkSelection` array, so a row ticked on List
    // stays ticked when the admin switches to Compact and back. Layouts that
    // don't render checkboxes (card / people / gallery) ignore the array;
    // returning to a checkbox-aware layout restores the ticks.
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
