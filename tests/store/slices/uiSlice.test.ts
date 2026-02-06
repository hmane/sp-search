import { StoreApi } from 'zustand/vanilla';
import { ISearchStore } from '../../../src/libraries/spSearchStore/interfaces';
import { createMockStore, createMockSearchResult } from '../../utils/testHelpers';

describe('uiSlice', () => {
  let store: StoreApi<ISearchStore>;

  beforeEach(() => {
    store = createMockStore();
  });

  describe('initial state', () => {
    it('should have default layout key "list"', () => {
      expect(store.getState().activeLayoutKey).toBe('list');
    });

    it('should have search manager closed', () => {
      expect(store.getState().isSearchManagerOpen).toBe(false);
    });

    it('should have preview panel closed with no item', () => {
      expect(store.getState().previewPanel).toEqual({
        isOpen: false,
        item: undefined,
      });
    });

    it('should have empty bulk selection', () => {
      expect(store.getState().bulkSelection).toEqual([]);
    });
  });

  describe('setLayout', () => {
    it('should change activeLayoutKey', () => {
      store.getState().setLayout('grid');
      expect(store.getState().activeLayoutKey).toBe('grid');
    });

    it('should handle all built-in layout keys', () => {
      const layouts = ['list', 'grid', 'card', 'compact', 'people', 'gallery'];
      for (const layout of layouts) {
        store.getState().setLayout(layout);
        expect(store.getState().activeLayoutKey).toBe(layout);
      }
    });

    it('should handle custom layout key', () => {
      store.getState().setLayout('custom-layout-v2');
      expect(store.getState().activeLayoutKey).toBe('custom-layout-v2');
    });

    it('should replace previous layout', () => {
      store.getState().setLayout('grid');
      store.getState().setLayout('card');
      expect(store.getState().activeLayoutKey).toBe('card');
    });
  });

  describe('toggleSearchManager', () => {
    it('should toggle from closed to open', () => {
      store.getState().toggleSearchManager();
      expect(store.getState().isSearchManagerOpen).toBe(true);
    });

    it('should toggle from open to closed', () => {
      store.getState().toggleSearchManager(); // open
      store.getState().toggleSearchManager(); // close
      expect(store.getState().isSearchManagerOpen).toBe(false);
    });

    it('should force open when isOpen=true', () => {
      store.getState().toggleSearchManager(true);
      expect(store.getState().isSearchManagerOpen).toBe(true);
    });

    it('should force close when isOpen=false', () => {
      store.getState().toggleSearchManager(true); // open
      store.getState().toggleSearchManager(false); // force close
      expect(store.getState().isSearchManagerOpen).toBe(false);
    });

    it('should stay open when already open and isOpen=true', () => {
      store.getState().toggleSearchManager(true);
      store.getState().toggleSearchManager(true);
      expect(store.getState().isSearchManagerOpen).toBe(true);
    });

    it('should stay closed when already closed and isOpen=false', () => {
      store.getState().toggleSearchManager(false);
      expect(store.getState().isSearchManagerOpen).toBe(false);
    });
  });

  describe('setPreviewItem', () => {
    it('should open the preview panel with an item', () => {
      const item = createMockSearchResult({ key: 'preview-1', title: 'Preview Doc' });
      store.getState().setPreviewItem(item);

      expect(store.getState().previewPanel.isOpen).toBe(true);
      expect(store.getState().previewPanel.item).toEqual(item);
    });

    it('should close the preview panel when item is undefined', () => {
      const item = createMockSearchResult();
      store.getState().setPreviewItem(item);
      store.getState().setPreviewItem(undefined);

      expect(store.getState().previewPanel.isOpen).toBe(false);
      expect(store.getState().previewPanel.item).toBeUndefined();
    });

    it('should replace the preview item when a new one is set', () => {
      const item1 = createMockSearchResult({ key: '1', title: 'First' });
      const item2 = createMockSearchResult({ key: '2', title: 'Second' });

      store.getState().setPreviewItem(item1);
      store.getState().setPreviewItem(item2);

      expect(store.getState().previewPanel.isOpen).toBe(true);
      expect(store.getState().previewPanel.item!.key).toBe('2');
    });
  });

  describe('toggleSelection', () => {
    describe('single-select mode (multiSelect=false)', () => {
      it('should select a single item', () => {
        store.getState().toggleSelection('item-1', false);
        expect(store.getState().bulkSelection).toEqual(['item-1']);
      });

      it('should replace selection with a new item', () => {
        store.getState().toggleSelection('item-1', false);
        store.getState().toggleSelection('item-2', false);
        expect(store.getState().bulkSelection).toEqual(['item-2']);
      });

      it('should deselect an already-selected item', () => {
        store.getState().toggleSelection('item-1', false);
        store.getState().toggleSelection('item-1', false);
        expect(store.getState().bulkSelection).toEqual([]);
      });
    });

    describe('multi-select mode (multiSelect=true)', () => {
      it('should add items to the selection', () => {
        store.getState().toggleSelection('item-1', true);
        store.getState().toggleSelection('item-2', true);
        store.getState().toggleSelection('item-3', true);
        expect(store.getState().bulkSelection).toEqual(['item-1', 'item-2', 'item-3']);
      });

      it('should deselect an already-selected item', () => {
        store.getState().toggleSelection('item-1', true);
        store.getState().toggleSelection('item-2', true);
        store.getState().toggleSelection('item-1', true); // deselect
        expect(store.getState().bulkSelection).toEqual(['item-2']);
      });

      it('should handle toggling the same item multiple times', () => {
        store.getState().toggleSelection('item-1', true);
        store.getState().toggleSelection('item-1', true);
        store.getState().toggleSelection('item-1', true);
        // Toggle: add -> remove -> add
        expect(store.getState().bulkSelection).toEqual(['item-1']);
      });
    });

    describe('mixed mode interactions', () => {
      it('should switch from multi to single correctly', () => {
        store.getState().toggleSelection('item-1', true);
        store.getState().toggleSelection('item-2', true);
        // Now single-select a new item: replaces all
        store.getState().toggleSelection('item-3', false);
        expect(store.getState().bulkSelection).toEqual(['item-3']);
      });
    });
  });

  describe('clearSelection', () => {
    it('should clear all selected items', () => {
      store.getState().toggleSelection('item-1', true);
      store.getState().toggleSelection('item-2', true);
      store.getState().toggleSelection('item-3', true);

      store.getState().clearSelection();
      expect(store.getState().bulkSelection).toEqual([]);
    });

    it('should be a no-op on empty selection', () => {
      store.getState().clearSelection();
      expect(store.getState().bulkSelection).toEqual([]);
    });
  });
});
