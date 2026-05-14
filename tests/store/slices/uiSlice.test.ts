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

  // bulkSelection / toggleSelection / clearSelection retired alongside
  // the BulkActionsToolbar surface. The per-row ECB menu replaces the
  // checkbox-driven bulk-action flow.
});
