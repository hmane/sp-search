import { StoreApi } from 'zustand/vanilla';
import { ISearchStore } from '../../../src/libraries/spSearchStore/interfaces';
import { createMockStore } from '../../utils/testHelpers';

describe('querySlice', () => {
  let store: StoreApi<ISearchStore>;

  beforeEach(() => {
    store = createMockStore();
  });

  describe('initial state', () => {
    it('should have empty queryText', () => {
      expect(store.getState().queryText).toBe('');
    });

    it('should have default queryTemplate', () => {
      expect(store.getState().queryTemplate).toBe('{searchTerms}');
    });

    it('should have default scope "All SharePoint"', () => {
      expect(store.getState().scope).toEqual({
        id: 'all',
        label: 'All SharePoint',
      });
    });

    it('should have empty suggestions array', () => {
      expect(store.getState().suggestions).toEqual([]);
    });

    it('should not be searching initially', () => {
      expect(store.getState().isSearching).toBe(false);
    });

    it('should have undefined abortController', () => {
      expect(store.getState().abortController).toBeUndefined();
    });
  });

  describe('setQueryText', () => {
    it('should update queryText', () => {
      store.getState().setQueryText('annual report');
      expect(store.getState().queryText).toBe('annual report');
    });

    it('should handle empty string', () => {
      store.getState().setQueryText('hello');
      store.getState().setQueryText('');
      expect(store.getState().queryText).toBe('');
    });

    it('should handle special characters', () => {
      store.getState().setQueryText('C# "unit tests" path:*.ts');
      expect(store.getState().queryText).toBe('C# "unit tests" path:*.ts');
    });

    it('should handle very long query text', () => {
      const longText = 'a'.repeat(1000);
      store.getState().setQueryText(longText);
      expect(store.getState().queryText).toBe(longText);
    });

    it('should not affect other slice state', () => {
      store.getState().setQueryText('test');
      expect(store.getState().activeFilters).toEqual([]);
      expect(store.getState().items).toEqual([]);
      expect(store.getState().currentVerticalKey).toBe('all');
    });
  });

  describe('setScope', () => {
    it('should update scope', () => {
      const newScope = { id: 'site', label: 'Current Site' };
      store.getState().setScope(newScope);
      expect(store.getState().scope).toEqual(newScope);
    });

    it('should handle scope with kqlPath', () => {
      const scopeWithPath = {
        id: 'hr',
        label: 'HR Site',
        kqlPath: 'Path:https://contoso.sharepoint.com/sites/hr',
      };
      store.getState().setScope(scopeWithPath);
      expect(store.getState().scope).toEqual(scopeWithPath);
    });

    it('should handle scope with resultSourceId', () => {
      const scopeWithRS = {
        id: 'people',
        label: 'People',
        resultSourceId: 'b09a7990-05ea-4af9-81ef-edfab16c4e31',
      };
      store.getState().setScope(scopeWithRS);
      expect(store.getState().scope).toEqual(scopeWithRS);
    });
  });

  describe('setSuggestions', () => {
    it('should update suggestions array', () => {
      const suggestions = [
        { displayText: 'annual report 2024', groupName: 'Recent' },
        { displayText: 'annual budget', groupName: 'Trending' },
      ];
      store.getState().setSuggestions(suggestions);
      expect(store.getState().suggestions).toEqual(suggestions);
    });

    it('should replace existing suggestions', () => {
      store.getState().setSuggestions([
        { displayText: 'old suggestion', groupName: 'Recent' },
      ]);
      store.getState().setSuggestions([
        { displayText: 'new suggestion', groupName: 'Trending' },
      ]);
      expect(store.getState().suggestions).toHaveLength(1);
      expect(store.getState().suggestions[0].displayText).toBe('new suggestion');
    });

    it('should handle empty suggestions array', () => {
      store.getState().setSuggestions([
        { displayText: 'some suggestion', groupName: 'Recent' },
      ]);
      store.getState().setSuggestions([]);
      expect(store.getState().suggestions).toEqual([]);
    });

    it('should handle suggestions with optional properties', () => {
      const suggestions = [
        {
          displayText: 'test suggestion',
          groupName: 'Files',
          iconName: 'Document',
          action: (): void => { /* no-op */ },
        },
      ];
      store.getState().setSuggestions(suggestions);
      expect(store.getState().suggestions[0].iconName).toBe('Document');
      expect(store.getState().suggestions[0].action).toBeDefined();
    });
  });

  describe('cancelSearch', () => {
    it('should abort the in-flight controller', () => {
      const controller = new AbortController();
      store.setState({ abortController: controller, isSearching: true });

      expect(controller.signal.aborted).toBe(false);

      store.getState().cancelSearch();

      expect(controller.signal.aborted).toBe(true);
    });

    it('should clear abortController and set isSearching to false', () => {
      const controller = new AbortController();
      store.setState({ abortController: controller, isSearching: true });

      store.getState().cancelSearch();

      expect(store.getState().abortController).toBeUndefined();
      expect(store.getState().isSearching).toBe(false);
    });

    it('should be a no-op when there is no controller', () => {
      store.setState({ isSearching: false, abortController: undefined });

      // Should not throw
      expect(() => store.getState().cancelSearch()).not.toThrow();

      expect(store.getState().abortController).toBeUndefined();
      expect(store.getState().isSearching).toBe(false);
    });

    it('should handle already-aborted controller gracefully', () => {
      const controller = new AbortController();
      controller.abort(); // Already aborted
      store.setState({ abortController: controller, isSearching: true });

      // Should not throw
      expect(() => store.getState().cancelSearch()).not.toThrow();
      expect(store.getState().isSearching).toBe(false);
    });
  });
});
