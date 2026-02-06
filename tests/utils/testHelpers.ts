import { StoreApi } from 'zustand/vanilla';
import { createSearchStore } from '../../src/libraries/spSearchStore/store/createStore';
import { createRegistryContainer } from '../../src/libraries/spSearchStore/registries';
import {
  ISearchStore,
  ISearchResult,
  IRefiner,
  IRefinerValue,
  IActiveFilter,
  IRegistryContainer,
  IPersonaInfo,
  ISavedSearch,
  ISearchHistoryEntry,
  IClickedItem,
} from '../../src/libraries/spSearchStore/interfaces';
import { ITokenContext } from '../../src/libraries/spSearchStore/services/TokenService';

/**
 * Create a real Zustand store with fresh registries for testing.
 * Each test gets its own isolated store instance.
 */
export function createMockStore(): StoreApi<ISearchStore> {
  const registries = createRegistryContainer();
  return createSearchStore(registries);
}

/**
 * Create a real Zustand store with a provided registry container.
 */
export function createMockStoreWithRegistries(
  registries: IRegistryContainer
): StoreApi<ISearchStore> {
  return createSearchStore(registries);
}

/**
 * Factory for creating mock ISearchResult objects.
 * Provide overrides to customize specific properties.
 */
export function createMockSearchResult(
  overrides?: Partial<ISearchResult>
): ISearchResult {
  const defaults: ISearchResult = {
    key: `result-${Math.random().toString(36).substring(2, 8)}`,
    title: 'Test Document',
    url: 'https://contoso.sharepoint.com/sites/test/Documents/test.docx',
    summary: 'This is a <mark>test</mark> document summary',
    author: createMockPersona(),
    created: '2024-01-15T10:30:00Z',
    modified: '2024-06-20T14:45:00Z',
    fileType: 'docx',
    fileSize: 45678,
    siteName: 'Test Site',
    siteUrl: 'https://contoso.sharepoint.com/sites/test',
    thumbnailUrl: 'https://contoso.sharepoint.com/_api/v2.0/thumbnails/test',
    properties: {},
  };

  return { ...defaults, ...overrides };
}

/**
 * Factory for creating mock IPersonaInfo objects.
 */
export function createMockPersona(
  overrides?: Partial<IPersonaInfo>
): IPersonaInfo {
  const defaults: IPersonaInfo = {
    displayText: 'John Doe',
    email: 'john.doe@contoso.com',
    imageUrl: 'https://contoso.sharepoint.com/_layouts/15/userphoto.aspx?size=S&accountname=john.doe@contoso.com',
  };
  return { ...defaults, ...overrides };
}

/**
 * Factory for creating mock IRefiner objects.
 */
export function createMockRefiner(
  overrides?: Partial<IRefiner>
): IRefiner {
  const defaults: IRefiner = {
    filterName: 'FileType',
    values: [
      createMockRefinerValue({ name: 'docx', value: '"ǂǂ646f6378"', count: 42 }),
      createMockRefinerValue({ name: 'pptx', value: '"ǂǂ70707478"', count: 15 }),
      createMockRefinerValue({ name: 'pdf', value: '"ǂǂ706466"', count: 8 }),
    ],
  };

  return { ...defaults, ...overrides };
}

/**
 * Factory for creating mock IRefinerValue objects.
 */
export function createMockRefinerValue(
  overrides?: Partial<IRefinerValue>
): IRefinerValue {
  const defaults: IRefinerValue = {
    name: 'docx',
    value: '"ǂǂ646f6378"',
    count: 10,
    isSelected: false,
  };

  return { ...defaults, ...overrides };
}

/**
 * Factory for creating mock IActiveFilter objects.
 */
export function createMockActiveFilter(
  overrides?: Partial<IActiveFilter>
): IActiveFilter {
  const defaults: IActiveFilter = {
    filterName: 'FileType',
    value: '"docx"',
    operator: 'OR',
  };

  return { ...defaults, ...overrides };
}

/**
 * Factory for creating a mock ITokenContext.
 */
export function createMockTokenContext(
  overrides?: Partial<ITokenContext>
): ITokenContext {
  const defaults: ITokenContext = {
    queryText: 'annual report',
    siteId: 'b5c3e9a1-2d4f-4e6a-8b7c-1d2e3f4a5b6c',
    siteUrl: 'https://contoso.sharepoint.com/sites/intranet',
    webId: 'a1b2c3d4-e5f6-7890-abcd-ef1234567890',
    webUrl: 'https://contoso.sharepoint.com/sites/intranet/subweb',
    hubSiteUrl: 'https://contoso.sharepoint.com/sites/hub',
    userDisplayName: 'Jane Smith',
    userEmail: 'jane.smith@contoso.com',
    listId: '',
  };

  return { ...defaults, ...overrides };
}

/**
 * Factory for creating mock ISavedSearch objects.
 */
export function createMockSavedSearch(
  overrides?: Partial<ISavedSearch>
): ISavedSearch {
  const defaults: ISavedSearch = {
    id: 1,
    title: 'My Saved Search',
    queryText: 'annual report',
    searchState: '{}',
    searchUrl: '?q=annual+report',
    entryType: 'SavedSearch',
    category: 'Work',
    sharedWith: [],
    resultCount: 42,
    lastUsed: new Date('2024-06-20'),
    created: new Date('2024-01-15'),
    author: createMockPersona(),
  };

  return { ...defaults, ...overrides };
}

/**
 * Factory for creating mock ISearchHistoryEntry objects.
 */
export function createMockHistoryEntry(
  overrides?: Partial<ISearchHistoryEntry>
): ISearchHistoryEntry {
  const defaults: ISearchHistoryEntry = {
    id: 1,
    queryHash: 'abc123def456',
    queryText: 'budget report',
    vertical: 'all',
    scope: 'all',
    resultCount: 25,
    clickedItems: [],
    searchTimestamp: new Date('2024-06-20T10:30:00Z'),
  };

  return { ...defaults, ...overrides };
}

/**
 * Factory for creating mock IClickedItem objects.
 */
export function createMockClickedItem(
  overrides?: Partial<IClickedItem>
): IClickedItem {
  const defaults: IClickedItem = {
    url: 'https://contoso.sharepoint.com/sites/test/Documents/report.docx',
    title: 'Budget Report 2024',
    position: 1,
    timestamp: new Date('2024-06-20T10:31:00Z'),
  };

  return { ...defaults, ...overrides };
}
