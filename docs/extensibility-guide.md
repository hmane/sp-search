# SP Search — Extensibility Guide

SP Search uses a pluggable **provider/registry model** with 5 extension points. Custom implementations can be registered alongside (or in place of) built-in providers without modifying core code.

---

## Architecture Overview

Each search context (identified by `searchContextId`) gets its own isolated registry container:

```
IRegistryContainer
├── dataProviders:  IRegistry<ISearchDataProvider>
├── suggestions:    IRegistry<ISuggestionProvider>
├── actions:        IRegistry<IActionProvider>
├── layouts:        IRegistry<ILayoutDefinition>
└── filterTypes:    IRegistry<IFilterTypeDefinition>
```

### Registry Rules

1. **First registration wins** — duplicate IDs log a warning and are skipped
2. **Force override** — pass `force: true` to replace an existing registration
3. **Registries freeze after first search** — prevents mid-session mutations that cause race conditions
4. **Registration happens in `onInit()`** — before any search executes

### Registry API

```typescript
interface IRegistry<T extends { id: string }> {
  register(provider: T, force?: boolean): void;
  get(id: string): T | undefined;
  getAll(): T[];
  freeze(): void;
  isFrozen(): boolean;
}
```

### Accessing Registries

```typescript
import { getStore } from 'sp-search-store';

const store = getStore('my-context-id');
const registries = store.getState().registries;

// Register a custom provider
registries.dataProviders.register(new MyCustomProvider());
```

---

## 1. Custom Data Provider

Data providers abstract over search backends. The default `SharePointSearchProvider` uses PnPjs; you can add providers for Microsoft Graph, external APIs, or mock data.

### Interface

```typescript
interface ISearchDataProvider {
  id: string;
  displayName: string;
  supportsRefiners: boolean;
  supportsCollapsing: boolean;
  supportsSorting: boolean;

  execute(query: ISearchQuery, signal: AbortSignal): Promise<ISearchResponse>;
  getSuggestions?(query: string, signal: AbortSignal): Promise<ISuggestion[]>;
  getSchema?(): Promise<IManagedProperty[]>;
}
```

### ISearchQuery (Input)

```typescript
interface ISearchQuery {
  queryText: string;             // User's search text (or '*')
  queryTemplate: string;         // KQL template with {searchTerms}, {Site.ID}, etc.
  scope: ISearchScope;           // Search scope (All, Current Site, Hub, custom)
  filters: IActiveFilter[];      // Active refinement filters
  sort: ISortField | undefined;  // Sort field + direction
  page: number;                  // Current page (1-based)
  pageSize: number;              // Results per page
  selectedProperties: string[];  // Managed properties to retrieve
  refiners: string[];            // Properties to get refiner data for
  resultSourceId?: string;       // SharePoint Result Source GUID
  trimDuplicates?: boolean;
}
```

### ISearchResponse (Output)

```typescript
interface ISearchResponse {
  items: ISearchResult[];
  totalCount: number;
  refiners: IRefiner[];
  promotedResults: IPromotedResultItem[];
  querySuggestion?: string;     // "Did you mean..."
}
```

### Example Implementation

```typescript
import { ISearchDataProvider, ISearchQuery, ISearchResponse } from 'sp-search-store';

export class ExternalApiProvider implements ISearchDataProvider {
  public readonly id = 'external-api';
  public readonly displayName = 'External Search';
  public readonly supportsRefiners = false;
  public readonly supportsCollapsing = false;
  public readonly supportsSorting = true;

  public async execute(query: ISearchQuery, signal: AbortSignal): Promise<ISearchResponse> {
    const response = await fetch('https://api.example.com/search', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ q: query.queryText, page: query.page, size: query.pageSize }),
      signal, // Pass AbortSignal for cancellation support
    });

    const data = await response.json();

    return {
      items: data.results.map((r: any) => ({
        key: r.id,
        title: r.title,
        url: r.url,
        summary: r.excerpt,
        author: { displayName: r.author, email: '', loginName: '' },
        created: r.created,
        modified: r.modified,
        fileType: r.type,
        fileSize: r.size || 0,
        siteName: '',
        siteUrl: '',
        thumbnailUrl: '',
        properties: r,
      })),
      totalCount: data.total,
      refiners: [],
      promotedResults: [],
    };
  }
}
```

### Registration

```typescript
// In your web part's onInit()
const store = getStore(this.properties.searchContextId);
const registry = store.getState().registries.dataProviders;

if (!registry.get('external-api')) {
  registry.register(new ExternalApiProvider());
}
```

### Per-Vertical Override

Different verticals can use different providers by checking `query.resultSourceId` or by registering multiple providers and selecting in the orchestrator configuration.

---

## 2. Custom Suggestion Provider

Suggestion providers populate the search box dropdown. Multiple providers run in parallel, grouped by priority.

### Interface

```typescript
interface ISuggestionProvider {
  id: string;
  displayName: string;    // Section label in dropdown
  priority: number;       // Lower = shown first
  maxResults: number;

  getSuggestions(query: string, context: ISearchContext): Promise<ISuggestion[]>;
  isEnabled(context: ISearchContext): boolean;
}

interface ISearchContext {
  searchContextId: string;
  siteUrl: string;
  scope: ISearchScope;
}

interface ISuggestion {
  displayText: string;
  groupName: string;      // Section label: "Recent", "Trending", "Files", etc.
  iconName?: string;      // Fluent UI icon name
  action?: () => void;    // Optional callback on selection
}
```

### Built-in Providers

| Provider | Priority | Description |
|----------|----------|-------------|
| `RecentSearchProvider` | 10 | User's recent searches from SearchHistory list |
| `TrendingQueryProvider` | 20 | Popular searches across the organization |
| `ManagedPropertyProvider` | 30 | Property:value suggestions (e.g., `Author: John`) |

### Example Implementation

```typescript
export class BookmarkSuggestionProvider implements ISuggestionProvider {
  public readonly id = 'bookmarks';
  public readonly displayName = 'Bookmarks';
  public readonly priority = 5; // Show above recent searches
  public readonly maxResults = 3;

  public isEnabled(_context: ISearchContext): boolean {
    return true;
  }

  public async getSuggestions(query: string): Promise<ISuggestion[]> {
    const bookmarks = await fetchUserBookmarks(query);
    return bookmarks.map((b) => ({
      displayText: b.title,
      groupName: 'Bookmarks',
      iconName: 'FavoriteStar',
    }));
  }
}
```

### Registration

```typescript
const registry = store.getState().registries.suggestions;
if (!registry.get('bookmarks')) {
  registry.register(new BookmarkSuggestionProvider());
}
```

---

## 3. Custom Action Provider

Action providers add quick actions to search results — toolbar buttons, context menu items, and bulk action options.

### Interface

```typescript
interface IActionProvider {
  id: string;
  label: string;
  iconName: string;                              // Fluent UI icon
  position: 'toolbar' | 'contextMenu' | 'both';
  isBulkEnabled: boolean;

  isApplicable(item: ISearchResult): boolean;
  execute(items: ISearchResult[], context: ISearchContext): Promise<void>;
}
```

### Built-in Actions

| ID | Label | Position | Bulk | Description |
|----|-------|----------|------|-------------|
| `open` | Open | both | Yes | Open result in new tab |
| `preview` | Preview | contextMenu | No | Open detail panel |
| `share` | Share | toolbar | Yes | Share via URL/email/Teams |
| `pin` | Pin | both | Yes | Pin to collection |
| `copyLink` | Copy link | both | Yes | Copy URL to clipboard |
| `download` | Download | toolbar | Yes | Download file |
| `compare` | Compare | toolbar | Yes | Compare selected versions |
| `exportCsv` | Export CSV | toolbar | No | Export results to CSV |

### Example Implementation

```typescript
export class SendToTeamsAction implements IActionProvider {
  public readonly id = 'send-to-teams';
  public readonly label = 'Send to Teams';
  public readonly iconName = 'TeamsLogo';
  public readonly position: 'toolbar' = 'toolbar';
  public readonly isBulkEnabled = true;

  public isApplicable(item: ISearchResult): boolean {
    return !!item.url; // Only for items with URLs
  }

  public async execute(items: ISearchResult[]): Promise<void> {
    const urls = items.map((i) => i.url).join('\n');
    const encoded = encodeURIComponent(urls);
    window.open(
      'https://teams.microsoft.com/l/chat/0/0?message=' + encoded,
      '_blank'
    );
  }
}
```

### Registration

```typescript
const registry = store.getState().registries.actions;
if (!registry.get('send-to-teams')) {
  registry.register(new SendToTeamsAction());
}
```

---

## 4. Custom Layout

Layouts define how search results are rendered. Each layout is code-split via `React.lazy()` for optimal bundle size.

### Interface

```typescript
interface ILayoutDefinition {
  id: string;
  displayName: string;
  iconName: string;                                              // Fluent UI icon
  component: React.LazyExoticComponent<React.ComponentType<any>>;
  supportsPaging: 'numbered' | 'infinite' | 'both';
  supportsBulkSelect: boolean;
  supportsVirtualization: boolean;
  defaultSortable: boolean;
}
```

### Built-in Layouts

| ID | Name | Paging | Bulk | Virtual | Sort |
|----|------|--------|------|---------|------|
| `list` | List | numbered | Yes | Yes | Yes |
| `compact` | Compact | numbered | Yes | Yes | Yes |
| `datagrid` | Data Grid | numbered | Yes | Yes | Yes |
| `card` | Card | numbered | Yes | No | No |
| `people` | People | numbered | No | No | No |
| `gallery` | Gallery | infinite | Yes | No | No |

### Layout Component Props

Your layout component receives (matching the actual `IListLayoutProps` pattern used by built-in layouts):

```typescript
interface ILayoutProps {
  items: ISearchResult[];
  enableSelection: boolean;
  selectedKeys: string[];
  onToggleSelection: (key: string, multiSelect: boolean) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
}
```

### Example Implementation

```typescript
// TimelineLayout.tsx
import * as React from 'react';
import type { ISearchResult } from 'sp-search-store';

export interface ITimelineLayoutProps {
  items: ISearchResult[];
  enableSelection: boolean;
  selectedKeys: string[];
  onToggleSelection: (key: string, multiSelect: boolean) => void;
  onItemClick?: (item: ISearchResult, position: number) => void;
}

const TimelineLayout: React.FC<ITimelineLayoutProps> = function TimelineLayout(props) {
  return (
    <div className="timeline-layout">
      {props.items.map(function (item, index) {
        return (
          <div
            key={item.key}
            className="timeline-entry"
            onClick={function () {
              if (props.onItemClick) { props.onItemClick(item, index); }
            }}
          >
            <div className="timeline-date">{item.modified}</div>
            <div className="timeline-title">{item.title}</div>
            <div className="timeline-summary">{item.summary}</div>
          </div>
        );
      })}
    </div>
  );
};

export default TimelineLayout;
```

### Registration

```typescript
const registry = store.getState().registries.layouts;

registry.register({
  id: 'timeline',
  displayName: 'Timeline',
  iconName: 'Timeline',
  component: React.lazy(function () { return import('./layouts/TimelineLayout'); }),
  supportsPaging: 'numbered',
  supportsBulkSelect: false,
  supportsVirtualization: false,
  defaultSortable: false,
});
```

---

## 5. Custom Filter Type

Filter types define the UI control, serialization, and refinement token building for a filter field.

### Interface

```typescript
interface IFilterTypeDefinition {
  id: string;
  displayName: string;
  component: React.ComponentType<any>;

  // URL serialization (for deep linking)
  serializeValue(value: unknown): string;
  deserializeValue(raw: string): unknown;

  // Search API token (KQL/FQL)
  buildRefinementToken(value: unknown, managedProperty: string): string;
}
```

### Built-in Filter Types

| ID | Name | Token Format |
|----|------|-------------|
| `checkbox` | Checkbox List | Quoted KQL `"value"` |
| `daterange` | Date Range | FQL `range(datetime("..."), datetime("..."))` |
| `slider` | Slider | FQL `range(decimal(...), decimal(...))` |
| `people` | People Picker | User claim string |
| `taxonomy` | Taxonomy Tree | `GP0\|#GUID` format |
| `tagbox` | Tag Box | Quoted KQL `"value"` |
| `toggle` | Toggle | Quoted KQL `"0"` / `"1"` |

### Supporting Interface: IFilterValueFormatter

For complex value display (taxonomy GUIDs to labels, claim strings to names):

```typescript
interface IFilterValueFormatter {
  id: string;
  formatForDisplay(rawValue: string, config: IFilterConfig): string | Promise<string>;
  formatForQuery(displayValue: unknown, config: IFilterConfig): string;
  formatForUrl(rawValue: string): string;
  parseFromUrl(urlValue: string): string;
}
```

### Filter Component Props

Your filter component receives:

```typescript
interface IFilterComponentProps {
  config: IFilterConfig;
  availableValues: IRefinerValue[];
  selectedValues: string[];
  onValueChange: (values: string[]) => void;
}
```

### Example Implementation

```typescript
// RatingFilter.tsx — star rating filter (1-5)
import * as React from 'react';
import { Rating, RatingSize } from '@fluentui/react/lib/Rating';
import type { IFilterComponentProps } from 'sp-search-store';

const RatingFilter: React.FC<IFilterComponentProps> = function RatingFilter(props) {
  const current = props.selectedValues.length > 0 ? parseInt(props.selectedValues[0], 10) : 0;

  return (
    <Rating
      min={1}
      max={5}
      size={RatingSize.Large}
      rating={current}
      onChange={function (_ev: unknown, rating?: number) {
        if (rating) {
          props.onValueChange([String(rating)]);
        }
      }}
      ariaLabel="Filter by minimum rating"
    />
  );
};

export default RatingFilter;
```

### Registration

```typescript
import { createLazyComponent } from 'spfx-toolkit/lib/utilities/lazyLoader';

const registry = store.getState().registries.filterTypes;

registry.register({
  id: 'rating',
  displayName: 'Star Rating',
  component: createLazyComponent(
    function () { return import('./filters/RatingFilter') as any; },
    { errorMessage: 'Failed to load rating filter' }
  ),
  serializeValue: function (value: unknown): string {
    return String(value);
  },
  deserializeValue: function (raw: string): unknown {
    return raw;
  },
  buildRefinementToken: function (value: unknown, _prop: string): string {
    // FQL range: rating >= selected value
    return 'range(' + String(value) + ', max)';
  },
});
```

**Note on `createLazyComponent`:** When using `createLazyComponent` from spfx-toolkit, the dynamic import must be cast `as any` due to a `@types/react` version mismatch between sp-search and spfx-toolkit. The utility bundles `Suspense` + error boundary internally — do NOT wrap its output in `<React.Suspense>`.

---

## Lifecycle: Registration to Execution

```
1. Web Part onInit()
   └── getStore(searchContextId) → creates or retrieves shared store
   └── store.getState().registries.* → access registries
   └── registry.register(provider) → register custom implementations

2. First Search Triggered
   └── SearchOrchestrator freezes all 5 registries
   └── No further registrations accepted (warnings logged)

3. Search Execution
   └── Orchestrator calls dataProvider.execute(query, signal)
   └── Results dispatched to store
   └── Filter types render refiner UI
   └── Actions appear on results

4. Web Part Dispose
   └── disposeStore(searchContextId) tears down store + subscriptions
```

---

## Best Practices

1. **Always check before registering** — `if (!registry.get(id)) { registry.register(...) }` prevents duplicate warnings.
2. **Support AbortSignal** — pass `signal` to all fetch/API calls for proper cancellation.
3. **Lazy-load heavy components** — use `React.lazy()` for layouts and `createLazyComponent()` for filter types.
4. **Use tree-shakable imports** — `@fluentui/react/lib/...` and `spfx-toolkit/lib/...`, never barrel imports.
5. **Test serialization round-trips** — `deserializeValue(serializeValue(x))` must return the original value for URL deep linking to work.
6. **Handle taxonomy tokens** — GP0|#GUID format requires async label resolution; cache resolved labels.
7. **Use FQL for date/numeric ranges** — never raw KQL date comparisons; they fail silently.
