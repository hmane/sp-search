# SP Search — Admin Configuration Guide

This guide covers property pane configuration for all 5 SP Search web parts. Each web part connects to a shared Zustand store via the `searchContextId` property — web parts with the same ID share state; different IDs create isolated search experiences.

---

## Prerequisites

1. SP Search `.sppkg` deployed to App Catalog (site-level or tenant-level)
2. Hidden lists provisioned via `Provision-SPSearchLists.ps1` (see [provisioning-guide.md](./provisioning-guide.md))
3. App added to the target site collection

---

## 1. Search Box Web Part (`SPSearchBoxWebPart`)

The query input. Dispatches search text and scope to the shared store.

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `searchContextId` | Text | `"default"` | **Required.** Shared store identifier. Must match across all connected web parts. |
| `placeholder` | Text | `"Search..."` | Placeholder text shown in the empty search input. |
| `debounceMs` | Slider (100–2000) | `300` | Milliseconds to wait after typing before triggering search. Lower = more responsive but more API calls. |
| `searchBehavior` | Choice | `"both"` | When to trigger search: **onEnter** (Enter key only), **onButton** (click only), **both** (either). |
| `enableScopeSelector` | Toggle | `false` | Show scope dropdown (All SharePoint / Current Site / Current Hub / custom). |
| `enableSuggestions` | Toggle | `true` | Show suggestion dropdown with recent, trending, and property-value suggestions. |
| `enableQueryBuilder` | Toggle | `false` | Show "Advanced" toggle to expand the visual query builder panel. |
| `enableSearchManager` | Toggle | `true` | Show the Search Manager icon button (saved searches, history, collections). |

### Configuration Tips

- Set `debounceMs` to **500–800** on high-traffic sites to reduce API load.
- `enableSuggestions` requires the SearchHistory list to be provisioned.
- `enableSearchManager` requires all 4 hidden lists to be provisioned.
- `enableQueryBuilder` populates property dropdowns from the Search Administration API — users need at least Search Admin or Site Collection Admin permissions for the schema browser.

---

## 2. Search Results Web Part (`SPSearchResultsWebPart`)

Displays search results with multiple layout options, sorting, and bulk actions.

### Data Group

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `searchContextId` | Text | `"default"` | **Required.** Must match the Search Box web part. |
| `queryTemplate` | Schema Helper | `"{searchTerms}"` | KQL query template. Use `{searchTerms}` for the user's query. Supports tokens: `{Site.ID}`, `{Site.URL}`, `{Hub}`, `{Today}`, `{User.*}`, `{PageContext.*}`. The "Browse Schema" button opens a managed property picker filtered to queryable properties. |
| `selectedProperties` | Schema Helper (multiline) | *(see below)* | Comma-separated managed properties to retrieve. The "Browse Schema" button opens a picker filtered to retrievable properties. |
| `pageSize` | Slider (5–100) | `25` | Results per page. Lower values improve perceived speed. |

**Default `selectedProperties`:** `Title,Path,HitHighlightedSummary,Author,LastModifiedTime,Created,FileType,SPSiteURL,SiteName,FileExtension,SecondaryFileExtension,ContentTypeId,UniqueId,NormSiteID,NormWebID,NormListID,NormUniqueID,ServerRedirectedURL,ServerRedirectedEmbedURL,ServerRedirectedPreviewURL,PictureThumbnailURL`

### Display Group

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `defaultLayout` | Choice | `"list"` | Initial layout: **list** (Google-style cards) or **compact** (single-line rows). Additional layouts (DataGrid, Card, People, Gallery) are available via the layout switcher. |
| `showResultCount` | Toggle | `true` | Show "N results found" above results. |
| `showSortDropdown` | Toggle | `true` | Show sort dropdown (Relevance, Date, Author, etc.). |
| `enableSelection` | Toggle | `false` | Enable checkbox selection on results for bulk actions (share, download, pin, export). |

### Schema Helper Control

The `queryTemplate` and `selectedProperties` fields use a custom **Schema Helper** control. Clicking "Browse Schema" opens a panel listing all managed properties from the search schema with:
- Name, alias, type columns
- Queryable / Retrievable / Refinable / Sortable flag indicators
- Filter tabs: All | Refinable | Sortable | Retrievable
- Search-within to find properties by name
- Click a property to insert it

**Permissions:** The schema browser requires Search Admin or Site Collection Admin permissions. If the current user lacks access, the "Browse Schema" button remains enabled but displays an info MessageBar stating "Contact your SharePoint admin for managed property names." The tooltip on the button also updates to indicate the permission requirement.

---

## 3. Search Filters Web Part (`SPSearchFiltersWebPart`)

Displays refinement filters. Reads refiner data from the search response and renders filter controls.

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `searchContextId` | Text | `"default"` | **Required.** Must match the Search Results web part. |
| `applyMode` | Choice | `"instant"` | **instant**: filter changes trigger search immediately. **manual**: user clicks "Apply" to execute. |
| `operatorBetweenFilters` | Choice | `"AND"` | Logic combining multiple filter selections: **AND** (all must match) or **OR** (any can match). |
| `showClearAll` | Toggle | `true` | Show "Clear All" button to reset all filter selections. |
| `enableVisualFilterBuilder` | Toggle | `false` | Show the visual filter builder UI (DevExtreme-style AND/OR group editor). |

### Filter Configuration

Filter types are configured in the Search Results web part's `queryTemplate` and through the store's `filterConfig` array. Each filter entry defines:

| Setting | Options | Description |
|---------|---------|-------------|
| `managedProperty` | Any refinable property | SharePoint managed property to refine on. |
| `displayName` | Free text | Label shown above the filter group. |
| `filterType` | `checkbox`, `daterange`, `slider`, `people`, `taxonomy`, `tagbox`, `toggle` | UI control type. |
| `operator` | `OR`, `AND` | Logic within a single filter (multi-value). |
| `maxValues` | Number | Max values to show before "Show more" (0 = unlimited). |
| `defaultExpanded` | Boolean | Whether the filter group starts expanded. |
| `showCount` | Boolean | Show result counts next to each value. |
| `sortBy` | `count`, `name` | Sort filter values by count (descending) or name (alphabetical). |

### Filter Types Reference

| Type | Best For | Notes |
|------|----------|-------|
| **checkbox** | Discrete values (File Type, Content Type, Site) | Multi-select with counts. Supports "search within" text filter. |
| **daterange** | Dates (Created, Modified) | Preset buttons (Today, This Week, Month, Year) + custom range. Uses FQL `range()` tokens. |
| **slider** | Numeric ranges (File Size, View Count) | Range slider with min/max. Supports file size formatting (KB/MB/GB). |
| **people** | Person fields (Author, Editor) | PnP PeoplePicker with Azure AD type-ahead. Resolves claim strings. |
| **taxonomy** | Managed metadata (Department, Category) | Hierarchical tree with expand/collapse. Resolves GP0|#GUID to term labels. |
| **tagbox** | Multi-value text (Tags, Keywords) | Tag-style chips with search. |
| **toggle** | Boolean fields (IsDocument, IsContainer) | Three-state: All / Yes / No. |

---

## 4. Search Verticals Web Part (`SPSearchVerticalsWebPart`)

Tab navigation that switches search verticals (All, Documents, People, Sites, etc.).

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `searchContextId` | Text | — | **Required.** Must match the other web parts. |
| `verticals` | Text (multiline JSON) | — | JSON array of vertical definitions (see below). |
| `showCounts` | Toggle | — | Show result count badge on each tab. |
| `hideEmptyVerticals` | Toggle | — | Hide or dim tabs with zero results. |
| `tabStyle` | Choice | `"tabs"` | Visual style: **tabs**, **pills**, or **underline**. |

### Vertical Definition Format

The `verticals` property expects a JSON array:

```json
[
  {
    "key": "all",
    "label": "All",
    "iconName": "Search",
    "queryTemplate": "{searchTerms}",
    "sortOrder": 1
  },
  {
    "key": "documents",
    "label": "Documents",
    "iconName": "Document",
    "queryTemplate": "{searchTerms} IsDocument:1",
    "resultSourceId": "e7ec8cee-eeee-eeee-eeee-eeeeeeeeeeee",
    "sortOrder": 2
  },
  {
    "key": "people",
    "label": "People",
    "iconName": "People",
    "queryTemplate": "{searchTerms}",
    "resultSourceId": "b09a7990-05ea-4af9-81ef-edfab16c4e31",
    "dataProviderId": "graph-search",
    "audienceGroups": ["group-guid-1"],
    "sortOrder": 3
  }
]
```

| Field | Required | Description |
|-------|----------|-------------|
| `key` | Yes | Unique identifier. Used in URL parameter `v=`. |
| `label` | Yes | Tab display text. |
| `sortOrder` | Yes | Display order (ascending). Required for stable tab ordering. |
| `iconName` | No | Fluent UI icon name. |
| `queryTemplate` | No | Override query template for this vertical. |
| `resultSourceId` | No | SharePoint Result Source GUID for this vertical. |
| `dataProviderId` | No | Override data provider for this vertical (e.g., `"graph-search"` for People). |
| `filterConfig` | No | Per-vertical filter configuration override. |
| `audienceGroups` | No | Azure AD security group GUIDs for audience targeting. Tab hidden from non-members. |

---

## 5. Search Manager Web Part (`SPSearchManagerWebPart`)

Saved searches, shared searches, collections/pinboards, and search history.

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `searchContextId` | Text | — | **Required.** Must match the other web parts. |
| `mode` | Choice | `"standalone"` | **standalone**: renders as a full web part on the page. **panel**: renders as a side panel triggered from the Search Box web part's icon button. |

### Tabs

| Tab | Description |
|-----|-------------|
| **Saved** | User's saved searches. Click to restore full state (query, filters, vertical, sort, layout). |
| **Shared** | Searches shared by other users. Item-level permissions control visibility. |
| **Collections** | Pinboards of individual search results. Supports tags, reordering, sharing. |
| **History** | Chronological search history with click tracking. Auto-logged, deduplicated via SHA-256 hash. |

### Requirements

- All 4 hidden lists must be provisioned (SearchSavedQueries, SearchHistory, SearchCollections, SearchConfiguration).
- The `mode: "panel"` option requires `enableSearchManager: true` on the Search Box web part.
- History cleanup is available via the `cleanupHistory(ttlDays)` API — there is no automatic background cleanup.

---

## Multi-Instance Configuration

To run two independent search experiences on the same page:

1. Add two sets of web parts
2. Set `searchContextId = "search1"` on the first set
3. Set `searchContextId = "search2"` on the second set
4. URL parameters are namespaced automatically: `?search1.q=budget&search2.q=reports`

---

## Promoted Results / Best Bets

Promoted results are configured in the **SearchConfiguration** list (requires admin write access).

Each promoted result is a **single list item** with one match rule and one or more promoted items:

1. Create a new item with `ConfigType = "PromotedResult"`
2. Set `ConfigValue` to JSON:

```json
{
  "matchType": "contains",
  "matchValue": "handbook",
  "promotedItems": [
    {
      "url": "https://contoso.sharepoint.com/handbook",
      "title": "Company Handbook",
      "description": "Official company policies and procedures",
      "imageUrl": "https://contoso.sharepoint.com/handbook/cover.png",
      "position": 1
    }
  ],
  "audienceGroups": ["all-employees-group-guid"],
  "startDate": "2026-01-01T00:00:00Z",
  "endDate": "2026-12-31T23:59:59Z",
  "verticalScope": ["all", "documents"],
  "isActive": true
}
```

3. Set `IsActive = Yes` and optionally set `ExpiresAt` for time-limited promotions.

### Promoted Result Fields

| Field | Required | Description |
|-------|----------|-------------|
| `matchType` | Yes | Rule type: `contains`, `equals`, `regex`, or `kql`. |
| `matchValue` | Yes | The value to match against the user's query text. |
| `promotedItems` | Yes | Array of items to promote when the rule matches. |
| `promotedItems[].url` | Yes | URL of the promoted result. |
| `promotedItems[].title` | Yes | Display title. |
| `promotedItems[].description` | No | Description text. |
| `promotedItems[].imageUrl` | No | Thumbnail image URL. |
| `promotedItems[].position` | Yes | Display order (ascending). |
| `audienceGroups` | No | Azure AD security group GUIDs. Empty = visible to all. |
| `startDate` | No | ISO date — rule active from this date. |
| `endDate` | No | ISO date — rule expires after this date. |
| `verticalScope` | No | Array of vertical keys this rule applies to. Empty = all verticals. |
| `isActive` | No | Set `false` to disable without deleting. |

To match multiple keywords, create separate list items — one rule per item.
