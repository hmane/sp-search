# SP Search — Admin Configuration Guide

This guide documents the current SharePoint search solution as shipped after the product cleanup and Sprint 4 work. The default authoring model is intentionally close to PnP Modern Search: a Search Box, Results, Filters, and optional Verticals web part all share the same `searchContextId`.

## Core Setup Rule

Every connected web part on the page must use the same `searchContextId`.

- Use a unique value such as `hr-search` or `policies-search` when a page hosts more than one search experience.
- `default` is safe for a single search experience on a page.
- The Results web part is the canonical source of truth. The Box, Filters, Verticals, and Manager web parts should match its context ID.

## Starter Experience

Fresh installs now provision a usable starter experience without immediate manual JSON editing.

| Web part | Starter default |
|----------|-----------------|
| Search Box | Placeholder `Search this site`, suggestions enabled, Search Manager enabled, query input transformation set to `{searchTerms}` |
| Search Results | Scope `This site`, query template `{searchTerms}`, 10 results per page, paging on, sort on, result count on, layouts `List`, `Compact`, `Grid` |
| Search Filters | File type, modified date, and author filters |
| Search Verticals | `All`, `Documents`, `Pages`, `Sites` |
| Search Manager | Panel mode, history/saved searches/collections enabled |

## Search Box

The Search Box owns input behavior and query transformation.

| Property | Default | Notes |
|----------|---------|-------|
| `searchContextId` | `default` | Must match the Results web part |
| `placeholder` | `Search this site` | Input placeholder text |
| `debounceMs` | `300` | Delay before search dispatch while typing |
| `searchBehavior` | `both` | Search on Enter and search button |
| `enableScopeSelector` | `false` | Uses custom search scopes configured on the box |
| `enableSuggestions` | `true` | Uses history/trending/property suggestions |
| `enableQueryBuilder` | `false` | Advanced builder UI |
| `enableKqlMode` | `false` | Exposes raw KQL input mode |
| `enableSearchManager` | `true` | Shows Saved/History/Insights entry point |
| `searchInNewPage` | `false` | Useful for landing-page search boxes |
| `newPageUrl` | empty | Required only when `searchInNewPage` is enabled |
| `queryInputTransformation` | `{searchTerms}` | Applied in the orchestrator before the query is built |

### Notes

- `queryInputTransformation` belongs to the execution path, not the input UI. Example: `Title:{searchTerms} OR Path:{searchTerms}`.
- Use `searchInNewPage` for home-page hero search boxes that should navigate to a dedicated results page.

## Search Results

The Results web part owns query scope, returned properties, layouts, sorting, paging, and scenario presets.

### Data and Query

| Property | Default | Notes |
|----------|---------|-------|
| `searchContextId` | `default` | Required |
| `searchScope` | `currentsite` | Starter behavior is site-scoped, similar to common PnP site-search pages |
| `searchScopePath` | empty | Used only when scope is `custom` |
| `queryTemplate` | `{searchTerms}` | Supports browse scenarios too |
| `resultSourceId` | empty | Optional SharePoint result source |
| `selectedPropertiesCollection` | Starter columns | Admin-configured result/grid columns |
| `refinementFiltersCollection` | empty | Pre-applied FQL/KQL-style restrictions |
| `collapseSpecification` | empty | SharePoint collapse spec |
| `enableQueryRules` | `true` | SharePoint query rules |
| `trimDuplicates` | `true` | SharePoint duplicate trimming |

### Starter Selected Properties

The starter results configuration exposes these admin-facing columns:

- `Title`
- `Author`
- `LastModifiedTime`
- `FileType`
- `Size`
- `Path`
- `SiteName`

The runtime also merges core system properties such as `HitHighlightedSummary`, preview URLs, IDs, and thumbnail fields, so list/card previews still work even if they are not shown as visible columns.

### Layouts and Display

| Property | Default | Notes |
|----------|---------|-------|
| `layoutPreset` | `general` | Applies a starter scenario profile |
| `defaultLayout` | `list` | Initial layout |
| `showListLayout` | `true` | Default primary reading view |
| `showCompactLayout` | `true` | Dense scan view |
| `showGridLayout` | `true` | Power-user table view |
| `showCardLayout` | `false` | Opt-in |
| `showPeopleLayout` | `false` | Opt-in |
| `showGalleryLayout` | `false` | Opt-in |
| `showResultCount` | `true` | Shows total results above results |
| `showSortDropdown` | `true` | Shows sort picker |

### Paging and Sort

| Property | Default | Notes |
|----------|---------|-------|
| `pageSize` | `10` | Starter value aligned to common search pages |
| `showPaging` | `true` | Server-side results paging |
| `pageRange` | `5` | Number of pages shown in pager |
| `sortablePropertiesCollection` | Modified desc, Title asc | Starter sort options |

### Scenario Presets

Built-in scenario presets:

- `general`
- `documents`
- `people`
- `news`
- `media`
- `hub-search`
- `knowledge-base`
- `policy-search`
- `custom`

Preset selection updates:

- query template
- default layout
- available layouts
- selected properties
- sort fields

Changing any individual layout toggle or the default layout reverts the preset selection to `custom`.

### DataGrid Notes

The DataGrid is meant for power users and includes:

- dynamic columns from `selectedPropertiesCollection`
- column chooser
- column resize
- fullscreen view
- row selection and bulk actions
- CSV and XLSX export
- persisted column state in local storage

The removed DevExtreme filter row is intentional. It only filtered the current loaded page, not the full result set, which was misleading in a search product.

## Search Filters

The Filters web part owns refiner configuration and how multiple filters combine.

| Property | Default | Notes |
|----------|---------|-------|
| `searchContextId` | `default` | Must match Results |
| `applyMode` | `instant` | `instant` or `manual` |
| `operatorBetweenFilters` | `AND` | Cross-filter logic |
| `showClearAll` | `true` | Renders clear-all action |
| `enableVisualFilterBuilder` | `false` | Advanced builder UI |
| `filtersCollection` | File type, modified date, author | Starter refiners |

### Starter Filters

| Managed property | Label | Filter type |
|------------------|-------|-------------|
| `FileType` | File type | `checkbox` |
| `LastModifiedTime` | Modified date | `daterange` |
| `AuthorOWSUSER` | Author | `people` |

### Notes

- People filters should use `AuthorOWSUSER`, not `Author`. The filter web part normalizes legacy `Author` people filters to `AuthorOWSUSER`.
- Date range, people, and toggle filters can render without returned refiner buckets, which allows useful starter filters even on clean pages.
- `operatorBetweenFilters = OR` is implemented in the provider path and produces cross-property `or(...)` FQL for SharePoint search.

## Search Verticals

The Verticals web part owns tabs, per-vertical query overrides, per-vertical provider routing, and optional default-layout switching.

| Property | Default | Notes |
|----------|---------|-------|
| `searchContextId` | `default` | Must match Results |
| `verticalsCollection` | Starter tabs | Preferred over legacy JSON |
| `defaultVertical` | `all` | Activated on initial load |
| `showCounts` | `true` | Count badges on tabs |
| `hideEmptyVerticals` | `false` | Keep or hide zero-count tabs |
| `tabStyle` | `tabs` | `tabs`, `pills`, or `underline` |

### Starter Verticals

| Key | Query template |
|-----|----------------|
| `all` | `{searchTerms}` |
| `documents` | `{searchTerms} IsDocument:1` |
| `pages` | `{searchTerms} (contentclass:STS_ListItem_WebPageLibrary OR contentclass:STS_ListItem_PublishingPages)` |
| `sites` | `{searchTerms} contentclass:STS_Site` |

### Per-Vertical Overrides

Each vertical can optionally set:

- `queryTemplate`
- `resultSourceId`
- `dataProviderId`
- `defaultLayout`
- external-link behavior
- audience targeting

Example:

- `dataProviderId = graph-people`
- `defaultLayout = people`

This gives a true Graph-backed People vertical instead of a SharePoint-file result set filtered to people-like properties.

## Search Manager

The Search Manager is not a PnP parity feature. It is a product extension that consolidates saved searches, history, collections, zero-result health, and insights.

| Property | Default | Notes |
|----------|---------|-------|
| `searchContextId` | `default` | Must match Results |
| `mode` | `panel` | `panel` or `standalone` |
| `enableSavedSearches` | `true` | Saved searches tab |
| `enableSharedSearches` | `true` | Shared searches tab |
| `enableCollections` | `true` | Collections tab |
| `enableHistory` | `true` | History tab |
| `enableAnnotations` | `false` | Extra annotations surface |
| `maxHistoryItems` | `50` | History page size |

### Tabs

- `Saved Searches`
- `History`
- `Collections`
- `Health`
- `Insights`

## Graph Requirements

Graph-backed People search and org-chart traversal require Microsoft Graph permissions.

| Capability | Requirement |
|------------|-------------|
| Graph People vertical | `People.Read` / configured Graph search permission path for `/search/query` |
| Org chart manager/direct reports | `User.Read.All` |

If Graph permission is not approved:

- SharePoint search still works
- Graph people verticals fall back to registered SharePoint providers where possible
- org-chart UI hides itself gracefully

## Recommended Authoring Patterns

### General site search

- Search Box
- Verticals
- Filters
- Results
- Search Manager in panel mode

### Document center

- Results preset: `documents`
- Filters: file type, modified date, author, site
- Layouts: `list`, `compact`, `grid`

### People directory

- Results preset: `people`
- A People vertical with `dataProviderId = graph-people`
- Default vertical layout: `people`

## Validation and Edit-Mode Warnings

The Results web part shows edit-mode `MessageBar` warnings for common misconfigurations, including:

- default layout not enabled
- grid enabled with no columns
- sparse grid columns
- query template without `{searchTerms}`
- card/gallery without thumbnail property
- people layout without profile fields
- invalid managed property names

These warnings are advisory. They do not block rendering, but they should be resolved before production rollout.
