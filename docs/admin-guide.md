# SP Search — Admin Configuration Guide

This guide documents the current SharePoint search solution as shipped through Sprint 4 (audit fixes). The default authoring model is intentionally close to PnP Modern Search: a Search Box, Results, Filters, and optional Verticals web part all share the same `searchContextId`.

**Sprint 4 highlights:** operatorBetweenFilters now functional, queryInputTransformation triggers re-search, Clear All Filters button, XLSX export in DataGrid, scope collection editor in property pane, accessibility improvements, sovereign cloud Teams sharing.

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

### Result Link Behavior

The Results web part lets the admin pick what happens when a user clicks a result title.

| Property | Options | Default | Notes |
|----------|---------|---------|-------|
| `resultClickTarget` | `panel` / `newTab` / `sameTab` / `sidePanel` | `panel` | See table below |
| `documentLinkMode` | `file` / `propertiesForm` | `file` | For document results: open the file (Office Online / PDF viewer) or the SharePoint properties form |
| `listItemLinkMode` | `displayForm` / `editForm` | `displayForm` | For list-item results: open the read-only display form or the edit form |

| `clickTarget` | Previewable files (Office / PDF / image / txt / csv / json / xml) | Everything else |
|---|---|---|
| `panel` (default) | Opens an in-page **Modal popup preview** | New tab |
| `newTab` | New tab | New tab |
| `sameTab` | Current tab (replaces the search page) | Current tab |
| `sidePanel` | Opens the result Detail Panel on the right of the page (requires `enablePreviewPanel = true`) | Opens the Detail Panel |

Behind the scenes:

- All result anchors carry `data-interception="off"` so SharePoint Modern's SPA navigation hijacker does not intercept the click. Without this, `target="_blank"` and `e.preventDefault()` would both be ignored by the SP shell. (Internal — admins don't configure this; documented here so support knows the mechanism.)
- The Modal popup uses `<embed type="application/pdf">` for PDFs and a sandboxed `<iframe>` for Office docs. Sandbox tokens deliberately omit `allow-top-navigation` so a misbehaving Office Online runtime can't break out of the Modal.

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

## Search Manager (user-facing)

The Search Manager is not a PnP parity feature. It is a product extension that consolidates saved searches, shared searches, history, and collections. The manager surface forked into two web parts — the **user-facing** Search Manager (this section) and the **SP Search Admin Manager** (see [Admin Manager](#admin-manager)). The user-facing variant surfaces only end-user tabs; admin diagnostics live in the Admin Manager.

| Property | Default | Notes |
|----------|---------|-------|
| `searchContextId` | `default` | Must match Results |
| `mode` | `panel` | `panel` or `standalone` |
| `defaultTab` | `saved` | One of `saved` / `history` / `collections` |
| `enableSavedSearches` | `true` | Saved searches tab |
| `enableSharedSearches` | `true` | Shared searches tab |
| `enableCollections` | `true` | Collections tab |
| `enableHistory` | `true` | History tab |
| `enableAnnotations` | `false` | Extra annotations surface |
| `maxHistoryItems` | `50` | History page size |

### Tabs (user variant)

- `Saved Searches`
- `History`
- `Collections`

## Admin Manager

The SP Search Admin Manager is the admin-only diagnostics web part. It renders **only** when the current user has `ManageWeb` (Owner/Admin) permission; everyone else sees nothing. It shares the store with the rest of the page via `searchContextId`, so the diagnostics reflect the active search experience.

| Property | Default | Notes |
|----------|---------|-------|
| `searchContextId` | `default` | Match the Results web part to inspect that experience |
| `defaultTab` | `dashboard` | One of `dashboard` / `health` / `insights` |
| `enableDashboard` | `true` | Admin Dashboard tab (Content Coverage, Search Quality, Zero-Result Queries) |
| `enableHealth` | `true` | Health tab — zero-result queries replay panel |
| `enableInsights` | `true` | Insights tab — stat cards, top queries, CTR, daily volume |
| `coverageProfilesCollection` | seeded by `Setup-SPSearchSite.ps1` | See [Coverage Profiles](#coverage-profiles-admin-manager) |
| `expectedSiteUrls` | empty | Drives gap-analysis on the Dashboard tab |
| `audienceGroups` | empty | Azure AD group object IDs; leave blank to show to all Owners/Admins |

### Tabs (admin variant)

| Tab | Purpose |
|---|---|
| `Dashboard` | Aggregated metrics: Content Coverage (item count, freshness, file-type breakdown, site distribution), Search Quality (CTR, zero-result rate), Zero-Result Queries |
| `Health` | Zero-result queries with one-click replay and the underlying ranked list |
| `Insights` | Search activity over time: stat cards, top queries, daily volume |
| `Pre-Flight` | Tenant readiness checklist — Graph permission, hidden lists, SearchHistory item-level security, schema mappings, content source. Single-page diagnostic admins run after install. |

Pre-Flight is admin-variant-only and always renders for the admin variant — there's no toggle for it.

## Coverage Profiles (Admin Manager)

The Admin Manager's **Dashboard → Content Coverage** section visualises item
count, freshness, and per-site gap analysis. It is driven by the
`coverageProfilesCollection` property pane field — a list of profiles that
each name one or more SharePoint URLs to enumerate.

### Default seeding (T4.D4)

When you run `Setup-SPSearchSite.ps1`, the script seeds **one tenant-aware
coverage profile** that points at the actual top-5 document libraries on the
target site (any list with `BaseTemplate = 101` that is not hidden). The
discovery uses `Get-PnPList -Includes BaseTemplate, Hidden, RootFolder` and
converts each `RootFolder.ServerRelativeUrl` to an absolute URL against the
tenant root.

```powershell
# Default — discover and seed top-5 actual libraries on the site
.\scripts\Setup-SPSearchSite.ps1 -SiteUrl <site> -ClientId <id>

# Configure how many libraries to seed (1-50)
.\scripts\Setup-SPSearchSite.ps1 -SiteUrl <site> -ClientId <id> -MaxSeededLibraries 10

# Legacy test-tenant flag (kept for backward compatibility; expects libraries
# named per the retired test-data fixture)
.\scripts\Setup-SPSearchSite.ps1 -SiteUrl <site> -ClientId <id> -UseTestData
```

### Empty state (no profiles configured)

If you add the Admin Manager web part to a page **without** running
`Setup-SPSearchSite.ps1`, the manifest default for
`coverageProfilesCollection` is `[]` — and the Dashboard's Content Coverage
section renders a help message:

> **No coverage profiles configured.**
> Configure coverage profiles in the web part property pane to begin
> monitoring item count, freshness, and gap analysis against your expected
> sites.

To configure profiles by hand, open the web part property pane → **Coverage
profiles** → **Manage profiles**, then add entries with at least a `title`
and one or more `sourceUrls` (comma-separated).

## Graph Requirements

Graph-backed People search, org-chart traversal, and audience targeting require Microsoft Graph permissions.

| Capability | Requirement |
|------------|-------------|
| Graph People vertical | `People.Read` / configured Graph search permission path for `/search/query` |
| Org chart manager/direct reports | `User.Read.All` |
| Audience targeting (verticals, refiners, web parts, promoted results) | `User.Read` — least-privilege scope for `/me/memberOf` per [Microsoft Learn](https://learn.microsoft.com/en-us/graph/api/user-list-memberof?view=graph-rest-1.0) |

Approve each permission at **SharePoint admin centre → Advanced → API access**. Pending approval, audience-targeted content stays hidden (fail-closed): verticals / refiners / web parts gated to specific Azure AD groups will be invisible to every user until the scope is approved.

If Graph permission is not approved:

- SharePoint search still works
- Graph people verticals fall back to registered SharePoint providers where possible
- org-chart UI hides itself gracefully
- audience-targeted items hide for all users (fail-closed); non-targeted items remain visible

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

## Property pane help anchors (T4.D11)

Each property pane group on the SP Search web parts renders a "Help:
&lt;topic&gt;" link as its first field. The link opens this guide at the
relevant section. Anchors used by the help links:

<a id="quick-start"></a>

### quick-start

Results web part → **Get started** group. Documents scenario presets
(general / documents / news / people / media / hub-search /
knowledge-base / policy-search / account-documents) and how the
preset picker rewrites layouts + selected properties + filter
suggestions in one step. See [Starter Experience](#starter-experience).

<a id="results-data"></a>

### results-data

Results web part → **Data** group. Covers search scope (all / site /
hub / list-by-url), query template tokens (`{searchTerms}`,
`{Site.URL}`), and managed property pickers (selected, compact,
grid, sortable, refinement). See [Search Results](#search-results)
and [Validation and Edit-Mode Warnings](#validation-and-edit-mode-warnings).

<a id="results-layouts"></a>

### results-layouts

Results web part → **Layouts** group. Documents the six layouts
(List, Compact, Card, People, Grid, Gallery), which preset enables
which layout, and how the DataGrid column chooser works. See
[Search Results](#search-results).

<a id="box-search"></a>

### box-search

Search Box → **Search** group. Placeholder text, search-on-Enter vs
search-on-button, debounce, scope selector. See [Search Box](#search-box).

<a id="box-navigation"></a>

### box-navigation

Search Box → **Navigation** group. Toggle between same-page search
(replaces results in place) and new-page navigation (sends the
query to a dedicated results page via the configured query
parameter name). See [Search Box](#search-box).

<a id="box-suggestions"></a>

### box-suggestions

Search Box → **Suggestions** group. Recent searches, frequent
queries (per-user), SharePoint search suggestions, managed-property
shortcuts, quick results. See [Search Box](#search-box).

<a id="filters-config"></a>

### filters-config

Filters → **Filters** group. Manage the refiner collection
(checkbox / dropdown / date-range / slider / people / taxonomy /
tag-box / toggle). Configure managed property, display name, URL
alias, max values, sort, dependencies. See [Search Filters](#search-filters).

<a id="filters-behavior"></a>

### filters-behavior

Filters → **Behavior** group. Apply mode (Instant vs Manual),
Show Clear All button, operator between filters (AND vs OR),
visual filter builder toggle. See [Search Filters](#search-filters).

<a id="verticals-config"></a>

### verticals-config

Verticals → **Verticals** group. Configure the vertical collection
(key, label, icon, KQL query template, result-source ID, data
provider id). See [Search Verticals](#search-verticals).

<a id="manager-user-tabs"></a>

### manager-user-tabs

Search Manager → **User tabs** group. Toggle the four user-facing
tabs: Saved Searches, Shared Searches, Collections, History. See
[Search Manager (user-facing)](#search-manager-user-facing).

<a id="adminmgr-coverage"></a>

### adminmgr-coverage

Admin Manager → **Monitoring** group. Add and configure coverage
profiles. Each profile names one or more SharePoint URLs and a
query template; the Dashboard tab's Content Coverage section reports
item count and freshness per profile. See [Coverage Profiles (Admin
Manager)](#coverage-profiles-admin-manager).

> Help links surface a subset of groups today (Quick Start, Data,
> Layouts on Results). Coverage will expand to every group on every
> web part in future passes; the helper `propertyPaneGroupHelp` is
> the durable contract for adding new ones —
> `propertyPaneGroupHelp('anchor-id', 'Help: <topic>')` returns a
> `PropertyPaneLink` field admin authors paste at the top of a
> group's `groupFields` array.

To override the link base URL (e.g. for tenants that mirror SP
Search docs internally), call
`setPropertyPaneHelpBaseUrl('https://intranet.contoso.com/wikis/sp-search/admin-guide')`
from the web part's `onInit()` before the property pane builds.
