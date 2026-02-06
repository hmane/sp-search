**SP SEARCHEnterprise SharePoint Search SolutionRequirements & Design Specification**



| Version: | 1.4 |
| --- | --- |
| Date: | February 5, 2026 |
| Author: | Hemant Mane |
| Status: | Draft |


# Table of Contents


# 1. Executive Summary


## 1.1 Purpose

SP Search is an enterprise-grade SharePoint search solution built as a set of SPFx web parts. It replaces and extends the capabilities of PnP Modern Search v4 with a modern component architecture, rich interactive UI powered by DevExtreme and the spfx-toolkit library, and unique features including saved searches, search sharing, search collections, and a comprehensive result detail experience.

## 1.2 Problem Statement

The current PnP Modern Search v4 solution, while feature-rich, presents several limitations for organizations requiring a tightly integrated, highly interactive search experience:
- Handlebars-based templating limits interactivity and makes complex UI patterns difficult to implement.
- Integration with enterprise component libraries (DevExtreme, custom SPFx toolkit components) is not natively supported.
- No built-in support for saved searches, search sharing, or search result collections.
- Dynamic data connections between web parts are unreliable and difficult to debug.
- Limited extensibility model often leads to errors when attempting customization.
- No user-centric search features such as search history, pinboards, or result annotations.

## 1.3 Solution Overview

SP Search delivers five interconnected web parts that communicate via a shared Zustand store distributed through an SPFx library component. The architecture functions as a **"Search Operating System"** — decoupling state (Zustand), UI (React), transport (SPFx Library Component), and data fetching (pluggable `ISearchDataProvider`) into clean layers. Data retrieval is abstracted behind a provider model: SharePoint Search API via PnPjs ships as the default provider, with Microsoft Graph Search as an optional provider for people, Teams messages, and external connector content. DevExtreme powers rich data presentation, and the spfx-toolkit component library provides enhanced UI/UX patterns.


| # | Web Part | Description |
| --- | --- | --- |
| 1 | Search Box | Query input with smart suggestions, recent searches, query builder, and search scope selector. |
| 2 | Search Results | Rich result display with configurable layouts (DataGrid, Cards, List, Compact, People, Gallery), detail panel, bulk actions, and inline actions. |
| 3 | Search Filters | Refinement filters using DevExtreme controls with visual filter builder, date range, people picker, and taxonomy tree. |
| 4 | Search Verticals | Tab-based content scoping with badge counts, audience targeting, and dynamic configuration. |
| 5 | Search Manager | Saved searches, shared searches, search collections/pinboards, search history, and result annotations. Available as both a standalone web part and an integrated panel from the Search Box. |


## 1.4 Key Differentiators from PnP Modern Search


| Capability | PnP Modern Search v4 | SP Search |
| --- | --- | --- |
| Component Library | Handlebars + Adaptive Cards | DevExtreme + spfx-toolkit + Fluent UI |
| Result Layouts | Custom Handlebars templates | 6 built-in configurable layouts + custom via LayoutRegistry |
| Inter-WP Communication | SPFx Dynamic Data | Shared Zustand store via library component (multi-instance isolated) |
| Data Source | SharePoint Search only | Pluggable ISearchDataProvider — SharePoint Search (default) + Microsoft Graph + custom |
| Saved Searches | Not supported | Full save/share/manage lifecycle |
| Search Sharing | URL copy only | URL, Email, Teams, User-specific sharing |
| Result Detail | Custom template only | Built-in panel with preview, metadata, version history, actions |
| Result Collections | Not supported | Pinboards with named collections |
| Result Collapsing | Manual KQL only | Built-in CollapseSpecification with "Show N more" UI |
| Promoted Results | Server-side Query Rules only | Client-side Best Bets with audience targeting, scheduling, and per-vertical scoping |
| Bulk Operations | Not supported | Multi-select with share, download, copy, compare |
| Export | Not built-in | Excel, CSV, PDF via DevExtreme |
| Query Builder | KQL text only | Visual query builder + KQL |
| Schema Helper | Not available | Managed Property Picker in property pane (admin) |
| Extensibility | Web components + Handlebars | Typed provider/registry model (data, suggestions, actions, layouts, filters) |


## 1.5 Reference Repositories

SP Search draws architecture patterns and search logic from PnP Modern Search v4. Claude Code should clone and reference these repositories during development to reuse proven patterns rather than re-inventing them.
- **PnP Modern Search v4 (primary reference):** https://github.com/microsoft-search/pnp-modern-search — Source code in search-parts/src/. Key directories: webparts/ (all four web part implementations), dataSources/ (SharePointSearchDataSource — query construction, refinement token handling, result mapping), services/ (TokenService, SearchService, SuggestionService), layouts/ (built-in layout templates). Study the data source layer and token resolution system. We reuse equivalent logic via PnPjs.
- **PnP Modern Search Documentation:** https://microsoft-search.github.io/pnp-modern-search/ — Key pages: /usage/search-results/ (result config, data sources, slots, tokens), /usage/search-results/data-sources/sharepoint-search/ (SharePoint Search API specifics), /usage/search-box/ (search box config, suggestions), /usage/search-filters/ (Refiner vs Static, deep linking via URL param f, multi-source merging), /usage/search-verticals/ (vertical config, audience targeting, count queries), /usage/search-results/tokens/ (query template token syntax).
- **spfx-toolkit (internal library):** https://github.com/hmane/spfx-toolkit — Key docs: README.md (critical import patterns for bundle size), SPFX-Toolkit-Usage-Guide.md (component API), CLAUDE.md (AI dev instructions). Components: Card, VersionHistory, DocumentLink, UserPersona, ErrorBoundary, Toast, FormContainer, WorkflowStepper. Hooks: useLocalStorage, useViewport, useCardController, useErrorHandler. Utilities: SPContext, BatchBuilder, createPermissionHelper, createSPExtractor. Always use direct path imports.
- **PnP Extensibility Samples:** https://github.com/microsoft-search/pnp-modern-search-extensibility-samples — Custom layouts, web components, data sources, query modifiers.
- **PnP Scenarios:** https://microsoft-search.github.io/pnp-modern-search/scenarios/ — Ready-made walkthroughs: search pages, filters, verticals, people search, promoted results, query string integration.

# 2. Technology Stack


## 2.1 Core Framework


| Technology | Version | Purpose |
| --- | --- | --- |
| SharePoint Framework | 1.21.1 | SPFx web part platform |
| React | 17.0.1 | UI component framework |
| TypeScript | 4.7+ | Type-safe development |
| PnPjs (SP) | 3.x | SharePoint Search API client (default data provider) |
| Microsoft Graph Client | 3.x | Graph Search API client (people, Teams, external connectors) |


## 2.2 UI Libraries


| Library | Version | Usage |
| --- | --- | --- |
| spfx-toolkit | Latest | Card, VersionHistory, DocumentLink, ErrorBoundary, Toast, UserPersona, FormContainer, hooks, utilities |
| DevExtreme | 22.2.x | DataGrid, FilterBuilder, TagBox, TreeView, DateRangeBox |
| devextreme-react | 22.2.x | React wrappers for DevExtreme components |
| Fluent UI v8 | 8.106.x | Panel, CommandBar, Persona, Shimmer, Icons, Theme |
| @pnp/spfx-controls-react | 3.x | PeoplePicker, TaxonomyPicker, FilePicker |


## 2.3 State Management & Utilities


| Library | Version | Purpose |
| --- | --- | --- |
| Zustand | 4.x | Shared state across web parts + URL sync |
| React Hook Form | 7.x | Search configuration forms, admin settings |


## 2.4 Architecture Pattern

The solution uses an SPFx Library Component to distribute Zustand store instances across all SP Search web parts on a page. A **store registry** pattern supports **multi-instance isolation** — multiple independent search contexts can coexist on the same page (e.g., "Policies Search" + "People Search") without cross-talk.

- **Library Component:** sp-search-store — exposes `getStore(searchContextId): Store` and `disposeStore(searchContextId): void`
- **Store Registry:** Each web part declares a `searchContextId` in its property pane (default: auto-generated GUID). All web parts sharing the same `searchContextId` share one Zustand store instance. Web parts with different IDs are fully isolated.
- **Store Slices:** querySlice, filterSlice, verticalSlice, resultSlice, uiSlice, userSlice
- **URL Sync:** Bi-directional sync between store and URL hash/query parameters. Multi-context pages use namespaced params: `?ctx1.q=budget&ctx1.f=...&ctx2.q=john` (context prefix derived from `searchContextId`). Single-context pages use short params (`q`, `f`, `v`, etc.) for clean URLs.
- **Lazy Loading:** Heavy components (DevExtreme DataGrid, Preview Panel) loaded on demand
- **Tree Shaking:** All spfx-toolkit imports use direct path imports for minimal bundle size


# 3. Web Part Specifications


## 3.1 Search Box Web Part

The Search Box is the primary query entry point. It goes beyond a simple text input by providing smart suggestions, recent search history, scope selection, and access to both the visual query builder and the Search Manager panel.

### 3.1.1 Core Features

- **Debounced Search Input:** Text input with configurable debounce (default 300ms). Supports free text and KQL queries.
- **Smart Suggestions:** As the user types, show suggestions from three sources: recent searches by the user, popular/trending queries across the organization (from SearchSavedQueries list), and SharePoint managed property values matching the input.
- **Recent Searches:** Dropdown showing the user's last N searches (configurable, default 10). Each entry shows the query text, timestamp, and result count. Stored per-user in the SearchSavedQueries list with a type discriminator.
- **Search Scope Selector:** A dropdown or pill selector allowing users to narrow search scope before executing. Preconfigured scopes: Current Site, Current Hub, All SharePoint, Specific Library (admin-configured). Scopes map to KQL path restrictions or result source IDs.
- **Visual Query Builder Toggle:** A button that expands an advanced query builder panel below the search box. Uses a DevExtreme-inspired filter builder UI with property dropdowns (from available managed properties), operator selection, and value pickers.
- **Search Manager Button:** Icon button that opens the Search Manager as a Fluent UI Panel from the right side. Provides quick access to saved searches, collections, and history without needing a separate web part on the page.
- **Clear and Reset:** Clear button resets query, filters, verticals, and sort to default state.
**PnP Reference:** Study the PnP Search Box web part at https://microsoft-search.github.io/pnp-modern-search/usage/search-box/ for suggestion providers, query string passthrough, and dynamic data source connections. Source: search-parts/src/webparts/searchBox/. Reuse the token replacement logic from search-parts/src/services/tokenService/ for dynamic query templates ({searchTerms}, {Site.ID}, {Today}, etc.).
- **Search on Enter:** Configurable behavior: search on Enter key, search on button click, or both.

### 3.1.2 Property Pane Configuration


| Property | Type | Default | Description |
| --- | --- | --- | --- |
| placeholder | string | "Search SharePoint..." | Placeholder text |
| debounceMs | number | 300 | Debounce delay in milliseconds |
| enableSuggestions | boolean | true | Enable smart suggestions dropdown |
| enableRecentSearches | boolean | true | Show recent searches |
| recentSearchCount | number | 10 | Max recent searches to display |
| enableScopeSelector | boolean | true | Show search scope selector |
| searchScopes | ISearchScope[] | Default scopes | Configurable search scopes |
| enableQueryBuilder | boolean | false | Show query builder toggle |
| enableSearchManager | boolean | true | Show Search Manager panel button |
| searchBehavior | enum | "both" | "onEnter" | "onButton" | "both" |


## 3.2 Search Results Web Part

The Search Results web part is the core display engine. It retrieves data through an **abstracted data provider** (`ISearchDataProvider`) and renders results using one of six configurable layouts. It integrates deeply with spfx-toolkit components for a rich, interactive experience.

### 3.2.1 Data Source: ISearchDataProvider Abstraction

The Search Results web part does **not** directly call PnPjs or Graph. Instead, it delegates all data fetching to an `ISearchDataProvider` implementation registered via the DataProviderRegistry (see Section 4.4.5). This allows different verticals to use different backends — for example, "Documents" can use SharePoint Search (better custom refiners) while "People" uses Microsoft Graph (better org chart data, presence, Teams integration).

**Built-in Data Providers:**

**SharePointSearchProvider (default):**
- Uses PnPjs `sp.search()` for KQL-based queries
- Full support for refiners, result sources, managed properties, CollapseSpecification
- Best for: document search, site content, custom managed properties, complex refinement

**GraphSearchProvider (Phase 4+):**
- Uses Microsoft Graph Search API (`/search/query`) via the SPFx `MSGraphClientV3`
- Supports: files, sites, messages, events, externalItems (Graph connectors), people, acronyms
- Best for: people search (richer profile data, presence), cross-workload search (Teams messages, Outlook), external connector content
- **API Permission:** Requires `Sites.Read.All` and `ExternalItem.Read.All` (managed via SPFx API permission requests)

**Per-Vertical Data Provider Override:** Each vertical definition (`IVerticalDefinition`) can specify a `dataProviderId` to override the default. This means a single search page can mix SharePoint Search and Graph results across different vertical tabs seamlessly.

**Common Query Configuration (applies to all providers):**
- **Query Template:** Configurable query template with token support (e.g., {searchTerms}, {Page.URL}, {User.Name}).
- **Result Source ID:** Optional result source for pre-scoped queries (SharePoint provider only).
- **Selected Properties:** Admin selects which properties to retrieve. These populate layout columns and detail panel fields.
- **Sort Configuration:** Default sort property and direction, with user-overridable sorting in the UI.
- **Paging:** Configurable page size (default 25). Supports both numbered pagination and infinite scroll depending on layout.
- **Query Modification:** Pre-query and post-query hooks for programmatic query manipulation.
- **Trimming & Duplicates:** Configurable result trimming and duplicate removal.
- **Result Collapsing:** Configurable `CollapseSpecification` to group duplicate results (see Section 3.2.6).

### 3.2.2 Result Layouts

**PnP Reference:** Reuse query construction logic from SharePointSearchDataSource.ts (search-parts/src/dataSources/). Handles: KQL query template assembly with token replacement, refinement filter token encoding/decoding (including FQL range() operators), result property mapping, refiner aggregation parsing, and sort handling. Docs: https://microsoft-search.github.io/pnp-modern-search/usage/search-results/data-sources/sharepoint-search/ — Token reference: https://microsoft-search.github.io/pnp-modern-search/usage/search-results/tokens/
SP Search ships with six built-in layouts. Admins configure which layouts are available for the web part instance, and end users can switch between available layouts at runtime via a layout toggle in the toolbar.


**Layout 1: DataGrid (DevExtreme)**


A full-featured data table powered by DevExtreme DataGrid. Best for power users who need to sort, filter, group, and export search results. The grid provides a familiar spreadsheet-like experience with enterprise-grade performance for large result sets.

**Core Features:**
- Column configuration from selected managed properties
- Server-side sorting (via Search API `SortList`) + client-side secondary sort
- Client-side column filtering with type-aware filter editors (see below)
- Column grouping with drag-to-group area
- Column reordering via drag-and-drop
- Column visibility toggle (column chooser)
- Column resizing with min/max width constraints
**PnP Reference:** PnP v4 ships built-in layouts in search-parts/src/layouts/: List, Cards, Details List, People, Custom. Study the layout switching pattern and result slot mapping. Docs: https://microsoft-search.github.io/pnp-modern-search/usage/search-results/layouts/ — SP Search replaces Handlebars templates with React components using DevExtreme DataGrid and spfx-toolkit Card (import from 'spfx-toolkit/lib/components/Card') for accordion grouping, maximize/restore, and responsive grid.
- Row selection (single and multi-select) with Shift+Click range select
- Export to Excel and CSV (via DevExtreme `exportDataGrid`)
- Virtual scrolling for large result sets (DevExtreme `scrolling.mode: 'virtual'`)
- Master-detail row expansion showing inline document preview (same WOPI frame as Detail Panel)
- Fixed columns (pin first/last columns for wide tables)
- Row alternating colors for readability
- Keyboard navigation between cells

**Type-Aware Cell Renderers:**

Every column in the DataGrid uses a custom `cellRender` function matched to the managed property type. No raw values are ever displayed.

| Property Type | Cell Renderer | Behavior |
| --- | --- | --- |
| Title / Link | `TitleCellRenderer` | File type icon (16px) + clickable title link. Click opens Detail Panel. Hover shows full path tooltip. |
| Person / User | `PersonaCellRenderer` | Mini Fluent UI Persona: 24px avatar + display name. Hover shows PersonaCard with email, title, department. Claim strings are resolved and cached. |
| DateTime | `DateCellRenderer` | Relative format ("3 days ago") with absolute date/time in tooltip ("Jan 15, 2026, 3:42 PM"). Uses `Intl.RelativeTimeFormat`. Column sort uses raw ISO value. |
| File Size | `FileSizeCellRenderer` | Human-readable with auto-scaled units: "2.4 MB", "156 KB", "1.2 GB". Column sort uses raw byte value. |
| File Type | `FileTypeCellRenderer` | Fluent UI file type icon (24px) + extension label. Color-coded by category (blue for Office, red for PDF, green for images). |
| URL | `UrlCellRenderer` | spfx-toolkit DocumentLink with truncated display URL. Click opens in new tab. |
| Taxonomy / MMD | `TaxonomyCellRenderer` | Term label as a subtle chip/tag. Hover shows full term path ("Departments > Marketing > Digital"). Multiple terms shown as stacked chips. |
| Boolean | `BooleanCellRenderer` | ✅ / ❌ icon (not text "true"/"false"). |
| Number / Currency | `NumberCellRenderer` | Locale-formatted number via `Intl.NumberFormat`. Currency symbol if configured. Right-aligned. |
| Multi-Value String | `TagsCellRenderer` | Horizontal tag chips with overflow "+N more" indicator. |
| Thumbnail | `ThumbnailCellRenderer` | 40px × 40px thumbnail image from SharePoint preview API. Fallback to file type icon. |
| Generic Text | `TextCellRenderer` | Truncated with ellipsis and full text on hover tooltip. Hit-highlight markup preserved (bold keywords). |

**DataGrid Column Filtering (Client-Side):**

DevExtreme DataGrid supports column-level filter rows. SP Search provides type-aware filter editors:

| Column Type | Filter Editor | Behavior |
| --- | --- | --- |
| Text (Title, etc.) | Text input with "contains" operator | Case-insensitive substring match |
| DateTime | DevExtreme DateBox with range support | Filter to exact date or date range |
| Person | Mini PeoplePicker dropdown | Type-ahead against resolved persona cache, then filter rows |
| File Type | Multi-select dropdown | Checkbox list of unique file types in current results |
| Number | Range inputs (from/to) | Numeric range filter |
| Boolean | Three-state toggle (All / Yes / No) | Filter to true, false, or all |
| Taxonomy | Searchable dropdown | Filter against resolved term labels |

**NOTE:** DataGrid column filtering is **client-side only** — it filters the currently loaded page of results. Server-side filtering is handled by the Search Filters web part via the Zustand store. The DataGrid column filter is a secondary "within results" refinement that doesn't trigger new API calls.

**DataGrid Sorting:**

- **Primary sort (server-side):** Clicking a column header dispatches `setSort()` to the resultSlice, which triggers a new search API call with the updated `SortList` parameter. The sort icon (▲/▼) appears in the column header.
- **Secondary sort (client-side):** Shift+Click on a second column adds a client-side secondary sort within the server-sorted results.
- **Sort persistence:** Sort state is included in URL synchronization (`&s=ModifiedOWSDTM:desc`) so bookmarked/shared URLs preserve sort order.
- **Non-sortable columns:** Columns mapped to non-sortable managed properties have sorting disabled (no hover cursor, no click handler). The Schema Helper marks these.

**DataGrid Responsive Behavior:**
- On screens < 768px, the DataGrid automatically switches to a "card mode" where each row renders as a stacked card (DevExtreme `columnHidingEnabled: true`)
- Priority columns (title, author, modified) remain visible; lower-priority columns hide into an expand row
- The user can manually toggle between grid and card mode on mobile


**Layout 2: Card Layout (spfx-toolkit Card)**


Uses the spfx-toolkit Card component to display each result as a rich, interactive card. Best for content browsing and visual exploration.
- Card with header showing document title, icon, and quick actions
- Card content showing configurable metadata fields
- Accordion pattern for grouping cards by content type, site, or custom property
- Card maximize to expand a single result into a full detail view
- Lazy loading of card content when scrolled into view
- Responsive grid: 1 column on mobile, 2 on tablet, 3-4 on desktop
- Highlight animation when card state changes


**Layout 3: List Layout**


A compact, Google-style search result list. Each result shows title (as link), URL breadcrumb, and a brief excerpt/summary. Best for text-heavy searches where scanning speed matters. This is the default layout.

**Result Card Anatomy:**
```
┌──────────────────────────────────────────────────────────────┐
│  📄 Quarterly Financial Report Q4 2025              [⋯]     │
│  finance.contoso.com > Shared Documents > Reports           │
│  ...total revenue increased by 12% compared to Q3, with     │
│  **operating margins** reaching a record high of 28.4%...   │
│                                                              │
│  👤 John Doe  ·  Modified Jan 15, 2026  ·  DOCX  ·  2.4 MB │
└──────────────────────────────────────────────────────────────┘
```

**Formatting Rules:**
- **Title:** 16px semibold, Fluent UI file type icon (16px) left of title. Clickable — opens Detail Panel (not the document itself). Color: theme primary on hover.
- **URL Breadcrumb:** 12px muted text, shows site > library > folder path. Truncated from the middle for long paths ("finance.contoso.com > ... > Reports"). Each segment is clickable (navigates to that folder).
- **Excerpt:** 14px regular text, 2–3 lines max with ellipsis. **Hit-highlighted keywords** rendered as `<mark>` with subtle background highlight (theme tertiary at 20% opacity). HTML sanitized (no raw `<ddd>` tags from SharePoint).
- **Metadata Line:** 12px muted text, dot-separated: Author persona (mini 20px avatar + name), relative modified date ("3 days ago" with absolute tooltip), file type extension (uppercase), file size (human-readable). All metadata values use the same type-aware formatters as the DataGrid cell renderers.
- **Quick Actions:** Hidden by default, appear as icon row on hover (right side): Open, Preview, Share, Pin, More (⋯). The ⋯ menu contains: Copy link, Download, Open in desktop app, View in library.
- **Selection:** Left-side checkbox appears on hover or when any item is selected (enters multi-select mode). Selected items have a subtle background highlight.
- **Keyboard Navigation:** Arrow keys move focus between results. Enter opens Detail Panel. Space toggles selection. Tab moves to quick actions.


**Layout 4: Compact Layout**


A high-density, single-line-per-result layout. Best for known-item searches or when users need to scan many results quickly.
- Single line: icon + title + author + modified date + file type
- Hover expands to show 2-line excerpt
- Checkbox for multi-select
- Inline quick actions on hover


**Layout 5: People Layout**


A dedicated layout for people search results using Fluent UI Persona components and spfx-toolkit UserPersona.
- Profile photo with presence indicator
- Name, job title, department, office location
- Contact actions: email, chat (Teams deep link), call
- Org chart position (direct reports count, manager name)
- Recent documents by this person (expandable section)
- Profile card on hover (Fluent UI PersonaCard)


**Layout 6: Document Gallery**


A thumbnail grid layout optimized for image and document browsing. Best for media libraries, image searches, and document galleries.
- Thumbnail preview for images, PDFs, Office documents (using SharePoint preview thumbnails API)
- Configurable thumbnail sizes: small (120px), medium (200px), large (300px)
- Overlay on hover showing title, file type, modified date
- Lightbox view for images
- Masonry or fixed grid option
- Infinite scroll pagination

### 3.2.3 Result Detail Panel

When a user clicks on a search result (across any layout), a Fluent UI Panel opens from the right side with comprehensive detail and action capabilities. The panel is the "single source of truth" for a document — everything a user needs to know before deciding to open it.

**Panel Layout (top to bottom):**

```
┌─────────────────────────────────────────────────────┐
│  📄 Quarterly Report Q4 2025.docx           [✕]    │
│  /sites/finance/Shared Documents/Reports            │
│  ─────────────────────────────────────────────────  │
│  [Open ▾]  [Download]  [Copy Link]  [Share]  [Pin] │
│  ─────────────────────────────────────────────────  │
│  ┌─────────────────────────────────────────────┐    │
│  │                                             │    │
│  │        DOCUMENT PREVIEW (WOPI Frame)        │    │
│  │         or Image Preview                    │    │
│  │         or File Type Icon + Download        │    │
│  │                                             │    │
│  └─────────────────────────────────────────────┘    │
│                                                     │
│  METADATA                                           │
│  Author         John Doe [avatar]                   │
│  Modified       Jan 15, 2026, 3:42 PM              │
│  Created        Oct 3, 2025, 9:15 AM               │
│  File Size      2.4 MB                             │
│  Site           Finance Portal                     │
│  Department     Accounting                          │
│  Content Type   Financial Report                    │
│  Tags           Q4, Annual, Board Review            │
│                                                     │
│  VERSION HISTORY                        [View All]  │
│  v3.0  John Doe    Jan 15, 2026  "Final version"   │
│  v2.0  Jane Smith  Jan 10, 2026  "Updated charts"  │
│  v1.0  John Doe    Oct 3, 2025   "Initial draft"   │
│                                                     │
│  RELATED DOCUMENTS                                  │
│  📄 Q3 Report 2025.docx     Finance Portal         │
│  📄 Annual Summary 2025.docx Finance Portal        │
└─────────────────────────────────────────────────────┘
```

**A. Document Preview Implementation**

The preview section is the most important part of the detail panel — it lets users verify a document without opening it.

**Office Documents (Word, Excel, PowerPoint, Visio):**
- Embed via WOPI frame URL: `{siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc={uniqueId}&action=interactivepreview`
- The `uniqueId` (document GUID) is obtained from the `UniqueId` managed property in search results
- Iframe rendered at 100% panel width × 400px height (resizable via drag handle)
- Fallback: If WOPI frame fails to load (timeout or 403), show thumbnail image from SharePoint's thumbnail API: `{siteUrl}/_api/v2.0/drives/{driveId}/items/{itemId}/thumbnails/0/c400x300/content`

**PDF Files:**
- Embed via SharePoint's native PDF viewer: `{siteUrl}/_layouts/15/WopiFrame.aspx?sourcedoc={uniqueId}&action=view`
- Alternative: Use the browser's native PDF renderer if the WOPI frame is blocked by CSP

**Images (PNG, JPG, GIF, SVG, WEBP):**
- Display full-size image inline with `object-fit: contain` and max-height 400px
- Click to open in lightbox overlay (full viewport)
- EXIF metadata shown below image if available (dimensions, camera, date taken)

**Video (MP4, MOV):**
- Embed SharePoint's native video player via Stream/WOPI integration
- Fallback: Show thumbnail with play icon → opens in new tab

**Unsupported File Types:**
- Large file type icon (48px, from Fluent UI file type icons) centered
- File name, size, and modified date displayed below
- "Download to preview" link as the primary action

**Preview Performance:**
- Preview iframe is lazy-loaded (only when the panel opens, not when results load)
- Loading state: Shimmer animation (Fluent UI Shimmer) in the preview area until the iframe signals `load` event
- Error state: File icon + "Preview unavailable" message + "Open in browser" button

**B. Metadata Display (Formatted)**

All metadata values are formatted using type-aware renderers. No raw values are ever shown to the user.

| Managed Property Type | Rendering | Example |
| --- | --- | --- |
| Person/User | Fluent UI Persona (avatar + name + title) | [👤] John Doe, Senior Analyst |
| DateTime | Relative + absolute (tooltip) | "3 days ago" (hover: "Jan 15, 2026, 3:42 PM") |
| File Size | Human-readable with units | "2.4 MB" (from raw bytes: 2,516,582) |
| URL/Link | spfx-toolkit DocumentLink with icon | 📄 Click to open (file type-aware icon) |
| Taxonomy/MMD | Term label with full path on hover | "Marketing" (hover: "Departments > Marketing > Digital") |
| Boolean | Fluent UI Icon + label | ✅ Yes / ❌ No |
| Number/Currency | Locale-formatted with units | "$12,500.00" or "1,234" |
| Multi-value (delimited) | Tag-style chips | [Finance] [Q4] [Board Review] |
| Rich Text / HTML | Sanitized HTML rendered inline | (stripped of scripts, iframes) |
| Empty/Null | Subtle "(Not set)" text in muted color | *(Not set)* |

**C. Version History**

Integrated via spfx-toolkit VersionHistory component:
- Shows up to 5 most recent versions by default, "View All" loads full history
- Each row: version number, author (Persona), date (relative), comment
- Click a version to view it in the preview pane (swaps the WOPI frame URL to the version-specific URL)
- Compare button: opens two versions side-by-side in a new browser tab using SharePoint's native version comparison

**D. Related Documents**

Related documents are loaded lazily after the panel opens (secondary priority to preview and metadata):
- Same library/folder documents: PnPjs query filtered to same `ParentLink` managed property
- Similar metadata documents: SharePoint search query using shared managed property values (same content type + same 2-3 tag values)
- Maximum 5 related documents shown, each as a compact row with icon + title + site name

**E. Quick Actions (Panel Toolbar)**

| Action | Behavior | Icon |
| --- | --- | --- |
| Open | Split button: default "Open in browser", dropdown "Open in desktop app" | OpenFile |
| Download | Direct download via `{siteUrl}/_layouts/15/download.aspx?UniqueId={uniqueId}` | Download |
| Copy Link | Copies sharing URL to clipboard, toast confirmation | Link |
| Share | Opens share flyout (URL / Email / Teams / Users) — reuses §3.5.2 patterns | Share |
| Pin | Opens collection picker flyout to pin to a collection | Pin |
| View in Library | Navigates to the parent document library folder | FolderOpen |

### 3.2.4 Bulk Actions Toolbar

When one or more results are selected, a contextual toolbar appears above the results with available batch operations.
- Share selected items (URL, email, Teams)
- Copy links to clipboard
- Download selected items as individual files
- Pin selected items to a collection
- Export selected metadata to Excel/CSV
- Compare metadata of 2-3 selected items side-by-side

### 3.2.5 Property Pane Configuration


| Property | Type | Default | Description |
| --- | --- | --- | --- |
| dataProviderId | string | "sharepoint" | Data provider to use: "sharepoint" (PnPjs) or "graph" (MS Graph). Overridden per-vertical if set. |
| queryTemplate | string | {searchTerms} | KQL query template |
| resultSourceId | string | (empty) | Result source GUID (SharePoint provider only) |
| selectedProperties | string[] | Default set | Managed properties to retrieve |
| enabledLayouts | LayoutType[] | All layouts | Which layouts are available |
| defaultLayout | LayoutType | "list" | Initial layout on load |
| pageSize | number | 25 | Results per page |
| enableDetailPanel | boolean | true | Enable result detail panel |
| enableBulkActions | boolean | true | Enable multi-select and bulk ops |
| enableExport | boolean | true | Enable export to Excel/CSV |
| sortableProperties | string[] | [] | Properties that allow sorting |
| defaultSort | ISortConfig | Relevance | Default sort property and direction |
| showResultCount | boolean | true | Display total result count |
| collapseSpecification | string | (empty) | SharePoint CollapseSpecification value for result grouping (see 3.2.6) |
| enableSchemaHelper | boolean | true | Show Managed Property Picker helper in property pane (admin only, see 3.2.7) |

### 3.2.6 Result Collapsing (Thread Folding)

Enterprise search results are often polluted by multiple versions of the same document or dozens of emails from the same thread. SP Search supports **result collapsing** to group duplicates and surface only the most relevant item per group.

**Implementation:** The SharePoint Search API supports a `CollapseSpecification` property that groups results by a specified managed property and returns only the top N results per group. Crucially, collapsing is done **server-side before pagination** — meaning the "Top 1" result per group is calculated before the page of 25 results is cut, ensuring accurate pagination and counts. Do NOT attempt client-side grouping in the React render cycle as it breaks pagination.

**⚠ Silent Failure Risk:** `CollapseSpecification` fails **silently** if the managed property isn't Sortable — it returns ungrouped results with no error. The `SharePointSearchProvider` must validate the property's sortability (via the Schema Helper / `getSchema()`) before including `CollapseSpecification` in the query. If the property isn't sortable, surface a warning in the property pane and skip the collapse parameter rather than sending a silently-ignored query.

**Configurable Collapse Fields:**
- `DocumentSignature` — groups near-duplicate documents (same content, different locations)
- `NormUniqueID` — groups document versions (same document, different versions)
- `ConversationID` — groups email threads
- Custom managed property — admin-configurable (e.g., group by `ProjectID` or `ContractNumber`)

**UI Behavior:**
- When results are collapsed, each group leader shows a "Show N more versions" or "N related items" expandable link
- Expanding a collapsed group loads the child results inline (lazy fetch or from cached response)
- The `ISearchResult` interface supports collapsing via: `isCollapsedGroup: boolean`, `childResults: ISearchResult[]`, `groupCount: number`
- Collapsing is disabled by default; admins enable it and select the collapse field via property pane

**Property Pane:**

| Property | Type | Default | Description |
| --- | --- | --- | --- |
| collapseSpecification | string | (empty) | Managed property to collapse on (e.g., "DocumentSignature") |
| collapseMaxResults | number | 1 | Max results per group (1 = show only top result) |
| showCollapseIndicator | boolean | true | Show "N more" expandable indicator |

### 3.2.7 Schema Helper (Managed Property Picker)

The #1 complaint from admins configuring search web parts is not knowing which `RefinableStringXX` maps to which business property. SP Search includes a **Schema Helper** in the property pane that makes configuration dramatically easier.

**Behavior:**
- When the admin opens the property pane for Search Results (or Search Filters), a "Browse Schema" button appears next to managed property fields
- Clicking it fetches the **Search Schema** via the SharePoint Search Administration API and displays a searchable/filterable list of managed properties with their aliases, types, and whether they're queryable/retrievable/refinable/sortable
- Admin can click a property to insert it into the configuration field
- **Permission Check:** The schema API requires Search Admin or Site Collection Admin permissions. If the current user lacks these, the helper falls back to a standard text input with a tooltip explaining "Contact your SharePoint admin for managed property names"

**Implementation Notes:**
- Schema data is cached in sessionStorage after first fetch (schema changes are rare)
- The helper is implemented as a reusable property pane control (`PropertyPaneSchemaHelper`) that can be used across Search Results, Search Filters, and Search Verticals property panes


## 3.3 Search Filters Web Part

The Search Filters web part provides refinement capabilities powered by DevExtreme filter controls and spfx-toolkit UI patterns. Filters connect to the shared Zustand store and immediately update Search Results web parts on the same page.

### 3.3.1 Filter Types


| Filter Type | Component | Best For | Features |
| --- | --- | --- | --- |
| Checkbox List | Fluent UI Checkbox | File type, content type | Multi-select, show count, search within filter values |
| Date Range | DevExtreme DateRangeBox | Modified date, created date | Preset ranges (Today, This Week, This Month, This Year, Custom) |
| People Picker | PnP PeoplePicker | Author, modified by, created by | Type-ahead, multi-select, resolve against AAD |
| Taxonomy Tree | DevExtreme TreeView | Managed metadata columns | Hierarchical expand/collapse, multi-select, search |
| Dropdown / ComboBox | DevExtreme TagBox | Site, department, category | Tag-style multi-select, search, custom items |
| Slider / Range | DevExtreme RangeSlider | File size, numeric fields | Min/max range, step configuration |
| Toggle / Boolean | Fluent UI Toggle | Is checked out, has attachments | Single on/off filter |


### 3.3.2 Filter Behavior

- **Operator Between Filters:** Configurable AND/OR logic between different filter groups.
- **Operator Within Filter:** Configurable AND/OR logic within multi-value selections of a single filter.
- **Filter Counts:** Each filter value shows the count of matching results. Counts update dynamically as filters are applied (the Search API returns updated refiner counts with each query).
- **Expand/Collapse:** Each filter group uses spfx-toolkit Card accordion pattern for clean expand/collapse with persistence.
- **Apply Mode:** Configurable: instant apply (filter on every change) or manual apply (user clicks an Apply button).
- **Clear Filters:** Individual filter clear and global clear all filters button.
- **Search Within Filters:** Text search within filter values for filters with many options.
- **Show More/Less:** Configurable initial count with expand to show all values.
**PnP Reference:** Study PnP Filters web part at https://microsoft-search.github.io/pnp-modern-search/usage/search-filters/ for two-way connection pattern, Refiner vs Static filter distinction, multi-source filter merging, URL deep linking via the f query parameter, and operator handling (OR/AND). Source: search-parts/src/webparts/searchFilters/. Filter layouts: https://microsoft-search.github.io/pnp-modern-search/usage/search-filters/layouts/ — SP Search replaces PnP filter templates with DevExtreme and Fluent UI equivalents.

### 3.3.3 Active Filter Pill Bar

When any filters are applied, an **Active Filter Pill Bar** renders at the top of the Search Results web part (not the Filters web part) as a horizontal strip of dismissible chips. This gives users immediate visibility into what's filtering their results and one-click removal without scrolling to the filter panel.

**UI Design:**

```
┌────────────────────────────────────────────────────────────────────┐
│ Filters:  [Author: John Doe ✕]  [Modified: Last 30 days ✕]       │
│           [File Type: PDF, DOCX ✕]  [Department: HR ✕]  Clear All │
└────────────────────────────────────────────────────────────────────┘
```

**Pill Rendering Rules:**
- Each pill shows: `{Filter Display Name}: {Human-Readable Value} ✕`
- Multi-value filters within the same group are **combined into one pill** with comma-separated values (e.g., "File Type: PDF, DOCX, XLSX" — not three separate pills)
- Clicking ✕ on a pill dispatches `removeRefiner(filterKey)` to the filterSlice and immediately re-executes the search
- "Clear All" link appears at the end, dispatches `clearAllFilters()` and re-executes
- Pill bar is sticky (stays visible when scrolling results) if the filter panel is in sidebar layout
- Pill bar animates smoothly on add/remove (Fluent UI motion tokens: `fadeIn`, `slideRight` for entry, `fadeOut` for removal)
- When no filters are active, the pill bar is hidden (zero height, not a blank row)
- Pills are color-coded by filter category: subtle background tint per filter group for visual grouping (uses the theme's semantic color tokens)

**Pill Value Formatting (Human-Readable Display):**

The raw refinement token from SharePoint is often unreadable (hex-encoded FQL, taxonomy GUIDs, date ranges). The pill bar uses `IFilterValueFormatter` (see §3.3.5) to display human-readable values:

| Field Type | Raw Token Example | Pill Display |
| --- | --- | --- |
| Date Range | `range(2024-01-01, 2024-12-31)` | "Jan 1, 2024 – Dec 31, 2024" |
| Date Preset | (computed range) | "Last 30 days" / "This year" |
| People/User | `i:0#.f|membership|john@contoso.com` | "John Doe" (resolved via User Profile) |
| Taxonomy | `GP0|#a1b2c3d4-...` | "Marketing > Digital" (resolved via Taxonomy API) |
| Checkbox | `"docx"` | "DOCX" (uppercase, friendly name mapping) |
| Slider Range | `range(1048576, max)` | "> 1 MB" (human-readable file size) |
| Boolean | `"true"` | "Yes" / filter display name |

**Technical Notes:**
- The pill bar is a React component rendered by the Search Results web part, not the Filters web part. It reads from `filterSlice.activeFilters` and writes back via `removeRefiner()`.
- User/People values are resolved on first display and cached in a `Map<string, IPersonaInfo>` in the userSlice (avoids repeated User Profile API calls).
- Taxonomy values are resolved via the PnP Taxonomy API on first display and cached similarly.
- The formatter registry (`FilterValueFormatterRegistry`) is extensible — custom filter types can register their own formatters for pill display.

### 3.3.4 Special Field Handling (Refiner Type Guide)

The #1 pain point in PnP Modern Search is that different SharePoint field types return refinement tokens in completely different formats, and the documentation doesn't explain how to configure, decode, or display them. SP Search provides first-class handling for every special field type.

**A. Date / DateTime Fields**

SharePoint Search returns date refiners as either:
- **Discrete date buckets:** Pre-computed ranges like "Past day", "Past week", "Past month", "Past year" when using `RefinableDate00-19`
- **Raw ISO timestamps:** When the refiner is configured with intervals

**SP Search Date Handling:**
- Date Range filter component renders **preset buttons** (Today, This Week, This Month, This Quarter, This Year, Custom Range) above a DevExtreme DateRangeBox calendar
- Preset buttons generate FQL `range()` tokens relative to the current date/time: e.g., "Last 30 days" → `range(2026-01-06T00:00:00Z, max)`
- Custom range uses the DateRangeBox picker, generates: `range(2026-01-01T00:00:00Z, 2026-01-31T23:59:59Z)`
- **Timezone handling:** All date comparisons use UTC. The UI displays dates in the user's local timezone (derived from the browser) but the FQL token always uses UTC.
- **Refinement token encoding:** Date range tokens for the Search API must use the FQL `range()` function, NOT raw KQL date comparisons. Example: `RefinableDate00:range(datetime("2026-01-01T00:00:00Z"), datetime("2026-12-31T23:59:59Z"))`

**B. People / User Fields**

SharePoint Search returns user refiners as claim-encoded strings (e.g., `i:0#.f|membership|john@contoso.com`). These are meaningless to users.

**SP Search User Field Handling:**
- The People Picker filter uses `@pnp/spfx-controls-react` PeoplePicker for selection (type-ahead against Azure AD)
- When displaying refiner values in checkbox/tagbox mode, the raw claim string is resolved to a display name via PnPjs `sp.profiles.getPropertiesFor()` (batched for performance)
- Resolved names are cached in a `Map<claimString, IPersonaInfo>` that persists for the page session
- The refinement token sent to the Search API keeps the original claim string — only the display is resolved
- **Person card on hover:** When hovering over a user refiner value, a mini Fluent UI PersonaCard shows photo, title, department
- **Multi-value user refiners:** When multiple authors are selected, the filter uses OR logic by default (AND for author makes no sense)

**C. Taxonomy / Managed Metadata Fields**

Taxonomy refiners are the most complex. SharePoint returns them as `GP0|#{GUID}` or `L0|#{GUID}` tokens with no human-readable labels.

**SP Search Taxonomy Handling:**
- Taxonomy filters render as a **DevExtreme TreeView** showing the full term hierarchy (Term Store > Term Group > Term Set > Term path)
- Term labels are resolved via PnP Taxonomy API: `sp.termStore.groups.getById().sets.getById().terms()` — resolved once on filter panel load and cached
- **Hierarchical selection:** Selecting a parent term automatically includes all child terms in the refinement (configurable: inclusive vs. exclusive)
- **GUID-to-Label mapping:** A `TaxonomyLabelCache` service maintains a `Map<termGuid, { label, path }>` populated on first use and refreshed on vertical switch
- **FQL refinement token format:** `owstaxIdRefinableString00:"GP0|#a1b2c3d4-e5f6-..."` — the GP0 prefix is mandatory for taxonomy refinement tokens
- **Search within taxonomy:** The TreeView supports type-ahead filtering within the term hierarchy for large term sets (1,000+ terms)
- **Orphaned terms:** If a term GUID in the refiner response can't be resolved (term was deleted), display it as "(Unknown term)" with the raw GUID in a tooltip, and log a warning

**D. Calculated Columns**

Calculated columns are **NOT refinable or sortable** in SharePoint Search. They appear in search results as text but cannot be used as refiner sources.

**SP Search Handling:**
- The Schema Helper (§3.2.7) marks calculated columns as non-refinable in the managed property picker UI, preventing admins from configuring them as filters
- If an admin manually enters a calculated column managed property as a filter, the `SharePointSearchProvider` detects it returns zero refiner values and surfaces a warning in the filter panel: "This property is not refinable. Calculated columns cannot be used as search filters."
- **Workaround guidance:** The property pane shows a tooltip: "To filter on calculated values, create a dedicated managed property mapped to a crawled property that stores the computed value."

**E. Number / Currency Fields**

Numeric refiners work well but require range-based filtering, not discrete values.

**SP Search Number Handling:**
- Slider/Range filter renders a DevExtreme RangeSlider with configurable min/max/step
- For currency, the slider shows formatted values with locale-appropriate currency symbol (uses `Intl.NumberFormat`)
- FQL token format: `range(decimal(1000), decimal(5000))` for the range 1000–5000
- For file size (common numeric refiner): automatically formats to human-readable units (KB, MB, GB) with the slider showing the raw byte value and the label showing the friendly size

**F. Yes/No (Boolean) Fields**

Boolean refiners return `"0"` and `"1"` string values.

**SP Search Boolean Handling:**
- Renders as a Fluent UI Toggle (not a checkbox — toggle is more intuitive for binary states)
- Maps `"1"` → "Yes" and `"0"` → "No" in the display label (or custom labels if configured: e.g., "Checked Out" / "Available")
- Three-state logic: toggle off = no filter applied (show all), toggle on = filter to "Yes", toggle explicitly set to "No" = filter to "No"

**G. Choice / Lookup Fields**

Choice and lookup fields return discrete string values as refiners. These are the easiest field type.

**SP Search Choice Handling:**
- Checkbox list with counts (default) or TagBox multi-select (configurable)
- Values displayed as-is (no transformation needed)
- "Other" or blank values rendered as "(No value)" with a count
- Sort order configurable: by count (default), alphabetical, or custom order defined in property pane

### 3.3.5 IFilterValueFormatter Interface

Each filter type provides a formatter that converts raw refinement tokens to human-readable pill display text and vice versa.

```typescript
export interface IFilterValueFormatter {
  /** Unique formatter ID, matches the filter type ID */
  id: string;
  /** Convert raw refiner token → human-readable display text for the pill bar */
  formatForDisplay: (rawValue: string, config: IFilterConfig) => string | Promise<string>;
  /** Convert user selection → FQL/KQL refinement token for the Search API */
  formatForQuery: (displayValue: unknown, config: IFilterConfig) => string;
  /** Convert raw refiner token → URL-safe string for deep linking */
  formatForUrl: (rawValue: string) => string;
  /** Restore from URL-safe string → raw refiner token */
  parseFromUrl: (urlValue: string) => string;
}
```

**Built-in Formatters:**
- `DateFilterFormatter` — handles date presets, custom ranges, FQL range() generation, timezone conversion
- `PeopleFilterFormatter` — resolves claim strings to display names, caches resolved profiles
- `TaxonomyFilterFormatter` — resolves GP0|#GUID tokens to term labels with path, caches term hierarchy
- `NumericFilterFormatter` — handles file size formatting (bytes → KB/MB/GB), currency formatting, range display
- `BooleanFilterFormatter` — maps 0/1 to Yes/No or custom labels
- `DefaultFilterFormatter` — pass-through for simple string values (checkbox, choice, lookup)

### 3.3.6 Visual Filter Builder

For power users, the Search Filters web part offers an advanced visual filter builder (toggled via a toolbar button). This uses a DevExtreme FilterBuilder-inspired interface where users can construct complex filter expressions with AND/OR grouping, property selection, operator selection, and value entry. The built expression is converted to KQL refinement queries.

### 3.3.7 Property Pane Configuration


| Property | Type | Default | Description |
| --- | --- | --- | --- |
| filters | IFilterConfig[] | [] | Array of filter definitions |
| filterLayout | enum | "vertical" | "vertical" \| "horizontal" \| "panel" |
| applyMode | enum | "instant" | "instant" \| "manual" |
| operatorBetweenFilters | enum | "AND" | "AND" \| "OR" |
| showCounts | boolean | true | Show result counts per filter value |
| initialDisplayCount | number | 5 | Filter values shown before Show More |
| enableFilterBuilder | boolean | false | Enable advanced visual filter builder |
| showActiveFilterPills | boolean | true | Show active filter pill bar above results |
| pillBarPosition | enum | "above-results" | "above-results" \| "below-toolbar" \| "sticky-top" |


## 3.4 Search Verticals Web Part

The Search Verticals web part provides tab-based navigation to scope search results to specific content types or data sources. Each vertical maps to a Search Results web part instance and can have its own query template, result source, and filter configuration.

### 3.4.1 Core Features

- **Tab Navigation:** Horizontal tab bar with configurable tab labels, icons (Fluent UI icons), and ordering.
- **Badge Counts:** Each tab shows the count of results for the current query. Counts are fetched in parallel when a query is executed.
- **Audience Targeting:** Individual tabs can be targeted to specific Azure AD security groups. Hidden tabs do not appear in the UI.
- **Dynamic Show/Hide:** Vertical tabs with zero results can be automatically hidden or shown dimmed (configurable).
- **Default Vertical:** Configurable default vertical selected on page load.
- **Vertical Scope:** Each vertical defines: query template override, result source override, selected managed properties, default sort, associated filter configuration, and **optional data provider override** (`dataProviderId` — e.g., "People" vertical uses GraphSearchProvider while "Documents" uses SharePointSearchProvider).
- **URL Sync:** Selected vertical is persisted in URL parameters for sharing and bookmarking.
- **Overflow Handling:** On narrow screens, excess tabs collapse into a more dropdown.

### 3.4.2 Property Pane Configuration


| Property | Type | Default | Description |
| --- | --- | --- | --- |
| verticals | IVerticalDefinition[] | [] | Array of vertical definitions (each includes optional dataProviderId, filterConfig[], audienceGroups[]) |
| showCounts | boolean | true | Show result count badges on tabs |
| hideEmptyVerticals | boolean | false | Hide tabs with zero results |
| defaultVerticalKey | string | "all" | Default selected vertical |
| tabStyle | enum | "tabs" | "tabs" \| "pills" \| "underline" |


**PnP Reference:** Study PnP Verticals web part at https://microsoft-search.github.io/pnp-modern-search/usage/search-verticals/ for vertical tab config, per-vertical query template overrides, result source targeting, badge count queries (parallel RowLimit=0 queries), and audience targeting via Azure AD groups. Source: search-parts/src/webparts/searchVerticals/. SP Search uses Zustand store instead of SPFx dynamic data, but reuse the vertical selection pattern and count query approach.

## 3.5 Search Manager Web Part

The Search Manager is the unique 5th web part that provides all user-centric search features. It is available in two modes: as a standalone web part placed on a page, and as a panel triggered from the Search Box web part. Both modes share the same component and state.

### 3.5.1 Saved Searches

- **Save Current Search:** User clicks Save and provides a name. The entire search state is serialized: query text, selected filters, active vertical, sort order, search scope, and the full URL with query parameters.
- **Saved Search List:** Shows all saved searches for the current user. Each entry displays: name, query preview, saved date, result count at save time.
- **Load Saved Search:** Clicking a saved search restores the full search state: query, filters, vertical, sort, scope. URL is updated accordingly.
- **Edit Saved Search:** Rename or update a saved search with the current search state.
- **Delete Saved Search:** Remove a saved search with confirmation dialog.
- **Saved Search Categories:** Users can organize saved searches into folders/categories.

### 3.5.2 Search Sharing

Users can share search configurations and specific result items with colleagues through multiple channels.


**Share via URL**


The complete search state is encoded in URL query parameters. Copy to clipboard generates a full URL that, when visited, restores the exact search state (query, filters, vertical, sort, page).


**Share via Email**


Opens the user's default mail client (mailto:) or composes via SharePoint send mail API. Email body includes: search description, direct link to the search, and optionally a summary of results (top N items with titles and links).


**Share to Specific Users**


Uses the SharedWith people column in the SearchSavedQueries list. The shared search appears in the recipient's Search Manager under a Shared With Me section. Uses PnP PeoplePicker for user selection.


**Share to Teams**


Generates a Teams deep link that opens a Teams chat or channel with the search URL and a preview message. Format: https://teams.microsoft.com/l/chat/0/0?message={encoded message with search URL}.

### 3.5.3 Search Collections (Pinboards)

- **Create Collection:** User creates a named collection (e.g., "Q4 Reports", "Compliance Docs").
- **Pin Results:** From any search result (via quick action or bulk action), users can pin items to one or more collections.
- **View Collection:** Opens a dedicated view showing all pinned items in the collection, with the same layout options as Search Results.
- **Share Collection:** Collections can be shared with specific users or via URL.
- **Manage Collections:** Rename, delete, reorder items within, or merge collections.

### 3.5.4 Search History

- **Automatic Logging:** Every search query is automatically logged with timestamp, query text, filter state, result count, and items clicked.
- **History View:** Chronological list of past searches with option to re-execute any historical search.
- **History Cleanup:** Configurable auto-cleanup (e.g., delete history older than 30 days) and manual clear all.

### 3.5.5 Result Annotations / Tags

- **Personal Tags:** Users can tag any search result with custom labels (e.g., "Reviewed", "Important", "Follow-up").
- **Shared Tags:** Optionally, tags can be shared with team members.
- **Tag-based Filtering:** In Search Manager, filter collections and history by tags.
- **Visual Indicators:** Tagged items show visual badges/labels in search result layouts.

### 3.5.6 Storage: Hidden SharePoint Lists

All Search Manager data is stored in hidden SharePoint lists provisioned during solution deployment. **Search history is stored in a dedicated list** to prevent scale issues — history grows orders of magnitude faster than saved/shared searches, and mixing them causes threshold pressure on the queries that matter.


**SearchSavedQueries List** (Saved + Shared searches only — no history)



| Column | Type | Indexed | Description |
| --- | --- | --- | --- |
| Title | Single line text | Yes | Display name of saved search |
| QueryText | Multiple lines | No | The search query text |
| SearchState | Multiple lines | No | JSON: filters, vertical, sort, scope |
| SearchUrl | Hyperlink | No | Full shareable URL with query params |
| EntryType | Choice | Yes | SavedSearch \| SharedSearch |
| Category | Single line text | Yes | Folder/category for organization |
| SharedWith | Person (multi) | Yes | Users this search is shared with |
| ResultCount | Number | No | Result count at time of save |
| LastUsed | Date/Time | Yes | Last time this search was executed |


**SearchHistory List** (Dedicated list with aggressive retention)

**⚠ CRITICAL — List View Threshold Risk:** In an active enterprise org, this list will exceed 5,000 items (the SharePoint List View Threshold) within weeks. The threshold is enforced **server-side before permissions are applied**. If queries are not filtered on an indexed column, they will fail for *everyone* once the list hits 5,000 items. The `Author` (Created By) built-in column and `SearchTimestamp` **MUST** be indexed at provisioning time. All CAML/REST queries against this list **MUST** filter on `Author eq [Me]` as the **first predicate** — clause ordering matters. If the query filters by `Created > [Today-30]` *before* filtering by Author, the query engine scans the full list before applying the Author filter, and it will throttle even if the user only has 10 history items. Always structure CAML as: `<Eq><FieldRef Name='Author' /><Value Type='Integer'><UserID/></Value></Eq>` as the outermost/first clause, with date/hash filters nested inside.

| Column | Type | Indexed | Description |
| --- | --- | --- | --- |
| Title | Single line text | No | Query text (truncated to 255 chars) |
| QueryHash | Single line text | Yes | SHA-256 hash of full query state for deduplication |
| Vertical | Single line text | Yes | Active vertical at time of search |
| Scope | Single line text | No | Search scope used |
| ResultCount | Number | No | Number of results returned |
| ClickedItems | Multiple lines | No | JSON array: [{url, title, position, timestamp}] — tracks which results user clicked |
| SearchTimestamp | Date/Time | Yes | When the search was executed (indexed for retention cleanup) |
| Author (Created By) | Person | **Yes** | Built-in column — **MUST be indexed at provisioning**. Primary filter for all queries. |

**SearchHistory Retention Policy:**
- Configurable TTL: 30 / 60 / 90 days (default 90), set in SearchConfiguration
- Cleanup via scheduled PnP PowerShell script or Azure Function: `DELETE WHERE SearchTimestamp < (Today - TTL)`
- Minimal columns keep item size small for high-volume writes
- User isolation: list permissions set to Add Items + Edit Own Items only (no Read All)
- **Provisioning script MUST:** (1) index Author, SearchTimestamp, QueryHash, and Vertical columns immediately, (2) verify index creation succeeded before marking provisioning complete


**SearchCollections List**



| Column | Type | Indexed | Description |
| --- | --- | --- | --- |
| Title | Single line text | Yes | Collection name |
| ItemUrl | Hyperlink | Yes | URL of the pinned item |
| ItemTitle | Single line text | No | Display title of the pinned item |
| ItemMetadata | Multiple lines | No | JSON: cached metadata snapshot |
| CollectionName | Single line text | Yes | Grouping key for the collection |
| Tags | Multiple lines | No | JSON array of user tags/annotations |
| SharedWith | Person (multi) | Yes | Users this collection is shared with |
| SortOrder | Number | No | Item order within collection |


**SearchConfiguration List**



| Column | Type | Indexed | Description |
| --- | --- | --- | --- |
| Title | Single line text | Yes | Configuration key name |
| ConfigType | Choice | Yes | Scope \| VerticalPreset \| LayoutMapping \| ManagedPropertyMap \| PromotedResult \| StateSnapshot |
| ConfigData | Multiple lines | No | JSON payload for the configuration |
| IsActive | Yes/No | Yes | Whether this config is currently active |
| SortOrder | Number | No | Display order for scopes/verticals |
| ExpiresAt | Date/Time | Yes | TTL for StateSnapshot entries (used by sid= deep links). Null = no expiry. |
| AudienceGroups | Multiple lines | No | JSON array of Azure AD security group IDs for audience targeting (PromotedResult entries) |


## 3.6 Promoted Results / Best Bets

Promoted Results transform SP Search from "just a query box" to a **curated, intent-aware search experience**. Admins define rules that surface specific documents or URLs above organic results when queries match defined patterns. This is the feature that makes stakeholders and execs say "this is better than what we had."

**Design Decision — Client-Side Injection vs. Server-Side Query Rules:** SP Search implements promoted results as a **client-side visual injection** (evaluated in the Zustand store after results load) rather than using SharePoint's server-side Query Rules. This is a deliberate architectural choice: server-side Query Rules affect ranking scores of *all* results and are unpredictable to debug, while client-side injection is deterministic — promoted items appear in a distinct "Recommended" block at position #0, and organic results remain untouched. The tradeoff is that we cannot "boost" a result to position #3 in the organic list, but for enterprise UX, a predictable dedicated block is cleaner than invisible ranking manipulation.

### 3.6.1 Promoted Result Rules

Rules are stored as `ConfigType: PromotedResult` entries in the SearchConfiguration list (admin-only write). Each rule defines:

- **Match Condition:** How the rule triggers:
  - `contains` — query contains keyword(s)
  - `equals` — exact query match
  - `regex` — regular expression pattern
  - `kql` — KQL predicate match (for structured queries)
- **Promoted Items:** Array of URLs/documents to promote, each with:
  - `url` — target document or page URL
  - `title` — display title override (optional, falls back to search result title)
  - `description` — custom description (optional)
  - `imageUrl` — custom thumbnail (optional)
  - `position` — rank order within promoted results block
- **Audience Targeting:** Optional Azure AD security group IDs (same pattern as vertical audience targeting). When set, only users in those groups see the promoted result.
- **Schedule:** Optional start/end dates for time-limited promotions (e.g., "Open Enrollment" rules active only during enrollment period).
- **Vertical Scope:** Optional — restrict the promoted result to specific verticals only.

### 3.6.2 UI Rendering

- **"Recommended" Block:** Displayed above organic search results in every layout. Visually distinct (subtle background, "Recommended" badge) so users understand these are curated.
- **Layout-Adaptive:** The recommended block adapts to the active layout — card style in Card Layout, row style in DataGrid/List/Compact, etc.
- **Dismissible:** Users can dismiss promoted results for the session (stored in uiSlice, not persisted).
- **Maximum Display:** Configurable max promoted results per query (default 3) to avoid overwhelming organic results.

### 3.6.3 Admin Configuration

Promoted result rules are managed through the SearchConfiguration list. In Phase 5+, an admin UI (either a dedicated admin page or a section within the Search Manager property pane) provides a CRUD interface for rules without requiring direct list access.

### 3.6.4 Property Pane Configuration

| Property | Type | Default | Description |
| --- | --- | --- | --- |
| enablePromotedResults | boolean | true | Show promoted results block above organic results |
| maxPromotedResults | number | 3 | Maximum promoted items to display per query |
| promotedResultStyle | enum | "card" | "card" \| "banner" \| "inline" — visual treatment |


# 4. Shared Architecture


## 4.1 Zustand Store Design

The Zustand store is the backbone of inter-webpart communication. It is distributed via an SPFx Library Component (sp-search-store) that provides a **store registry** supporting multiple isolated search contexts on the same page.

**Store Registry Pattern:**
- `getStore(searchContextId: string): SearchStore` — returns existing store or creates new one
- `disposeStore(searchContextId: string): void` — cleanup on web part disposal
- Each web part reads `searchContextId` from its property pane and calls `getStore(searchContextId)` during initialization
- Web parts sharing the same `searchContextId` share a single store instance; different IDs create fully isolated stores
- Default `searchContextId`: auto-generated GUID (ensures isolation by default, admin configures shared ID to connect web parts)

### 4.1.1 Store Slices


| Slice | State | Purpose |
| --- | --- | --- |
| querySlice | queryText, queryTemplate, scope, suggestions, isSearching | Search query state and execution control |
| filterSlice | activeFilters, availableRefiners, filterConfig, filterBuilderState | Filter selections and refiner data |
| verticalSlice | activeVertical, verticals, verticalCounts | Vertical tab state and result counts |
| resultSlice | results, totalCount, currentPage, sort, selectedItems, isLoading | Search results data and pagination |
| uiSlice | activeLayout, detailPanelItem, isManagerOpen, bulkActionMode | UI presentation state |
| userSlice | savedSearches, collections, recentSearches, tags, history | User-specific data from hidden lists |


### 4.1.2 URL Synchronization

A Zustand middleware layer handles bi-directional synchronization between the store and URL query parameters. Every state change that affects the search context is reflected in the URL, and page load from a URL with search parameters restores the full state.

**Dual-Mode Deep Linking:**

The URL sync middleware supports two modes to handle varying state complexity:

1. **Short URL Mode (default):** Current state serialized into compact query parameters. Used when total URL length stays under 2,000 characters.
2. **StateId Mode (automatic fallback):** When URL would exceed the limit (complex filter builder expressions, taxonomy paths, multiple refiners), the middleware automatically saves the full state JSON to a hidden list item (SearchConfiguration) and replaces the URL with `?sid=<itemId>`. On page load, `sid` is detected and the state is restored from the list. This also enables expiring links (TTL column) and auditable sharing.

**State Schema Versioning:** All serialized state includes `sv=1` (state version). This allows safe URL migration between releases — older URLs are handled by version-specific deserializers.

**Multi-Context URL Namespacing:** When multiple search contexts exist on a page, parameters are namespaced by context: `?ctx1.q=budget&ctx1.f=...&ctx2.q=john`. Single-context pages use short params for clean URLs.

**URL Parameter Mapping:**

| Parameter | Store Property | Example |
| --- | --- | --- |
| q | querySlice.queryText | ?q=annual+report |
| f | filterSlice.activeFilters | &f=FileType:docx,pptx\|Author:JohnDoe |
| v | verticalSlice.activeVertical | &v=documents |
| s | resultSlice.sort | &s=LastModifiedTime:desc |
| p | resultSlice.currentPage | &p=3 |
| sc | querySlice.scope | &sc=currentsite |
| l | uiSlice.activeLayout | &l=grid |
| sv | (state version) | &sv=1 |
| sid | (state ID fallback) | ?sid=42 |


## 4.2 SPFx Library Component

The sp-search-store library component serves five purposes: it provides the store registry for multi-instance isolation, it exposes shared TypeScript interfaces and types, it contains shared utility functions for search query construction and URL encoding/decoding, it hosts all provider registries for extensibility, and it ships the built-in data providers (SharePointSearchProvider, GraphSearchProvider).

- Package name: sp-search-store
- Exports: `getStore()`, `disposeStore()`, all TypeScript interfaces, utility functions, provider registries (DataProviderRegistry, SuggestionProviderRegistry, ActionProviderRegistry, LayoutRegistry, FilterTypeRegistry), built-in providers
**Architecture Note:** PnP v4 uses SPFx Dynamic Data for inter-web-part communication (search-parts/src/webparts/*/dynamicData). SP Search replaces this with a shared Zustand store via SPFx Library Component for more predictable state flow. Study PnP v4 deep linking for filters (URL param f) and extend to all search state slices: https://microsoft-search.github.io/pnp-modern-search/usage/search-filters/ (Deep Linking section).
- Version-locked with the web part package for compatibility
- Deployed as part of the same .sppkg solution package

**API Permissions (webApiPermissionRequests in package-solution.json):**
- `Sites.Read.All` (Graph) — required for GraphSearchProvider to search files and sites
- `ExternalItem.Read.All` (Graph) — required for Graph connector content
- `People.Read` (Graph) — required for Graph people search
- These permissions are only requested when GraphSearchProvider is enabled. SharePointSearchProvider requires no additional permissions.

**Development Note — Library Component Hot-Reload Quirk:** SPFx Library Components do **not** support hot module replacement during `gulp serve`. Changes to sp-search-store code require a full page refresh in the workbench, not just a module replacement. **Mitigation:** Structure store actions, data providers, and all business logic as pure TypeScript classes/functions that are fully testable via Jest outside the SPFx workbench. Only the thin SPFx integration layer (the LibraryComponent class itself) should require workbench testing. This saves hundreds of hours of development time.

## 4.3 Performance Strategy


### 4.3.1 Bundle Optimization

- All spfx-toolkit imports use direct path imports (e.g., import { Card } from 'spfx-toolkit/lib/components/Card')
- DevExtreme components are lazy-loaded: DataGrid only loads when grid layout is selected
- Result detail panel loaded on first open, not on page load
- Search Manager panel content lazy-loaded on toggle
- Code splitting per layout: each layout is a separate chunk

### 4.3.2 Data Optimization

- Debounced search execution prevents excessive API calls
- Vertical counts fetched in parallel with main results query
- Filter refiner data cached and invalidated on query change
- Search history and saved searches loaded on demand, not on page load
- Virtualized rendering for large result sets (DevExtreme virtual scrolling, react-window for Card/List layouts)

### 4.3.3 Request Lifecycle Management

These three behaviors are critical for a professional search UX and must be implemented in Phase 1:

**Abort In-Flight Searches:** Every search execution creates an `AbortController`. When a new query fires (keystroke, filter change, vertical switch), the previous controller is aborted before the new request starts. This prevents race conditions where slow old results overwrite fast new results. The PnPjs `sp.search()` call must receive the abort signal via `ISearchQuery.signal`. Vertical count queries are also abortable — all parallel count requests share a single AbortController that cancels when the next query cycle begins.

**Request Coalescing:** When Search Results, vertical counts, and suggestion queries fire simultaneously (common on initial page load or rapid typing), the SearchService should resolve shared work once:
- Token resolution (replace `{searchTerms}`, `{Site.ID}`, `{User.Name}`, etc.) happens once per query cycle and is cached for the duration of that cycle.
- KQL query construction (base template + filters + sort + scope) is computed once and shared across the main results query and all vertical count queries.
- This avoids redundant token resolution and query assembly across 5-10+ parallel API calls.

**Refiner Stability Mode:** Optional setting (property pane toggle, default: on). When enabled, the filter panel retains the previous refiner values for a configurable window (default 500ms) while new refiners load. This prevents the jarring UI flicker where filter options disappear and reappear during rapid typing with instant-apply mode. Implementation: the filterSlice maintains both `currentRefiners` (latest from API) and `displayRefiners` (what the UI shows), with a debounced transition between them.

### 4.3.4 Caching Strategy

- Session storage for search history and recent searches (per tab lifecycle)
- In-memory cache for refiner data with TTL
- Managed property metadata cached on first load (rarely changes)
- Thumbnail URLs cached with result data

## 4.4 Extensibility: Provider & Registry Model

To avoid the "PnP extensibility gotchas" (error-prone customization, limited extension points, breaking changes on update), SP Search formalizes a **provider/registry pattern** that keeps the core stable while letting teams extend safely. All registries are hosted in the sp-search-store library component and are available to all web parts in a search context.

### 4.4.1 DataProvider Registry (Most Critical)

The DataProvider Registry is the most important extensibility point — it abstracts the actual search execution away from the UI layer. The Search Results web part never calls PnPjs or Graph directly; it always delegates to the registered `ISearchDataProvider`.

**ISearchDataProvider Interface:**
- `id: string` — unique provider identifier (e.g., "sharepoint", "graph", "custom-elastic")
- `displayName: string` — admin-facing label in property pane
- `execute(query: ISearchQuery, signal: AbortSignal): Promise<ISearchResponse>` — execute a search and return normalized results
- `getSuggestions?(query: string, signal: AbortSignal): Promise<ISuggestion[]>` — optional: provider-specific suggestions
- `getSchema?(): Promise<IManagedProperty[]>` — optional: fetch available properties for Schema Helper
- `supportsRefiners: boolean` — whether the provider returns refiner data
- `supportsCollapsing: boolean` — whether the provider supports CollapseSpecification
- `supportsSorting: boolean` — whether the provider supports server-side sorting

**ISearchQuery (normalized input):**
- `queryText: string`, `queryTemplate: string`, `scope: ISearchScope`
- `filters: IActiveFilter[]`, `sort: ISortField | null`
- `page: number`, `pageSize: number`
- `selectedProperties: string[]`, `refiners: string[]`
- `collapseSpecification?: string`
- `resultSourceId?: string` (SharePoint-specific, ignored by Graph)
- `trimDuplicates?: boolean`

**ISearchResponse (normalized output):**
- `items: ISearchResult[]`, `totalCount: number`
- `refiners: IRefiner[]`
- `promotedResults: IPromotedResult[]` (server-side promoted results, if any)
- `querySuggestion?: string` (did-you-mean)

**Built-in Providers:**
- `SharePointSearchProvider` — PnPjs `sp.search()`. Default for all verticals. Full refiner/collapsing/sorting support.
- `GraphSearchProvider` — MS Graph `/search/query`. Better for people, Teams messages, external connectors, acronyms. Limited refiner support (Graph returns aggregations differently).

**Per-Vertical Override:** Each `IVerticalDefinition` includes an optional `dataProviderId: string`. When set, the Search Results web part uses that provider instead of the default for the active vertical. This enables hybrid search pages where "Documents" uses SharePoint Search and "People" uses Graph on the same page.

**Custom Provider Example:** An org could register an `ElasticsearchProvider` for a custom index, or a `ServiceNowProvider` that queries their ITSM platform and maps results to `ISearchResult`.

### 4.4.2 SuggestionProvider Registry

The Search Box uses an array of `ISuggestionProvider` implementations to populate the suggestions dropdown. Each provider is queried in parallel and results are merged and ranked.

**ISuggestionProvider Interface:**
- `id: string` — unique provider identifier
- `displayName: string` — section label in dropdown (e.g., "Recent", "Trending", "Properties")
- `priority: number` — sort order in dropdown (lower = higher)
- `maxResults: number` — max suggestions from this provider
- `getSuggestions(query: string, context: ISearchContext): Promise<ISuggestion[]>` — async search
- `isEnabled(context: ISearchContext): boolean` — conditional activation

**Built-in Providers:**
- `RecentSearchProvider` — user's recent searches from SearchHistory list
- `TrendingQueryProvider` — popular queries across org (aggregated from SearchHistory)
- `ManagedPropertyProvider` — suggests managed property values matching input (e.g., typing "john" suggests "Author: John Doe")

**Custom Provider Example:** An org could register a `JiraTicketProvider` that suggests Jira tickets matching the query, or a `PeopleProvider` that queries the User Profile Service directly.

### 4.4.3 ActionProvider Registry

Result actions (share, pin, preview, open, etc.) are registered via `IActionProvider` implementations. This allows orgs to add custom actions without modifying the core web part code.

**IActionProvider Interface:**
- `id: string` — unique action identifier
- `label: string` — display label
- `iconName: string` — Fluent UI icon name
- `position: 'toolbar' | 'contextMenu' | 'both'` — where the action appears
- `isApplicable(item: ISearchResult): boolean` — conditional visibility (e.g., only show "Open in AutoCAD" for .dwg files)
- `execute(items: ISearchResult[], context: ISearchContext): Promise<void>` — action handler
- `isBulkEnabled: boolean` — whether action supports multi-select

**Built-in Providers:**
- `OpenAction` — open document in browser or desktop app
- `PreviewAction` — open in Result Detail Panel
- `ShareAction` — share via URL, email, Teams, or to specific users
- `PinAction` — pin to a collection
- `CopyLinkAction` — copy document URL to clipboard
- `DownloadAction` — download file

**Custom Provider Example:** `ApproveAction` for document approval workflows, `SendToArchiveAction` for records management, `OpenInLineOfBusinessApp` for ERP/CRM links.

### 4.4.4 LayoutRegistry

Result layouts are registered via `ILayoutDefinition` and rendered by the Search Results web part. This enables teams to ship custom layouts without forking the core.

**ILayoutDefinition Interface:**
- `id: string` — unique layout identifier (e.g., "datagrid", "cards", "custom-timeline")
- `displayName: string` — label in layout switcher
- `iconName: string` — Fluent UI icon for layout switcher
- `component: React.LazyExoticComponent` — lazy-loaded React component
- `supportsPaging: 'numbered' | 'infinite' | 'both'` — pagination model
- `supportsBulkSelect: boolean` — whether layout supports multi-select
- `supportsVirtualization: boolean` — whether layout supports virtual scrolling
- `defaultSortable: boolean` — whether layout supports client-side sorting

**Built-in Layouts:** DataGrid, Card, List, Compact, People, Document Gallery (6 total as specified in Section 3.2.2)

**Custom Layout Example:** A "Timeline Layout" that renders results on a chronological timeline, or a "Map Layout" that plots geotagged documents.

### 4.4.5 FilterTypeRegistry

Filter types are registered via `IFilterTypeDefinition`. This allows teams to add domain-specific filter controls beyond the built-in set.

**IFilterTypeDefinition Interface:**
- `id: string` — unique filter type identifier (e.g., "checkbox", "daterange", "custom-status")
- `displayName: string` — label in property pane filter type dropdown
- `component: React.LazyExoticComponent` — lazy-loaded filter UI component
- `serializeValue(value: unknown): string` — convert filter value to URL-safe string for deep linking
- `deserializeValue(raw: string): unknown` — restore filter value from URL param
- `buildRefinementToken(value: unknown, managedProperty: string): string` — convert to KQL/FQL refinement token

**Built-in Types:** Checkbox, Date Range, Slider, People Picker, Taxonomy Tree, TagBox, Toggle (as specified in Section 3.3.1)

**Custom Type Example:** A "Project Status" filter with org-specific statuses, or a "Cost Center" filter backed by a custom data source.

### 4.4.6 Registration Pattern

All providers are registered during web part initialization in `onInit()`. The library component exposes typed registry objects:

```
// In a custom extension or web part onInit():
import { getStore } from 'sp-search-store';

const store = getStore(this.properties.searchContextId);
store.registries.dataProviders.register(new ServiceNowProvider());
store.registries.suggestions.register(new JiraTicketProvider());
store.registries.actions.register(new ApproveAction());
store.registries.layouts.register(myTimelineLayout);
store.registries.filterTypes.register(myStatusFilter);
```

**Registration Rules:**
- Duplicate IDs throw a warning and the first registration wins (no silent overwrite)
- Built-in providers are registered first; custom providers can override by using the same ID with a `force: true` flag
- All registries are per-store-instance (scoped to `searchContextId`), so different search contexts on the same page can have different provider sets
- Registries are frozen after the first search execution to prevent mid-session mutations


# 5. Data Flow


## 5.1 Search Execution Flow

The following describes the end-to-end data flow when a user performs a search:
- User enters query in Search Box or loads page with URL parameters.
- Search Box dispatches setQuery action to Zustand querySlice.
- URL sync middleware updates URL query parameters.
- **Any in-flight search AbortController is aborted** before the new cycle begins.
- Search Results web part (subscribed to querySlice, filterSlice, verticalSlice) detects state change.
- Search Results constructs KQL query: merge query template + query text + active filters + vertical scope + sort. **Token resolution and query construction are computed once** and shared across results + count queries (request coalescing).
- **Promoted Results rules are evaluated** against the query. Matching promoted items are fetched from SearchConfiguration.
- PnPjs sp.search() is called with constructed query, selected properties, paging parameters, and **AbortController signal**.
- Results and refiners are dispatched to resultSlice and filterSlice respectively. If refiner stability mode is on, **displayRefiners update after debounce window**.
- In parallel, vertical count queries are dispatched for each vertical tab (using rowLimit=0, selectProperties=[], **sharing the same AbortController**).
- Search Filters web part re-renders with new refiner data from filterSlice.
- Search Verticals web part re-renders with new counts from verticalSlice.
- Search history entry is logged to userSlice and persisted to **SearchHistory list** asynchronously.

## 5.2 Filter Interaction Flow

- User selects/deselects a filter value in Search Filters web part.
- Filter web part dispatches setFilter action to filterSlice with the updated refinement.
- If applyMode is "instant": Search Results detects filterSlice change and re-executes query.
- If applyMode is "manual": Filter changes are staged. User clicks Apply to dispatch applyFilters, triggering re-execution.
- URL parameters updated to reflect new filter state.

## 5.3 Save Search Flow

- User clicks Save in Search Manager (panel or web part).
- Dialog prompts for search name and optional category.
- Current Zustand store state is serialized into SearchState JSON.
- Current URL (with all query parameters) is captured as SearchUrl.
- New item is created in SearchSavedQueries list with EntryType = "SavedSearch".
- userSlice is updated and UI reflects the new saved search.


# 6. spfx-toolkit Integration Map

The following maps spfx-toolkit components, hooks, and utilities to their usage within SP Search web parts.

## 6.1 Components


| Component | Used In | Usage |
| --- | --- | --- |
| Card + Header + Content | Search Results | Card layout for results. Accordion for grouped results. Maximize for detail expand. |
| Card (Accordion) | Search Filters | Each filter group in a collapsible card with persistence. |
| VersionHistory | Result Detail Panel | Full version history in detail panel with download and comparison. |
| DocumentLink | All layouts, Detail Panel | Smart document link rendering with file type icon awareness. |
| UserPersona | People Layout, Detail Panel | User profile card with photo, presence, contact actions. |
| ErrorBoundary | All web parts | Wraps each web part root for graceful error handling. |
| Toast / ToastProvider | All web parts | Success/error notifications for save, share, export, delete operations. |
| FormContainer / FormItem | Detail Panel, Search Manager | Structured metadata display and search configuration forms. |
| WorkflowStepper | Detail Panel (optional) | Show document workflow status if workflow metadata available. |


## 6.2 Hooks


| Hook | Used In | Usage |
| --- | --- | --- |
| useLocalStorage | Search Box, Filters | Persist UI preferences (last layout, collapsed filters) |
| useViewport | All layouts | Responsive layout switching based on viewport size |
| useCardController | Card Layout, Filters | Programmatic card expand/collapse and scroll-to |
| useErrorHandler | All web parts | Centralized error handling across search operations |


## 6.3 Utilities


| Utility | Used In | Usage |
| --- | --- | --- |
| SPContext | All web parts | SharePoint context initialization for PnPjs |
| BatchBuilder | Search Manager | Batch operations for saving/updating multiple list items |
| createPermissionHelper | Search Manager | Check user permissions on hidden lists |
| createSPExtractor | Search Manager | Extract and transform list item data for saved searches |


# 7. Deployment & Provisioning

**PnP Reference:** Study PnP installation at https://microsoft-search.github.io/pnp-modern-search/installation/ for packaging patterns, app catalog deployment (tenant vs site-level), and API permission management. SP Search uses site-level app catalog via CI/CD.

## 7.1 Solution Package Structure

The solution is deployed as a single .sppkg file containing all five web parts and the library component. The package is deployed to a site-level app catalog via CI/CD pipeline.
- sp-search.sppkg containing:
- SPSearchBoxWebPart
**Repository:** https://github.com/hmane/spfx-toolkit — Always use direct path imports to avoid pulling DevExtreme into the bundle (~500KB+). See README.md for import patterns, SPFX-Toolkit-Usage-Guide.md for component APIs, and CLAUDE.md for AI development instructions.
- SPSearchResultsWebPart
- SPSearchFiltersWebPart
- SPSearchVerticalsWebPart
- SPSearchManagerWebPart
- SPSearchStoreLibrary (SPFx Library Component)

## 7.2 Hidden List Provisioning

Hidden lists are provisioned via PnP PowerShell script executed as a post-deployment step in the CI/CD pipeline.
- Script: Provision-SPSearchLists.ps1
- Creates four hidden lists: SearchSavedQueries, SearchCollections, SearchHistory, SearchConfiguration
- Sets list property: Hidden = true
- **Per-list permission model (see Section 8.2 for details):**
  - SearchSavedQueries & SearchCollections: All authenticated users have Add Items. Item-level permissions enforced — author gets full control, shared recipients get Read.
  - SearchHistory: All authenticated users have Add Items + Edit Own Items. No cross-user visibility.
  - SearchConfiguration: Site Collection Admins only (or dedicated SP Search Admins security group). Regular users have Read access.
- Creates required columns, content types, and indexed columns
- Seeds default configuration entries (search scopes, default layout mappings, promoted result rules)
- Idempotent: safe to run multiple times (checks existence before creating)

## 7.3 Pipeline Steps

- Build solution: gulp bundle --ship && gulp package-solution --ship
- Deploy .sppkg to site-level app catalog
- Run PnP PowerShell provisioning script for hidden lists
- Verify deployment: check web part availability and list creation

# 8. Security & Permissions


## 8.1 API Permissions

SP Search uses the SharePoint Search API through PnPjs, which operates under the current user's permissions. No additional Azure AD API permissions are required beyond standard SharePoint access.

## 8.2 Data Access & Item-Level Security

Search results are security-trimmed by SharePoint (users only see results they have access to). Hidden list data is protected using **item-level permissions** — not just query-level filtering — to ensure data isolation even if users access lists via REST API directly.

**SearchSavedQueries & SearchCollections — Item-Level Permissions:**
- On item creation, the provisioning service breaks permission inheritance on the new item.
- **Author-owned items (EntryType: SavedSearch, RecentSearch):** Only the author has Full Control. No other users can read.
- **Shared items (EntryType: SharedSearch, shared collections):** Author gets Full Control. Each user in `SharedWith` column gets Read permission. A SharePoint event receiver or PnP webhook triggers permission updates when `SharedWith` changes.
- **Implementation:** The SearchManagerService handles `breakRoleInheritance()` and `addRoleAssignment()` calls via PnPjs after item creation. This adds ~200ms latency per share operation but provides real security that passes audit.

**SearchHistory — User Isolation:**
- All authenticated users have Add Items + Edit Own Items (no Read All).
- CAML queries filter by `Author eq [Me]` as a convenience layer, but the underlying list permissions prevent cross-user enumeration even via REST.
- Retention policy: items older than configurable TTL (default 90 days) are cleaned up by a scheduled PnP PowerShell script or Azure Function.

**SearchConfiguration — Admin-Only Write:**
- Regular users have Read access only (needed to load scopes, vertical presets, promoted results).
- Write access restricted to Site Collection Admins or a dedicated "SP Search Admins" security group.
- This prevents users from modifying search scopes, promoted results, or layout mappings.

**General:**
- No data leaves the SharePoint tenant. All processing is client-side within the browser.
- JSON payloads in list columns are validated against TypeScript interfaces before save.

## 8.3 Content Security

- All user inputs (query text, tags, collection names) are sanitized before storage.
- JSON payloads in list columns are validated against TypeScript interfaces before save.
- No script injection possible through search results (content rendered via React, not innerHTML).
- Embedded Office Online previews use SharePoint's native preview infrastructure.


# 9. Implementation Phases


## Phase 1: Foundation

**Priority:** High  |  Focus: Core search functionality + architecture
- SPFx Library Component with Zustand **store registry** (`getStore(searchContextId)`) and URL sync middleware (including `sv=1` state versioning)
- **Provider/Registry scaffolding:** ISuggestionProvider, IActionProvider, ILayoutDefinition, IFilterTypeDefinition interfaces + registry classes (empty custom slots, built-in providers registered)
- Search Box web part with basic query input, debounce, scope selector, and `searchContextId` property pane field
- Search Results web part with List Layout and Compact Layout (registered via LayoutRegistry)
- Search Filters web part with Checkbox and Date Range filter types (registered via FilterTypeRegistry)
- Search Verticals web part with tab navigation and badge counts
- Hidden list provisioning PowerShell script (4 lists: SearchSavedQueries, SearchCollections, SearchHistory, SearchConfiguration) with **item-level permission scaffolding**
- PnPjs search service layer with query construction, **AbortController integration**, and **request coalescing** (shared token resolution + query construction)
- SearchHistory list with aggressive retention policy

## Phase 2: Rich Layouts

**Priority:** High  |  Focus: Advanced result presentation
- DataGrid Layout (DevExtreme) with sort, filter, group, export (registered via LayoutRegistry)
- Card Layout (spfx-toolkit) with accordion and maximize (registered via LayoutRegistry)
- People Layout with UserPersona integration (registered via LayoutRegistry)
- Document Gallery Layout with thumbnails (registered via LayoutRegistry)
- Result Detail Panel with preview, metadata, version history, actions
- Layout switcher UI in Search Results toolbar
- **Refiner stability mode** (debounced displayRefiners transition)

## Phase 3: User Features

**Priority:** Medium  |  Focus: Search Manager, sharing, and security
- Search Manager web part (standalone and panel mode)
- Saved searches: save, load, edit, delete, categorize
- Search sharing: URL, email, Teams, user-specific
- Search collections/pinboards: create, pin, manage, share
- Search history logging and view (dedicated SearchHistory list)
- Recent searches in Search Box suggestions (RecentSearchProvider)
- **Item-level permission enforcement:** `breakRoleInheritance()` + `addRoleAssignment()` on save/share operations
- **StateId deep link fallback:** automatic `?sid=` mode when URL exceeds limit
- **Promoted Results / Best Bets:** admin-defined rules with "Recommended" block above results

## Phase 4: Power Features

**Priority:** Lower  |  Focus: Advanced query, annotation, and extensibility
- Visual Query Builder in Search Box
- Advanced filter types: Taxonomy Tree, People Picker, Slider, TagBox (registered via FilterTypeRegistry)
- Visual Filter Builder in Search Filters
- Result annotations / tags (personal and shared)
- Bulk actions toolbar: share, download, compare, export (all via ActionProvider registry)
- Smart suggestions: TrendingQueryProvider + ManagedPropertyProvider
- Audience targeting for verticals and promoted results
- **Custom provider documentation:** developer guide for registering custom SuggestionProvider, ActionProvider, Layout, and FilterType

## Phase 5: Polish & Optimization

**Priority:** Lower  |  Focus: Performance and UX refinement
- Bundle size optimization and lazy loading verification
- Accessibility audit (WCAG 2.1 AA)
- Responsive design testing across devices
- Performance profiling and optimization
- Search Analytics dashboard (admin usage insights)
- Comprehensive error handling and empty states
- Documentation and admin guide


# 10. Appendix


## 10.1 Key TypeScript Interfaces

The following are the primary interfaces shared across all web parts via the library component. These are defined in the sp-search-store package. Interfaces include both **state properties** and **action methods** that mutate the store. This is the "shape" of the application — copy-paste-ready for the sp-search-store package.

### ISearchStore (Root Store)

```typescript
export interface ISearchStore {
  query: IQuerySlice;
  filters: IFilterSlice;
  verticals: IVerticalSlice;
  results: IResultSlice;
  ui: IUISlice;
  user: IUserSlice;
  registries: IRegistryContainer;
  reset: () => void;
  dispose: () => void;
}
```

### IQuerySlice

```typescript
export interface IQuerySlice {
  queryText: string;
  queryTemplate: string;       // e.g. "{searchTerms} Path:{Site.URL}"
  scope: ISearchScope;
  suggestions: ISuggestion[];
  isSearching: boolean;
  abortController: AbortController | null;
  // Actions
  setQueryText: (text: string) => void;
  setScope: (scope: ISearchScope) => void;
  setSuggestions: (suggestions: ISuggestion[]) => void;
  cancelSearch: () => void;    // Aborts current controller
}

export interface ISearchScope {
  id: string;
  label: string;
  kqlPath?: string;            // e.g. "Path:https://contoso.sharepoint.com/sites/hr"
  resultSourceId?: string;
}

export interface ISuggestion {
  displayText: string;
  groupName: string;           // "Recent", "People", "Files"
  iconName?: string;
  action?: () => void;
}
```

### IFilterSlice

```typescript
export interface IFilterSlice {
  activeFilters: IActiveFilter[];
  availableRefiners: IRefiner[];
  displayRefiners: IRefiner[];    // Refiner stability mode: debounced version for UI
  filterConfig: IFilterConfig[];
  isRefining: boolean;
  // Actions
  setRefiner: (filter: IActiveFilter) => void;
  removeRefiner: (filterKey: string, value?: string) => void;
  clearAllFilters: () => void;
  setAvailableRefiners: (refiners: IRefiner[]) => void;
}

export interface IActiveFilter {
  filterName: string;            // Managed Property Name
  value: string;
  operator: 'AND' | 'OR';
}

export interface IRefiner {
  filterName: string;
  values: IRefinerValue[];
}

export interface IRefinerValue {
  name: string;
  value: string;                 // Token for query (e.g., encoded hex)
  count: number;
  isSelected: boolean;
}
```

### IResultSlice

```typescript
export interface IResultSlice {
  items: ISearchResult[];
  totalCount: number;
  currentPage: number;
  pageSize: number;
  sort: ISortField | null;
  promotedResults: IPromotedResult[];
  isLoading: boolean;
  error: string | null;
  // Actions
  setResults: (items: ISearchResult[], total: number) => void;
  setPage: (page: number) => void;
  setSort: (sort: ISortField) => void;
}

export interface ISearchResult {
  key: string;                   // Unique key (WorkId or DocId)
  title: string;
  url: string;
  summary: string;               // HitHighlightedSummary
  author: IPersonaInfo;          // Structured author (not just a string)
  created: string;               // ISO Date
  modified: string;              // ISO Date
  fileType: string;
  fileSize: number;
  siteName: string;
  siteUrl: string;
  thumbnailUrl: string;
  properties: Record<string, unknown>;  // Dynamic managed property bag
  // Collapsing / Threading
  isCollapsedGroup?: boolean;
  childResults?: ISearchResult[];
  groupCount?: number;
}

export interface IPersonaInfo {
  displayText: string;
  email: string;
  imageUrl?: string;
}

export interface ISortField {
  property: string;
  direction: 'Ascending' | 'Descending';
}
```

### IVerticalSlice

```typescript
export interface IVerticalSlice {
  currentVerticalKey: string;
  verticals: IVerticalDefinition[];
  verticalCounts: Record<string, number>;  // { "All": 105, "Files": 40 }
  // Actions
  setVertical: (key: string) => void;
  setVerticalCounts: (counts: Record<string, number>) => void;
}

export interface IVerticalDefinition {
  key: string;
  label: string;
  iconName?: string;
  queryTemplate?: string;        // Overrides global template
  resultSourceId?: string;
  dataProviderId?: string;       // Per-vertical data provider override
  filterConfig?: IFilterConfig[];  // Vertical-specific filters
  audienceGroups?: string[];     // Azure AD security group IDs
  sortOrder: number;
}
```

### IUISlice

```typescript
export interface IUISlice {
  activeLayoutKey: string;
  isSearchManagerOpen: boolean;
  previewPanel: {
    isOpen: boolean;
    item: ISearchResult | null;
  };
  bulkSelection: string[];       // Array of selected item keys
  // Actions
  setLayout: (key: string) => void;
  toggleSearchManager: (isOpen?: boolean) => void;
  setPreviewItem: (item: ISearchResult | null) => void;
  toggleSelection: (itemKey: string, multiSelect: boolean) => void;
}
```

### IUserSlice

```typescript
export interface IUserSlice {
  savedSearches: ISavedSearch[];
  searchHistory: ISearchHistoryEntry[];
  collections: ISearchCollection[];
  // Actions
  saveSearch: (search: ISavedSearch) => Promise<void>;
  loadHistory: () => Promise<void>;
  addToHistory: (entry: ISearchHistoryEntry) => void;
}
```

### Data & Persistence Interfaces

```typescript
export interface ISavedSearch {
  id: number;
  title: string;
  queryText: string;
  searchState: string;           // JSON: serialized Query+Filter+Vertical slices
  searchUrl: string;
  entryType: 'SavedSearch' | 'SharedSearch';
  category: string;
  sharedWith: IPersonaInfo[];
  resultCount: number;
  lastUsed: Date;
  created: Date;
  author: IPersonaInfo;
}

export interface ISearchCollection {
  id: number;
  collectionName: string;
  items: ICollectionItem[];
  sharedWith: IPersonaInfo[];
  created: Date;
  author: IPersonaInfo;
}

export interface ISearchHistoryEntry {
  id: number;
  queryHash: string;
  queryText: string;
  vertical: string;
  scope: string;
  resultCount: number;
  clickedItems: IClickedItem[];
  searchTimestamp: Date;
}

export interface IClickedItem {
  url: string;
  title: string;
  position: number;
  timestamp: Date;
}
```

### Filter & Config Interfaces

```typescript
export interface IFilterConfig {
  id: string;
  displayName: string;
  managedProperty: string;
  filterType: FilterType;
  operator: FilterOperator;
  maxValues: number;
  defaultExpanded: boolean;
  showCount: boolean;
  sortBy: SortBy;
  sortDirection: SortDirection;
}
```

### Data Provider Interface (Extensibility — Most Critical)

```typescript
export interface ISearchDataProvider {
  id: string;
  displayName: string;
  execute: (query: ISearchQuery, signal: AbortSignal) => Promise<ISearchResponse>;
  getSuggestions?: (query: string, signal: AbortSignal) => Promise<ISuggestion[]>;
  getSchema?: () => Promise<IManagedProperty[]>;
  supportsRefiners: boolean;
  supportsCollapsing: boolean;
  supportsSorting: boolean;
}

export interface ISearchQuery {
  queryText: string;
  queryTemplate: string;
  scope: ISearchScope;
  filters: IActiveFilter[];
  sort: ISortField | null;
  page: number;
  pageSize: number;
  selectedProperties: string[];
  refiners: string[];
  collapseSpecification?: string;
  resultSourceId?: string;       // SharePoint-specific
  trimDuplicates?: boolean;
}

export interface ISearchResponse {
  items: ISearchResult[];
  totalCount: number;
  refiners: IRefiner[];
  promotedResults: IPromotedResult[];
  querySuggestion?: string;      // "Did you mean..."
}

export interface IManagedProperty {
  name: string;
  type: string;                  // Text, DateTime, Integer, etc.
  alias?: string;                // Human-readable alias
  queryable: boolean;
  retrievable: boolean;
  refinable: boolean;
  sortable: boolean;
}
```

### Extensibility Provider Interfaces

```typescript
export interface ISuggestionProvider {
  id: string;
  displayName: string;
  priority: number;
  maxResults: number;
  getSuggestions: (query: string, context: ISearchContext) => Promise<ISuggestion[]>;
  isEnabled: (context: ISearchContext) => boolean;
}

export interface IActionProvider {
  id: string;
  label: string;
  iconName: string;
  position: 'toolbar' | 'contextMenu' | 'both';
  isApplicable: (item: ISearchResult) => boolean;
  execute: (items: ISearchResult[], context: ISearchContext) => Promise<void>;
  isBulkEnabled: boolean;
}

export interface ILayoutDefinition {
  id: string;
  displayName: string;
  iconName: string;
  component: React.LazyExoticComponent<any>;
  supportsPaging: 'numbered' | 'infinite' | 'both';
  supportsBulkSelect: boolean;
  supportsVirtualization: boolean;
  defaultSortable: boolean;
}

export interface IFilterTypeDefinition {
  id: string;
  displayName: string;
  component: React.LazyExoticComponent<any>;
  serializeValue: (value: unknown) => string;
  deserializeValue: (raw: string) => unknown;
  buildRefinementToken: (value: unknown, managedProperty: string) => string;
}
```

### Promoted Results Interfaces

```typescript
export interface IPromotedResult {
  title: string;
  url: string;
  description?: string;
  iconUrl?: string;
}

export interface IPromotedResultRule {
  id: number;
  matchType: 'contains' | 'equals' | 'regex' | 'kql';
  matchValue: string;
  promotedItems: IPromotedResult[];
  audienceGroups: string[];
  startDate: Date | null;
  endDate: Date | null;
  verticalScope: string[] | null;
  isActive: boolean;
}
```

### Registry Container Interface

```typescript
export interface IRegistryContainer {
  dataProviders: Registry<ISearchDataProvider>;
  suggestions: Registry<ISuggestionProvider>;
  actions: Registry<IActionProvider>;
  layouts: Registry<ILayoutDefinition>;
  filterTypes: Registry<IFilterTypeDefinition>;
}

export interface Registry<T extends { id: string }> {
  register: (provider: T, force?: boolean) => void;
  get: (id: string) => T | undefined;
  getAll: () => T[];
  freeze: () => void;  // Locks after first search execution
}
```

## 10.2 Naming Conventions


| Element | Convention | Example |
| --- | --- | --- |
| Web Part class | SP[Name]WebPart | SPSearchBoxWebPart |
| React component | PascalCase | SearchResultsGrid |
| Zustand slice | camelCase + Slice | querySlice, filterSlice |
| Interface | I + PascalCase | ISearchResult, IFilterConfig |
| Hook | use + PascalCase | useSearchStore, useFilterState |
| Hidden list | PascalCase | SearchSavedQueries, SearchHistory |
| URL parameter | Short lowercase | q, f, v, s, p, sc, l, sv, sid |
| Provider class | PascalCase + Provider | RecentSearchProvider, ShareAction |
| Registry | PascalCase + Registry | LayoutRegistry, FilterTypeRegistry |


## 10.3 Browser Support

- Microsoft Edge (Chromium) — Primary
- Google Chrome — Primary
- Mozilla Firefox — Supported
- Safari — Supported (limited testing)

## 10.4 Document History


| Version | Date | Author | Changes |
| --- | --- | --- | --- |
| 1.0 | Feb 4, 2026 | Hemant Mane | Initial requirements specification |
| 1.1 | Feb 5, 2026 | Hemant Mane | Added: multi-instance store registry, dual-mode deep linking, item-level security, SearchHistory list, promoted results/best bets, abort/coalescing/refiner stability, provider/registry extensibility model |
| 1.2 | Feb 5, 2026 | Hemant Mane | Added: ISearchDataProvider abstraction (SharePoint + Graph providers), result collapsing/thread folding, Schema Helper property pane control, List View Threshold mitigation, per-vertical data provider override, IRegistryContainer, full typed store interfaces with actions |
| 1.3 | Feb 5, 2026 | Hemant Mane | Refinements: CollapseSpecification silent failure validation, CAML clause ordering for threshold safety, implementation file structure for sp-search-store |
| 1.4 | Feb 5, 2026 | Hemant Mane | Major UX additions: Active Filter Pill Bar (§3.3.3), Special Field Handling guide for 7 field types (§3.3.4), IFilterValueFormatter interface (§3.3.5), enhanced Result Detail Panel with WOPI preview, metadata formatting, version history UI (§3.2.3), DataGrid type-aware cell renderers for 12 property types (§3.2.2), DataGrid column filtering and sorting behavior, List Layout result card anatomy, ICellRendererConfig and caching interfaces |


## 10.5 Reference URLs for Development

The following URLs are essential references for Claude Code during development. Clone the PnP Modern Search repository and study source code patterns before implementing SP Search equivalents.

### PnP Modern Search v4 — Source Code

- Repository: https://github.com/microsoft-search/pnp-modern-search
- Web Parts: search-parts/src/webparts/ — SearchBoxWebPart.ts, SearchResultsWebPart.ts, SearchFiltersWebPart.ts, SearchVerticalsWebPart.ts
- Data Sources: search-parts/src/dataSources/SharePointSearchDataSource.ts — KQL query assembly, refinement token encoding, result mapping, refiner parsing. Reuse for PnPjs search service.
- Token Service: search-parts/src/services/tokenService/ — Dynamic token replacement ({searchTerms}, {Site.ID}, {Hub}, {Today}, {PageContext}, {User}). Port this logic.
- Layouts: search-parts/src/layouts/ — List, Cards, Details List, People, Custom. Reference for result-to-layout data mapping.
- Models/Interfaces: search-parts/src/models/ — IDataSource, ILayout, IFilterConfig. Study abstraction patterns.

### PnP Modern Search v4 — Documentation

- Home: https://microsoft-search.github.io/pnp-modern-search/
- Search Results: https://microsoft-search.github.io/pnp-modern-search/usage/search-results/
- SP Search Data Source: https://microsoft-search.github.io/pnp-modern-search/usage/search-results/data-sources/sharepoint-search/
- Query Tokens: https://microsoft-search.github.io/pnp-modern-search/usage/search-results/tokens/
- Result Layouts: https://microsoft-search.github.io/pnp-modern-search/usage/search-results/layouts/
- Search Box: https://microsoft-search.github.io/pnp-modern-search/usage/search-box/
- Search Filters: https://microsoft-search.github.io/pnp-modern-search/usage/search-filters/
- Filter Layouts: https://microsoft-search.github.io/pnp-modern-search/usage/search-filters/layouts/
- Search Verticals: https://microsoft-search.github.io/pnp-modern-search/usage/search-verticals/
- Installation: https://microsoft-search.github.io/pnp-modern-search/installation/
- Scenarios: https://microsoft-search.github.io/pnp-modern-search/scenarios/
- Custom Data Sources: https://microsoft-search.github.io/pnp-modern-search/extensibility/custom_data_sources/
- Extensibility Samples: https://github.com/microsoft-search/pnp-modern-search-extensibility-samples

### spfx-toolkit — Internal Component Library

- Repository: https://github.com/hmane/spfx-toolkit
- README.md — Critical bundle size import patterns (ALWAYS use direct path imports)
- SPFX-Toolkit-Usage-Guide.md — Full component API reference
- CLAUDE.md — AI development instructions for Claude Code
- src/components/ — Card, VersionHistory, DocumentLink, UserPersona, ErrorBoundary, Toast, FormContainer, WorkflowStepper
- src/hooks/ — useLocalStorage, useViewport, useCardController, useErrorHandler
- src/utilities/ — SPContext, BatchBuilder, createPermissionHelper, createSPExtractor

### Reuse Strategy from PnP Modern Search

The following PnP v4 patterns should be studied and adapted (not copied verbatim) for SP Search:
- Query Construction: SharePointSearchDataSource.ts — KQL assembly from template + user input + filters + sort. Port to SearchService using PnPjs sp.search().
- Refinement Token Handling: Filter value encoding/decoding, FQL range() operators, multi-value refinement. Critical for filter-to-query translation.
- Token Resolution: TokenService pattern for {searchTerms}, {Site.ID}, {Hub}, {Today+N}, {PageContext.*}, {User.*}. Port to utility for Zustand querySlice.
- Result Slot Mapping: Search result properties to display slots (Title, Summary, Path, Author). Adapt for ISearchResult interface.
- Filter Deep Linking: URL param f encoding for filter state. Extend to all Zustand slices via URL sync middleware.
- Vertical Count Queries: Parallel RowLimit=0 queries for per-vertical badge counts. Reuse directly.
- Suggestion Provider: Search box auto-suggest patterns. Adapt for debounced suggestion mechanism.