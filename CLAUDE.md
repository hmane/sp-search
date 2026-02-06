# CLAUDE.md - SP Search Development Guide

This file provides comprehensive guidance for Claude Code when working with the SP Search project. It is the primary reference for architecture, patterns, conventions, and rules.

## Quick Reference

1. **SPFx 1.21.1 solution** — 5 web parts + 1 library component in a single .sppkg
2. **React 17 + TypeScript 4.7+** — Functional components only, strict mode
3. **Zustand store via Library Component** — Shared state across web parts, NOT SPFx Dynamic Data
4. **Multi-instance isolation** — `searchContextId` property on every web part; same ID = shared store
5. **spfx-toolkit is at `/Users/hemantmane/Development/spfx-toolkit`** — ALWAYS use direct path imports
6. **Bundle size is critical** — Tree-shake spfx-toolkit, Fluent UI, and DevExtreme
7. **ISearchDataProvider abstraction** — Web parts never call PnPjs/Graph directly; always go through providers
8. **PnP Modern Search v4 is the reference** — Study patterns, don't copy verbatim
9. **No additional npm packages** — Only use the dependencies listed in the tech stack

---

## Project Overview

**SP Search** is an enterprise-grade SharePoint search solution built as SPFx web parts. It replaces PnP Modern Search v4 with:
- Modern React component architecture (not Handlebars)
- Rich UI via DevExtreme + spfx-toolkit + Fluent UI v8
- Shared Zustand store via SPFx Library Component (replaces Dynamic Data)
- Pluggable provider/registry model for data, suggestions, actions, layouts, filters
- Saved searches, search sharing, collections/pinboards, search history
- Client-side promoted results / best bets with audience targeting

### Solution Components

| # | Package | Type | Description |
|---|---------|------|-------------|
| 1 | sp-search-store | SPFx Library Component | Zustand store registry, interfaces, providers, registries |
| 2 | SPSearchBoxWebPart | Web Part | Query input, suggestions, scope selector, query builder |
| 3 | SPSearchResultsWebPart | Web Part | Result display with 6 layouts, detail panel, bulk actions |
| 4 | SPSearchFiltersWebPart | Web Part | Refinement filters with 7 filter types |
| 5 | SPSearchVerticalsWebPart | Web Part | Tab navigation with badge counts |
| 6 | SPSearchManagerWebPart | Web Part | Saved searches, sharing, collections, history |

---

## Technology Stack

### Core

| Technology | Version | Purpose |
|-----------|---------|---------|
| SharePoint Framework | 1.21.1 | SPFx web part platform |
| React | 17.0.1 | UI framework |
| TypeScript | 4.7+ | Type safety |
| PnPjs (SP) | 3.x | SharePoint Search API (default provider) |
| Microsoft Graph Client | 3.x | Graph Search API (optional provider) |

### UI Libraries

| Library | Version | Usage |
|---------|---------|-------|
| spfx-toolkit | Latest | Card, VersionHistory, DocumentLink, ErrorBoundary, Toast, UserPersona, FormContainer, hooks, utilities |
| DevExtreme | 22.2.x | DataGrid, FilterBuilder, TagBox, TreeView, DateRangeBox |
| devextreme-react | 22.2.x | React wrappers for DevExtreme |
| Fluent UI v8 | 8.106.x | Panel, CommandBar, Persona, Shimmer, Icons, Theme |
| @pnp/spfx-controls-react | 3.x | PeoplePicker, TaxonomyPicker |

### State & Utilities

| Library | Version | Purpose |
|---------|---------|---------|
| Zustand | 4.x | Shared state via library component + URL sync |
| React Hook Form | 7.x | Property pane config forms |

---

## Critical Rules

### Import Patterns (MANDATORY)

```typescript
// spfx-toolkit — ALWAYS direct path imports
import { Card, Header, Content } from 'spfx-toolkit/lib/components/Card';
import { ErrorBoundary } from 'spfx-toolkit/lib/components/ErrorBoundary';
import { useLocalStorage, useViewport } from 'spfx-toolkit/lib/hooks';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { BatchBuilder } from 'spfx-toolkit/lib/utilities/batchBuilder';
// NEVER: import { Card } from 'spfx-toolkit';

// Fluent UI v8 — ALWAYS tree-shakable imports
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { CommandBar } from '@fluentui/react/lib/CommandBar';
import { Persona, PersonaSize } from '@fluentui/react/lib/Persona';
import { Shimmer } from '@fluentui/react/lib/Shimmer';
import { Icon } from '@fluentui/react/lib/Icon';
// NEVER: import { Panel, Icon, Persona } from '@fluentui/react';

// DevExtreme — Lazy load heavy components
const DataGrid = React.lazy(() => import('devextreme-react/data-grid'));
const FilterBuilder = React.lazy(() => import('devextreme-react/filter-builder'));
const TreeView = React.lazy(() => import('devextreme-react/tree-view'));
// Direct import ONLY for lightweight components:
import { TagBox } from 'devextreme-react/tag-box';
import { DateRangeBox } from 'devextreme-react/date-range-box';
```

### Architecture Rules

1. **Web parts NEVER call PnPjs or Graph directly** — Always through `ISearchDataProvider`
2. **All inter-webpart communication via Zustand store** — No SPFx Dynamic Data, no events, no DOM
3. **Store is accessed via `getStore(searchContextId)`** — From the sp-search-store library component
4. **AbortController on every search** — Cancel in-flight requests before starting new ones
5. **Request coalescing** — Token resolution + KQL construction computed once per query cycle
6. **URL sync is bi-directional** — State changes update URL; URL load restores state
7. **Lazy load heavy components** — DataGrid, Preview Panel, Search Manager panel
8. **Code split per layout** — Each layout is a separate chunk via `React.lazy()`

### Data Rules

1. **SearchHistory list will exceed 5,000 items** — ALWAYS filter by `Author eq [Me]` as FIRST CAML predicate
2. **CollapseSpecification fails silently** if managed property isn't Sortable — Validate before sending
3. **Taxonomy refiner tokens use GP0|#GUID format** — Must resolve to labels via PnP Taxonomy API
4. **User claim strings (i:0#.f|...)** — Must resolve to display names, cache in Map
5. **Date refiners use FQL range()** — NOT raw KQL date comparisons
6. **Item-level permissions** on saved/shared searches — `breakRoleInheritance()` + `addRoleAssignment()`

### Security Rules

1. **All inputs sanitized** — Query text, tags, collection names
2. **No innerHTML** — Always render via React (XSS prevention)
3. **JSON validated against interfaces** before storage
4. **No data leaves the SharePoint tenant** — All processing client-side
5. **SearchConfiguration list is admin-only write** — Regular users Read only

---

## Architecture

### Store Registry Pattern

```typescript
// Library component exports
export function getStore(searchContextId: string): SearchStore;
export function disposeStore(searchContextId: string): void;

// Web part initialization
const store = getStore(this.properties.searchContextId);

// Same searchContextId = shared store instance
// Different searchContextId = fully isolated stores
```

### Store Slices

| Slice | Purpose | Key State |
|-------|---------|-----------|
| querySlice | Search query & execution | queryText, scope, suggestions, abortController |
| filterSlice | Filter selections & refiners | activeFilters, availableRefiners, displayRefiners |
| verticalSlice | Vertical tabs & counts | currentVerticalKey, verticals, verticalCounts |
| resultSlice | Results & pagination | items, totalCount, currentPage, sort, promotedResults |
| uiSlice | UI presentation | activeLayoutKey, previewPanel, bulkSelection |
| userSlice | User data from lists | savedSearches, searchHistory, collections |

### URL Sync Parameters

| Param | Slice Property | Example |
|-------|---------------|---------|
| q | querySlice.queryText | ?q=annual+report |
| f | filterSlice.activeFilters | &f=FileType:docx,pptx |
| v | verticalSlice.activeVertical | &v=documents |
| s | resultSlice.sort | &s=LastModifiedTime:desc |
| p | resultSlice.currentPage | &p=3 |
| sc | querySlice.scope | &sc=currentsite |
| l | uiSlice.activeLayout | &l=grid |
| sv | state version | &sv=1 |
| sid | state ID (fallback) | ?sid=42 |

Multi-context pages namespace params: `?ctx1.q=budget&ctx2.q=john`

### Provider/Registry Model

All registries are per-store-instance, hosted in sp-search-store:

| Registry | Interface | Built-in Providers |
|----------|-----------|-------------------|
| DataProviderRegistry | ISearchDataProvider | SharePointSearchProvider, GraphSearchProvider |
| SuggestionProviderRegistry | ISuggestionProvider | RecentSearchProvider, TrendingQueryProvider, ManagedPropertyProvider |
| ActionProviderRegistry | IActionProvider | OpenAction, PreviewAction, ShareAction, PinAction, CopyLinkAction, DownloadAction |
| LayoutRegistry | ILayoutDefinition | DataGrid, Card, List, Compact, People, DocumentGallery |
| FilterTypeRegistry | IFilterTypeDefinition | Checkbox, DateRange, Slider, PeoplePicker, TaxonomyTree, TagBox, Toggle |

Registration happens in `onInit()`. Registries freeze after first search execution.

---

## File Structure

```
sp-search/
├── CLAUDE.md
├── docs/
│   └── sp-search-requirements.md
├── src/
│   ├── library/                          # SPFx Library Component
│   │   └── sp-search-store/
│   │       ├── SpSearchStoreLibrary.ts   # Library component class
│   │       ├── store/
│   │       │   ├── createStore.ts        # Zustand store factory
│   │       │   ├── registry.ts           # Store registry (getStore/disposeStore)
│   │       │   └── slices/
│   │       │       ├── querySlice.ts
│   │       │       ├── filterSlice.ts
│   │       │       ├── verticalSlice.ts
│   │       │       ├── resultSlice.ts
│   │       │       ├── uiSlice.ts
│   │       │       └── userSlice.ts
│   │       ├── middleware/
│   │       │   └── urlSyncMiddleware.ts  # Bi-directional URL sync
│   │       ├── providers/
│   │       │   ├── data/
│   │       │   │   ├── SharePointSearchProvider.ts
│   │       │   │   └── GraphSearchProvider.ts
│   │       │   ├── suggestions/
│   │       │   │   ├── RecentSearchProvider.ts
│   │       │   │   ├── TrendingQueryProvider.ts
│   │       │   │   └── ManagedPropertyProvider.ts
│   │       │   └── actions/
│   │       │       ├── OpenAction.ts
│   │       │       ├── PreviewAction.ts
│   │       │       ├── ShareAction.ts
│   │       │       ├── PinAction.ts
│   │       │       ├── CopyLinkAction.ts
│   │       │       └── DownloadAction.ts
│   │       ├── registries/
│   │       │   ├── Registry.ts           # Generic Registry<T> class
│   │       │   ├── DataProviderRegistry.ts
│   │       │   ├── SuggestionProviderRegistry.ts
│   │       │   ├── ActionProviderRegistry.ts
│   │       │   ├── LayoutRegistry.ts
│   │       │   └── FilterTypeRegistry.ts
│   │       ├── services/
│   │       │   ├── SearchService.ts      # Query construction, token resolution, coalescing
│   │       │   ├── TokenService.ts       # {searchTerms}, {Site.ID}, etc.
│   │       │   └── SearchManagerService.ts # CRUD for saved searches, collections, history
│   │       ├── interfaces/
│   │       │   ├── ISearchStore.ts
│   │       │   ├── ISearchDataProvider.ts
│   │       │   ├── ISearchResult.ts
│   │       │   ├── IFilterConfig.ts
│   │       │   ├── IVerticalDefinition.ts
│   │       │   ├── ISuggestionProvider.ts
│   │       │   ├── IActionProvider.ts
│   │       │   ├── ILayoutDefinition.ts
│   │       │   ├── IFilterTypeDefinition.ts
│   │       │   ├── IFilterValueFormatter.ts
│   │       │   ├── IPromotedResult.ts
│   │       │   └── index.ts
│   │       ├── utils/
│   │       │   ├── urlEncoder.ts         # Filter/state URL encoding/decoding
│   │       │   ├── refinementTokens.ts   # FQL token encoding/decoding
│   │       │   ├── queryHash.ts          # SHA-256 for history dedup
│   │       │   └── formatters.ts         # Type-aware value formatters
│   │       └── index.ts                  # Public API exports
│   │
│   ├── webparts/
│   │   ├── searchBox/
│   │   │   ├── SPSearchBoxWebPart.ts
│   │   │   ├── components/
│   │   │   │   ├── SearchBox.tsx
│   │   │   │   ├── SuggestionDropdown.tsx
│   │   │   │   ├── ScopeSelector.tsx
│   │   │   │   └── QueryBuilder.tsx
│   │   │   ├── loc/
│   │   │   └── SPSearchBoxWebPart.manifest.json
│   │   │
│   │   ├── searchResults/
│   │   │   ├── SPSearchResultsWebPart.ts
│   │   │   ├── components/
│   │   │   │   ├── SearchResults.tsx
│   │   │   │   ├── ResultToolbar.tsx
│   │   │   │   ├── ActiveFilterPillBar.tsx
│   │   │   │   ├── BulkActionsToolbar.tsx
│   │   │   │   ├── PromotedResultsBlock.tsx
│   │   │   │   └── DetailPanel/
│   │   │   │       ├── ResultDetailPanel.tsx
│   │   │   │       ├── DocumentPreview.tsx
│   │   │   │       ├── MetadataDisplay.tsx
│   │   │   │       └── RelatedDocuments.tsx
│   │   │   ├── layouts/
│   │   │   │   ├── DataGridLayout.tsx
│   │   │   │   ├── CardLayout.tsx
│   │   │   │   ├── ListLayout.tsx
│   │   │   │   ├── CompactLayout.tsx
│   │   │   │   ├── PeopleLayout.tsx
│   │   │   │   └── DocumentGalleryLayout.tsx
│   │   │   ├── cellRenderers/
│   │   │   │   ├── TitleCellRenderer.tsx
│   │   │   │   ├── PersonaCellRenderer.tsx
│   │   │   │   ├── DateCellRenderer.tsx
│   │   │   │   ├── FileSizeCellRenderer.tsx
│   │   │   │   ├── FileTypeCellRenderer.tsx
│   │   │   │   ├── UrlCellRenderer.tsx
│   │   │   │   ├── TaxonomyCellRenderer.tsx
│   │   │   │   ├── BooleanCellRenderer.tsx
│   │   │   │   ├── NumberCellRenderer.tsx
│   │   │   │   ├── TagsCellRenderer.tsx
│   │   │   │   ├── ThumbnailCellRenderer.tsx
│   │   │   │   └── TextCellRenderer.tsx
│   │   │   ├── loc/
│   │   │   └── SPSearchResultsWebPart.manifest.json
│   │   │
│   │   ├── searchFilters/
│   │   │   ├── SPSearchFiltersWebPart.ts
│   │   │   ├── components/
│   │   │   │   ├── SearchFilters.tsx
│   │   │   │   ├── FilterGroup.tsx
│   │   │   │   └── VisualFilterBuilder.tsx
│   │   │   ├── filterTypes/
│   │   │   │   ├── CheckboxFilter.tsx
│   │   │   │   ├── DateRangeFilter.tsx
│   │   │   │   ├── PeoplePickerFilter.tsx
│   │   │   │   ├── TaxonomyTreeFilter.tsx
│   │   │   │   ├── TagBoxFilter.tsx
│   │   │   │   ├── SliderFilter.tsx
│   │   │   │   └── ToggleFilter.tsx
│   │   │   ├── formatters/
│   │   │   │   ├── DateFilterFormatter.ts
│   │   │   │   ├── PeopleFilterFormatter.ts
│   │   │   │   ├── TaxonomyFilterFormatter.ts
│   │   │   │   ├── NumericFilterFormatter.ts
│   │   │   │   ├── BooleanFilterFormatter.ts
│   │   │   │   └── DefaultFilterFormatter.ts
│   │   │   ├── loc/
│   │   │   └── SPSearchFiltersWebPart.manifest.json
│   │   │
│   │   ├── searchVerticals/
│   │   │   ├── SPSearchVerticalsWebPart.ts
│   │   │   ├── components/
│   │   │   │   ├── SearchVerticals.tsx
│   │   │   │   └── VerticalTab.tsx
│   │   │   ├── loc/
│   │   │   └── SPSearchVerticalsWebPart.manifest.json
│   │   │
│   │   └── searchManager/
│   │       ├── SPSearchManagerWebPart.ts
│   │       ├── components/
│   │       │   ├── SearchManager.tsx
│   │       │   ├── SavedSearchList.tsx
│   │       │   ├── SearchCollections.tsx
│   │       │   ├── SearchHistory.tsx
│   │       │   ├── ShareSearchDialog.tsx
│   │       │   └── ResultAnnotations.tsx
│   │       ├── loc/
│   │       └── SPSearchManagerWebPart.manifest.json
│   │
│   └── propertyPaneControls/
│       ├── PropertyPaneSchemaHelper.ts     # Managed property picker
│       └── PropertyPaneSearchContextId.ts  # Context ID config
│
├── scripts/
│   └── Provision-SPSearchLists.ps1         # Hidden list provisioning
│
├── config/                                 # SPFx config
│   ├── config.json
│   ├── deploy-azure-storage.json
│   ├── package-solution.json
│   └── serve.json
│
├── tests/
│   ├── store/
│   ├── providers/
│   ├── services/
│   └── utils/
│
├── package.json
├── tsconfig.json
├── gulpfile.js
└── .gitignore
```

---

## Naming Conventions

| Element | Convention | Example |
|---------|-----------|---------|
| Web Part class | SP[Name]WebPart | SPSearchBoxWebPart |
| React component | PascalCase | SearchResultsGrid |
| Zustand slice | camelCase + Slice | querySlice, filterSlice |
| Interface | I + PascalCase | ISearchResult, IFilterConfig |
| Hook | use + PascalCase | useSearchStore, useFilterState |
| Hidden list | PascalCase | SearchSavedQueries, SearchHistory |
| URL parameter | Short lowercase | q, f, v, s, p, sc, l, sv, sid |
| Provider class | PascalCase + Provider | SharePointSearchProvider |
| Action class | PascalCase + Action | ShareAction, PinAction |
| Registry | PascalCase + Registry | LayoutRegistry, FilterTypeRegistry |
| Service | PascalCase + Service | SearchService, TokenService |
| Cell renderer | PascalCase + CellRenderer | DateCellRenderer |
| Filter formatter | PascalCase + FilterFormatter | TaxonomyFilterFormatter |
| CSS class | kebab-case, scoped | .sp-search-results, .sp-filter-pill |

---

## Implementation Phases

### Phase 1: Foundation (Current)
- Library component with Zustand store registry + URL sync
- Provider/registry scaffolding (all interfaces + registry classes)
- Search Box: basic input, debounce, scope selector
- Search Results: List Layout + Compact Layout
- Search Filters: Checkbox + Date Range
- Search Verticals: tabs + badge counts
- PnPjs search service with AbortController + request coalescing
- Hidden list provisioning script

### Phase 2: Rich Layouts
- DataGrid Layout (DevExtreme), Card Layout (spfx-toolkit), People Layout, Gallery Layout
- Result Detail Panel with WOPI preview, metadata, version history
- Layout switcher, refiner stability mode

### Phase 3: User Features
- Search Manager (standalone + panel), saved/shared searches
- Collections/pinboards, search history, item-level permissions
- Promoted Results / Best Bets, StateId deep link fallback

### Phase 4: Power Features
- Visual Query Builder, advanced filter types, visual filter builder
- Result annotations, bulk actions, smart suggestions, audience targeting

### Phase 5: Polish
- Bundle optimization, accessibility audit, responsive testing, analytics

---

## Reference Repositories

- **PnP Modern Search v4:** https://github.com/microsoft-search/pnp-modern-search — Study query construction, token resolution, refinement tokens, layout switching
- **PnP Modern Search Docs:** https://microsoft-search.github.io/pnp-modern-search/
- **spfx-toolkit (local):** `/Users/hemantmane/Development/spfx-toolkit` — Components, hooks, utilities. See README.md, CLAUDE.md, SPFX-Toolkit-Usage-Guide.md
- **PnP Extensibility Samples:** https://github.com/microsoft-search/pnp-modern-search-extensibility-samples

---

## spfx-toolkit Integration Map

### Components Used

| Component | Web Part | Usage |
|-----------|----------|-------|
| Card + Header + Content | Search Results | Card layout, accordion grouping, maximize |
| Card (Accordion) | Search Filters | Collapsible filter groups with persistence |
| VersionHistory | Detail Panel | Version history with download/compare |
| DocumentLink | All layouts, Detail Panel | File type-aware document links |
| UserPersona | People Layout, Detail Panel | User profile with photo, presence |
| ErrorBoundary | All web parts | Root-level error wrapping |
| Toast / ToastProvider | All web parts | Save/share/export notifications |
| FormContainer / FormItem | Detail Panel, Search Manager | Metadata display, config forms |
| WorkflowStepper | Detail Panel | Workflow status display |

### Hooks Used

| Hook | Web Part | Usage |
|------|----------|-------|
| useLocalStorage | Search Box, Filters | Persist UI preferences |
| useViewport | All layouts | Responsive layout switching |
| useCardController | Card Layout, Filters | Programmatic card control |
| useErrorHandler | All web parts | Centralized error handling |

### Utilities Used

| Utility | Web Part | Usage |
|---------|----------|-------|
| SPContext | All web parts | SharePoint context for PnPjs |
| BatchBuilder | Search Manager | Batch list item operations |
| createPermissionHelper | Search Manager | Check permissions on hidden lists |
| createSPExtractor | Search Manager | Extract list item data |

---

## Common Commands

```bash
# Development
gulp serve                                    # Start local workbench
gulp bundle --ship && gulp package-solution --ship  # Production build

# Testing
npx jest                                      # Run unit tests
npx jest --watch                              # Watch mode

# spfx-toolkit (in toolkit directory)
cd /Users/hemantmane/Development/spfx-toolkit && npm run build
```

---

## Key Design Decisions

1. **Zustand over SPFx Dynamic Data** — More predictable state flow, multi-instance isolation, URL sync middleware
2. **ISearchDataProvider abstraction** — Allows mixing SharePoint Search + Graph on same page via per-vertical overrides
3. **Client-side promoted results** — Deterministic "Recommended" block at position #0, not invisible server-side ranking manipulation
4. **Dual-mode deep linking** — Short URL params (default) with automatic `?sid=` fallback for complex state
5. **Separate SearchHistory list** — Prevents threshold pressure on saved searches from high-volume history writes
6. **Provider registries freeze after first search** — Prevents mid-session mutations that cause race conditions
7. **Refiner stability mode** — Debounced displayRefiners prevents jarring filter option flicker during rapid typing

---

## Performance Checklist

- [ ] All spfx-toolkit imports use direct paths (`spfx-toolkit/lib/...`)
- [ ] All Fluent UI imports use tree-shakable paths (`@fluentui/react/lib/...`)
- [ ] DevExtreme DataGrid lazy-loaded via React.lazy()
- [ ] Detail panel lazy-loaded on first open
- [ ] Search Manager panel lazy-loaded on toggle
- [ ] Each layout is a separate code-split chunk
- [ ] AbortController cancels in-flight requests before new ones
- [ ] Token resolution computed once per query cycle (coalesced)
- [ ] Vertical count queries share AbortController with main query
- [ ] Thumbnail URLs cached with result data
- [ ] User claim strings resolved and cached in Map
- [ ] Taxonomy term GUIDs resolved and cached in Map
