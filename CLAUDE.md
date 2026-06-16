# CLAUDE.md - SP Search Development Guide

This file provides comprehensive guidance for Claude Code when working with the SP Search project. It is the primary reference for architecture, patterns, conventions, and rules.

## Quick Reference

1. **SPFx 1.22.2 + Heft solution** — 7 web parts + 1 library component in a single .sppkg
2. **React 17 + TypeScript 5.3+** — Functional components only, strict mode
3. **Zustand store via Library Component** — Shared state across web parts, NOT SPFx Dynamic Data
4. **Multi-instance isolation** — `searchContextId` property on every web part; same ID = shared store
5. **spfx-toolkit is a sibling clone** — `package.json` resolves it via `file:../spfx-toolkit`. Clone https://github.com/dodgeandcox/spfx-toolkit alongside this repo. ALWAYS use direct path imports (`spfx-toolkit/lib/...`)
6. **Bundle size is critical** — budgets enforced via `config/bundle-budgets.json` + `scripts/check-bundle-sizes.js`
7. **ISearchDataProvider abstraction** — Web parts never call PnPjs/Graph directly; always go through providers
8. **PnP Modern Search v4 is the reference** — Study patterns, don't copy verbatim
9. **No additional npm packages** — Only use the dependencies listed in the tech stack
10. **URL sanitization is mandatory** — every web part's `onInit()` must call `configureLegacyPnPBaseUrl(this.context)` after `SPContext.basic(...)` to strip `/_layouts/15` contamination from the PnP v2 base URL bundled with `@pnp/spfx-controls-react`
11. **Cancellation is owned by the orchestrator** — use `getOrchestrator(contextId).cancelPending()`. The slice has no `cancelSearch` action
12. **Registry freeze is lazy** — fires on first `_executeSearch`, NOT in `Results.onInit`. SPFx provides no init-order guarantee, so freezing too early creates a race

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
| 1 | spSearchStore | SPFx Library Component | Zustand store registry, orchestrator, services, providers, registries |
| 2 | SpSearchBoxWebPart | Web Part | Query input, suggestions, scope selector, query builder (KQL mode) |
| 3 | SpSearchResultsWebPart | Web Part | Result display with 6 layouts, detail panel, per-row ECB menu |
| 4 | SpSearchFiltersWebPart | Web Part | Refinement filters with 9 registered filter types, phone-width drawer |
| 5 | SpSearchExperienceWebPart | Web Part | Optional full-width wrapper that renders Results + Filters with one property bag |
| 6 | SpSearchVerticalsWebPart | Web Part | Tab navigation with badge counts, JS-measured overflow menu |
| 7 | SpSearchManagerWebPart | Web Part | Saved searches, sharing, collections, history (end-user variant) |
| 8 | SpSearchAdminManagerWebPart | Web Part | Subclass of Manager — Dashboard / Health / Insights / Pre-Flight (admin variant, gated by `manageWeb`) |

---

## Technology Stack

### Core

| Technology | Version | Purpose |
|-----------|---------|---------|
| SharePoint Framework | 1.22.2 | SPFx web part platform |
| React | 17.0.1 | UI framework |
| TypeScript | 5.3.x | Type safety |
| PnPjs (SP) | 3.x | SharePoint Search API (default provider) |
| Microsoft Graph Client | 3.x | Graph Search API (optional provider) |

### UI Libraries

| Library | Version | Usage |
|---------|---------|-------|
| spfx-toolkit | Latest | Card, VersionHistory, DocumentLink, ErrorBoundary, UserPersona, FormContainer, hooks, utilities |
| DevExtreme | 22.2.x | DataGrid, FilterBuilder, TagBox, DateBox, RangeSlider |
| devextreme-react | 22.2.x | React wrappers for DevExtreme |
| Fluent UI v8 | 8.106.x | Panel, CommandBar, Persona, Shimmer, Icons, Theme |
| @pnp/spfx-controls-react | 3.x | FileTypeIcon, Search Manager share-dialog PeoplePicker |
| @pnp/spfx-property-controls | 3.x | Utility property controls only. Do not reintroduce PnP `PropertyFieldCollectionData`; use the local `PropertyPaneCollectionData` replacement. |

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
// Direct import ONLY for lightweight components:
import { TagBox } from 'devextreme-react/tag-box';
import { DateBox } from 'devextreme-react/date-box';
import { RangeSlider } from 'devextreme-react/range-slider';
```

### Architecture Rules

1. **Web parts NEVER call PnPjs or Graph directly** — Always through `ISearchDataProvider`
2. **All inter-webpart communication via Zustand store** — No SPFx Dynamic Data, no events, no DOM
3. **Store is accessed via `getStore(searchContextId)`** — From the `spSearchStore` library component
4. **AbortController on every search** — Cancel in-flight requests before starting new ones
5. **Request coalescing** — Token resolution + KQL construction computed once per query cycle
6. **URL sync is bi-directional** — State changes update URL; URL load restores state
7. **Lazy load heavy components** — DataGrid, Preview Panel, Search Manager panel
8. **Code split per layout** — Each layout is a separate chunk via `React.lazy()`
9. **External CSS from node_modules must exclude sp-css-loader** — SPFx registers an `sp-css-loader` rule that matches all non-module `.css` files. `sp-css-loader` uses css-loader's `urlParser` internally and will try to import binary font files (woff2 etc.) as webpack modules. Any external CSS library (DevExtreme, etc.) must be excluded from the sp-css-loader rule in `gulpfile.js`, then handled by a dedicated `css-loader { url: false, import: false }` rule. See `gulpfile.js` `additionalConfiguration` for the pattern. Failing to do this produces `Module parse failed: Unexpected character` errors on binary font files.
10. **Stale webpack filesystem cache causes phantom `ENOENT` errors** — When `npm install` changes a package's entry point, the dev-mode filesystem cache (`node_modules/.cache/webpack`) retains the old path. Run `npm run clean:cache` (which invokes `rimraf node_modules/.cache`) whenever `@pnp/*` or other dependency packages are updated.

### Data Rules

1. **SearchHistory list will exceed 5,000 items** — ALWAYS filter by `Author eq [Me]` as FIRST CAML predicate for user queries. Admin cross-user queries (`loadZeroResultQueries`, `loadAllHistoryForInsights`) use `SearchTimestamp >= cutoff` as FIRST predicate instead — safe because SearchTimestamp is indexed.
2. **CollapseSpecification fails silently** if managed property isn't Sortable — Validate before sending
3. **Taxonomy refiner tokens use GP0|#GUID format** — Must resolve to labels via PnP Taxonomy API
4. **User claim strings (i:0#.f|...)** — Must resolve to display names, cache in Map
5. **Date refiners use FQL range()** — NOT raw KQL date comparisons
6. **Item-level permissions** on saved/shared searches — `breakRoleInheritance()` + `addRoleAssignment()`
7. **`IsZeroResult` Boolean field in SearchHistory** — Added in Sprint 3. Any pre-Sprint 3 install needs the field added before Health/Insights tabs populate. `mapToHistoryEntry` reads it with `ext.boolean('IsZeroResult', false)` — note `boolean()` not `bool()`.
8. **Graph permissions are capability-specific** — `People.Read` powers the People vertical; `User.Read` powers audience targeting through `/me/memberOf`; `User.Read.All` is optional for org-chart manager/direct reports. `Sites.Read.All` is not sufficient for Graph people search.

### Security Rules

1. **All inputs sanitized** — Query text, tags, collection names
2. **No innerHTML** — Always render via React (XSS prevention)
3. **JSON validated against interfaces** before storage
4. **No data leaves the SharePoint tenant** — All processing client-side
5. **Promoted results can be provider-mapped or client-evaluated** — SharePoint Query Rules arrive as `SpecialTermResults`; client-side promoted-result rules can also be evaluated by `PromotedResultsService`. Keep URLs/text sanitized and audience targeting fail-closed.

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
| uiSlice | UI presentation | activeLayoutKey, availableLayouts, previewPanel, currentUserGroups |
| userSlice | User data from lists | savedSearches, searchHistory, collections |

### URL Sync Parameters

| Param | Slice Property | Example |
|-------|---------------|---------|
| q | querySlice.queryText | ?q=annual+report |
| v | verticalSlice.activeVertical | &v=documents |
| s | resultSlice.sort | &s=LastModifiedTime:desc |
| p | resultSlice.currentPage | &p=3 |
| l | uiSlice.activeLayout | &l=grid |
| x | state version | &x=1 |
| i | state ID (fallback) | ?i=42 |
| \<alias\> | filter values (per filterConfig) | &ft=docx,pptx |

Multi-context pages namespace params: `?ctx1.q=budget&ctx2.q=john`

### Provider/Registry Model

All registries are per-store-instance, hosted in the `spSearchStore` library:

| Registry | Interface | Built-in Providers |
|----------|-----------|-------------------|
| DataProviderRegistry | ISearchDataProvider | SharePointSearchProvider, GraphSearchProvider |
| SuggestionProviderRegistry | ISuggestionProvider | RecentSearchProvider, TrendingQueryProvider, ManagedPropertyProvider |
| ActionProviderRegistry | IActionProvider | OpenAction, PreviewAction, ShareAction, PinAction, CopyLinkAction, DownloadAction |
| LayoutRegistry | ILayoutDefinition | DataGrid, Card, List, Compact, People, DocumentGallery |
| FilterTypeRegistry | IFilterTypeDefinition | Checkbox, Dropdown, DateRange, Text, Toggle, TagBox, Slider, Taxonomy TagBox, People |

Registration happens in `onInit()`. Registries freeze after first search execution.

---

## File Structure

```
sp-search/
├── CLAUDE.md
├── docs/
│   └── sp-search-requirements.md
├── src/
│   ├── libraries/
│   │   └── spSearchStore/                # SPFx Library Component
│   │       ├── SpSearchStoreLibrary.ts
│   │       ├── store/                    # Zustand store factory + slices
│   │       ├── orchestrator/             # SearchOrchestrator
│   │       ├── providers/                # SharePoint, Graph, suggestions
│   │       ├── providers/actions/        # Open, copy, download, pin/action providers
│   │       ├── registries/               # Generic Registry<T>
│   │       ├── services/                 # Search, token, audience, manager, coverage
│   │       ├── interfaces/               # Shared contracts
│   │       ├── configValidation/         # Shared edit-mode validators
│   │       ├── telemetry/
│   │       └── utils/
│   │
│   ├── webparts/
│   │   ├── spSearchBox/
│   │   │   ├── SpSearchBoxWebPart.ts
│   │   │   ├── components/
│   │   │   │   ├── SpSearchBox.tsx
│   │   │   │   ├── SuggestionDropdown.tsx
│   │   │   │   ├── ScopeSelector.tsx
│   │   │   │   └── QueryBuilder.tsx
│   │   │   ├── loc/
│   │   │   └── SpSearchBoxWebPart.manifest.json
│   │   │
│   │   ├── spSearchResults/
│   │   │   ├── SpSearchResultsWebPart.ts
│   │   │   ├── components/
│   │   │   │   ├── SpSearchResults.tsx
│   │   │   │   ├── ResultToolbar.tsx
│   │   │   │   ├── ActiveFilterPillBar.tsx
│   │   │   │   ├── DataGridLayout.tsx
│   │   │   │   ├── ListLayout.tsx
│   │   │   │   ├── CompactLayout.tsx
│   │   │   │   ├── CardLayout.tsx
│   │   │   │   ├── PeopleLayout.tsx
│   │   │   │   ├── GalleryLayout.tsx
│   │   │   │   ├── ResultDetailPanel.tsx
│   │   │   │   └── buildRowActionMenu.ts
│   │   │   ├── loc/
│   │   │   └── SpSearchResultsWebPart.manifest.json
│   │   │
│   │   ├── spSearchFilters/
│   │   │   ├── SpSearchFiltersWebPart.ts
│   │   │   ├── components/
│   │   │   │   ├── SpSearchFilters.tsx
│   │   │   │   ├── FilterGroup.tsx
│   │   │   │   ├── CheckboxFilter.tsx
│   │   │   │   ├── DropdownFilter.tsx
│   │   │   │   ├── DateRangeFilter.tsx
│   │   │   │   ├── TextFilter.tsx
│   │   │   │   ├── PeoplePickerFilter.tsx
│   │   │   │   ├── TaxonomyTreeFilter.tsx
│   │   │   │   ├── TagBoxFilter.tsx
│   │   │   │   ├── SliderFilter.tsx
│   │   │   │   ├── ToggleFilter.tsx
│   │   │   │   └── VisualFilterBuilder.tsx
│   │   │   ├── formatters/
│   │   │   ├── registerBuiltInFilterTypes.ts
│   │   │   ├── loc/
│   │   │   └── SpSearchFiltersWebPart.manifest.json
│   │   │
│   │   ├── spSearchExperience/
│   │   │   ├── SpSearchExperienceWebPart.ts
│   │   │   ├── components/
│   │   │   │   ├── SpSearchExperience.tsx
│   │   │   │   └── SpSearchExperience.module.scss
│   │   │   ├── loc/
│   │   │   └── SpSearchExperienceWebPart.manifest.json
│   │   │
│   │   ├── spSearchVerticals/
│   │   │   ├── SpSearchVerticalsWebPart.ts
│   │   │   ├── components/
│   │   │   │   └── SpSearchVerticals.tsx
│   │   │   ├── loc/
│   │   │   └── SpSearchVerticalsWebPart.manifest.json
│   │   │
│   │   ├── spSearchManager/
│   │   │   ├── SpSearchManagerWebPart.ts
│   │   │   ├── components/
│   │   │   │   ├── SpSearchManager.tsx
│   │   │   │   ├── SavedSearchList.tsx
│   │   │   │   ├── SearchCollections.tsx
│   │   │   │   ├── SearchHistory.tsx
│   │   │   │   ├── ShareSearchDialog.tsx
│   │   │   │   ├── AdminDashboard.tsx
│   │   │   │   └── PreFlightPanel.tsx
│   │   │   ├── loc/
│   │   │   └── SpSearchManagerWebPart.manifest.json
│   │   │
│   │   └── spSearchAdminManager/
│   │       ├── SpSearchAdminManagerWebPart.ts
│   │       ├── components/
│   │       ├── loc/
│   │       └── SpSearchAdminManagerWebPart.manifest.json
│   │
│   └── propertyPaneControls/
│       ├── PropertyPaneSchemaHelper.ts     # Managed property picker
│       ├── PropertyPaneSearchContextIdField.ts
│       ├── propertyPaneGroupHelp.tsx       # Local help modal topics
│       ├── collectionData/                 # PnP CollectionData replacement
│       └── filtersCollection/              # Refiner editor
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
| Web Part class | Sp[Name]WebPart | SpSearchBoxWebPart |
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

## Release State

**v1.0.0** is the current GA tag. The historical Phase 1-5 model is retired; see `CHANGELOG.md` for the GA changelog and `docs/sp-search-launch-readiness-audit.md` for the launch-readiness audit. Every audit P0/P1/P2 item has either landed on main or is documented as won't-fix with rationale in the commit message.

**Open follow-ups** (deferred, not derivable from code) are tracked in `docs/known-issues.md` — currently: cascade-filtered refiner counts, and cross-site version history (the spfx-toolkit `VersionHistory` is current-web-only).

Shipped capabilities (one-line each):
- 7 web parts + 1 library component, single .sppkg via Heft
- 6 layouts (DataGrid, Card, List, Compact, People, Gallery) with type-aware cell renderers; the DataGrid `tags` renderer ("Tags / badges") supports a configurable split delimiter + a colored-badge display style with admin value→color mapping (`ColumnConfigField/columnConfig.ts`, `renderCell.tsx`)
- 9 registered filter types (Checkbox, Dropdown, DateRange, Text, Toggle, TagBox, Slider, Taxonomy TagBox, People) + visual filter builder. Default-valued toggle filters are implicit (excluded from URL + history; see `utils/toggleDefaults.ts`); refiner tokens are decoded for display via `utils/refinerDisplay.ts`
- Two data providers (SharePoint + Graph) with per-vertical `dataProviderId` routing
- Search Manager (end-user + admin variant) with saved/shared/collections/history/promoted results
- AdminManager Dashboard / Health / Insights / Pre-Flight tabs
- Multi-context URL sync, scenario presets, layout switcher, detail panel with next/prev navigation
- Global keyboard shortcuts (`/`, `?`, `Esc`, `Enter`, `Alt+←/→`) via cross-bundle singleton host
- Cross-bundle DebugFab + Panel singleton (T5.D1)
- spLog PII-redacting logger + Pre-Flight tenant-readiness scan

---

## Reference Repositories

- **PnP Modern Search v4:** https://github.com/microsoft-search/pnp-modern-search — Study query construction, token resolution, refinement tokens, layout switching
- **PnP Modern Search Docs:** https://microsoft-search.github.io/pnp-modern-search/
- **spfx-toolkit (sibling clone):** https://github.com/dodgeandcox/spfx-toolkit — Resolved via `file:../spfx-toolkit` from this repo. See its README.md, CLAUDE.md, SPFX-Toolkit-Usage-Guide.md for component APIs.
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
| UserPersona | List, Grid cells, Detail Panel | User profile display with photo/name |
| ErrorBoundary | All web parts | Root-level error wrapping |
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
npm start                                     # heft start --clean (local workbench)
npm run package                               # heft build --clean --production && heft package-solution --production

# Testing
npm test                                      # heft test (Heft-managed Jest invocation)
npm test -- --watch                           # watch mode
npm test -- --test-path-pattern <pattern>     # filtered run

# spfx-toolkit (sibling directory — adjust path to wherever you cloned it)
cd ../spfx-toolkit && npm run build
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
