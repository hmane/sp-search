# Admin Search Dashboard — Design Spec

## Problem

The Admin Search Manager has 3 separate tabs (Coverage, Health, Insights) that require tab-switching to get the full picture. The Coverage tab requires manual profile configuration that duplicates settings already present in the Results web part. Admins need a single-glance dashboard that auto-detects search config and surfaces actionable metrics.

## Solution

Replace the 3 admin tabs with a single unified **Dashboard** tab. The dashboard reads search config directly from the shared Zustand store (same `searchContextId`, same page), eliminating manual coverage profile configuration. Four scrollable sections provide complete admin visibility.

## Constraints

- Admin Manager web part is on the same page as the search web parts
- Single search context per page (`searchContextId`)
- Visible only to site collection admins (`SPPermission.manageWeb`)
- No additional npm packages — uses existing PnPjs, Fluent UI, DevExtreme
- No schema changes to SearchHistory list — all new metrics derived from existing fields

## Dashboard Sections

### 1. Coverage Stats

Four data points fetched by running queries against the store's search config (queryTemplate + scope + resultSourceId + refinementFilters):

| Metric | How | Display |
|--------|-----|---------|
| Item Count | `*` query, `RowLimit: 1`, read `TotalRows` from response | Large number with label |
| Freshness | `*` query sorted by `LastModifiedTime` desc (RowLimit:1) + asc (RowLimit:1) | "Newest: 2 hours ago" / "Oldest: 3 years ago". Green < 24h, yellow < 7d, red > 7d |
| File Type Breakdown | `*` query with `FileType` refiner | Horizontal bar chart, top 10 types with counts |
| Gap Analysis | Compare expected sites vs. `SPWebUrl` refiner results | List with Found/Missing status badges |

All 4 queries run in parallel on dashboard load. Single refresh button re-runs all. All queries accept an `AbortSignal` — cancelled on component unmount or refresh.

**Freshness fallback:** If `LastModifiedTime` is not sortable (sort returns no results), fall back to unsorted query and extract `LastModifiedTime` from the first result's properties. Log a warning in the debug panel.

#### Gap Analysis Data Sources

**Expected sites (two sources combined):**
- Auto-discover: `/_api/web/webs` (current site collection's direct child subsites — shallow, not recursive). For most SharePoint Online tenancies deep nesting is rare; this provides sufficient coverage.
- Manual overrides: admin adds/removes URLs in property pane (`expectedSiteUrls` string array) — covers nested subsites or external sites not auto-discovered.

**Actual sites:**
- `SPWebUrl` refiner from the `*` coverage query (subsite-level granularity, not site collection level)

**Display:**
- Each expected site shown with status badge: "Found" (green) or "Missing" (red)
- Missing sites highlighted — indicates content not being indexed

### 2. Quality Metrics

**Time Range Selector:** 30d / 60d / 90d toggle (applies to all quality metrics, zero-result table, and top queries).

**Stat Cards Row** — 5 cards, horizontal layout:

| Card | Computation | Warning |
|------|-------------|---------|
| Total Searches | history entry count | — |
| Zero-Result Rate | (isZeroResult entries / total) * 100 | > 20% red |
| Click-Through Rate | (entries with clickedItems / total) * 100 | < 30% yellow |
| Repeat Query Rate | (entries where UseCount > 1 / total) * 100. If UseCount field is unavailable (pre-Sprint 3 installs), show "N/A" with tooltip explaining the field is required. | > 40% yellow |
| Top Vertical | vertical with highest count | — (informational) |

**Vertical Usage Breakdown** — Horizontal bar chart below stat cards. Groups history entries by `vertical` field, shows search volume per vertical. Entries with empty vertical are labeled "All" (the default vertical).

**Sampling note:** Quality metrics are computed from the most recent 500 history entries (the `loadAllHistoryForInsights` row limit). This is a sample, not exhaustive — the dashboard displays a note: "Based on last 500 searches" next to the time range selector.

### 3. Zero-Result Queries

ZeroResultsPanel embedded as a collapsible section (expanded by default):
- Columns: Query text, Occurrences, Verticals, Last seen, "Try it" action
- Groups by queryText, aggregates count + verticals
- `daysBack` prop passed from the dashboard's time range selector (replaces the hardcoded 90-day default)

### 4. Top Queries & Engagement

SearchInsightsPanel charts embedded as collapsible sections:
- Top 10 queries by frequency (bar chart, clickable to re-run)
- Top 10 clicked results (bar chart with links)
- Daily volume sparkline (last 30 days)
- `daysBack` prop passed from dashboard; panel's own time range ChoiceGroup suppressed

**New — Repeat Queries Table:**
- Shows queries with UseCount > 2 (searched 3+ times)
- Columns: Query text, Total searches, Last seen
- Sorted by total searches descending
- Helps identify candidates for promoted results or query rules
- Computed from the same 500-entry sample; may miss older repeat queries
- Hidden when UseCount field is unavailable

## Architecture

### Store Integration

The dashboard reads search config from the shared Zustand store via `getStore(searchContextId)`:

```
Store State Used:
  - queryTemplate      → coverage query template
  - scope              → site/path restriction (includes scope.resultSourceId)
  - resultSourceId     → explicit result source GUID (takes priority over scope.resultSourceId)
  - refinementFilters  → persistent admin filters
  - selectedProperties → property list
```

Result source priority: `resultSourceId` (from resultSlice) > `scope.resultSourceId` (from ISearchScope). This matches the existing `SearchOrchestrator._buildQuery()` logic.

No manual profile configuration needed — the Results web part's config IS the coverage profile.

### Data Flow

```
Dashboard Load
  ├── Coverage Queries (parallel, via CoverageService, with AbortSignal)
  │   ├── Item count query (*, RowLimit: 1, read TotalRows)
  │   ├── Freshness query (*, sorted LastModifiedTime, RowLimit: 1 x2)
  │   ├── File type refiner query (*, Refiners: FileType)
  │   ├── Site refiner query (*, Refiners: SPWebUrl)
  │   └── Subsite discovery (/_api/web/webs, shallow)
  │
  └── Quality Queries (via existing SearchManagerService)
      ├── loadAllHistoryForInsights(daysBack, 500)
      └── loadZeroResultQueries(daysBack, 200)

Client-side computation:
  - Stat card metrics from history entries
  - Vertical usage grouping (empty vertical → "All")
  - Repeat query filtering (UseCount > 2, hidden if UseCount unavailable)
  - Gap analysis (expected vs. actual site URLs)
```

### CoverageService

New service that runs search queries against the store's config for coverage analysis. All methods accept `AbortSignal`.

```
class CoverageService {
  constructor(storeConfig: ICoverageConfig)

  getItemCount(signal: AbortSignal): Promise<number>
  getFreshness(signal: AbortSignal): Promise<{ newest: Date | undefined; oldest: Date | undefined }>
  getFileTypeBreakdown(signal: AbortSignal): Promise<Array<{ type: string; count: number }>>
  getSiteDistribution(signal: AbortSignal): Promise<Array<{ url: string; count: number }>>
  discoverExpectedSites(signal: AbortSignal): Promise<string[]>
  runAll(signal: AbortSignal): Promise<ICoverageResult>  // parallel execution of all above
}

interface ICoverageConfig {
  queryTemplate: string;
  scope: ISearchScope;
  resultSourceId: string | undefined;  // takes priority over scope.resultSourceId
  refinementFilters: string | undefined;
}

interface ICoverageResult {
  itemCount: number;
  newest: Date | undefined;
  oldest: Date | undefined;
  fileTypes: Array<{ type: string; count: number }>;
  actualSites: Array<{ url: string; count: number }>;
  expectedSites: string[];
  gapSites: string[];       // expected but not in actual
  timestamp: number;
}
```

## File Changes

| Action | Path | Responsibility |
|--------|------|---------------|
| Create | `src/libraries/spSearchStore/services/CoverageService.ts` | Coverage queries against store config |
| Create | `src/webparts/spSearchManager/components/AdminDashboard.tsx` | Main dashboard with 4 sections, time range state |
| Create | `src/webparts/spSearchManager/components/CoverageStatsSection.tsx` | Coverage stats + gap analysis UI |
| Create | `src/webparts/spSearchManager/components/QualityMetricsSection.tsx` | Stat cards + vertical usage chart |
| Modify | `src/webparts/spSearchManager/components/SpSearchManager.tsx` | Replace 3 admin tabs with single Dashboard tab, add `'dashboard'` to tab key type |
| Modify | `src/webparts/spSearchManager/components/ISpSearchManagerProps.ts` | Add `'dashboard'` to `defaultTab` type union |
| Modify | `src/webparts/spSearchManager/SpSearchManagerWebPart.ts` | Replace 3 enable toggles with `enableDashboard`, add `expectedSiteUrls` |
| Modify | `src/webparts/spSearchAdminManager/SpSearchAdminManagerWebPart.manifest.json` | Update defaults: `enableDashboard: true`, `defaultTab: 'dashboard'` |
| Modify | `src/webparts/spSearchManager/components/ZeroResultsPanel.tsx` | Accept `daysBack` prop (replaces hardcoded 90) |
| Modify | `src/webparts/spSearchManager/components/SearchInsightsPanel.tsx` | Accept `daysBack` prop, add `hideTimeRange` prop to suppress built-in ChoiceGroup; add repeat queries table + vertical usage chart |
| Remove | `src/webparts/spSearchManager/components/CoveragePanel.tsx` | Replaced by CoverageStatsSection |

**Note on CoveragePanel removal:** The existing CoveragePanel provides per-source-URL comparison and duplicate analysis via manually configured profiles. This functionality is replaced by the auto-detected coverage stats + gap analysis approach. The manual profile workflow is intentionally dropped — the store-based auto-detection provides equivalent coverage monitoring without admin configuration overhead.

## Error Handling

- Each dashboard section loads independently. If one section fails, it shows an inline error message while other sections continue to display data.
- Coverage query failures show "Unable to load coverage data" with a retry button per section.
- Quality data failures show "Unable to load search history" — covers both stat cards and zero-result table since they share the data source.
- Gap analysis gracefully handles `/_api/web/webs` failure by showing only manual overrides (with a note that auto-discovery failed).

## UI Design

- **Section headers** — collapsible with chevron icon, expanded by default
- **Stat cards** — same visual pattern as existing Insights cards (icon, label, value, optional warning color)
- **Bar charts** — same horizontal bar row pattern as existing Insights (animated width, label + count)
- **Gap analysis** — simple list with green/red status badges
- **Freshness indicator** — colored dot (green/yellow/red) next to date text
- **Refresh** — single button top-right refreshes all sections
- **Time range** — ChoiceGroup (30d/60d/90d) top of quality section, controls quality + zero-result + top queries
- **Sampling note** — small muted text "Based on last 500 searches" next to time range

## Performance

- Coverage queries run in parallel via `Promise.all`, all with AbortSignal
- Coverage results cached in component state — only re-fetched on manual refresh
- Quality data uses existing service methods (no new API calls)
- Subsite discovery (`/_api/web/webs`) cached — typically < 50 subsites
- No additional list queries beyond what Health + Insights already do
- AbortController cancels in-flight coverage queries on unmount or refresh
