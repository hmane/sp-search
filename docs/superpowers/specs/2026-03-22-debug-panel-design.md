# SP Search Debug Panel — Design Spec

## Problem

Developers and QA have no visibility into what sp-search is doing under the hood. When results are unexpected, there's no way to see the constructed KQL, active filters, timing, errors, or web part configuration without attaching a debugger.

## Solution

A debug/diagnostic panel that appears only when `?debug=1` is in the URL. A floating button in the bottom-right corner toggles a slide-up panel with 4 tabs showing query details, store state, web part config, and an event log.

## Activation

- Panel FAB appears only when `window.location.search` contains `debug=1`
- All debug collection is no-op when debug mode is inactive (zero overhead in production)
- FAB is a small icon button, bottom-right, `z-index: 10000`

## Tabs

### 1. Query

Displays details of the last search execution:

| Field | Source |
|-------|--------|
| Constructed KQL | SearchOrchestrator (after token resolution + scope + filters) |
| Query template | Web part config `queryTemplate` |
| Result source ID | Provider config |
| Refinement filters | FQL array sent to SharePoint |
| Timing | Request start timestamp, duration (ms) |
| Results | totalCount, items returned, currentPage, pageSize |
| Refiners | Collapsible tree: refiner name -> values with counts |
| Errors | Any errors caught during the search cycle |
| Provider | Which ISearchDataProvider handled the request |

### 2. State

Live Zustand store snapshot, updating on every state change:

- **querySlice** — queryText, scope
- **filterSlice** — activeFilters, displayRefiners count, filterConfig
- **verticalSlice** — currentVerticalKey, verticals list, counts
- **resultSlice** — item count, totalCount, currentPage, sort, promotedResults count
- **uiSlice** — activeLayoutKey, availableLayouts
- **Registries** — registered provider/layout/filter type names

Each slice is a collapsible JSON tree. Values that changed in the last 2 seconds get a highlight flash.

### 3. Config

Web part property pane settings, grouped by web part:

- Each web part registers `{ componentName, properties }` during `onInit()`
- Displayed as collapsible JSON tree per web part
- Shows searchContextId, scope, layout config, filter configs, column configs, etc.

### 4. Log

Timestamped event stream (newest first, max 200 entries):

| Event Type | Captured When |
|------------|---------------|
| `SEARCH` | Query executed (includes duration, result count) |
| `FILTER` | Filter applied/removed (name + value) |
| `VERTICAL` | Vertical tab changed |
| `URL` | URL sync push or popstate restore |
| `ERROR` | Any error caught in orchestrator/providers/services |
| `INIT` | Web part initialized, provider registered |

Filterable by event type checkboxes.

## Architecture

### DebugCollector (singleton)

Location: `src/libraries/spSearchStore/debug/DebugCollector.ts`

Window-backed singleton (same pattern as store registry) to survive webpack entry-point duplication.

```
class DebugCollector {
  static isActive(): boolean        // checks ?debug=1
  logEvent(type, data): void        // adds timestamped entry (no-op if inactive)
  setLastQuery(queryInfo): void     // captures Query tab data
  registerWebPart(name, config): void // captures web part config
  getStoreRef(): StoreApi | null    // returns current store reference
  getEvents(): DebugEvent[]         // returns event log
  getLastQuery(): QueryDebugInfo    // returns last query info
  getWebPartConfigs(): Map          // returns registered configs
}
```

### DebugPanel (React component)

Location: `src/webparts/spSearchResults/components/DebugPanel.tsx`

- Lazy-loaded via `React.lazy()` (only loaded when debug=1)
- Rendered by the first web part that detects debug mode
- Subscribes to DebugCollector for data
- Dark theme: dark background (#1e1e1e), monospace font, syntax-highlighted JSON
- Slide-up from bottom, 60% viewport height, rounded top corners
- Resizable via drag handle, close/minimize buttons

### Integration Points

| Location | Call |
|----------|------|
| Each web part `onInit()` | `DebugCollector.registerWebPart(name, this.properties)` |
| SearchOrchestrator (before search) | `DebugCollector.setLastQuery({ kql, template, filters, ... })` |
| SearchOrchestrator (after search) | `DebugCollector.logEvent('SEARCH', { duration, resultCount, ... })` |
| SearchOrchestrator (on error) | `DebugCollector.logEvent('ERROR', { message, stack })` |
| filterSlice (on filter change) | `DebugCollector.logEvent('FILTER', { action, filterName, value })` |
| verticalSlice (on vertical change) | `DebugCollector.logEvent('VERTICAL', { key })` |
| urlSyncMiddleware (on push/pop) | `DebugCollector.logEvent('URL', { action, params })` |

### Performance

- All DebugCollector methods are no-ops when `debug=1` is not in the URL
- DebugPanel component is never imported/loaded in non-debug mode
- Event log capped at 200 entries
- No additional API calls — all data captured from existing code paths

## UI Design

- **FAB**: 40x40px circle, bottom-right (16px inset), bug icon, semi-transparent background, pulses briefly on new errors
- **Panel**: slides up from bottom, dark theme (#1e1e1e bg, #d4d4d4 text), monospace font (Cascadia Code / Consolas fallback)
- **Tabs**: horizontal tab bar at top of panel
- **JSON trees**: collapsible with indent guides, keys in muted blue, strings in orange, numbers in green, booleans in purple
- **Timing badges**: green < 500ms, yellow 500-2000ms, red > 2000ms
- **Error entries**: red left border accent in the log
