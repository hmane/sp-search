# Debug Panel Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a developer debug panel that appears when `?debug=1` is in the URL, showing query details, store state, web part config, and an event log.

**Architecture:** A window-backed `DebugCollector` singleton (same pattern as `storeRegistry.ts`) collects data from existing code paths. A lazy-loaded `DebugPanel` React component renders the UI. All DebugCollector methods are no-ops when debug mode is inactive — zero production overhead.

**Tech Stack:** React 17, TypeScript, Zustand (for store snapshots), Fluent UI v8 (Icon, Pivot), CSS modules

**Spec:** `docs/superpowers/specs/2026-03-22-debug-panel-design.md`

---

## File Structure

| Action | Path | Responsibility |
|--------|------|---------------|
| Create | `src/libraries/spSearchStore/debug/DebugCollector.ts` | Window-backed singleton: event log, query info, web part configs |
| Create | `src/libraries/spSearchStore/debug/IDebugTypes.ts` | Interfaces: `IDebugEvent`, `IQueryDebugInfo`, `IWebPartDebugConfig` |
| Create | `src/libraries/spSearchStore/debug/index.ts` | Re-exports |
| Create | `src/webparts/spSearchResults/components/DebugPanel.tsx` | React UI with 4 tabs (Query, State, Config, Log) |
| Create | `src/webparts/spSearchResults/components/DebugPanel.module.scss` | Dark theme styles |
| Create | `src/webparts/spSearchResults/components/DebugFab.tsx` | Floating action button (shown when debug=1) |
| Modify | `src/libraries/spSearchStore/orchestrator/SearchOrchestrator.ts` | Add DebugCollector calls for SEARCH, ERROR events + setLastQuery |
| Modify | `src/libraries/spSearchStore/store/slices/filterSlice.ts` | Add DebugCollector.logEvent('FILTER') on filter changes |
| Modify | `src/libraries/spSearchStore/store/slices/verticalSlice.ts` | Add DebugCollector.logEvent('VERTICAL') on vertical changes |
| Modify | `src/libraries/spSearchStore/store/middleware/urlSyncMiddleware.ts` | Add DebugCollector.logEvent('URL') on push/popstate |
| Modify | `src/webparts/spSearchResults/SpSearchResultsWebPart.ts` | registerWebPart() + render DebugFab/DebugPanel |
| Modify | `src/webparts/spSearchBox/SpSearchBoxWebPart.ts` | registerWebPart() in onInit |
| Modify | `src/webparts/spSearchFilters/SpSearchFiltersWebPart.ts` | registerWebPart() in onInit |
| Modify | `src/webparts/spSearchVerticals/SpSearchVerticalsWebPart.ts` | registerWebPart() in onInit |
| Modify | `src/webparts/spSearchManager/SpSearchManagerWebPart.ts` | registerWebPart() in onInit |
| Modify | `src/webparts/spSearchAdminManager/SpSearchAdminManagerWebPart.ts` | registerWebPart() in onInit |

---

## Task 1: Debug Interfaces

**Files:**
- Create: `src/libraries/spSearchStore/debug/IDebugTypes.ts`

- [ ] **Step 1: Create debug type definitions**

```typescript
// src/libraries/spSearchStore/debug/IDebugTypes.ts

export type DebugEventType = 'SEARCH' | 'FILTER' | 'VERTICAL' | 'URL' | 'ERROR' | 'INIT';

export interface IDebugEvent {
  readonly id: number;
  readonly type: DebugEventType;
  readonly timestamp: number;
  readonly data: Record<string, unknown>;
}

export interface IQueryDebugInfo {
  readonly kql: string;
  readonly queryTemplate: string;
  readonly resultSourceId: string | undefined;
  readonly refinementFilters: string[];
  readonly providerId: string;
  readonly startTime: number;
  readonly duration: number | undefined;
  readonly totalCount: number | undefined;
  readonly itemsReturned: number | undefined;
  readonly currentPage: number;
  readonly pageSize: number;
  readonly refiners: Array<{ name: string; values: Array<{ value: string; count: number }> }>;
  readonly error: string | undefined;
}

export interface IWebPartDebugConfig {
  readonly componentName: string;
  readonly properties: Record<string, unknown>;
  readonly registeredAt: number;
}
```

- [ ] **Step 2: Commit**

```bash
git add src/libraries/spSearchStore/debug/IDebugTypes.ts
git commit -m "feat(debug): add debug panel type definitions"
```

---

## Task 2: DebugCollector Singleton

**Files:**
- Create: `src/libraries/spSearchStore/debug/DebugCollector.ts`
- Create: `src/libraries/spSearchStore/debug/index.ts`

- [ ] **Step 1: Create DebugCollector**

The DebugCollector follows the same window-backed singleton pattern as `storeRegistry.ts`. All methods are no-ops when `?debug=1` is absent.

```typescript
// src/libraries/spSearchStore/debug/DebugCollector.ts
import type { IDebugEvent, IQueryDebugInfo, IWebPartDebugConfig, DebugEventType } from './IDebugTypes';

const COLLECTOR_KEY = '__sp_search_debug_collector__';
const MAX_EVENTS = 200;

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const _win = window as any;

interface IDebugCollectorState {
  events: IDebugEvent[];
  lastQuery: IQueryDebugInfo | undefined;
  webPartConfigs: Map<string, IWebPartDebugConfig>;
  nextEventId: number;
  listeners: Set<() => void>;
}

function getState(): IDebugCollectorState {
  if (!_win[COLLECTOR_KEY]) {
    _win[COLLECTOR_KEY] = {
      events: [],
      lastQuery: undefined,
      webPartConfigs: new Map(),
      nextEventId: 1,
      listeners: new Set(),
    };
  }
  return _win[COLLECTOR_KEY];
}

function notify(): void {
  const state = getState();
  state.listeners.forEach((fn) => fn());
}

let _active: boolean | undefined;

export const DebugCollector = {
  isActive(): boolean {
    if (_active === undefined) {
      try {
        _active = window.location.search.indexOf('debug=1') >= 0;
      } catch {
        _active = false;
      }
    }
    return _active;
  },

  logEvent(type: DebugEventType, data: Record<string, unknown>): void {
    if (!DebugCollector.isActive()) return;
    const state = getState();
    const event: IDebugEvent = {
      id: state.nextEventId++,
      type,
      timestamp: Date.now(),
      data,
    };
    state.events.unshift(event);
    if (state.events.length > MAX_EVENTS) {
      state.events.length = MAX_EVENTS;
    }
    notify();
  },

  setLastQuery(info: IQueryDebugInfo): void {
    if (!DebugCollector.isActive()) return;
    const state = getState();
    state.lastQuery = info;
    notify();
  },

  registerWebPart(name: string, properties: Record<string, unknown>): void {
    if (!DebugCollector.isActive()) return;
    const state = getState();
    state.webPartConfigs.set(name, {
      componentName: name,
      properties: { ...properties },
      registeredAt: Date.now(),
    });
    DebugCollector.logEvent('INIT', { componentName: name });
    notify();
  },

  getEvents(): IDebugEvent[] {
    return getState().events;
  },

  getLastQuery(): IQueryDebugInfo | undefined {
    return getState().lastQuery;
  },

  getWebPartConfigs(): Map<string, IWebPartDebugConfig> {
    return getState().webPartConfigs;
  },

  subscribe(listener: () => void): () => void {
    const state = getState();
    state.listeners.add(listener);
    return (): void => { state.listeners.delete(listener); };
  },
};
```

- [ ] **Step 2: Create index.ts re-exports**

```typescript
// src/libraries/spSearchStore/debug/index.ts
export { DebugCollector } from './DebugCollector';
export type { IDebugEvent, IQueryDebugInfo, IWebPartDebugConfig, DebugEventType } from './IDebugTypes';
```

- [ ] **Step 3: Commit**

```bash
git add src/libraries/spSearchStore/debug/
git commit -m "feat(debug): add DebugCollector window-backed singleton"
```

---

## Task 3: Instrument SearchOrchestrator

**Files:**
- Modify: `src/libraries/spSearchStore/orchestrator/SearchOrchestrator.ts`

- [ ] **Step 1: Add DebugCollector import**

At the top of the file, after the existing imports, add:

```typescript
import { DebugCollector } from '../debug';
```

- [ ] **Step 2: Instrument _executeSearch — before API call**

Inside `_executeSearch()`, after the `const query = this._buildQuery(state);` line (line ~309), add:

```typescript
      // Debug: capture query info before execution
      DebugCollector.setLastQuery({
        kql: query.queryText,
        queryTemplate: query.queryTemplate,
        resultSourceId: query.resultSourceId,
        refinementFilters: query.filters.map((f) => f.filterName + ':' + f.value),
        providerId: provider.id,
        startTime: searchStart,
        duration: undefined,
        totalCount: undefined,
        itemsReturned: undefined,
        currentPage: query.page,
        pageSize: query.pageSize,
        refiners: [],
        error: undefined,
      });
```

- [ ] **Step 3: Instrument _executeSearch — after success**

After the `const elapsed = Math.round(performance.now() - searchStart);` line (line ~388), add:

```typescript
      // Debug: update query info with results + log SEARCH event
      DebugCollector.setLastQuery({
        kql: query.queryText,
        queryTemplate: query.queryTemplate,
        resultSourceId: query.resultSourceId,
        refinementFilters: query.filters.map((f) => f.filterName + ':' + f.value),
        providerId: provider.id,
        startTime: searchStart,
        duration: elapsed,
        totalCount: adjustedTotal,
        itemsReturned: response.items.length,
        currentPage: query.page,
        pageSize: query.pageSize,
        refiners: response.refiners.map((r) => ({
          name: r.refinerName,
          values: r.values.map((v) => ({ value: v.value, count: v.count })),
        })),
        error: undefined,
      });
      DebugCollector.logEvent('SEARCH', {
        duration: elapsed,
        resultCount: response.items.length,
        totalCount: adjustedTotal,
        providerId: provider.id,
        query: query.queryText,
      });
```

- [ ] **Step 4: Instrument _executeSearch — on error**

Inside the catch block (line ~412), after the `state.setError(message);` line, add:

```typescript
      DebugCollector.logEvent('ERROR', {
        message,
        providerId: provider.id,
        query: state.queryText || '*',
        stack: error instanceof Error ? error.stack : undefined,
      });
```

- [ ] **Step 5: Verify build compiles**

Run: `cd /Users/hemantmane/Development/sp-search && npx tsc --noEmit --skipLibCheck 2>&1 | head -20`

- [ ] **Step 6: Commit**

```bash
git add src/libraries/spSearchStore/orchestrator/SearchOrchestrator.ts
git commit -m "feat(debug): instrument SearchOrchestrator with DebugCollector"
```

---

## Task 4: Instrument Filter Slice

**Files:**
- Modify: `src/libraries/spSearchStore/store/slices/filterSlice.ts`

- [ ] **Step 1: Add import and instrument setRefiner/removeRefiner/clearAllFilters**

Add at the top of the file:

```typescript
import { DebugCollector } from '../../debug';
```

Inside `setRefiner` — at the very end of the function body (after all the `set()` calls), add:

```typescript
    DebugCollector.logEvent('FILTER', { action: 'set', filterName: filter.filterName, value: filter.value });
```

Inside `removeRefiner` — after the `set()` call, add:

```typescript
    DebugCollector.logEvent('FILTER', { action: 'remove', filterName: filterKey, value: value || '*' });
```

Inside `clearAllFilters` — after the `set()` call, add:

```typescript
    DebugCollector.logEvent('FILTER', { action: 'clearAll' });
```

- [ ] **Step 2: Commit**

```bash
git add src/libraries/spSearchStore/store/slices/filterSlice.ts
git commit -m "feat(debug): instrument filterSlice with DebugCollector"
```

---

## Task 5: Instrument Vertical Slice

**Files:**
- Modify: `src/libraries/spSearchStore/store/slices/verticalSlice.ts`

- [ ] **Step 1: Add import and instrument setVertical**

Add at the top:

```typescript
import { DebugCollector } from '../../debug';
```

Inside `setVertical` — after the `set()` call, add:

```typescript
    DebugCollector.logEvent('VERTICAL', { key });
```

- [ ] **Step 2: Commit**

```bash
git add src/libraries/spSearchStore/store/slices/verticalSlice.ts
git commit -m "feat(debug): instrument verticalSlice with DebugCollector"
```

---

## Task 6: Instrument URL Sync Middleware

**Files:**
- Modify: `src/libraries/spSearchStore/store/middleware/urlSyncMiddleware.ts`

- [ ] **Step 1: Add import**

Add at the top:

```typescript
import { DebugCollector } from '../../debug';
```

- [ ] **Step 2: Instrument pushStateToUrl**

Inside the `pushStateToUrl` function (the part that actually calls `history.pushState` or `history.replaceState`), add a DebugCollector call. Find the `history.pushState` / `history.replaceState` call and add immediately before it:

```typescript
    DebugCollector.logEvent('URL', { action: 'push', params: newSearch });
```

- [ ] **Step 3: Instrument popstate handler**

Inside the `onPopState` handler in `createUrlSyncSubscription` (line ~755), add after `deserializeFromUrl`:

```typescript
    DebugCollector.logEvent('URL', { action: 'popstate', params: window.location.search });
```

- [ ] **Step 4: Commit**

```bash
git add src/libraries/spSearchStore/store/middleware/urlSyncMiddleware.ts
git commit -m "feat(debug): instrument urlSyncMiddleware with DebugCollector"
```

---

## Task 7: Instrument Web Part onInit Methods

**Files:**
- Modify: `src/webparts/spSearchResults/SpSearchResultsWebPart.ts`
- Modify: `src/webparts/spSearchBox/SpSearchBoxWebPart.ts`
- Modify: `src/webparts/spSearchFilters/SpSearchFiltersWebPart.ts`
- Modify: `src/webparts/spSearchVerticals/SpSearchVerticalsWebPart.ts`
- Modify: `src/webparts/spSearchManager/SpSearchManagerWebPart.ts`
- Modify: `src/webparts/spSearchAdminManager/SpSearchAdminManagerWebPart.ts`

- [ ] **Step 1: Add DebugCollector.registerWebPart to each web part's onInit**

In each web part file, add at the top:

```typescript
import { DebugCollector } from '@store/debug';
```

The `@store/*` alias in `tsconfig.json` maps to `src/libraries/spSearchStore/*`, so `@store/debug` resolves to `src/libraries/spSearchStore/debug/index.ts`.

At the end of each `onInit()` method, add:

```typescript
    DebugCollector.registerWebPart('SPSearchResultsWebPart', this.properties as unknown as Record<string, unknown>);
```

(Replace the component name string for each web part: `SPSearchBoxWebPart`, `SPSearchFiltersWebPart`, `SPSearchVerticalsWebPart`, `SPSearchManagerWebPart`, `SPSearchAdminManagerWebPart`.)

- [ ] **Step 2: Verify build compiles**

Run: `cd /Users/hemantmane/Development/sp-search && npx tsc --noEmit --skipLibCheck 2>&1 | head -20`

- [ ] **Step 3: Commit**

```bash
git add src/webparts/
git commit -m "feat(debug): register web parts with DebugCollector in onInit"
```

---

## Task 8: DebugPanel Styles

**Files:**
- Create: `src/webparts/spSearchResults/components/DebugPanel.module.scss`

- [ ] **Step 1: Create SCSS module**

```scss
// src/webparts/spSearchResults/components/DebugPanel.module.scss

.debugFab {
  position: fixed;
  bottom: 16px;
  right: 16px;
  z-index: 10000;
  width: 40px;
  height: 40px;
  border-radius: 50%;
  background: rgba(30, 30, 30, 0.85);
  color: #d4d4d4;
  border: 1px solid #444;
  cursor: pointer;
  display: flex;
  align-items: center;
  justify-content: center;
  font-size: 18px;
  transition: background 0.2s;

  &:hover {
    background: rgba(50, 50, 50, 0.95);
  }

  &.hasError {
    animation: pulse 1.5s ease-in-out 3;
  }
}

@keyframes pulse {
  0%, 100% { box-shadow: 0 0 0 0 rgba(255, 80, 80, 0.4); }
  50% { box-shadow: 0 0 0 8px rgba(255, 80, 80, 0); }
}

.debugPanel {
  position: fixed;
  bottom: 0;
  left: 0;
  right: 0;
  z-index: 10001;
  background: #1e1e1e;
  color: #d4d4d4;
  font-family: 'Cascadia Code', Consolas, 'Courier New', monospace;
  font-size: 12px;
  border-top: 2px solid #444;
  border-radius: 8px 8px 0 0;
  display: flex;
  flex-direction: column;
  box-shadow: 0 -4px 20px rgba(0, 0, 0, 0.5);
}

.dragHandle {
  height: 6px;
  cursor: ns-resize;
  display: flex;
  align-items: center;
  justify-content: center;

  &::after {
    content: '';
    width: 40px;
    height: 3px;
    background: #555;
    border-radius: 2px;
  }
}

.header {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 4px 12px;
  border-bottom: 1px solid #333;
}

.headerTitle {
  font-weight: 600;
  font-size: 11px;
  text-transform: uppercase;
  letter-spacing: 0.5px;
  color: #888;
}

.headerActions {
  display: flex;
  gap: 4px;

  button {
    background: none;
    border: none;
    color: #888;
    cursor: pointer;
    padding: 2px 6px;
    font-size: 14px;

    &:hover {
      color: #d4d4d4;
    }
  }
}

.tabBar {
  display: flex;
  border-bottom: 1px solid #333;
}

.tab {
  padding: 6px 16px;
  cursor: pointer;
  border: none;
  background: none;
  color: #888;
  font-size: 12px;
  font-family: inherit;
  border-bottom: 2px solid transparent;

  &:hover {
    color: #d4d4d4;
  }
}

.tabActive {
  color: #569cd6;
  border-bottom-color: #569cd6;
}

.tabContent {
  flex: 1;
  overflow-y: auto;
  padding: 8px 12px;
}

// JSON tree styles
.jsonKey {
  color: #9cdcfe;
}

.jsonString {
  color: #ce9178;
}

.jsonNumber {
  color: #b5cea8;
}

.jsonBoolean {
  color: #c586c0;
}

.jsonNull {
  color: #808080;
  font-style: italic;
}

.jsonBracket {
  color: #808080;
}

.jsonToggle {
  cursor: pointer;
  user-select: none;
  color: #808080;
  margin-right: 4px;
}

.jsonRow {
  padding-left: 16px;
  line-height: 1.6;
}

.jsonHighlight {
  background: rgba(86, 156, 214, 0.15);
  transition: background 2s ease-out;
}

// Timing badges
.timingBadge {
  display: inline-block;
  padding: 1px 6px;
  border-radius: 3px;
  font-size: 11px;
  font-weight: 600;
  margin-left: 8px;
}

.timingGreen {
  background: rgba(80, 200, 120, 0.2);
  color: #50c878;
}

.timingYellow {
  background: rgba(255, 200, 50, 0.2);
  color: #ffc832;
}

.timingRed {
  background: rgba(255, 80, 80, 0.2);
  color: #ff5050;
}

// Log tab styles
.logEntry {
  padding: 4px 0;
  border-bottom: 1px solid #2a2a2a;
  display: flex;
  gap: 8px;
  align-items: flex-start;
}

.logEntryError {
  border-left: 3px solid #ff5050;
  padding-left: 8px;
}

.logTimestamp {
  color: #666;
  white-space: nowrap;
  flex-shrink: 0;
}

.logType {
  font-weight: 600;
  white-space: nowrap;
  flex-shrink: 0;
  padding: 0 4px;
  border-radius: 2px;
  font-size: 10px;
  text-transform: uppercase;
}

.logTypeSearch { color: #569cd6; }
.logTypeFilter { color: #dcdcaa; }
.logTypeVertical { color: #c586c0; }
.logTypeUrl { color: #4ec9b0; }
.logTypeError { color: #ff5050; }
.logTypeInit { color: #808080; }

.logData {
  color: #999;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}

.logFilters {
  display: flex;
  gap: 8px;
  padding: 4px 0 8px;
  border-bottom: 1px solid #333;
  flex-wrap: wrap;
}

.logFilterCheckbox {
  display: flex;
  align-items: center;
  gap: 4px;
  color: #888;
  font-size: 11px;
  cursor: pointer;

  input {
    cursor: pointer;
  }
}

// State tab collapsible sections
.stateSection {
  margin-bottom: 8px;
}

.stateSectionHeader {
  cursor: pointer;
  padding: 4px 0;
  font-weight: 600;
  color: #569cd6;
  user-select: none;

  &:hover {
    color: #7ec8e3;
  }
}

// Query tab field rows
.queryField {
  display: flex;
  gap: 12px;
  padding: 4px 0;
  border-bottom: 1px solid #2a2a2a;
}

.queryLabel {
  color: #888;
  min-width: 140px;
  flex-shrink: 0;
}

.queryValue {
  word-break: break-all;
}
```

- [ ] **Step 2: Commit**

```bash
git add src/webparts/spSearchResults/components/DebugPanel.module.scss
git commit -m "feat(debug): add DebugPanel dark theme styles"
```

---

## Task 9: DebugFab Component

**Files:**
- Create: `src/webparts/spSearchResults/components/DebugFab.tsx`

- [ ] **Step 1: Create the floating action button component**

```tsx
// src/webparts/spSearchResults/components/DebugFab.tsx
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { DebugCollector } from '@store/debug';
import styles from './DebugPanel.module.scss';

export interface IDebugFabProps {
  onClick: () => void;
}

const DebugFab: React.FC<IDebugFabProps> = ({ onClick }) => {
  const [hasError, setHasError] = React.useState(false);

  React.useEffect(() => {
    return DebugCollector.subscribe(() => {
      const events = DebugCollector.getEvents();
      const recentError = events.length > 0 &&
        events[0].type === 'ERROR' &&
        Date.now() - events[0].timestamp < 5000;
      setHasError(!!recentError);
    });
  }, []);

  return (
    <button
      className={`${styles.debugFab}${hasError ? ' ' + styles.hasError : ''}`}
      onClick={onClick}
      title="SP Search Debug Panel"
      type="button"
    >
      <Icon iconName="Bug" />
    </button>
  );
};

export default DebugFab;
```

- [ ] **Step 2: Commit**

```bash
git add src/webparts/spSearchResults/components/DebugFab.tsx
git commit -m "feat(debug): add DebugFab floating button component"
```

---

## Task 10: DebugPanel Component

**Files:**
- Create: `src/webparts/spSearchResults/components/DebugPanel.tsx`

This is the largest task. The panel has 4 tabs: Query, State, Config, Log.

- [ ] **Step 1: Create DebugPanel with tab infrastructure and JsonTree helper**

```tsx
// src/webparts/spSearchResults/components/DebugPanel.tsx
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { DebugCollector } from '@store/debug';
import type { IDebugEvent, IQueryDebugInfo, DebugEventType } from '@store/debug';
import type { StoreApi } from 'zustand/vanilla';
import type { ISearchStore } from '@interfaces/index';
import styles from './DebugPanel.module.scss';

// ─── Helpers ─────────────────────────────────────────────

function formatTime(ts: number): string {
  const d = new Date(ts);
  return d.toLocaleTimeString([], { hour12: false }) + '.' + String(d.getMilliseconds()).padStart(3, '0');
}

function timingClass(ms: number): string {
  if (ms < 500) return styles.timingGreen;
  if (ms <= 2000) return styles.timingYellow;
  return styles.timingRed;
}

const LOG_TYPE_CLASSES: Record<DebugEventType, string> = {
  SEARCH: styles.logTypeSearch,
  FILTER: styles.logTypeFilter,
  VERTICAL: styles.logTypeVertical,
  URL: styles.logTypeUrl,
  ERROR: styles.logTypeError,
  INIT: styles.logTypeInit,
};

// ─── Collapsible JSON Tree ───────────────────────────────

interface IJsonTreeProps {
  data: unknown;
  depth?: number;
  defaultExpanded?: boolean;
}

const JsonTree: React.FC<IJsonTreeProps> = ({ data, depth = 0, defaultExpanded = true }) => {
  const [expanded, setExpanded] = React.useState(defaultExpanded);

  if (data === null || data === undefined) {
    return <span className={styles.jsonNull}>{String(data)}</span>;
  }

  if (typeof data === 'string') {
    return <span className={styles.jsonString}>"{data}"</span>;
  }

  if (typeof data === 'number') {
    return <span className={styles.jsonNumber}>{data}</span>;
  }

  if (typeof data === 'boolean') {
    return <span className={styles.jsonBoolean}>{String(data)}</span>;
  }

  if (Array.isArray(data)) {
    if (data.length === 0) {
      return <span className={styles.jsonBracket}>[]</span>;
    }
    return (
      <span>
        <span className={styles.jsonToggle} onClick={() => setExpanded(!expanded)}>
          {expanded ? '\u25BC' : '\u25B6'}
        </span>
        <span className={styles.jsonBracket}>[{data.length}]</span>
        {expanded && data.map((item, i) => (
          <div key={i} className={styles.jsonRow}>
            <span className={styles.jsonKey}>{i}: </span>
            <JsonTree data={item} depth={depth + 1} defaultExpanded={depth < 1} />
          </div>
        ))}
      </span>
    );
  }

  if (typeof data === 'object') {
    const entries = Object.entries(data as Record<string, unknown>);
    if (entries.length === 0) {
      return <span className={styles.jsonBracket}>{'{}'}</span>;
    }
    return (
      <span>
        <span className={styles.jsonToggle} onClick={() => setExpanded(!expanded)}>
          {expanded ? '\u25BC' : '\u25B6'}
        </span>
        <span className={styles.jsonBracket}>{'{'}{entries.length}{'}'}</span>
        {expanded && entries.map(([key, val]) => (
          <div key={key} className={styles.jsonRow}>
            <span className={styles.jsonKey}>{key}: </span>
            <JsonTree data={val} depth={depth + 1} defaultExpanded={depth < 1} />
          </div>
        ))}
      </span>
    );
  }

  return <span>{String(data)}</span>;
};

// ─── Query Tab ───────────────────────────────────────────

const QueryTab: React.FC = () => {
  const [query, setQuery] = React.useState<IQueryDebugInfo | undefined>(DebugCollector.getLastQuery());

  React.useEffect(() => {
    return DebugCollector.subscribe(() => {
      setQuery(DebugCollector.getLastQuery());
    });
  }, []);

  if (!query) {
    return <div style={{ color: '#888', padding: 16 }}>No search executed yet.</div>;
  }

  return (
    <div>
      <div className={styles.queryField}>
        <span className={styles.queryLabel}>KQL</span>
        <span className={styles.queryValue}>{query.kql || '(empty)'}</span>
      </div>
      <div className={styles.queryField}>
        <span className={styles.queryLabel}>Query Template</span>
        <span className={styles.queryValue}>{query.queryTemplate}</span>
      </div>
      <div className={styles.queryField}>
        <span className={styles.queryLabel}>Result Source ID</span>
        <span className={styles.queryValue}>{query.resultSourceId || '(default)'}</span>
      </div>
      <div className={styles.queryField}>
        <span className={styles.queryLabel}>Provider</span>
        <span className={styles.queryValue}>{query.providerId}</span>
      </div>
      <div className={styles.queryField}>
        <span className={styles.queryLabel}>Timing</span>
        <span className={styles.queryValue}>
          {query.duration !== undefined ? (
            <>
              {query.duration}ms
              <span className={`${styles.timingBadge} ${timingClass(query.duration)}`}>
                {query.duration}ms
              </span>
            </>
          ) : 'In progress...'}
        </span>
      </div>
      <div className={styles.queryField}>
        <span className={styles.queryLabel}>Results</span>
        <span className={styles.queryValue}>
          {query.itemsReturned !== undefined
            ? `${query.itemsReturned} of ${query.totalCount} (page ${query.currentPage}, size ${query.pageSize})`
            : 'Pending...'}
        </span>
      </div>
      {query.refinementFilters.length > 0 && (
        <div className={styles.queryField}>
          <span className={styles.queryLabel}>Refinement Filters</span>
          <span className={styles.queryValue}>
            <JsonTree data={query.refinementFilters} defaultExpanded={true} />
          </span>
        </div>
      )}
      {query.refiners.length > 0 && (
        <div className={styles.queryField}>
          <span className={styles.queryLabel}>Refiners</span>
          <span className={styles.queryValue}>
            <JsonTree data={query.refiners} defaultExpanded={false} />
          </span>
        </div>
      )}
      {query.error && (
        <div className={styles.queryField}>
          <span className={styles.queryLabel} style={{ color: '#ff5050' }}>Error</span>
          <span className={styles.queryValue} style={{ color: '#ff5050' }}>{query.error}</span>
        </div>
      )}
    </div>
  );
};

// ─── State Tab ───────────────────────────────────────────

interface IStateTabProps {
  store: StoreApi<ISearchStore>;
}

const StateTab: React.FC<IStateTabProps> = ({ store }) => {
  const [snapshot, setSnapshot] = React.useState<Record<string, unknown>>({});
  const [changedKeys, setChangedKeys] = React.useState<Set<string>>(new Set());
  const prevSnapshotRef = React.useRef<Record<string, unknown>>({});

  React.useEffect(() => {
    function update(): void {
      const s = store.getState();
      const next: Record<string, unknown> = {
        querySlice: {
          queryText: s.queryText,
          scope: s.scope,
          queryTemplate: s.queryTemplate,
        },
        filterSlice: {
          activeFilters: s.activeFilters,
          displayRefinersCount: s.displayRefiners.length,
          filterConfigCount: s.filterConfig.length,
          operatorBetweenFilters: s.operatorBetweenFilters,
        },
        verticalSlice: {
          currentVerticalKey: s.currentVerticalKey,
          verticalsCount: s.verticals.length,
          verticalCounts: s.verticalCounts,
        },
        resultSlice: {
          itemCount: s.items.length,
          totalCount: s.totalCount,
          currentPage: s.currentPage,
          pageSize: s.pageSize,
          sort: s.sort,
          promotedResultsCount: s.promotedResults.length,
          isLoading: s.isLoading,
        },
        uiSlice: {
          activeLayoutKey: s.activeLayoutKey,
          availableLayouts: s.availableLayouts,
        },
        registries: {
          dataProviders: s.registries.dataProviders.getAll().map((p) => p.id),
          actions: s.registries.actions.getAll().map((a) => a.id),
          layouts: s.registries.layouts.getAll().map((l) => l.key),
          filterTypes: s.registries.filterTypes.getAll().map((f) => f.key),
          suggestions: s.registries.suggestions.getAll().map((sp) => sp.id),
        },
      };

      // Detect changed top-level slice keys for highlight flash
      const prev = prevSnapshotRef.current;
      const newChanged = new Set<string>();
      for (const key of Object.keys(next)) {
        if (JSON.stringify(prev[key]) !== JSON.stringify(next[key])) {
          newChanged.add(key);
        }
      }
      prevSnapshotRef.current = next;
      setSnapshot(next);

      if (newChanged.size > 0) {
        setChangedKeys(newChanged);
        // Clear highlights after 2 seconds
        setTimeout(() => setChangedKeys(new Set()), 2000);
      }
    }
    update();
    return store.subscribe(update);
  }, [store]);

  return (
    <div>
      {Object.entries(snapshot).map(([key, value]) => (
        <div
          key={key}
          className={`${styles.stateSection}${changedKeys.has(key) ? ' ' + styles.jsonHighlight : ''}`}
        >
          <JsonTree data={{ [key]: value }} defaultExpanded={true} />
        </div>
      ))}
    </div>
  );
};

// ─── Config Tab ──────────────────────────────────────────

const ConfigTab: React.FC = () => {
  const [configs, setConfigs] = React.useState<Array<{ name: string; properties: Record<string, unknown> }>>([]);

  React.useEffect(() => {
    function update(): void {
      const map = DebugCollector.getWebPartConfigs();
      const arr: Array<{ name: string; properties: Record<string, unknown> }> = [];
      map.forEach((config) => {
        arr.push({ name: config.componentName, properties: config.properties });
      });
      setConfigs(arr);
    }
    update();
    return DebugCollector.subscribe(update);
  }, []);

  if (configs.length === 0) {
    return <div style={{ color: '#888', padding: 16 }}>No web parts registered yet.</div>;
  }

  return (
    <div>
      {configs.map((config) => (
        <div key={config.name} className={styles.stateSection}>
          <div className={styles.stateSectionHeader}>{config.name}</div>
          <JsonTree data={config.properties} defaultExpanded={false} />
        </div>
      ))}
    </div>
  );
};

// ─── Log Tab ─────────────────────────────────────────────

const ALL_EVENT_TYPES: DebugEventType[] = ['SEARCH', 'FILTER', 'VERTICAL', 'URL', 'ERROR', 'INIT'];

const LogTab: React.FC = () => {
  const [events, setEvents] = React.useState<IDebugEvent[]>(DebugCollector.getEvents());
  const [enabledTypes, setEnabledTypes] = React.useState<Set<DebugEventType>>(new Set(ALL_EVENT_TYPES));

  React.useEffect(() => {
    return DebugCollector.subscribe(() => {
      setEvents([...DebugCollector.getEvents()]);
    });
  }, []);

  const toggleType = (type: DebugEventType): void => {
    setEnabledTypes((prev) => {
      const next = new Set(prev);
      if (next.has(type)) {
        next.delete(type);
      } else {
        next.add(type);
      }
      return next;
    });
  };

  const filtered = events.filter((e) => enabledTypes.has(e.type));

  return (
    <div>
      <div className={styles.logFilters}>
        {ALL_EVENT_TYPES.map((type) => (
          <label key={type} className={styles.logFilterCheckbox}>
            <input
              type="checkbox"
              checked={enabledTypes.has(type)}
              onChange={() => toggleType(type)}
            />
            <span className={LOG_TYPE_CLASSES[type]}>{type}</span>
          </label>
        ))}
      </div>
      {filtered.map((event) => (
        <div
          key={event.id}
          className={`${styles.logEntry}${event.type === 'ERROR' ? ' ' + styles.logEntryError : ''}`}
        >
          <span className={styles.logTimestamp}>{formatTime(event.timestamp)}</span>
          <span className={`${styles.logType} ${LOG_TYPE_CLASSES[event.type]}`}>{event.type}</span>
          <span className={styles.logData}>{JSON.stringify(event.data)}</span>
        </div>
      ))}
      {filtered.length === 0 && (
        <div style={{ color: '#888', padding: 16 }}>No events matching filter.</div>
      )}
    </div>
  );
};

// ─── Main Panel ──────────────────────────────────────────

type TabKey = 'query' | 'state' | 'config' | 'log';

interface IDebugPanelProps {
  store: StoreApi<ISearchStore>;
  onClose: () => void;
}

const TABS: Array<{ key: TabKey; label: string }> = [
  { key: 'query', label: 'Query' },
  { key: 'state', label: 'State' },
  { key: 'config', label: 'Config' },
  { key: 'log', label: 'Log' },
];

const DEFAULT_HEIGHT = 0.6; // 60% viewport height

const DebugPanel: React.FC<IDebugPanelProps> = ({ store, onClose }) => {
  const [activeTab, setActiveTab] = React.useState<TabKey>('query');
  const [height, setHeight] = React.useState(DEFAULT_HEIGHT);
  const panelRef = React.useRef<HTMLDivElement>(null);
  const dragRef = React.useRef<{ startY: number; startHeight: number } | null>(null);

  // Drag-to-resize handler
  const onDragStart = React.useCallback((e: React.MouseEvent) => {
    e.preventDefault();
    dragRef.current = { startY: e.clientY, startHeight: height };
    const onMove = (me: MouseEvent): void => {
      if (!dragRef.current) return;
      const delta = dragRef.current.startY - me.clientY;
      const newHeight = Math.min(0.9, Math.max(0.2, dragRef.current.startHeight + delta / window.innerHeight));
      setHeight(newHeight);
    };
    const onUp = (): void => {
      dragRef.current = null;
      document.removeEventListener('mousemove', onMove);
      document.removeEventListener('mouseup', onUp);
    };
    document.addEventListener('mousemove', onMove);
    document.addEventListener('mouseup', onUp);
  }, [height]);

  const renderTab = (): React.ReactElement => {
    switch (activeTab) {
      case 'query': return <QueryTab />;
      case 'state': return <StateTab store={store} />;
      case 'config': return <ConfigTab />;
      case 'log': return <LogTab />;
    }
  };

  return (
    <div
      ref={panelRef}
      className={styles.debugPanel}
      style={{ height: `${Math.round(height * 100)}vh` }}
    >
      <div className={styles.dragHandle} onMouseDown={onDragStart} />
      <div className={styles.header}>
        <span className={styles.headerTitle}>SP Search Debug</span>
        <div className={styles.headerActions}>
          <button onClick={onClose} title="Close" type="button">
            <Icon iconName="Cancel" />
          </button>
        </div>
      </div>
      <div className={styles.tabBar}>
        {TABS.map((tab) => (
          <button
            key={tab.key}
            className={`${styles.tab}${activeTab === tab.key ? ' ' + styles.tabActive : ''}`}
            onClick={() => setActiveTab(tab.key)}
            type="button"
          >
            {tab.label}
          </button>
        ))}
      </div>
      <div className={styles.tabContent}>
        {renderTab()}
      </div>
    </div>
  );
};

export default DebugPanel;
```

- [ ] **Step 2: Commit**

```bash
git add src/webparts/spSearchResults/components/DebugPanel.tsx
git commit -m "feat(debug): add DebugPanel component with 4 tabs"
```

---

## Task 11: Integrate DebugFab + DebugPanel into SpSearchResults

**Files:**
- Modify: `src/webparts/spSearchResults/components/SpSearchResults.tsx`

- [ ] **Step 1: Add lazy-loaded DebugPanel and DebugFab imports**

After the existing lazy layout imports at the top of the file, add:

```typescript
import { DebugCollector } from '@store/debug';

// Debug panel — only loaded when ?debug=1
const DebugFab = React.lazy(
  () => import(/* webpackChunkName: 'DebugPanel' */ './DebugFab') as unknown as Promise<{ default: React.ComponentType<Record<string, unknown>> }>
);
const DebugPanel = React.lazy(
  () => import(/* webpackChunkName: 'DebugPanel' */ './DebugPanel') as unknown as Promise<{ default: React.ComponentType<Record<string, unknown>> }>
);
```

- [ ] **Step 2: Add debug state and rendering inside the SpSearchResults component**

Inside the component function, add state for the debug panel:

```typescript
  const isDebugActive = DebugCollector.isActive();
  const [debugOpen, setDebugOpen] = React.useState(false);
```

- [ ] **Step 3: Add DebugFab + DebugPanel to the JSX return**

Just before the closing `</ErrorBoundary>` tag (line ~825), add:

```tsx
        {isDebugActive && (
          <React.Suspense fallback={null}>
            {!debugOpen && (
              <DebugFab onClick={() => setDebugOpen(true)} />
            )}
            {debugOpen && (
              <DebugPanel store={store} onClose={() => setDebugOpen(false)} />
            )}
          </React.Suspense>
        )}
```

Note: `store` is already available in `SpSearchResults` via `props.store` (type `StoreApi<ISearchStore>` from `ISpSearchResultsProps`). The `DebugFab` and `DebugPanel` are lazy-loaded using `React.lazy()` directly (NOT `lazyBridge`), wrapped in `<React.Suspense>`. Since these components are only loaded when `debug=1` is active, they won't affect production bundle size. The `as unknown as Promise<...>` cast follows the existing pattern in the file for @types/react version mismatches.

- [ ] **Step 4: Verify build compiles**

Run: `cd /Users/hemantmane/Development/sp-search && npx tsc --noEmit --skipLibCheck 2>&1 | head -30`

- [ ] **Step 5: Commit**

```bash
git add src/webparts/spSearchResults/components/SpSearchResults.tsx
git commit -m "feat(debug): integrate DebugFab and DebugPanel into SpSearchResults"
```

---

## Task 12: Final Verification

- [ ] **Step 1: Run dev build to verify webpack alias resolution**

```bash
cd /Users/hemantmane/Development/sp-search && gulp bundle 2>&1 | tail -30
```

If `@store/debug` fails to resolve, the `@store/*` alias in `tsconfig.json` maps to `src/libraries/spSearchStore/*`. The webpack aliases in `gulpfile.js` must match. If webpack doesn't have a matching alias, add `@store` to the webpack config in `gulpfile.js` (look for the existing `additionalConfiguration` section that sets up aliases).

- [ ] **Step 2: Run production build**

```bash
cd /Users/hemantmane/Development/sp-search && gulp bundle --ship 2>&1 | tail -20
```

- [ ] **Step 2: Verify debug code is tree-shaken in non-debug mode**

Confirm that `DebugCollector.isActive()` returns false when `?debug=1` is not in the URL, making all calls no-ops. The `DebugPanel` chunk should only appear in the webpack output as a separate chunk file (not inlined into the main bundle).

- [ ] **Step 3: Commit any final fixes**

```bash
git add -A
git commit -m "feat(debug): debug panel implementation complete"
```
