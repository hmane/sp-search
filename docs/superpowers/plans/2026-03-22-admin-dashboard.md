# Admin Search Dashboard Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Replace the Admin Manager's 3 tabs (Coverage/Health/Insights) with a unified Dashboard tab that auto-detects search config from the shared Zustand store.

**Architecture:** A new `CoverageStatsService` runs search queries against the store's config for item count, freshness, file types, and site distribution. An `AdminDashboard` component orchestrates 4 collapsible sections: Coverage Stats, Quality Metrics, Zero-Result Queries (existing), and Top Queries (existing + repeat queries). The dashboard controls a shared time range selector that flows down to child panels.

**Tech Stack:** React 17, TypeScript, PnPjs sp.search(), Fluent UI v8 (Pivot, ChoiceGroup, Icon), existing SCSS patterns

**Spec:** `docs/superpowers/specs/2026-03-22-admin-dashboard-design.md`

---

## File Structure

| Action | Path | Responsibility |
|--------|------|---------------|
| Create | `src/libraries/spSearchStore/services/CoverageStatsService.ts` | Runs coverage search queries (item count, freshness, file types, site distribution) |
| Create | `src/webparts/spSearchManager/components/AdminDashboard.tsx` | Dashboard shell with 4 collapsible sections, time range state, refresh |
| Create | `src/webparts/spSearchManager/components/CoverageStatsSection.tsx` | Coverage stats UI: item count, freshness, file type chart, gap analysis |
| Create | `src/webparts/spSearchManager/components/QualityMetricsSection.tsx` | Stat cards (5) + vertical usage bar chart |
| Modify | `src/webparts/spSearchManager/components/ZeroResultsPanel.tsx` | Accept `daysBack` prop instead of hardcoded 90 |
| Modify | `src/webparts/spSearchManager/components/SearchInsightsPanel.tsx` | Accept `daysBack` + `hideTimeRange` props; add repeat queries table + vertical usage chart |
| Modify | `src/webparts/spSearchManager/components/ISpSearchManagerProps.ts` | Add `'dashboard'` to defaultTab union, add `expectedSiteUrls`, add `enableDashboard` |
| Modify | `src/webparts/spSearchManager/components/SpSearchManager.tsx` | Replace 3 admin tabs with single Dashboard tab |
| Modify | `src/webparts/spSearchManager/SpSearchManagerWebPart.ts` | Replace 3 enable toggles with `enableDashboard`, add `expectedSiteUrls` property pane |
| Modify | `src/webparts/spSearchAdminManager/SpSearchAdminManagerWebPart.manifest.json` | Update defaults |
| Remove | `src/webparts/spSearchManager/components/CoveragePanel.tsx` | Replaced by CoverageStatsSection |

---

## Task 1: CoverageStatsService

**Files:**
- Create: `src/libraries/spSearchStore/services/CoverageStatsService.ts`

- [ ] **Step 1: Create CoverageStatsService**

```typescript
// src/libraries/spSearchStore/services/CoverageStatsService.ts
import 'spfx-toolkit/lib/utilities/context/pnpImports/search';
import type { ISearchQuery as IPnPSearchQuery, SearchResults } from '@pnp/sp/search';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import type { ISearchScope } from '@interfaces/index';

export interface ICoverageConfig {
  queryTemplate: string;
  scope: ISearchScope;
  resultSourceId: string | undefined;
  refinementFilters: string | undefined;
}

export interface ICoverageStatsResult {
  itemCount: number;
  newest: Date | undefined;
  oldest: Date | undefined;
  fileTypes: Array<{ type: string; count: number }>;
  actualSites: Array<{ url: string; count: number }>;
  timestamp: number;
}

export class CoverageStatsService {
  private readonly _config: ICoverageConfig;

  public constructor(config: ICoverageConfig) {
    this._config = config;
  }

  /**
   * Build base search request from config.
   */
  private _baseRequest(): IPnPSearchQuery {
    const req: IPnPSearchQuery = {
      Querytext: '*',
      QueryTemplate: this._config.queryTemplate || '{searchTerms}',
      RowLimit: 1,
      SelectProperties: ['Title', 'LastModifiedTime'],
      TrimDuplicates: false,
      ClientType: 'SPSearchCoverage',
    };

    // Apply scope
    if (this._config.scope && this._config.scope.kqlPath) {
      req.Querytext = '* ' + this._config.scope.kqlPath;
    }

    // Apply result source
    const sourceId = this._config.resultSourceId ||
      (this._config.scope && this._config.scope.resultSourceId) || undefined;
    if (sourceId) {
      req.SourceId = sourceId;
    }

    // Apply persistent refinement filters
    if (this._config.refinementFilters) {
      const filters = this._config.refinementFilters
        .split(',')
        .map(function (f: string): string { return f.trim(); })
        .filter(Boolean);
      if (filters.length > 0) {
        req.RefinementFilters = filters;
      }
    }

    return req;
  }

  public async getItemCount(signal: AbortSignal): Promise<number> {
    if (signal.aborted) throw new DOMException('Aborted', 'AbortError');
    const req = this._baseRequest();
    req.RowLimit = 1;
    req.SelectProperties = ['Title'];
    const results: SearchResults = await SPContext.sp.search(req);
    return results.TotalRows;
  }

  public async getFreshness(signal: AbortSignal): Promise<{ newest: Date | undefined; oldest: Date | undefined }> {
    if (signal.aborted) throw new DOMException('Aborted', 'AbortError');

    const newestReq = this._baseRequest();
    newestReq.RowLimit = 1;
    newestReq.SelectProperties = ['LastModifiedTime'];
    newestReq.SortList = [{ Property: 'LastModifiedTime', Direction: 1 }]; // Descending

    const oldestReq = this._baseRequest();
    oldestReq.RowLimit = 1;
    oldestReq.SelectProperties = ['LastModifiedTime'];
    oldestReq.SortList = [{ Property: 'LastModifiedTime', Direction: 0 }]; // Ascending

    const [newestResults, oldestResults] = await Promise.all([
      SPContext.sp.search(newestReq).catch(function (): undefined { return undefined; }),
      SPContext.sp.search(oldestReq).catch(function (): undefined { return undefined; }),
    ]);

    let newest: Date | undefined;
    let oldest: Date | undefined;

    if (newestResults && newestResults.PrimarySearchResults.length > 0) {
      const val = (newestResults.PrimarySearchResults[0] as Record<string, unknown>).LastModifiedTime;
      if (val) newest = new Date(val as string);
    }
    if (oldestResults && oldestResults.PrimarySearchResults.length > 0) {
      const val = (oldestResults.PrimarySearchResults[0] as Record<string, unknown>).LastModifiedTime;
      if (val) oldest = new Date(val as string);
    }

    // Fallback: if sorted queries returned no results (LastModifiedTime may not be sortable),
    // run an unsorted query and extract LastModifiedTime from the first result
    if (!newest && !oldest) {
      try {
        const fallbackReq = this._baseRequest();
        fallbackReq.RowLimit = 1;
        fallbackReq.SelectProperties = ['LastModifiedTime'];
        const fallbackResults: SearchResults = await SPContext.sp.search(fallbackReq);
        if (fallbackResults.PrimarySearchResults.length > 0) {
          const val = (fallbackResults.PrimarySearchResults[0] as Record<string, unknown>).LastModifiedTime;
          if (val) {
            newest = new Date(val as string);
            // Can't determine oldest without sort — leave undefined
          }
        }
        console.warn('[SP Search] LastModifiedTime may not be sortable — freshness data is limited');
      } catch {
        // Fallback failed — leave dates undefined
      }
    }

    return { newest, oldest };
  }

  public async getFileTypeBreakdown(signal: AbortSignal): Promise<Array<{ type: string; count: number }>> {
    if (signal.aborted) throw new DOMException('Aborted', 'AbortError');
    const req = this._baseRequest();
    req.RowLimit = 1;
    req.SelectProperties = ['Title'];
    req.Refiners = 'FileType';

    const results: SearchResults = await SPContext.sp.search(req);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const raw = results.RawSearchResults as any;
    const refiners = raw?.PrimaryQueryResult?.RefinementResults?.Refiners;
    if (!refiners || !Array.isArray(refiners) || refiners.length === 0) {
      return [];
    }

    const fileTypeRefiner = refiners[0];
    if (!fileTypeRefiner.Entries) return [];

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return fileTypeRefiner.Entries.map(function (e: any): { type: string; count: number } {
      return { type: e.RefinementName || e.RefinementValue || '', count: e.RefinementCount || 0 };
    }).sort(function (a: { count: number }, b: { count: number }): number { return b.count - a.count; });
  }

  public async getSiteDistribution(signal: AbortSignal): Promise<Array<{ url: string; count: number }>> {
    if (signal.aborted) throw new DOMException('Aborted', 'AbortError');
    const req = this._baseRequest();
    req.RowLimit = 1;
    req.SelectProperties = ['Title'];
    req.Refiners = 'SPWebUrl';

    const results: SearchResults = await SPContext.sp.search(req);
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const raw = results.RawSearchResults as any;
    const refiners = raw?.PrimaryQueryResult?.RefinementResults?.Refiners;
    if (!refiners || !Array.isArray(refiners) || refiners.length === 0) {
      return [];
    }

    const siteRefiner = refiners[0];
    if (!siteRefiner.Entries) return [];

    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    return siteRefiner.Entries.map(function (e: any): { url: string; count: number } {
      return { url: e.RefinementName || e.RefinementValue || '', count: e.RefinementCount || 0 };
    }).sort(function (a: { count: number }, b: { count: number }): number { return b.count - a.count; });
  }

  public async runAll(signal: AbortSignal): Promise<ICoverageStatsResult> {
    const [itemCount, freshness, fileTypes, actualSites] = await Promise.all([
      this.getItemCount(signal),
      this.getFreshness(signal),
      this.getFileTypeBreakdown(signal),
      this.getSiteDistribution(signal),
    ]);

    return {
      itemCount,
      newest: freshness.newest,
      oldest: freshness.oldest,
      fileTypes,
      actualSites,
      timestamp: Date.now(),
    };
  }
}
```

- [ ] **Step 2: Export from services index**

Add to `src/libraries/spSearchStore/services/index.ts`:

```typescript
export { CoverageStatsService } from './CoverageStatsService';
export type { ICoverageConfig, ICoverageStatsResult } from './CoverageStatsService';
```

- [ ] **Step 3: Verify build**

Run: `npm run build 2>&1 | tail -5`

- [ ] **Step 4: Commit**

```bash
git add src/libraries/spSearchStore/services/CoverageStatsService.ts src/libraries/spSearchStore/services/index.ts
git commit -m "feat(admin): add CoverageStatsService for search coverage analysis"
```

---

## Task 2: Update ZeroResultsPanel to accept daysBack prop

**Files:**
- Modify: `src/webparts/spSearchManager/components/ZeroResultsPanel.tsx`

- [ ] **Step 1: Add daysBack to props interface**

Change the props interface from:

```typescript
export interface IZeroResultsPanelProps {
  service: SearchManagerService;
  onRunQuery: (queryText: string, vertical: string) => void;
}
```

To:

```typescript
export interface IZeroResultsPanelProps {
  service: SearchManagerService;
  onRunQuery: (queryText: string, vertical: string) => void;
  daysBack?: number;
}
```

- [ ] **Step 2: Replace hardcoded DAYS_BACK with prop**

Change line 87 from:

```typescript
const DAYS_BACK = 90;
```

To:

```typescript
const DAYS_BACK = props.daysBack || 90;
```

And update the `load` dependency array to include `DAYS_BACK`. Since `DAYS_BACK` is now derived from props, restructure:

Replace the `const DAYS_BACK = 90;` and the `load` useCallback to use `props.daysBack`:

```typescript
  const effectiveDaysBack = props.daysBack || 90;

  const load = React.useCallback(function (): void {
    setIsLoading(true);
    setError(undefined);
    service.loadZeroResultQueries(effectiveDaysBack, 200)
      .then(function (entries: ISearchHistoryEntry[]): void {
        // ... existing logic
      })
      .catch(function (err: unknown): void {
        // ... existing logic
      });
  }, [service, effectiveDaysBack]);
```

Also add a `useEffect` to reload when `effectiveDaysBack` changes:

```typescript
  React.useEffect(function (): void {
    load();
  }, [load]);
```

If there's already a useEffect that calls `load()` on mount, replace it with this one (which also triggers on daysBack change).

- [ ] **Step 3: Commit**

```bash
git add src/webparts/spSearchManager/components/ZeroResultsPanel.tsx
git commit -m "feat(admin): make ZeroResultsPanel daysBack configurable via prop"
```

---

## Task 3: Update SearchInsightsPanel to accept daysBack + hideTimeRange props, add repeat queries + vertical usage

**Files:**
- Modify: `src/webparts/spSearchManager/components/SearchInsightsPanel.tsx`

- [ ] **Step 1: Update props interface**

```typescript
export interface ISearchInsightsPanelProps {
  service: SearchManagerService;
  onRunQuery: (queryText: string, vertical: string) => void;
  daysBack?: number;
  hideTimeRange?: boolean;
}
```

- [ ] **Step 2: Use external daysBack when provided**

In the component, change the `daysBack` state initialization:

```typescript
const [daysBack, setDaysBack] = React.useState<number>(props.daysBack || 30);
```

Add a `useEffect` to sync when the prop changes. Use a ref for the previous value to avoid stale closure issues:

```typescript
const prevDaysBackRef = React.useRef<number | undefined>(props.daysBack);
React.useEffect(function (): void {
  if (props.daysBack !== undefined && props.daysBack !== prevDaysBackRef.current) {
    prevDaysBackRef.current = props.daysBack;
    setDaysBack(props.daysBack);
    load(props.daysBack);
  }
}, [props.daysBack, load]);
```

- [ ] **Step 3: Conditionally hide the built-in ChoiceGroup**

Wrap the existing ChoiceGroup JSX with:

```typescript
{!props.hideTimeRange && (
  <ChoiceGroup ... />
)}
```

- [ ] **Step 4: Add repeat queries and vertical usage to IInsightMetrics**

Extend the metrics interface:

```typescript
interface IInsightMetrics {
  // ... existing fields ...
  repeatQueries: Array<{ queryText: string; totalSearches: number; lastSeen: Date }>;
  verticalUsage: Array<{ vertical: string; count: number }>;
}
```

- [ ] **Step 5: Add computation logic in computeMetrics**

Inside the `computeMetrics` function, add after existing computations:

```typescript
  // Repeat queries: entries where UseCount > 2
  const repeatMap = new Map<string, { totalSearches: number; lastSeen: Date }>();
  for (let i = 0; i < entries.length; i++) {
    const e = entries[i];
    if (e.useCount > 2) {
      const key = e.queryText.toLowerCase();
      const existing = repeatMap.get(key);
      if (existing) {
        existing.totalSearches += e.useCount;
        if (e.searchTimestamp > existing.lastSeen) {
          existing.lastSeen = e.searchTimestamp;
        }
      } else {
        repeatMap.set(key, { totalSearches: e.useCount, lastSeen: e.searchTimestamp });
      }
    }
  }
  const repeatQueries: Array<{ queryText: string; totalSearches: number; lastSeen: Date }> = [];
  repeatMap.forEach(function (val, key): void {
    repeatQueries.push({ queryText: key, totalSearches: val.totalSearches, lastSeen: val.lastSeen });
  });
  repeatQueries.sort(function (a, b): number { return b.totalSearches - a.totalSearches; });

  // Vertical usage
  const verticalMap = new Map<string, number>();
  for (let i = 0; i < entries.length; i++) {
    const v = entries[i].vertical || 'All';
    verticalMap.set(v, (verticalMap.get(v) || 0) + 1);
  }
  const verticalUsage: Array<{ vertical: string; count: number }> = [];
  verticalMap.forEach(function (count, vertical): void {
    verticalUsage.push({ vertical, count });
  });
  verticalUsage.sort(function (a, b): number { return b.count - a.count; });
```

Add `repeatQueries` and `verticalUsage` to the returned metrics object.

- [ ] **Step 6: Render repeat queries table and vertical usage chart**

After the existing top queries and volume chart sections, add:

**Vertical Usage:**
```tsx
{metrics.verticalUsage.length > 0 && (
  <div className={styles.insightSection}>
    <h3 className={styles.insightSectionTitle}>Vertical Usage</h3>
    <div className={styles.insightBarList}>
      {metrics.verticalUsage.map(function (v) {
        const maxCount = metrics.verticalUsage[0].count;
        return (
          <div key={v.vertical} className={styles.insightBarRow}>
            <span className={styles.insightBarLabel}>{v.vertical}</span>
            <div className={styles.insightBarTrack}>
              <div className={styles.insightBarFill} style={{ width: (v.count / maxCount * 100) + '%' }} />
            </div>
            <span className={styles.insightBarCount}>{v.count}</span>
          </div>
        );
      })}
    </div>
  </div>
)}
```

**Repeat Queries** (only show if data exists):
```tsx
{metrics.repeatQueries.length > 0 && (
  <div className={styles.insightSection}>
    <h3 className={styles.insightSectionTitle}>Repeat Queries</h3>
    <p className={styles.insightNote}>Queries searched 3+ times by the same user (candidates for promoted results)</p>
    <div className={styles.insightBarList}>
      {metrics.repeatQueries.slice(0, 10).map(function (rq) {
        const maxCount = metrics.repeatQueries[0].totalSearches;
        return (
          <div key={rq.queryText} className={styles.insightBarRow + ' ' + styles.insightBarRowClickable}
               onClick={function () { props.onRunQuery(rq.queryText, ''); }}>
            <span className={styles.insightBarLabel}>{rq.queryText}</span>
            <div className={styles.insightBarTrack}>
              <div className={styles.insightBarFill} style={{ width: (rq.totalSearches / maxCount * 100) + '%' }} />
            </div>
            <span className={styles.insightBarCount}>{rq.totalSearches}</span>
          </div>
        );
      })}
    </div>
  </div>
)}
```

- [ ] **Step 7: Commit**

```bash
git add src/webparts/spSearchManager/components/SearchInsightsPanel.tsx
git commit -m "feat(admin): add daysBack/hideTimeRange props, repeat queries, vertical usage to InsightsPanel"
```

---

## Task 4: CoverageStatsSection component

**Files:**
- Create: `src/webparts/spSearchManager/components/CoverageStatsSection.tsx`

- [ ] **Step 1: Create component**

```tsx
// src/webparts/spSearchManager/components/CoverageStatsSection.tsx
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import type { ICoverageStatsResult } from '@services/index';
import styles from './SpSearchManager.module.scss';

export interface ICoverageStatsSectionProps {
  coverage: ICoverageStatsResult | undefined;
  expectedSiteUrls: string[];
  isLoading: boolean;
  error: string | undefined;
}

function formatRelativeDate(date: Date): string {
  const now = Date.now();
  const diff = now - date.getTime();
  const hours = Math.floor(diff / 3600000);
  if (hours < 1) return 'Less than an hour ago';
  if (hours < 24) return hours + ' hours ago';
  const days = Math.floor(hours / 24);
  if (days < 7) return days + ' days ago';
  if (days < 30) return Math.floor(days / 7) + ' weeks ago';
  if (days < 365) return Math.floor(days / 30) + ' months ago';
  return Math.floor(days / 365) + ' years ago';
}

function freshnessColor(date: Date | undefined): string {
  if (!date) return '#808080';
  const hours = (Date.now() - date.getTime()) / 3600000;
  if (hours < 24) return '#50c878';   // green
  if (hours < 168) return '#ffc832';  // yellow (7 days)
  return '#ff5050';                    // red
}

const CoverageStatsSection: React.FC<ICoverageStatsSectionProps> = (props) => {
  const { coverage, expectedSiteUrls, isLoading, error } = props;

  if (isLoading) {
    return <Spinner size={SpinnerSize.medium} label="Loading coverage data..." />;
  }

  if (error) {
    return <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;
  }

  if (!coverage) {
    return <div style={{ color: '#888', padding: 12 }}>No coverage data available.</div>;
  }

  // Gap analysis
  const actualSiteUrls = new Set(coverage.actualSites.map(function (s) { return s.url.toLowerCase(); }));
  const gapAnalysis = expectedSiteUrls.map(function (url) {
    return {
      url: url,
      found: actualSiteUrls.has(url.toLowerCase()),
    };
  });

  return (
    <div>
      {/* Item Count + Freshness stat cards */}
      <div className={styles.insightStatCards}>
        <div className={styles.insightStatCard}>
          <div className={styles.insightStatValue}>{coverage.itemCount.toLocaleString()}</div>
          <div className={styles.insightStatLabel}>
            <Icon iconName="NumberField" /> Indexed Items
          </div>
        </div>
        <div className={styles.insightStatCard}>
          <div className={styles.insightStatValue}>
            <span style={{ color: freshnessColor(coverage.newest) }}>{'\u25CF'} </span>
            {coverage.newest ? formatRelativeDate(coverage.newest) : 'Unknown'}
          </div>
          <div className={styles.insightStatLabel}>
            <Icon iconName="Recent" /> Newest Item
          </div>
        </div>
        <div className={styles.insightStatCard}>
          <div className={styles.insightStatValue}>
            {coverage.oldest ? formatRelativeDate(coverage.oldest) : 'Unknown'}
          </div>
          <div className={styles.insightStatLabel}>
            <Icon iconName="History" /> Oldest Item
          </div>
        </div>
        <div className={styles.insightStatCard}>
          <div className={styles.insightStatValue}>{coverage.fileTypes.length}</div>
          <div className={styles.insightStatLabel}>
            <Icon iconName="Page" /> File Types
          </div>
        </div>
      </div>

      {/* File Type Breakdown */}
      {coverage.fileTypes.length > 0 && (
        <div style={{ marginBottom: 20 }}>
          <h3 className={styles.insightSectionTitle}>File Type Breakdown</h3>
          <div className={styles.insightBarList}>
            {coverage.fileTypes.slice(0, 10).map(function (ft) {
              const maxCount = coverage.fileTypes[0].count;
              return (
                <div key={ft.type} className={styles.insightBarRow}>
                  <span className={styles.insightBarLabel}>{ft.type || '(none)'}</span>
                  <div className={styles.insightBarTrack}>
                    <div className={styles.insightBarFill} style={{ width: (ft.count / maxCount * 100) + '%' }} />
                  </div>
                  <span className={styles.insightBarCount}>{ft.count.toLocaleString()}</span>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* Gap Analysis */}
      {expectedSiteUrls.length > 0 ? (
        <div>
          <h3 className={styles.insightSectionTitle}>Content Gap Analysis</h3>
          <div className={styles.insightBarList}>
            {gapAnalysis.map(function (site) {
              return (
                <div key={site.url} className={styles.insightBarRow}>
                  <span className={styles.insightBarLabel} title={site.url}>
                    {site.url.replace(/^https?:\/\/[^/]+/, '')}
                  </span>
                  <span style={{
                    color: site.found ? '#50c878' : '#ff5050',
                    fontWeight: 600,
                    fontSize: 12,
                  }}>
                    {site.found ? 'Found' : 'Missing'}
                  </span>
                </div>
              );
            })}
          </div>
        </div>
      ) : (
        <div style={{ color: '#888', fontSize: 12, padding: '8px 0' }}>
          <Icon iconName="Info" /> No expected sites configured. Add site URLs in the property pane to enable gap analysis.
        </div>
      )}
    </div>
  );
};

export default CoverageStatsSection;
```

- [ ] **Step 2: Commit**

```bash
git add src/webparts/spSearchManager/components/CoverageStatsSection.tsx
git commit -m "feat(admin): add CoverageStatsSection component"
```

---

## Task 5: QualityMetricsSection component

**Files:**
- Create: `src/webparts/spSearchManager/components/QualityMetricsSection.tsx`

- [ ] **Step 1: Create component**

This component renders the 5 stat cards + the vertical usage bar chart. It receives pre-computed metrics from the parent dashboard.

```tsx
// src/webparts/spSearchManager/components/QualityMetricsSection.tsx
import * as React from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import styles from './SpSearchManager.module.scss';

export interface IQualityMetrics {
  totalSearches: number;
  zeroResultRate: number;
  clickThroughRate: number;
  repeatQueryRate: number;
  hasUseCountField: boolean;
  topVertical: string;
  verticalUsage: Array<{ vertical: string; count: number }>;
}

export interface IQualityMetricsSectionProps {
  metrics: IQualityMetrics | undefined;
  isLoading: boolean;
  error: string | undefined;
  samplingNote: string;
}

const QualityMetricsSection: React.FC<IQualityMetricsSectionProps> = (props) => {
  const { metrics, isLoading, error, samplingNote } = props;

  if (isLoading) {
    return <Spinner size={SpinnerSize.medium} label="Loading quality metrics..." />;
  }

  if (error) {
    return <MessageBar messageBarType={MessageBarType.error}>{error}</MessageBar>;
  }

  if (!metrics) {
    return <div style={{ color: '#888', padding: 12 }}>No search history data available.</div>;
  }

  return (
    <div>
      {/* Sampling note */}
      <div style={{ color: '#888', fontSize: 11, marginBottom: 8 }}>{samplingNote}</div>

      {/* Stat Cards */}
      <div className={styles.insightStatCards}>
        <div className={styles.insightStatCard}>
          <div className={styles.insightStatValue}>{metrics.totalSearches.toLocaleString()}</div>
          <div className={styles.insightStatLabel}>
            <Icon iconName="Search" /> Total Searches
          </div>
        </div>
        <div className={`${styles.insightStatCard}${metrics.zeroResultRate > 20 ? ' ' + styles.insightStatCardWarning : ''}`}>
          <div className={styles.insightStatValue}>{metrics.zeroResultRate.toFixed(1)}%</div>
          <div className={styles.insightStatLabel}>
            <Icon iconName="SearchIssue" /> Zero-Result Rate
          </div>
        </div>
        <div className={`${styles.insightStatCard}${metrics.clickThroughRate < 30 ? ' ' + styles.insightStatCardWarning : ''}`}>
          <div className={styles.insightStatValue}>{metrics.clickThroughRate.toFixed(1)}%</div>
          <div className={styles.insightStatLabel}>
            <Icon iconName="TouchPointer" /> Click-Through Rate
          </div>
        </div>
        <div className={`${styles.insightStatCard}${metrics.hasUseCountField && metrics.repeatQueryRate > 40 ? ' ' + styles.insightStatCardWarning : ''}`}>
          <div className={styles.insightStatValue}>
            {metrics.hasUseCountField ? metrics.repeatQueryRate.toFixed(1) + '%' : 'N/A'}
          </div>
          <div className={styles.insightStatLabel}>
            <Icon iconName="Refresh" /> Repeat Query Rate
            {!metrics.hasUseCountField && (
              <span title="UseCount field not available on SearchHistory list" style={{ cursor: 'help' }}> (?)</span>
            )}
          </div>
        </div>
        <div className={styles.insightStatCard}>
          <div className={styles.insightStatValue}>{metrics.topVertical || 'All'}</div>
          <div className={styles.insightStatLabel}>
            <Icon iconName="ViewAll" /> Top Vertical
          </div>
        </div>
      </div>

      {/* Vertical Usage */}
      {metrics.verticalUsage.length > 1 && (
        <div style={{ marginTop: 12 }}>
          <h3 className={styles.insightSectionTitle}>Vertical Usage</h3>
          <div className={styles.insightBarList}>
            {metrics.verticalUsage.map(function (v) {
              const maxCount = metrics.verticalUsage[0].count;
              return (
                <div key={v.vertical} className={styles.insightBarRow}>
                  <span className={styles.insightBarLabel}>{v.vertical}</span>
                  <div className={styles.insightBarTrack}>
                    <div className={styles.insightBarFill} style={{ width: (v.count / maxCount * 100) + '%' }} />
                  </div>
                  <span className={styles.insightBarCount}>{v.count}</span>
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
};

export default QualityMetricsSection;
```

- [ ] **Step 2: Commit**

```bash
git add src/webparts/spSearchManager/components/QualityMetricsSection.tsx
git commit -m "feat(admin): add QualityMetricsSection component"
```

---

## Task 6: AdminDashboard component

**Files:**
- Create: `src/webparts/spSearchManager/components/AdminDashboard.tsx`

- [ ] **Step 1: Create the dashboard shell**

This component orchestrates all 4 sections, manages time range state, and runs coverage queries via CoverageStatsService.

```tsx
// src/webparts/spSearchManager/components/AdminDashboard.tsx
import * as React from 'react';
import { IconButton } from '@fluentui/react/lib/Button';
import { ChoiceGroup, IChoiceGroupOption } from '@fluentui/react/lib/ChoiceGroup';
import { Icon } from '@fluentui/react/lib/Icon';
import type { StoreApi } from 'zustand/vanilla';
import type { ISearchStore } from '@interfaces/index';
import { CoverageStatsService } from '@services/index';
import type { ICoverageStatsResult } from '@services/index';
import { SearchManagerService } from '@services/index';
import type { ISearchHistoryEntry } from '@interfaces/index';
import CoverageStatsSection from './CoverageStatsSection';
import QualityMetricsSection from './QualityMetricsSection';
import type { IQualityMetrics } from './QualityMetricsSection';
import ZeroResultsPanel from './ZeroResultsPanel';
import SearchInsightsPanel from './SearchInsightsPanel';
import styles from './SpSearchManager.module.scss';

export interface IAdminDashboardProps {
  store: StoreApi<ISearchStore>;
  service: SearchManagerService;
  expectedSiteUrls: string[];
  onRunQuery: (queryText: string, vertical: string) => void;
}

const RANGE_OPTIONS: IChoiceGroupOption[] = [
  { key: '30', text: '30d' },
  { key: '60', text: '60d' },
  { key: '90', text: '90d' },
];

function computeQualityMetrics(entries: ISearchHistoryEntry[]): IQualityMetrics {
  const total = entries.length;
  if (total === 0) {
    return {
      totalSearches: 0,
      zeroResultRate: 0,
      clickThroughRate: 0,
      repeatQueryRate: 0,
      hasUseCountField: false,
      topVertical: 'All',
      verticalUsage: [],
    };
  }

  let zeroCount = 0;
  let clickedCount = 0;
  let repeatCount = 0;
  let hasUseCount = false;
  const verticalMap = new Map<string, number>();

  for (let i = 0; i < entries.length; i++) {
    const e = entries[i];
    if (e.isZeroResult) zeroCount++;
    if (e.clickedItems && e.clickedItems.length > 0) clickedCount++;
    if (e.useCount > 1) {
      repeatCount++;
      hasUseCount = true;
    }
    if (e.useCount === 1) {
      // Could still have UseCount field — just value is 1
      // We detect field availability by checking if ANY entry has useCount > 1
    }
    const v = e.vertical || 'All';
    verticalMap.set(v, (verticalMap.get(v) || 0) + 1);
  }

  const verticalUsage: Array<{ vertical: string; count: number }> = [];
  verticalMap.forEach(function (count, vertical): void {
    verticalUsage.push({ vertical, count });
  });
  verticalUsage.sort(function (a, b): number { return b.count - a.count; });

  return {
    totalSearches: total,
    zeroResultRate: (zeroCount / total) * 100,
    clickThroughRate: (clickedCount / total) * 100,
    repeatQueryRate: hasUseCount ? (repeatCount / total) * 100 : 0,
    hasUseCountField: hasUseCount,
    topVertical: verticalUsage.length > 0 ? verticalUsage[0].vertical : 'All',
    verticalUsage,
  };
}

const AdminDashboard: React.FC<IAdminDashboardProps> = (props) => {
  const { store, service, expectedSiteUrls, onRunQuery } = props;

  // Time range state
  const [daysBack, setDaysBack] = React.useState<number>(30);

  // Coverage state
  const [coverage, setCoverage] = React.useState<ICoverageStatsResult | undefined>(undefined);
  const [coverageLoading, setCoverageLoading] = React.useState(true);
  const [coverageError, setCoverageError] = React.useState<string | undefined>(undefined);

  // Quality state
  const [qualityMetrics, setQualityMetrics] = React.useState<IQualityMetrics | undefined>(undefined);
  const [qualityLoading, setQualityLoading] = React.useState(true);
  const [qualityError, setQualityError] = React.useState<string | undefined>(undefined);

  // Collapsible sections
  const [coverageExpanded, setCoverageExpanded] = React.useState(true);
  const [qualityExpanded, setQualityExpanded] = React.useState(true);
  const [zeroResultExpanded, setZeroResultExpanded] = React.useState(true);
  const [insightsExpanded, setInsightsExpanded] = React.useState(true);

  // AbortController ref
  const abortRef = React.useRef<AbortController | undefined>(undefined);

  // Load coverage data
  const loadCoverage = React.useCallback(function (): void {
    if (abortRef.current) abortRef.current.abort();
    const controller = new AbortController();
    abortRef.current = controller;

    setCoverageLoading(true);
    setCoverageError(undefined);

    const state = store.getState();
    const coverageService = new CoverageStatsService({
      queryTemplate: state.queryTemplate || '{searchTerms}',
      scope: state.scope,
      resultSourceId: state.resultSourceId || undefined,
      refinementFilters: state.refinementFilters || undefined,
    });

    coverageService.runAll(controller.signal)
      .then(function (result: ICoverageStatsResult): void {
        setCoverage(result);
        setCoverageLoading(false);
      })
      .catch(function (err: unknown): void {
        if (err instanceof DOMException && err.name === 'AbortError') return;
        setCoverageError(err instanceof Error ? err.message : 'Failed to load coverage data');
        setCoverageLoading(false);
      });
  }, [store]);

  // Load quality data
  const loadQuality = React.useCallback(function (days: number): void {
    setQualityLoading(true);
    setQualityError(undefined);

    service.loadAllHistoryForInsights(days, 500)
      .then(function (entries: ISearchHistoryEntry[]): void {
        setQualityMetrics(computeQualityMetrics(entries));
        setQualityLoading(false);
      })
      .catch(function (err: unknown): void {
        setQualityError(err instanceof Error ? err.message : 'Failed to load search history');
        setQualityLoading(false);
      });
  }, [service]);

  // Initial load
  React.useEffect(function (): void {
    loadCoverage();
    loadQuality(daysBack);
    return function (): void {
      if (abortRef.current) abortRef.current.abort();
    };
  }, []); // eslint-disable-line react-hooks/exhaustive-deps

  // Handle time range change
  const handleRangeChange = React.useCallback(function (_: unknown, option?: IChoiceGroupOption): void {
    if (!option) return;
    const days = parseInt(option.key, 10);
    setDaysBack(days);
    loadQuality(days);
  }, [loadQuality]);

  // Refresh all
  const handleRefresh = React.useCallback(function (): void {
    loadCoverage();
    loadQuality(daysBack);
  }, [loadCoverage, loadQuality, daysBack]);

  return (
    <div className={styles.healthPanel}>
      {/* Header toolbar */}
      <div className={styles.healthToolbar}>
        <span style={{ fontWeight: 600 }}>Admin Dashboard</span>
        <div style={{ display: 'flex', alignItems: 'center', gap: 12 }}>
          <ChoiceGroup
            options={RANGE_OPTIONS}
            selectedKey={String(daysBack)}
            onChange={handleRangeChange}
            styles={{ flexContainer: { display: 'flex', gap: 8 } }}
          />
          <IconButton
            iconProps={{ iconName: 'Refresh' }}
            title="Refresh all"
            onClick={handleRefresh}
          />
        </div>
      </div>

      {/* Section 1: Coverage Stats */}
      <div style={{ marginBottom: 20 }}>
        <button
          type="button"
          className={styles.insightSectionTitle}
          onClick={function (): void { setCoverageExpanded(!coverageExpanded); }}
          style={{ background: 'none', border: 'none', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6, padding: 0, width: '100%', textAlign: 'left' }}
        >
          <Icon iconName={coverageExpanded ? 'ChevronDown' : 'ChevronRight'} />
          Content Coverage
        </button>
        {coverageExpanded && (
          <CoverageStatsSection
            coverage={coverage}
            expectedSiteUrls={expectedSiteUrls}
            isLoading={coverageLoading}
            error={coverageError}
          />
        )}
      </div>

      {/* Section 2: Quality Metrics */}
      <div style={{ marginBottom: 20 }}>
        <button
          type="button"
          className={styles.insightSectionTitle}
          onClick={function (): void { setQualityExpanded(!qualityExpanded); }}
          style={{ background: 'none', border: 'none', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6, padding: 0, width: '100%', textAlign: 'left' }}
        >
          <Icon iconName={qualityExpanded ? 'ChevronDown' : 'ChevronRight'} />
          Search Quality
        </button>
        {qualityExpanded && (
          <QualityMetricsSection
            metrics={qualityMetrics}
            isLoading={qualityLoading}
            error={qualityError}
            samplingNote={'Based on last 500 searches'}
          />
        )}
      </div>

      {/* Section 3: Zero-Result Queries */}
      <div style={{ marginBottom: 20 }}>
        <button
          type="button"
          className={styles.insightSectionTitle}
          onClick={function (): void { setZeroResultExpanded(!zeroResultExpanded); }}
          style={{ background: 'none', border: 'none', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6, padding: 0, width: '100%', textAlign: 'left' }}
        >
          <Icon iconName={zeroResultExpanded ? 'ChevronDown' : 'ChevronRight'} />
          Zero-Result Queries
        </button>
        {zeroResultExpanded && (
          <ZeroResultsPanel
            service={service}
            onRunQuery={onRunQuery}
            daysBack={daysBack}
          />
        )}
      </div>

      {/* Section 4: Top Queries & Engagement */}
      <div style={{ marginBottom: 20 }}>
        <button
          type="button"
          className={styles.insightSectionTitle}
          onClick={function (): void { setInsightsExpanded(!insightsExpanded); }}
          style={{ background: 'none', border: 'none', cursor: 'pointer', display: 'flex', alignItems: 'center', gap: 6, padding: 0, width: '100%', textAlign: 'left' }}
        >
          <Icon iconName={insightsExpanded ? 'ChevronDown' : 'ChevronRight'} />
          Top Queries & Engagement
        </button>
        {insightsExpanded && (
          <SearchInsightsPanel
            service={service}
            onRunQuery={onRunQuery}
            daysBack={daysBack}
            hideTimeRange={true}
          />
        )}
      </div>
    </div>
  );
};

export default AdminDashboard;
```

- [ ] **Step 2: Commit**

```bash
git add src/webparts/spSearchManager/components/AdminDashboard.tsx
git commit -m "feat(admin): add AdminDashboard orchestrator component"
```

---

## Task 7: Update props interface and SpSearchManager tab routing

**Files:**
- Modify: `src/webparts/spSearchManager/components/ISpSearchManagerProps.ts`
- Modify: `src/webparts/spSearchManager/components/SpSearchManager.tsx`

- [ ] **Step 1: Update ISpSearchManagerProps**

Add `'dashboard'` to defaultTab union. Replace coverage/health/insights toggles with `enableDashboard`. Add `expectedSiteUrls`.

Change `defaultTab` type from:
```typescript
defaultTab?: 'saved' | 'history' | 'collections' | 'coverage' | 'health' | 'insights';
```
To:
```typescript
defaultTab?: 'saved' | 'history' | 'collections' | 'coverage' | 'health' | 'insights' | 'dashboard';
```

Add new props:
```typescript
  enableDashboard?: boolean;
  expectedSiteUrls?: string[];
```

Keep `enableCoverage`, `enableHealth`, `enableInsights` for backward compatibility (they'll be ignored when `enableDashboard` is true).

- [ ] **Step 2: Update SpSearchManager tab routing**

In `SpSearchManager.tsx`:

1. Update the `SearchManagerTabKey` type alias to include `'dashboard'`.

2. In the admin variant config block (around line 193-214), when `variant === 'admin'` and `enableDashboard` is true, set `defaultTab` to `'dashboard'` and disable the old individual tabs.

3. In the `availableTabs` computation, add dashboard as an available tab when `config.enableDashboard` is true:
```typescript
if (config.enableDashboard) {
  tabs.push({ key: 'dashboard', label: 'Dashboard', icon: 'ViewDashboard' });
}
```

4. In the Pivot tab rendering section, add the dashboard case:
```tsx
{selectedTabKey === 'dashboard' && (
  <AdminDashboard
    store={props.store}
    service={props.service}
    expectedSiteUrls={props.expectedSiteUrls || []}
    onRunQuery={handleRunZeroResultQuery}
  />
)}
```

Import `AdminDashboard` at the top (lazy-loaded):
```typescript
const AdminDashboard = React.lazy(
  () => import(/* webpackChunkName: 'AdminDashboard' */ './AdminDashboard') as unknown as Promise<{ default: React.ComponentType<Record<string, unknown>> }>
);
```

Wrap the dashboard rendering in `<React.Suspense fallback={<Spinner />}>`.

- [ ] **Step 3: Commit**

```bash
git add src/webparts/spSearchManager/components/ISpSearchManagerProps.ts src/webparts/spSearchManager/components/SpSearchManager.tsx
git commit -m "feat(admin): route Dashboard tab in SpSearchManager"
```

---

## Task 8: Update SpSearchManagerWebPart property pane + manifest

**Files:**
- Modify: `src/webparts/spSearchManager/SpSearchManagerWebPart.ts`
- Modify: `src/webparts/spSearchAdminManager/SpSearchAdminManagerWebPart.manifest.json`

- [ ] **Step 1: Add enableDashboard and expectedSiteUrls to web part props interface**

Add to `ISpSearchManagerWebPartProps`:
```typescript
  enableDashboard: boolean;
  expectedSiteUrls: string[];
```

- [ ] **Step 2: Update property pane**

In the Sections property pane group, add a toggle for `enableDashboard` above the existing coverage/health/insights toggles:

```typescript
PropertyPaneToggle('enableDashboard', {
  label: 'Enable Admin Dashboard',
  checked: true,
}),
```

Add a new property pane group for expected site URLs:

```typescript
{
  groupName: 'Gap Analysis',
  groupFields: [
    PropertyPaneTextField('expectedSiteUrls', {
      label: 'Expected Site URLs (one per line)',
      multiline: true,
      rows: 5,
      description: 'Enter site URLs to monitor for content coverage. One URL per line.',
    }),
  ],
}
```

Note: `expectedSiteUrls` is stored as a newline-separated string in the property pane, then split into an array when passed to the component:

```typescript
expectedSiteUrls: (this.properties.expectedSiteUrls || '')
  .split('\n')
  .map(function (s: string): string { return s.trim(); })
  .filter(Boolean),
```

- [ ] **Step 3: Pass new props to SpSearchManager component**

In the render method, add to the props:
```typescript
enableDashboard: this.properties.enableDashboard,
expectedSiteUrls: (this.properties.expectedSiteUrls || '').split('\n').map(function (s: string) { return s.trim(); }).filter(Boolean),
```

- [ ] **Step 4: Update admin manifest defaults**

In `SpSearchAdminManagerWebPart.manifest.json`, update the preconfigured properties:

Change:
```json
"defaultTab": "coverage",
"enableCoverage": true,
```

To:
```json
"defaultTab": "dashboard",
"enableDashboard": true,
"enableCoverage": false,
```

Add:
```json
"expectedSiteUrls": "",
```

- [ ] **Step 5: Verify build**

Run: `npm run build 2>&1 | tail -5`

- [ ] **Step 6: Commit**

```bash
git add src/webparts/spSearchManager/SpSearchManagerWebPart.ts src/webparts/spSearchAdminManager/SpSearchAdminManagerWebPart.manifest.json
git commit -m "feat(admin): add enableDashboard toggle and expectedSiteUrls property pane"
```

---

## Task 9: Remove CoveragePanel

**Files:**
- Remove: `src/webparts/spSearchManager/components/CoveragePanel.tsx`

- [ ] **Step 1: Remove CoveragePanel import from SpSearchManager.tsx**

Find and remove the import of `CoveragePanel` and the `<CoveragePanel>` JSX from the coverage tab case. The coverage tab can remain as a dead branch (the `enableCoverage` toggle defaults to false) or be fully removed from the tab routing.

- [ ] **Step 2: Delete CoveragePanel.tsx**

```bash
rm src/webparts/spSearchManager/components/CoveragePanel.tsx
```

- [ ] **Step 3: Verify build**

Run: `npm run build 2>&1 | tail -5`

- [ ] **Step 4: Commit**

```bash
git add src/webparts/spSearchManager/components/SpSearchManager.tsx
git rm src/webparts/spSearchManager/components/CoveragePanel.tsx
git commit -m "refactor(admin): remove CoveragePanel (replaced by AdminDashboard)"
```

---

## Task 10: Final verification

- [ ] **Step 1: Run full build**

```bash
npm run build 2>&1 | tail -10
```

- [ ] **Step 2: Run production build**

```bash
npm run build:ship 2>&1 | tail -10
```

- [ ] **Step 3: Verify AdminDashboard chunk exists**

Check that webpack output includes an `AdminDashboard` chunk (code-split via React.lazy).

- [ ] **Step 4: Commit any final fixes**

```bash
git add -A
git commit -m "feat(admin): admin dashboard implementation complete"
```
