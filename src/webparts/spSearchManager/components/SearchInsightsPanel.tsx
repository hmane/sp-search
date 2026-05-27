import * as React from 'react';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { ChoiceGroup, IChoiceGroupOption } from '@fluentui/react/lib/ChoiceGroup';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { ISearchHistoryEntry } from '@interfaces/index';
import { SearchManagerService } from '@services/index';
import styles from './SpSearchManager.module.scss';

export interface ISearchInsightsPanelProps {
  service: SearchManagerService;
  /** Called when user clicks a top query to run it */
  onRunQuery: (queryText: string, vertical: string) => void;
  /** Override the default 30-day lookback window */
  daysBack?: number;
  /** Hide the built-in time range ChoiceGroup (when parent controls daysBack) */
  hideTimeRange?: boolean;
}

// ─── Types ────────────────────────────────────────────────────────────────────

interface IInsightMetrics {
  totalSearches: number;
  zeroResultCount: number;
  zeroResultRate: number;
  clickedSearchCount: number;
  clickThroughRate: number;
  avgResultCount: number;
  topQueries: Array<{ queryText: string; count: number; vertical: string }>;
  topClickedItems: Array<{ url: string; title: string; clicks: number }>;
  volumeByDay: Array<{ dateLabel: string; count: number }>;
  // T4.D7 / UX-007 — per-day CTR for the 7-day rolling sparkline. Each
  // entry's `ctr` is `clicks/total` for that calendar day, computed once
  // at aggregation time. `dateLabel` matches the format used by
  // `volumeByDay` so the two surfaces line up visually.
  ctrByDay: Array<{ dateLabel: string; ctr: number; total: number; clicked: number }>;
  repeatQueries: Array<{ queryText: string; totalSearches: number; lastSeen: Date }>;
  verticalUsage: Array<{ vertical: string; count: number }>;
}

// ─── Computation ──────────────────────────────────────────────────────────────

function computeMetrics(entries: ISearchHistoryEntry[]): IInsightMetrics {
  const total = entries.length;
  if (total === 0) {
    return {
      totalSearches: 0,
      zeroResultCount: 0,
      zeroResultRate: 0,
      clickedSearchCount: 0,
      clickThroughRate: 0,
      avgResultCount: 0,
      topQueries: [],
      topClickedItems: [],
      volumeByDay: [],
      ctrByDay: [],
      repeatQueries: [],
      verticalUsage: [],
    };
  }

  // ── Scalar metrics ────────────────────────────────────────
  let zeroCount = 0;
  let clickedCount = 0;
  let resultCountSum = 0;

  // ── Query frequency map ───────────────────────────────────
  const queryMap = new Map<string, { count: number; vertical: string }>();

  // ── Clicked item frequency map ────────────────────────────
  const clickMap = new Map<string, { title: string; clicks: number }>();

  // ── Volume by day (ISO date string → count) ───────────────
  const dayMap = new Map<string, number>();
  // T4.D7 / UX-007 — per-day clicked-search count for the CTR sparkline.
  // Shares its date keys with `dayMap` so the two zip cleanly at the end.
  const clickedDayMap = new Map<string, number>();

  for (let i = 0; i < entries.length; i++) {
    const e = entries[i];
    if (e.isZeroResult) {
      zeroCount++;
    }
    const hasClicks = !!(e.clickedItems && e.clickedItems.length > 0);
    if (hasClicks) {
      clickedCount++;
    }
    resultCountSum += e.resultCount || 0;

    // Query frequency
    const qKey = e.queryText.trim().toLowerCase();
    if (qKey) {
      const existing = queryMap.get(qKey);
      if (existing) {
        existing.count++;
      } else {
        queryMap.set(qKey, { count: 1, vertical: e.vertical || '' });
      }
    }

    // Clicked item frequency
    if (e.clickedItems) {
      for (let j = 0; j < e.clickedItems.length; j++) {
        const ci = e.clickedItems[j];
        const ciKey = ci.url;
        const existingClick = clickMap.get(ciKey);
        if (existingClick) {
          existingClick.clicks++;
        } else {
          clickMap.set(ciKey, { title: ci.title || ci.url, clicks: 1 });
        }
      }
    }

    // Volume by day
    if (e.searchTimestamp) {
      const dayKey = e.searchTimestamp.toISOString().substring(0, 10); // YYYY-MM-DD
      dayMap.set(dayKey, (dayMap.get(dayKey) || 0) + 1);
      if (hasClicks) {
        clickedDayMap.set(dayKey, (clickedDayMap.get(dayKey) || 0) + 1);
      }
    }
  }

  // ── Top 10 queries ────────────────────────────────────────
  const topQueries = Array.from(queryMap.entries())
    .map(([queryText, val]) => ({ queryText, count: val.count, vertical: val.vertical }))
    .sort((a, b) => b.count - a.count)
    .slice(0, 10);

  // ── Top 10 clicked items ──────────────────────────────────
  const topClickedItems = Array.from(clickMap.entries())
    .map(([url, val]) => ({ url, title: val.title, clicks: val.clicks }))
    .sort((a, b) => b.clicks - a.clicks)
    .slice(0, 10);

  // ── Volume by day — sorted ascending, last 30 entries ─────
  const volumeByDay = Array.from(dayMap.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .slice(-30)
    .map(([dateKey, count]) => {
      // Format "Mon 12" style
      const d = new Date(dateKey + 'T00:00:00');
      const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
      const dateLabel = dayNames[d.getDay()] + ' ' + String(d.getDate());
      return { dateLabel, count };
    });

  // T4.D7 / UX-007 — 7-day rolling CTR sparkline. Zip `dayMap` totals with
  // `clickedDayMap` per-day click counts, take the last 7 calendar days
  // that had any searches, and compute ctr = clicked / total. Days with
  // zero searches are skipped (ctr undefined). Audit acceptance:
  // "Insights panel shows a 7-day rolling CTR sparkline alongside the
  // window-wide number."
  const ctrByDay = Array.from(dayMap.entries())
    .sort((a, b) => a[0].localeCompare(b[0]))
    .slice(-7)
    .map(([dateKey, total]) => {
      const d = new Date(dateKey + 'T00:00:00');
      const dayNames = ['Sun', 'Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat'];
      const dateLabel = dayNames[d.getDay()] + ' ' + String(d.getDate());
      const clicked = clickedDayMap.get(dateKey) || 0;
      const ctr = total > 0 ? (clicked / total) * 100 : 0;
      return { dateLabel, ctr, total, clicked };
    });

  // ── Repeat queries: entries where UseCount > 2 ──────────
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

  // ── Vertical usage ─────────────────────────────────────
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

  return {
    totalSearches: total,
    zeroResultCount: zeroCount,
    zeroResultRate: total > 0 ? Math.round((zeroCount / total) * 100) : 0,
    clickedSearchCount: clickedCount,
    clickThroughRate: total > 0 ? Math.round((clickedCount / total) * 100) : 0,
    avgResultCount: total > 0 ? Math.round(resultCountSum / total) : 0,
    topQueries,
    topClickedItems,
    volumeByDay,
    ctrByDay,
    repeatQueries,
    verticalUsage,
  };
}

// ─── Sub-components ───────────────────────────────────────────────────────────

/**
 * Single stat card with icon, label, value, and optional colour hint.
 */
const StatCard: React.FC<{
  label: string;
  value: string;
  iconName: string;
  highlight?: boolean;
  warning?: boolean;
}> = (p) => (
  <div className={styles.insightStatCard + (p.warning ? ' ' + styles.insightStatCardWarning : '')}>
    <div className={styles.insightStatIconRow}>
      <Icon
        iconName={p.iconName}
        className={p.warning ? styles.insightStatIconWarning : styles.insightStatIcon}
      />
    </div>
    <div className={styles.insightStatValue + (p.highlight ? ' ' + styles.insightStatValueHighlight : '')}>
      {p.value}
    </div>
    <div className={styles.insightStatLabel}>{p.label}</div>
  </div>
);

/**
 * Horizontal mini-bar chart row: label | bar | value label.
 */
const BarRow: React.FC<{
  label: string;
  count: number;
  max: number;
  onClick?: () => void;
  href?: string;
}> = (p) => {
  const pct = p.max > 0 ? Math.max(2, Math.round((p.count / p.max) * 100)) : 2;
  const inner = (
    <>
      <span className={styles.insightBarLabel} title={p.label}>{p.label}</span>
      <div className={styles.insightBarTrack}>
        <div className={styles.insightBarFill} style={{ width: pct + '%' }} />
      </div>
      <span className={styles.insightBarCount}>{String(p.count)}</span>
    </>
  );

  if (p.onClick) {
    return (
      <div className={styles.insightBarRow + ' ' + styles.insightBarRowClickable} onClick={p.onClick} role="button" tabIndex={0}
        onKeyDown={(e): void => { if (e.key === 'Enter' || e.key === ' ') { p.onClick!(); } }}>
        {inner}
      </div>
    );
  }
  if (p.href) {
    return (
      <a className={styles.insightBarRow + ' ' + styles.insightBarRowClickable} href={p.href} target="_blank" rel="noopener noreferrer">
        {inner}
      </a>
    );
  }
  return <div className={styles.insightBarRow}>{inner}</div>;
};

/**
 * Simple column sparkline for daily volume.
 */
const VolumeChart: React.FC<{ days: IInsightMetrics['volumeByDay'] }> = ({ days }) => {
  if (days.length === 0) {
    return (
      <p className={styles.insightNoData}>No volume data for this period.</p>
    );
  }
  const max = Math.max(...days.map((d) => d.count), 1);
  return (
    <div className={styles.insightVolumeChart} aria-label="Search volume by day">
      {days.map((day) => {
        const pct = Math.max(4, Math.round((day.count / max) * 100));
        return (
          <div key={day.dateLabel} className={styles.insightVolumeColumn}>
            <span className={styles.insightVolumeCount}>{String(day.count)}</span>
            <div className={styles.insightVolumeBarTrack}>
              <div className={styles.insightVolumeBarFill} style={{ height: pct + '%' }} />
            </div>
            <span className={styles.insightVolumeLabel}>{day.dateLabel}</span>
          </div>
        );
      })}
    </div>
  );
};

// ─── Day-range options ────────────────────────────────────────────────────────

const RANGE_OPTIONS: IChoiceGroupOption[] = [
  { key: '30', text: '30d' },
  { key: '60', text: '60d' },
  { key: '90', text: '90d' },
];

// ─── Main component ───────────────────────────────────────────────────────────

/**
 * SearchInsightsPanel — admin analytics surface for SP Search.
 *
 * Loads all search history in a configurable rolling window (cross-user, no Author
 * filter), computes aggregate metrics client-side, and renders:
 *   - Summary stat cards (volume, zero-result %, CTR, avg results)
 *   - Top 10 queries ranked by frequency (clickable to re-run)
 *   - Top 10 clicked results ranked by engagement
 *   - Daily volume sparkline for trend visibility
 *
 * Lazy-loads on first tab open so it doesn't slow down Search Manager startup.
 */
const SearchInsightsPanel: React.FC<ISearchInsightsPanelProps> = (props) => {
  const { service, onRunQuery } = props;

  const [daysBack, setDaysBack] = React.useState<number>(props.daysBack || 30);
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [metrics, setMetrics] = React.useState<IInsightMetrics | undefined>(undefined);

  const load = React.useCallback(function (days: number): void {
    setIsLoading(true);
    setError(undefined);
    service.loadAllHistoryForInsights(days, 500)
      .then(function (entries): void {
        setMetrics(computeMetrics(entries));
        setIsLoading(false);
      })
      .catch(function (err): void {
        setError(err instanceof Error ? err.message : 'Failed to load insights data');
        setIsLoading(false);
      });
  }, [service]);

  // Load on first mount
  React.useEffect(function (): void {
    load(daysBack);
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Sync when parent changes daysBack prop
  const prevDaysBackRef = React.useRef<number | undefined>(props.daysBack);
  React.useEffect(function (): void {
    if (props.daysBack !== undefined && props.daysBack !== prevDaysBackRef.current) {
      prevDaysBackRef.current = props.daysBack;
      setDaysBack(props.daysBack);
      load(props.daysBack);
    }
  }, [props.daysBack, load]);

  function handleRangeChange(_: unknown, option?: IChoiceGroupOption): void {
    if (!option) {
      return;
    }
    const days = parseInt(option.key, 10);
    setDaysBack(days);
    load(days);
  }

  function handleRefresh(): void {
    load(daysBack);
  }

  // ── Loading ───────────────────────────────────────────────
  if (isLoading) {
    return (
      <div className={styles.loadingContainer}>
        <Spinner size={SpinnerSize.large} label="Loading insights..." />
      </div>
    );
  }

  // ── Error ─────────────────────────────────────────────────
  if (error) {
    return (
      <div className={styles.healthPanel}>
        <div className={styles.errorContainer} role="alert">
          <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>{error}</MessageBar>
        </div>
        <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Retry" onClick={handleRefresh} />
      </div>
    );
  }

  // ── Empty ─────────────────────────────────────────────────
  if (!metrics || metrics.totalSearches === 0) {
    return (
      <div className={styles.insightsPanel}>
        <div className={styles.insightsToolbar}>
          {!props.hideTimeRange && (
            <ChoiceGroup
              options={RANGE_OPTIONS}
              selectedKey={String(daysBack)}
              onChange={handleRangeChange}
              styles={{ flexContainer: { display: 'flex', gap: 12 } }}
            />
          )}
          <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh" onClick={handleRefresh} />
        </div>
        <div className={styles.emptyState}>
          <div className={styles.emptyIcon}><Icon iconName="BarChart4" /></div>
          <h3 className={styles.emptyTitle}>No data yet</h3>
          <p className={styles.emptyDescription}>
            Search history for the last {daysBack} days will appear here once users start searching.
          </p>
        </div>
      </div>
    );
  }

  const topQueryMax = metrics.topQueries.length > 0 ? metrics.topQueries[0].count : 1;
  const topClickMax = metrics.topClickedItems.length > 0 ? metrics.topClickedItems[0].clicks : 1;

  return (
    <div className={styles.insightsPanel}>
      {/* ── Toolbar ──────────────────────────────────────────── */}
      <div className={styles.insightsToolbar}>
        {!props.hideTimeRange && (
          <ChoiceGroup
            options={RANGE_OPTIONS}
            selectedKey={String(daysBack)}
            onChange={handleRangeChange}
            styles={{ flexContainer: { display: 'flex', gap: 12 }, label: { display: 'none' } }}
            label=""
          />
        )}
        <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh" onClick={handleRefresh} />
      </div>

      {/* ── Stat cards ───────────────────────────────────────── */}
      <div className={styles.insightStatCards}>
        <StatCard
          label="Total searches"
          value={String(metrics.totalSearches)}
          iconName="Search"
        />
        <StatCard
          label="Zero-result rate"
          value={String(metrics.zeroResultRate) + '%'}
          iconName="SearchIssue"
          warning={metrics.zeroResultRate > 20}
        />
        <StatCard
          label="Click-through rate"
          value={String(metrics.clickThroughRate) + '%'}
          iconName="TouchPointer"
          highlight={metrics.clickThroughRate > 50}
        />
        <StatCard
          label="Avg result count"
          value={String(metrics.avgResultCount)}
          iconName="NumberField"
        />
      </div>

      {/* T4.D7 / UX-007 — 7-day rolling CTR sparkline. Sits between the
          window-wide CTR stat card and the split section so the admin can
          compare today's trend against the time-window aggregate. */}
      {metrics.ctrByDay.length > 0 && (
        <div className={styles.insightSection} style={{ marginTop: 4, marginBottom: 12 }}>
          <h3 className={styles.insightSectionTitle}>
            <Icon iconName="LineChart" className={styles.insightSectionIcon} />
            CTR trend (last 7 days)
          </h3>
          <div style={{ display: 'flex', alignItems: 'flex-end', gap: 8 }}>
            {metrics.ctrByDay.map((day, idx) => {
              // Bars scale to a 100% ceiling (CTR cap) so the visual is
              // comparable day-over-day rather than self-normalised.
              const pct = Math.max(2, Math.min(100, day.ctr));
              const color = day.ctr >= 50 ? '#107c10' : day.ctr >= 20 ? '#0078d4' : '#a4262c';
              return (
                <div
                  key={String(idx) + '-' + day.dateLabel}
                  style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', flex: 1 }}
                  title={day.dateLabel + ': ' + day.ctr.toFixed(1) + '% (' + day.clicked + '/' + day.total + ')'}
                >
                  <span style={{ fontSize: 11, color: '#605e5c' }}>{day.ctr.toFixed(0)}%</span>
                  <div style={{ width: '100%', height: 40, display: 'flex', alignItems: 'flex-end' }}>
                    <div style={{ width: '100%', height: pct + '%', backgroundColor: color, borderRadius: 2 }} />
                  </div>
                  <span style={{ fontSize: 11, color: '#605e5c', marginTop: 2 }}>{day.dateLabel}</span>
                </div>
              );
            })}
          </div>
        </div>
      )}

      {/* ── Two-column section: top queries + top clicked ─────── */}
      <div className={styles.insightsSplit}>

        {/* Top queries */}
        <div className={styles.insightSection}>
          <h3 className={styles.insightSectionTitle}>
            <Icon iconName="Search" className={styles.insightSectionIcon} />
            Top queries
          </h3>
          {metrics.topQueries.length === 0 ? (
            <p className={styles.insightNoData}>No query data yet.</p>
          ) : (
            <div className={styles.insightBarList}>
              {metrics.topQueries.map((q) => (
                <BarRow
                  key={q.queryText}
                  label={q.queryText}
                  count={q.count}
                  max={topQueryMax}
                  onClick={(): void => onRunQuery(q.queryText, q.vertical)}
                />
              ))}
            </div>
          )}
        </div>

        {/* Top clicked */}
        <div className={styles.insightSection}>
          <h3 className={styles.insightSectionTitle}>
            <Icon iconName="TouchPointer" className={styles.insightSectionIcon} />
            Top clicked results
          </h3>
          {metrics.topClickedItems.length === 0 ? (
            <p className={styles.insightNoData}>
              No clicks recorded yet. Click tracking activates automatically when users open results.
            </p>
          ) : (
            <div className={styles.insightBarList}>
              {metrics.topClickedItems.map((item) => (
                <BarRow
                  key={item.url}
                  label={item.title}
                  count={item.clicks}
                  max={topClickMax}
                  href={item.url}
                />
              ))}
            </div>
          )}
        </div>
      </div>

      {/* ── Daily volume ─────────────────────────────────────── */}
      <div className={styles.insightSection}>
        <h3 className={styles.insightSectionTitle}>
          <Icon iconName="BarChart4" className={styles.insightSectionIcon} />
          Daily volume
        </h3>
        <VolumeChart days={metrics.volumeByDay} />
      </div>

      {/* ── Vertical usage ──────────────────────────────────── */}
      {metrics.verticalUsage.length > 0 && (
        <div className={styles.insightSection}>
          <h3 className={styles.insightSectionTitle}>
            <Icon iconName="TabCenter" className={styles.insightSectionIcon} />
            Vertical Usage
          </h3>
          <div className={styles.insightBarList}>
            {metrics.verticalUsage.map(function (v) {
              const maxCount = metrics.verticalUsage[0].count;
              return (
                <BarRow
                  key={v.vertical}
                  label={v.vertical}
                  count={v.count}
                  max={maxCount}
                />
              );
            })}
          </div>
        </div>
      )}

      {/* ── Repeat queries ──────────────────────────────────── */}
      {metrics.repeatQueries.length > 0 && (
        <div className={styles.insightSection}>
          <h3 className={styles.insightSectionTitle}>
            <Icon iconName="Sync" className={styles.insightSectionIcon} />
            Repeat Queries
          </h3>
          <p className={styles.insightNoData}>
            Queries searched 3+ times by the same user (candidates for promoted results)
          </p>
          <div className={styles.insightBarList}>
            {metrics.repeatQueries.slice(0, 10).map(function (rq) {
              const maxCount = metrics.repeatQueries[0].totalSearches;
              return (
                <BarRow
                  key={rq.queryText}
                  label={rq.queryText}
                  count={rq.totalSearches}
                  max={maxCount}
                  onClick={function () { onRunQuery(rq.queryText, ''); }}
                />
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
};

export default SearchInsightsPanel;
