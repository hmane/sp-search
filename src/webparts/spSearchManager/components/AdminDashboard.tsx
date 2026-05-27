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
  // T4.D4 — number of coverage profiles the admin has configured in the
  // property pane. Zero means the Coverage section renders a CTA empty
  // state instead of broken stat cards. The actual fetch in
  // `CoverageStatsService` uses the store's queryTemplate, not the
  // profile URLs, but a clean tenant install with `coverageProfilesCollection: []`
  // is the signal that the admin hasn't configured the surface yet.
  coverageProfileCount: number;
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
  // T4.D7 / UX-006 — track per-vertical totals + zero-result counts so the
  // QualityMetrics surface can render a sortable per-vertical zero-rate table.
  const verticalMap = new Map<string, { count: number; zeroCount: number }>();

  for (let i = 0; i < entries.length; i++) {
    const e = entries[i];
    if (e.isZeroResult) zeroCount++;
    if (e.clickedItems && e.clickedItems.length > 0) clickedCount++;
    if (e.useCount > 1) {
      repeatCount++;
      hasUseCount = true;
    }
    const v = e.vertical || 'All';
    const bucket = verticalMap.get(v) || { count: 0, zeroCount: 0 };
    bucket.count++;
    if (e.isZeroResult) { bucket.zeroCount++; }
    verticalMap.set(v, bucket);
  }

  const verticalUsage: Array<{ vertical: string; count: number; zeroCount: number; zeroRate: number }> = [];
  verticalMap.forEach(function (bucket, vertical): void {
    verticalUsage.push({
      vertical,
      count: bucket.count,
      zeroCount: bucket.zeroCount,
      zeroRate: bucket.count > 0 ? (bucket.zeroCount / bucket.count) * 100 : 0,
    });
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
  const { store, service, expectedSiteUrls, coverageProfileCount, onRunQuery } = props;

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
  React.useEffect(function (): () => void {
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
          coverageProfileCount === 0 ? (
            // T4.D4 — empty-state CTA. Clean tenant install with no profiles
            // configured: render the helpful copy instead of broken stat
            // cards. SPFx can't programmatically open the web part's own
            // property pane from rendered output, so the CTA is descriptive
            // (admin opens edit → ellipsis → Edit web part).
            <div role="status" style={{ padding: 16, backgroundColor: '#faf9f8', borderRadius: 4, color: '#605e5c' }}>
              <Icon iconName="Info" style={{ marginRight: 8, verticalAlign: 'middle', color: '#0078d4' }} />
              <strong>No coverage profiles configured.</strong>
              <div style={{ marginTop: 8, fontSize: 13 }}>
                Configure coverage profiles in the web part property pane to begin monitoring item count, freshness, and gap analysis against your expected sites.
              </div>
            </div>
          ) : (
            <CoverageStatsSection
              coverage={coverage}
              expectedSiteUrls={expectedSiteUrls}
              isLoading={coverageLoading}
              error={coverageError}
            />
          )
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
