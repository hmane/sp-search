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
