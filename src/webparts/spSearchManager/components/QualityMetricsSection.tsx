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
  // T4.D7 / UX-006 — per-vertical breakdown for the zero-rate table. `zeroRate`
  // is computed at aggregation time so the renderer can sort by it without
  // re-deriving the ratio on every paint.
  verticalUsage: Array<{ vertical: string; count: number; zeroCount: number; zeroRate: number }>;
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

      {/* T4.D7 / UX-006 — per-vertical zero-result rate table. Sorted by
          zero-rate descending so the worst-performing verticals surface
          first. Empty verticals are skipped (the audit signal calls out
          "per-vertical zero-rate table sortable by rate descending"). */}
      {metrics.verticalUsage.length > 1 && (
        <div style={{ marginTop: 16 }}>
          <h3 className={styles.insightSectionTitle}>Zero-result rate by vertical</h3>
          <table style={{ width: '100%', borderCollapse: 'collapse', fontSize: 13 }}>
            <thead>
              <tr style={{ textAlign: 'left', borderBottom: '1px solid #edebe9', color: '#605e5c' }}>
                <th style={{ padding: '6px 8px' }}>Vertical</th>
                <th style={{ padding: '6px 8px', textAlign: 'right' }}>Searches</th>
                <th style={{ padding: '6px 8px', textAlign: 'right' }}>Zero-result</th>
                <th style={{ padding: '6px 8px', textAlign: 'right' }}>Rate</th>
              </tr>
            </thead>
            <tbody>
              {[...metrics.verticalUsage].sort((a, b) => b.zeroRate - a.zeroRate).map(function (v) {
                const isWarning = v.zeroRate > 20 && v.count >= 5;
                return (
                  <tr key={v.vertical} style={{ borderBottom: '1px solid #faf9f8' }}>
                    <td style={{ padding: '6px 8px' }}>{v.vertical}</td>
                    <td style={{ padding: '6px 8px', textAlign: 'right' }}>{v.count.toLocaleString()}</td>
                    <td style={{ padding: '6px 8px', textAlign: 'right' }}>{v.zeroCount.toLocaleString()}</td>
                    <td style={{ padding: '6px 8px', textAlign: 'right', color: isWarning ? '#a4262c' : '#323130', fontWeight: isWarning ? 600 : 400 }}>
                      {v.zeroRate.toFixed(1)}%
                    </td>
                  </tr>
                );
              })}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
};

export default QualityMetricsSection;
