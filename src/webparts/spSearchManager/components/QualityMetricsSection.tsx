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
