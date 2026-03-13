import * as React from 'react';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { ISearchHistoryEntry } from '@interfaces/index';
import { SearchManagerService } from '@services/index';
import styles from './SpSearchManager.module.scss';

export interface IZeroResultsPanelProps {
  service: SearchManagerService;
  /** Called when the user clicks "Try it" on a zero-result query */
  onRunQuery: (queryText: string, vertical: string, scope: string) => void;
}

// ─── Aggregation ─────────────────────────────────────────────────────────────

interface IZeroResultSummary {
  queryText: string;
  count: number;
  verticals: string[];
  scope: string;
  lastSeen: Date;
}

/**
 * Collapses raw history entries into per-query-text summaries, sorted by
 * occurrence count descending (most-broken queries first).
 */
function aggregateEntries(entries: ISearchHistoryEntry[]): IZeroResultSummary[] {
  const map = new Map<string, IZeroResultSummary>();

  for (let i = 0; i < entries.length; i++) {
    const entry = entries[i];
    const key = entry.queryText.trim().toLowerCase();
    if (!key) {
      continue;
    }

    const existing = map.get(key);
    if (existing) {
      existing.count++;
      if (entry.vertical && existing.verticals.indexOf(entry.vertical) < 0) {
        existing.verticals.push(entry.vertical);
      }
      if (entry.searchTimestamp > existing.lastSeen) {
        existing.lastSeen = entry.searchTimestamp;
      }
    } else {
      map.set(key, {
        queryText: entry.queryText.trim(),
        count: 1,
        verticals: entry.vertical ? [entry.vertical] : [],
        scope: entry.scope || '',
        lastSeen: entry.searchTimestamp,
      });
    }
  }

  const result = Array.from(map.values());
  result.sort((a, b) => b.count - a.count || b.lastSeen.getTime() - a.lastSeen.getTime());
  return result;
}

// ─── Date formatter ───────────────────────────────────────────────────────────

function formatShortDate(date: Date): string {
  if (!date || isNaN(date.getTime())) {
    return '—';
  }
  const now = new Date();
  const diffMs = now.getTime() - date.getTime();
  const diffDays = Math.floor(diffMs / (1000 * 60 * 60 * 24));
  if (diffDays === 0) {
    return 'Today';
  }
  if (diffDays === 1) {
    return 'Yesterday';
  }
  if (diffDays < 7) {
    return String(diffDays) + 'd ago';
  }
  return date.toLocaleDateString();
}

// ─── Component ────────────────────────────────────────────────────────────────

const DAYS_BACK = 90;

/**
 * ZeroResultsPanel — admin health surface for zero-result query tuning.
 *
 * Loads all zero-result queries from the last 90 days (cross-user, no Author
 * filter), aggregates them by query text, and displays a ranked table so
 * admins can identify which queries need query rules, synonyms, or content.
 *
 * Lazy-loads on first mount so it does not slow down Search Manager startup.
 */
const ZeroResultsPanel: React.FC<IZeroResultsPanelProps> = (props) => {
  const { service, onRunQuery } = props;

  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [summaries, setSummaries] = React.useState<IZeroResultSummary[]>([]);
  const [rawCount, setRawCount] = React.useState<number>(0);

  const load = React.useCallback(function (): void {
    setIsLoading(true);
    setError(undefined);
    service.loadZeroResultQueries(DAYS_BACK, 200)
      .then(function (entries): void {
        setRawCount(entries.length);
        setSummaries(aggregateEntries(entries));
        setIsLoading(false);
      })
      .catch(function (err): void {
        setError(err instanceof Error ? err.message : 'Failed to load zero-result data');
        setIsLoading(false);
      });
  }, [service]);

  // Load on first mount only
  React.useEffect(function (): void {
    load();
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  function handleRefresh(): void {
    load();
  }

  function handleTryQuery(summary: IZeroResultSummary): void {
    onRunQuery(summary.queryText, summary.verticals[0] || '', summary.scope);
  }

  // ─── Loading ────────────────────────────────────────────────────────────────

  if (isLoading) {
    return (
      <div className={styles.loadingContainer}>
        <Spinner size={SpinnerSize.large} label="Loading health data..." />
      </div>
    );
  }

  // ─── Error ──────────────────────────────────────────────────────────────────

  if (error) {
    return (
      <div className={styles.healthPanel}>
        <div className={styles.errorContainer} role="alert">
          <MessageBar
            messageBarType={MessageBarType.error}
            isMultiline={false}
          >
            {error}
          </MessageBar>
        </div>
        <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Retry" onClick={handleRefresh} />
      </div>
    );
  }

  // ─── Empty (healthy) ────────────────────────────────────────────────────────

  if (summaries.length === 0) {
    return (
      <div className={styles.healthPanel}>
        <div className={styles.healthToolbar}>
          <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh" onClick={handleRefresh} />
        </div>
        <div className={styles.emptyState}>
          <div className={styles.emptyIcon}>
            <Icon iconName="StatusCircleCheckmark" />
          </div>
          <h3 className={styles.emptyTitle}>All clear</h3>
          <p className={styles.emptyDescription}>
            No zero-result queries in the last {DAYS_BACK} days.
            Your search experience looks healthy.
          </p>
        </div>
      </div>
    );
  }

  // ─── Table ──────────────────────────────────────────────────────────────────

  return (
    <div className={styles.healthPanel}>
      {/* Toolbar */}
      <div className={styles.healthToolbar}>
        <p className={styles.healthSummary}>
          <Icon iconName="Warning" className={styles.healthSummaryIcon} />
          <strong>{String(summaries.length)}</strong> unique {'quer' + (summaries.length === 1 ? 'y' : 'ies')} returned
          no results across <strong>{String(rawCount)}</strong> searches in the last {DAYS_BACK} days.
        </p>
        <DefaultButton iconProps={{ iconName: 'Refresh' }} text="Refresh" onClick={handleRefresh} />
      </div>

      {/* Table */}
      <div className={styles.zeroResultsTable} role="table" aria-label="Zero-result queries">

        {/* Header row */}
        <div className={styles.zeroResultsHeader} role="row">
          <div className={styles.zeroResultsColQuery} role="columnheader">Query</div>
          <div className={styles.zeroResultsColCount} role="columnheader">Occurrences</div>
          <div className={styles.zeroResultsColVerticals} role="columnheader">Verticals</div>
          <div className={styles.zeroResultsColDate} role="columnheader">Last seen</div>
          <div className={styles.zeroResultsColAction} role="columnheader" />
        </div>

        {/* Data rows */}
        {summaries.map(function (summary, index): React.ReactElement {
          const verticalDisplay = summary.verticals.filter(Boolean).join(', ') || '—';
          return (
            <div key={summary.queryText + '-' + String(index)} className={styles.zeroResultsRow} role="row">
              <div className={styles.zeroResultsColQuery} role="cell">
                <Icon iconName="SearchIssue" className={styles.zeroResultsQueryIcon} />
                <span className={styles.zeroResultsQueryText}>{summary.queryText}</span>
              </div>
              <div className={styles.zeroResultsColCount} role="cell">
                <span className={styles.zeroResultsCountBadge}>
                  {String(summary.count)}
                </span>
              </div>
              <div className={styles.zeroResultsColVerticals} role="cell">
                <TooltipHost content={verticalDisplay}>
                  <span className={styles.zeroResultsVertical}>{verticalDisplay}</span>
                </TooltipHost>
              </div>
              <div className={styles.zeroResultsColDate} role="cell">
                {formatShortDate(summary.lastSeen)}
              </div>
              <div className={styles.zeroResultsColAction} role="cell">
                <DefaultButton
                  iconProps={{ iconName: 'Play' }}
                  text="Try it"
                  className={styles.zeroResultsTryBtn}
                  onClick={function (): void { handleTryQuery(summary); }}
                  title={'Re-run "' + summary.queryText + '" to see why it returns no results'}
                />
              </div>
            </div>
          );
        })}
      </div>
    </div>
  );
};

export default ZeroResultsPanel;
