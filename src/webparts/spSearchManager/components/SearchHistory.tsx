import * as React from 'react';
import { StoreApi } from 'zustand/vanilla';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import {
  ISearchHistoryEntry,
  ISearchStore
} from '@interfaces/index';
import { SearchManagerService } from '@services/index';
import styles from './SpSearchManager.module.scss';

export interface ISearchHistoryProps {
  store: StoreApi<ISearchStore>;
  service: SearchManagerService;
  history: ISearchHistoryEntry[];
  onDataChanged: () => void;
}

/**
 * Format a date as a relative timestamp string.
 */
function formatRelativeTimestamp(date: Date): string {
  const now = new Date();
  const diffMs = now.getTime() - date.getTime();
  const diffSec = Math.floor(diffMs / 1000);
  const diffMin = Math.floor(diffSec / 60);
  const diffHour = Math.floor(diffMin / 60);
  const diffDay = Math.floor(diffHour / 24);
  const diffWeek = Math.floor(diffDay / 7);

  if (diffSec < 60) {
    return 'just now';
  }
  if (diffMin < 60) {
    return diffMin === 1 ? '1 minute ago' : String(diffMin) + ' minutes ago';
  }
  if (diffHour < 24) {
    return diffHour === 1 ? '1 hour ago' : String(diffHour) + ' hours ago';
  }
  if (diffDay < 7) {
    return diffDay === 1 ? 'yesterday' : String(diffDay) + ' days ago';
  }
  if (diffWeek < 4) {
    return diffWeek === 1 ? '1 week ago' : String(diffWeek) + ' weeks ago';
  }

  return date.toLocaleDateString();
}

/**
 * SearchHistory -- displays a chronological list of past searches.
 * Each item shows the query text, vertical, result count, and relative timestamp.
 * Click to re-execute the search. Includes a "Clear history" button.
 */
const SearchHistory: React.FC<ISearchHistoryProps> = (props) => {
  const { store, service, history, onDataChanged } = props;

  // ─── Local state ──────────────────────────────────────────
  const [showClearDialog, setShowClearDialog] = React.useState<boolean>(false);
  const [isClearing, setIsClearing] = React.useState<boolean>(false);

  // ─── Handlers ─────────────────────────────────────────────

  function handleReExecuteSearch(entry: ISearchHistoryEntry): void {
    const storeState = store.getState();

    // Try to restore full search state from JSON
    try {
      const savedState: {
        queryText?: string;
        activeFilters?: Array<{ filterName: string; value: string; operator: 'AND' | 'OR' }>;
        currentVerticalKey?: string;
        sort?: { property: string; direction: 'Ascending' | 'Descending' };
        scope?: { id: string; label: string; kqlPath?: string; resultSourceId?: string };
        activeLayoutKey?: string;
      } = JSON.parse(entry.searchState || '{}');

      // Clear existing filters first to avoid stale state
      storeState.clearAllFilters();

      // Restore query text
      storeState.setQueryText(savedState.queryText || entry.queryText);

      // Restore filters
      if (savedState.activeFilters && savedState.activeFilters.length > 0) {
        for (let i = 0; i < savedState.activeFilters.length; i++) {
          storeState.setRefiner(savedState.activeFilters[i]);
        }
      }

      // Restore vertical
      if (savedState.currentVerticalKey) {
        storeState.setVertical(savedState.currentVerticalKey);
      } else if (entry.vertical) {
        storeState.setVertical(entry.vertical);
      }

      // Restore sort
      if (savedState.sort) {
        storeState.setSort(savedState.sort);
      }

      // Restore scope
      if (savedState.scope) {
        storeState.setScope(savedState.scope);
      } else if (entry.scope) {
        storeState.setScope({ id: entry.scope, label: entry.scope });
      }

      // Restore layout
      if (savedState.activeLayoutKey) {
        storeState.setLayout(savedState.activeLayoutKey);
      }
    } catch {
      // Fallback: just set query text
      storeState.setQueryText(entry.queryText);

      if (entry.vertical) {
        storeState.setVertical(entry.vertical);
      }
      if (entry.scope) {
        storeState.setScope({ id: entry.scope, label: entry.scope });
      }
    }
  }

  function handleClearClick(): void {
    setShowClearDialog(true);
  }

  function handleClearConfirm(): void {
    setIsClearing(true);
    service.clearHistory()
      .then(function (): void {
        setShowClearDialog(false);
        setIsClearing(false);
        onDataChanged();
      })
      .catch(function (): void {
        setIsClearing(false);
        setShowClearDialog(false);
      });
  }

  function handleClearCancel(): void {
    setShowClearDialog(false);
  }

  // ─── Empty state ──────────────────────────────────────────
  if (!history || history.length === 0) {
    return (
      <div className={styles.emptyState}>
        <div className={styles.emptyIcon}>
          <Icon iconName="History" />
        </div>
        <h3 className={styles.emptyTitle}>No search history</h3>
        <p className={styles.emptyDescription}>
          Your recent searches will appear here. Start searching to build your history.
        </p>
      </div>
    );
  }

  return (
    <div>
      {/* Toolbar with clear button */}
      <div className={styles.historyToolbar}>
        <DefaultButton
          iconProps={{ iconName: 'Delete' }}
          text="Clear history"
          onClick={handleClearClick}
        />
      </div>

      {/* History list */}
      <div className={styles.historyList}>
        {history.map(function (entry): React.ReactElement {
          return (
            <div
              key={entry.id}
              className={styles.historyItem}
              onClick={function (): void { handleReExecuteSearch(entry); }}
              role="button"
              aria-label={'Re-run search: ' + entry.queryText}
            >
              {/* Clock icon */}
              <div className={styles.historyIcon}>
                <Icon iconName="History" />
              </div>

              {/* Body: query + meta */}
              <div className={styles.historyBody}>
                <p className={styles.historyQuery}>{entry.queryText}</p>
                <div className={styles.historyMeta}>
                  {entry.vertical && (
                    <span>{entry.vertical}</span>
                  )}
                  {entry.vertical && <span className={styles.metaDot} />}
                  <span>{String(entry.resultCount) + ' results'}</span>
                </div>
              </div>

              {/* Timestamp */}
              <span className={styles.historyTimestamp}>
                {formatRelativeTimestamp(entry.searchTimestamp)}
              </span>
            </div>
          );
        })}
      </div>

      {/* Clear history confirmation dialog */}
      <Dialog
        hidden={!showClearDialog}
        onDismiss={handleClearCancel}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Clear search history',
          subText: 'Are you sure you want to clear your entire search history? This action cannot be undone.'
        }}
        modalProps={{ isBlocking: true }}
      >
        {isClearing && (
          <div className={styles.loadingContainer}>
            <Spinner size={SpinnerSize.medium} label="Clearing history..." />
          </div>
        )}
        <DialogFooter>
          <PrimaryButton
            onClick={handleClearConfirm}
            text="Clear all"
            disabled={isClearing}
          />
          <DefaultButton
            onClick={handleClearCancel}
            text="Cancel"
          />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default SearchHistory;
