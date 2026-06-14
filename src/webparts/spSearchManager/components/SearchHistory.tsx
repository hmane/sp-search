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
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { validateSearchState } from '@store/utils/searchStateSchema';
import { spLog } from '@store/utils/spLog';
import styles from './SpSearchManager.module.scss';
import { getHistoryDisplay } from './historyDisplay';
import { formatHistoryTime, groupSearchHistoryByDate } from './historyGrouping';

export interface ISearchHistoryProps {
  store: StoreApi<ISearchStore>;
  service: SearchManagerService;
  history: ISearchHistoryEntry[];
  onDataChanged: () => void;
  onSearchLoaded?: () => void;
}

/**
 * SearchHistory -- displays a chronological list of past searches.
 * Each item shows the query text, vertical, result count, and time.
 * Click to re-execute the search. Includes a "Clear history" button.
 */
const SearchHistory: React.FC<ISearchHistoryProps> = (props) => {
  const { store, service, history, onDataChanged, onSearchLoaded } = props;
  const historyGroups = React.useMemo(function () {
    return groupSearchHistoryByDate(history || []);
  }, [history]);

  // ─── Local state ──────────────────────────────────────────
  const [showClearDialog, setShowClearDialog] = React.useState<boolean>(false);
  const [isClearing, setIsClearing] = React.useState<boolean>(false);
  // T2.D3 — set when a history-entry restore fails schema validation.
  const [restoreError, setRestoreError] = React.useState<{ queryText: string; errors: string[] } | undefined>(undefined);

  // ─── Handlers ─────────────────────────────────────────────

  function handleReExecuteSearch(entry: ISearchHistoryEntry): void {
    // T2.D3 — schema-validate before applying. Malformed history rows are
    // surfaced via MessageBar instead of silently corrupting the store.
    const validation = validateSearchState(entry.searchState);
    if (!validation.ok) {
      setRestoreError({ queryText: entry.queryText, errors: validation.errors });
      spLog.warn('Skipping history-entry restore; schema validation failed', { errors: validation.errors });
      return;
    }
    const savedState = validation.state;

    // Set ALL state atomically via store.setState() so the orchestrator
    // sees a single change notification and fires ONE search with
    // complete state (including filters).
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const update: Record<string, any> = {
      queryText: savedState.queryText !== undefined ? savedState.queryText : entry.queryText,
      activeFilters: savedState.activeFilters || [],
      currentPage: 1,
    };

    if (savedState.currentVerticalKey) {
      update.currentVerticalKey = savedState.currentVerticalKey;
    } else if (entry.vertical) {
      update.currentVerticalKey = entry.vertical;
    }
    if (savedState.sort) {
      update.sort = savedState.sort;
    }
    store.setState(update);
    setRestoreError(undefined);

    // Notify parent (e.g., close panel in panel mode)
    if (onSearchLoaded) {
      onSearchLoaded();
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
      {restoreError && (
        <MessageBar
          messageBarType={MessageBarType.error}
          isMultiline={true}
          onDismiss={(): void => setRestoreError(undefined)}
          dismissButtonAriaLabel="Dismiss"
        >
          <strong>Could not restore the search for &quot;{restoreError.queryText}&quot;.</strong> The
          history entry is malformed and was not applied. Details: {restoreError.errors.join('; ')}
        </MessageBar>
      )}
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
        {historyGroups.map(function (group): React.ReactElement {
          return (
            <section key={group.key} className={styles.historyDateGroup}>
              <div className={styles.historyDateHeader}>
                <h3 className={styles.historyDateLabel}>{group.label}</h3>
                <span className={styles.historyDateCount}>
                  {String(group.count)} {group.count === 1 ? 'search' : 'searches'}
                </span>
              </div>
              {group.entries.map(function (entry): React.ReactElement {
                const display = getHistoryDisplay(entry);
                return (
                  <div
                    key={entry.id}
                    className={styles.historyItem}
                    onClick={function (): void { handleReExecuteSearch(entry); }}
                    role="button"
                    aria-label={'Re-run search: ' + display.title}
                  >
                    {/* Body: query + meta */}
                    <div className={styles.historyBody}>
                      <p className={styles.historyQuery} title={display.title}>{display.title}</p>
                      <div className={styles.historyMeta}>
                        {display.metaParts.map(function (part: string, index: number): React.ReactElement {
                          const isUsage = part.indexOf('Used ') === 0;
                          return (
                            <React.Fragment key={part + String(index)}>
                              {index > 0 && <span className={styles.metaDot} />}
                              <span className={isUsage ? styles.historyCountBadge : undefined}>
                                {part}
                              </span>
                            </React.Fragment>
                          );
                        })}
                      </div>
                    </div>

                    {/* Timestamp */}
                    <span className={styles.historyTimestamp}>
                      {formatHistoryTime(entry.searchTimestamp)}
                    </span>
                  </div>
                );
              })}
            </section>
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
