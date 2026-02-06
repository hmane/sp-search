import * as React from 'react';
import { StoreApi } from 'zustand/vanilla';
import { IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import {
  ISavedSearch,
  ISearchStore,
  IActiveFilter
} from '@interfaces/index';
import { SearchManagerService } from '@services/index';
import styles from './SpSearchManager.module.scss';

export interface ISavedSearchListProps {
  store: StoreApi<ISearchStore>;
  service: SearchManagerService;
  savedSearches: ISavedSearch[];
  onDataChanged: () => void;
  onShare: (search: ISavedSearch) => void;
}

/**
 * Format a date as a relative string ("just now", "2 hours ago", "3 days ago", etc.).
 */
function formatRelativeDate(date: Date): string {
  const now = new Date();
  const diffMs = now.getTime() - date.getTime();
  const diffSec = Math.floor(diffMs / 1000);
  const diffMin = Math.floor(diffSec / 60);
  const diffHour = Math.floor(diffMin / 60);
  const diffDay = Math.floor(diffHour / 24);

  if (diffSec < 60) {
    return 'just now';
  }
  if (diffMin < 60) {
    return diffMin === 1 ? '1 minute ago' : String(diffMin) + ' minutes ago';
  }
  if (diffHour < 24) {
    return diffHour === 1 ? '1 hour ago' : String(diffHour) + ' hours ago';
  }
  if (diffDay < 30) {
    return diffDay === 1 ? '1 day ago' : String(diffDay) + ' days ago';
  }

  // Older than 30 days -- show actual date
  return date.toLocaleDateString();
}

/**
 * Group saved searches by category.
 */
function groupByCategory(searches: ISavedSearch[]): Record<string, ISavedSearch[]> {
  const grouped: Record<string, ISavedSearch[]> = {};
  for (let i = 0; i < searches.length; i++) {
    const category = searches[i].category || 'Uncategorized';
    if (!grouped[category]) {
      grouped[category] = [];
    }
    grouped[category].push(searches[i]);
  }
  return grouped;
}

/**
 * SavedSearchList -- displays saved searches grouped by category with
 * collapsible sections. Supports click-to-load, inline rename, delete
 * confirmation, and share action.
 */
const SavedSearchList: React.FC<ISavedSearchListProps> = (props) => {
  const { store, service, savedSearches, onDataChanged, onShare } = props;

  // ─── Local state ──────────────────────────────────────────
  const [expandedCategories, setExpandedCategories] = React.useState<Record<string, boolean>>({});
  const [renamingId, setRenamingId] = React.useState<number | undefined>(undefined);
  const [renameValue, setRenameValue] = React.useState<string>('');
  const [deleteTarget, setDeleteTarget] = React.useState<ISavedSearch | undefined>(undefined);
  const [isDeleting, setIsDeleting] = React.useState<boolean>(false);
  const [isRenaming, setIsRenaming] = React.useState<boolean>(false);

  // ─── Initialize expanded categories ──────────────────────
  React.useEffect(() => {
    const initial: Record<string, boolean> = {};
    const grouped = groupByCategory(savedSearches);
    const keys = Object.keys(grouped);
    for (let i = 0; i < keys.length; i++) {
      if (expandedCategories[keys[i]] === undefined) {
        initial[keys[i]] = true;
      }
    }
    if (Object.keys(initial).length > 0) {
      setExpandedCategories((prev) => {
        const next = { ...prev };
        const initKeys = Object.keys(initial);
        for (let i = 0; i < initKeys.length; i++) {
          if (next[initKeys[i]] === undefined) {
            next[initKeys[i]] = initial[initKeys[i]];
          }
        }
        return next;
      });
    }
  }, [savedSearches]); // eslint-disable-line react-hooks/exhaustive-deps

  // ─── Handlers ─────────────────────────────────────────────

  function handleToggleCategory(category: string): void {
    setExpandedCategories((prev) => ({
      ...prev,
      [category]: !prev[category]
    }));
  }

  function handleLoadSearch(search: ISavedSearch): void {
    if (renamingId === search.id) {
      return; // Don't load if renaming
    }
    try {
      const state: {
        queryText?: string;
        activeFilters?: IActiveFilter[];
        currentVerticalKey?: string;
        sort?: { property: string; direction: 'Ascending' | 'Descending' };
        scope?: { id: string; label: string; kqlPath?: string; resultSourceId?: string };
        activeLayoutKey?: string;
      } = JSON.parse(search.searchState);

      const storeState = store.getState();

      // ALWAYS clear existing filters first to avoid stale filter state
      storeState.clearAllFilters();

      if (state.queryText !== undefined) {
        storeState.setQueryText(state.queryText);
      }
      // Restore filters only if the saved search has them
      if (state.activeFilters !== undefined && state.activeFilters.length > 0) {
        for (let i = 0; i < state.activeFilters.length; i++) {
          storeState.setRefiner(state.activeFilters[i]);
        }
      }
      if (state.currentVerticalKey !== undefined) {
        storeState.setVertical(state.currentVerticalKey);
      }
      if (state.sort !== undefined) {
        storeState.setSort(state.sort);
      }
      if (state.scope !== undefined) {
        storeState.setScope(state.scope);
      }
      if (state.activeLayoutKey !== undefined) {
        storeState.setLayout(state.activeLayoutKey);
      }
    } catch {
      // If searchState JSON is invalid, fall back to just setting query text
      const storeState = store.getState();
      storeState.clearAllFilters();
      storeState.setQueryText(search.queryText);
    }

    // Update lastUsed in the background
    service.updateSavedSearch(search.id, {}).catch(function noop(): void { /* swallow */ });
  }

  function handleStartRename(search: ISavedSearch, event: React.MouseEvent): void {
    event.stopPropagation();
    setRenamingId(search.id);
    setRenameValue(search.title);
  }

  function handleRenameChange(_event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>, newValue?: string): void {
    setRenameValue(newValue !== undefined ? newValue : '');
  }

  function handleRenameKeyDown(event: React.KeyboardEvent<HTMLInputElement | HTMLTextAreaElement>): void {
    if (event.key === 'Enter') {
      handleCommitRename();
    } else if (event.key === 'Escape') {
      handleCancelRename();
    }
  }

  function handleCommitRename(): void {
    if (renamingId === undefined || !renameValue.trim()) {
      handleCancelRename();
      return;
    }
    setIsRenaming(true);
    service.updateSavedSearch(renamingId, { title: renameValue.trim() })
      .then(function (): void {
        setRenamingId(undefined);
        setRenameValue('');
        setIsRenaming(false);
        onDataChanged();
      })
      .catch(function (): void {
        setIsRenaming(false);
      });
  }

  function handleCancelRename(): void {
    setRenamingId(undefined);
    setRenameValue('');
  }

  function handleDeleteClick(search: ISavedSearch, event: React.MouseEvent): void {
    event.stopPropagation();
    setDeleteTarget(search);
  }

  function handleDeleteConfirm(): void {
    if (!deleteTarget) {
      return;
    }
    setIsDeleting(true);
    service.deleteSavedSearch(deleteTarget.id)
      .then(function (): void {
        setDeleteTarget(undefined);
        setIsDeleting(false);
        onDataChanged();
      })
      .catch(function (): void {
        setIsDeleting(false);
      });
  }

  function handleDeleteCancel(): void {
    setDeleteTarget(undefined);
  }

  function handleShareClick(search: ISavedSearch, event: React.MouseEvent): void {
    event.stopPropagation();
    onShare(search);
  }

  // ─── Empty state ──────────────────────────────────────────
  if (!savedSearches || savedSearches.length === 0) {
    return (
      <div className={styles.emptyState}>
        <div className={styles.emptyIcon}>
          <Icon iconName="SearchBookmark" />
        </div>
        <h3 className={styles.emptyTitle}>No saved searches</h3>
        <p className={styles.emptyDescription}>
          Save your current search to quickly access it later. Use the Save button above to get started.
        </p>
      </div>
    );
  }

  // ─── Group by category ────────────────────────────────────
  const grouped = groupByCategory(savedSearches);
  const categoryKeys = Object.keys(grouped);

  return (
    <div className={styles.savedSearchList}>
      {categoryKeys.map(function (category): React.ReactElement {
        const items = grouped[category];
        const isExpanded = expandedCategories[category] !== false;

        return (
          <div key={category} className={styles.categoryGroup}>
            {/* Category header */}
            <div
              className={styles.categoryHeader}
              onClick={function (): void { handleToggleCategory(category); }}
              role="button"
              aria-expanded={isExpanded}
            >
              <Icon
                iconName="ChevronRight"
                className={
                  isExpanded
                    ? styles.categoryChevron + ' ' + styles.categoryChevronExpanded
                    : styles.categoryChevron
                }
              />
              <span className={styles.categoryLabel}>{category}</span>
              <span className={styles.categoryCount}>({String(items.length)})</span>
            </div>

            {/* Category items */}
            {isExpanded && (
              <div className={styles.categoryItems}>
                {items.map(function (search): React.ReactElement {
                  const isRenamingThis = renamingId === search.id;

                  return (
                    <div
                      key={search.id}
                      className={styles.savedSearchItem}
                      onClick={function (): void { handleLoadSearch(search); }}
                      role="button"
                      aria-label={'Load saved search: ' + search.title}
                    >
                      {/* Icon */}
                      <div className={styles.savedSearchIcon}>
                        <Icon iconName={search.entryType === 'SharedSearch' ? 'People' : 'SearchBookmark'} />
                      </div>

                      {/* Body: title + meta or inline rename */}
                      {isRenamingThis ? (
                        <div
                          className={styles.inlineRenameContainer}
                          onClick={function (e: React.MouseEvent): void { e.stopPropagation(); }}
                        >
                          <div className={styles.inlineRenameInput}>
                            <TextField
                              value={renameValue}
                              onChange={handleRenameChange}
                              onKeyDown={handleRenameKeyDown}
                              autoFocus={true}
                              borderless={false}
                            />
                          </div>
                          <IconButton
                            iconProps={{ iconName: 'Accept' }}
                            title="Save"
                            ariaLabel="Save rename"
                            onClick={handleCommitRename}
                            disabled={isRenaming}
                          />
                          <IconButton
                            iconProps={{ iconName: 'Cancel' }}
                            title="Cancel"
                            ariaLabel="Cancel rename"
                            onClick={handleCancelRename}
                          />
                        </div>
                      ) : (
                        <div className={styles.savedSearchBody}>
                          <p className={styles.savedSearchTitle}>{search.title}</p>
                          <p className={styles.savedSearchQuery}>{search.queryText}</p>
                          <div className={styles.savedSearchMeta}>
                            {search.category && (
                              <span className={styles.categoryBadge}>{search.category}</span>
                            )}
                            {search.category && <span className={styles.metaDot} />}
                            <span>{String(search.resultCount) + ' results'}</span>
                            <span className={styles.metaDot} />
                            <span>{formatRelativeDate(search.lastUsed)}</span>
                          </div>
                        </div>
                      )}

                      {/* Hover actions */}
                      {!isRenamingThis && (
                        <div className={styles.itemActions}>
                          <IconButton
                            iconProps={{ iconName: 'Edit' }}
                            title="Rename"
                            ariaLabel={'Rename ' + search.title}
                            onClick={function (e: React.MouseEvent<HTMLButtonElement>): void {
                              handleStartRename(search, e as unknown as React.MouseEvent);
                            }}
                          />
                          <IconButton
                            iconProps={{ iconName: 'Share' }}
                            title="Share"
                            ariaLabel={'Share ' + search.title}
                            onClick={function (e: React.MouseEvent<HTMLButtonElement>): void {
                              handleShareClick(search, e as unknown as React.MouseEvent);
                            }}
                          />
                          <IconButton
                            iconProps={{ iconName: 'Delete' }}
                            title="Delete"
                            ariaLabel={'Delete ' + search.title}
                            onClick={function (e: React.MouseEvent<HTMLButtonElement>): void {
                              handleDeleteClick(search, e as unknown as React.MouseEvent);
                            }}
                          />
                        </div>
                      )}
                    </div>
                  );
                })}
              </div>
            )}
          </div>
        );
      })}

      {/* Delete confirmation dialog */}
      <Dialog
        hidden={!deleteTarget}
        onDismiss={handleDeleteCancel}
        dialogContentProps={{
          type: DialogType.normal,
          title: 'Delete saved search',
          subText: deleteTarget
            ? 'Are you sure you want to delete "' + deleteTarget.title + '"? This action cannot be undone.'
            : ''
        }}
        modalProps={{ isBlocking: true }}
      >
        <DialogFooter>
          <PrimaryButton
            onClick={handleDeleteConfirm}
            text="Delete"
            disabled={isDeleting}
          />
          <DefaultButton
            onClick={handleDeleteCancel}
            text="Cancel"
          />
        </DialogFooter>
      </Dialog>
    </div>
  );
};

export default SavedSearchList;
