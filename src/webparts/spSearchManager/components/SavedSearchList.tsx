import * as React from 'react';
import { StoreApi } from 'zustand/vanilla';
import { IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { PrimaryButton, DefaultButton } from '@fluentui/react/lib/Button';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
// T2.D6 — Owned / Shared-with-me / All ownership toggle on the saved-search list.
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import {
  ISavedSearch,
  ISearchStore,
  IActiveFilter
} from '@interfaces/index';
import { SearchManagerService } from '@services/index';
import { validateSearchState } from '@store/utils/searchStateSchema';
import { spLog } from '@store/utils/spLog';
import styles from './SpSearchManager.module.scss';

export interface ISavedSearchListProps {
  store: StoreApi<ISearchStore>;
  service: SearchManagerService;
  savedSearches: ISavedSearch[];
  onDataChanged: () => void;
  allowSharing: boolean;
  onShare: (search: ISavedSearch) => void;
  onSearchLoaded?: () => void;
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
 * Parse the searchState JSON and extract a compact filter summary string.
 * Returns undefined if no filters are saved.
 */
function getFilterSummary(search: ISavedSearch): string | undefined {
  try {
    const state: { activeFilters?: IActiveFilter[] } = JSON.parse(search.searchState);
    if (!state.activeFilters || state.activeFilters.length === 0) {
      return undefined;
    }
    // Group by filterName, collect displayValues
    const grouped: Record<string, string[]> = {};
    for (let i = 0; i < state.activeFilters.length; i++) {
      const f: IActiveFilter = state.activeFilters[i];
      const name: string = f.filterName;
      const display: string = f.displayValue || f.value;
      if (!grouped[name]) {
        grouped[name] = [];
      }
      grouped[name].push(display);
    }
    // Build "FileType: docx, png | Author: John" format
    const parts: string[] = [];
    const keys = Object.keys(grouped);
    for (let i = 0; i < keys.length; i++) {
      parts.push(keys[i] + ': ' + grouped[keys[i]].join(', '));
    }
    return parts.join(' | ');
  } catch {
    return undefined;
  }
}

/**
 * SavedSearchList -- displays saved searches grouped by category with
 * collapsible sections. Supports click-to-load, inline rename, delete
 * confirmation, and share action.
 */
const SavedSearchList: React.FC<ISavedSearchListProps> = (props) => {
  const { store, service, savedSearches, onDataChanged, allowSharing, onShare, onSearchLoaded } = props;

  // ─── Local state ──────────────────────────────────────────
  const [expandedCategories, setExpandedCategories] = React.useState<Record<string, boolean>>({});
  const [renamingId, setRenamingId] = React.useState<number | undefined>(undefined);
  const [renameValue, setRenameValue] = React.useState<string>('');
  const [deleteTarget, setDeleteTarget] = React.useState<ISavedSearch | undefined>(undefined);
  const [isDeleting, setIsDeleting] = React.useState<boolean>(false);
  const [deleteError, setDeleteError] = React.useState<string | undefined>(undefined);
  const [isRenaming, setIsRenaming] = React.useState<boolean>(false);
  // T2.D3 — set when a saved-search restore fails schema validation; rendered as a MessageBar.
  const [restoreError, setRestoreError] = React.useState<{ title: string; errors: string[] } | undefined>(undefined);
  // T2.D6 — Owned / Shared-with-me / All ownership filter. "All" matches
  // today's behaviour; "Owned" filters to entryType === 'SavedSearch';
  // "Shared with me" filters to entryType === 'SharedSearch'.
  const [ownershipFilter, setOwnershipFilter] = React.useState<'all' | 'owned' | 'shared'>('all');

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

    // T2.D3 — schema-validate before applying. A malformed `searchState`
    // (wrong field types, prototype-pollution, etc.) is now flagged with a
    // MessageBar instead of silently corrupting the store.
    const validation = validateSearchState(search.searchState);
    if (!validation.ok) {
      setRestoreError({ title: search.title, errors: validation.errors });
      spLog.warn('Skipping saved-search restore; schema validation failed', { errors: validation.errors });
      return;
    }
    const state = validation.state;

    // Set ALL state atomically via store.setState() so the orchestrator
    // sees a single change notification and fires ONE search with
    // complete state. Calling individual methods (clearAllFilters,
    // setQueryText, setRefiner) would trigger multiple orchestrator
    // reactions, causing a search WITHOUT filters before filters are set.
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const update: Record<string, any> = {
      activeFilters: (state.activeFilters || []) as IActiveFilter[],
      currentPage: 1,
    };
    if (state.queryText !== undefined) {
      update.queryText = state.queryText;
    }
    if (state.currentVerticalKey !== undefined) {
      update.currentVerticalKey = state.currentVerticalKey;
    }
    if (state.sort !== undefined) {
      update.sort = state.sort;
    }
    store.setState(update);
    setRestoreError(undefined);

    // Update lastUsed in the background
    service.updateSavedSearch(search.id, {}).catch(function noop(): void { /* swallow */ });

    // Notify parent (e.g., close panel in panel mode)
    if (onSearchLoaded) {
      onSearchLoaded();
    }
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
    setDeleteError(undefined);
    const deletedId = deleteTarget.id;
    service.deleteSavedSearch(deletedId)
      .then(function (): void {
        // Optimistically remove from store immediately — don't wait for reload
        // (PnPjs caching may serve stale GET responses on immediate re-query)
        const current = store.getState().savedSearches;
        store.setState({
          savedSearches: current.filter(function (s: ISavedSearch): boolean { return s.id !== deletedId; })
        });
        setDeleteTarget(undefined);
        setIsDeleting(false);
        onDataChanged();
      })
      .catch(function (err: unknown): void {
        setIsDeleting(false);
        const message = err instanceof Error ? err.message : 'Failed to delete saved search';
        setDeleteError(message);
        spLog.error('deleteSavedSearch failed', { error: err });
      });
  }

  function handleDeleteCancel(): void {
    setDeleteTarget(undefined);
    setDeleteError(undefined);
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
          Save your current search setup to reuse the same query, filters, and vertical later. Collections are for saving result items, not the search itself.
        </p>
      </div>
    );
  }

  // T2.D6 — apply the ownership filter before grouping.
  const filteredSavedSearches: ISavedSearch[] = ownershipFilter === 'all'
    ? savedSearches
    : ownershipFilter === 'owned'
      ? savedSearches.filter((s) => s.entryType !== 'SharedSearch')
      : savedSearches.filter((s) => s.entryType === 'SharedSearch');

  // Counts for the toggle labels (computed against the unfiltered set).
  const ownedCount = savedSearches.filter((s) => s.entryType !== 'SharedSearch').length;
  const sharedCount = savedSearches.filter((s) => s.entryType === 'SharedSearch').length;

  // ─── Group by category ────────────────────────────────────
  const grouped = groupByCategory(filteredSavedSearches);
  const categoryKeys = Object.keys(grouped);

  return (
    <div className={styles.savedSearchList}>
      <div className={styles.sectionIntro}>
        <strong>Saved searches keep your search setup.</strong> Save the query, filters, and vertical so you can rerun the same search later or share that setup with someone else.
      </div>
      {restoreError && (
        <MessageBar
          messageBarType={MessageBarType.error}
          isMultiline={true}
          onDismiss={(): void => setRestoreError(undefined)}
          dismissButtonAriaLabel="Dismiss"
        >
          <strong>Could not restore &quot;{restoreError.title}&quot;.</strong> The saved
          search data is malformed and was not applied. Details: {restoreError.errors.join('; ')}
        </MessageBar>
      )}
      {/* T2.D6 — ownership toggle. Hidden when there are zero shared
          searches (a single-state toggle adds clutter without context). */}
      {sharedCount > 0 && (
        <Pivot
          selectedKey={ownershipFilter}
          onLinkClick={(item): void => {
            if (item && item.props.itemKey) {
              setOwnershipFilter(item.props.itemKey as 'all' | 'owned' | 'shared');
            }
          }}
          styles={{ root: { marginBottom: 12 } }}
        >
          <PivotItem itemKey="all" headerText={'All (' + (ownedCount + sharedCount) + ')'} />
          <PivotItem itemKey="owned" headerText={'Owned (' + ownedCount + ')'} />
          <PivotItem itemKey="shared" headerText={'Shared with me (' + sharedCount + ')'} />
        </Pivot>
      )}
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
                          {getFilterSummary(search) && (
                            <p className={styles.savedSearchFilters}>
                              <Icon iconName="Filter" className={styles.filterSummaryIcon} />
                              {getFilterSummary(search)}
                            </p>
                          )}
                          <div className={styles.savedSearchMeta}>
                            {search.category && (
                              <span className={styles.categoryBadge}>{search.category}</span>
                            )}
                            {search.category && <span className={styles.metaDot} />}
                            <span>{String(search.resultCount) + ' results'}</span>
                            <span className={styles.metaDot} />
                            <span>{formatRelativeDate(search.lastUsed)}</span>
                            {/* T2.D6 — "Shared by <Name>" badge surfaces
                                the sender on every shared-with-me row.
                                Skipped for owned rows. */}
                            {search.entryType === 'SharedSearch' && search.author && search.author.displayText && (
                              <>
                                <span className={styles.metaDot} />
                                <span style={{ display: 'inline-flex', alignItems: 'center', gap: 4, color: '#0078d4' }}>
                                  <Icon iconName="People" style={{ fontSize: 11 }} />
                                  Shared by {search.author.displayText}
                                </span>
                              </>
                            )}
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
                          {allowSharing && (
                            <IconButton
                              iconProps={{ iconName: 'Share' }}
                              title="Share"
                              ariaLabel={'Share ' + search.title}
                              onClick={function (e: React.MouseEvent<HTMLButtonElement>): void {
                                handleShareClick(search, e as unknown as React.MouseEvent);
                              }}
                            />
                          )}
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
        {deleteError && (
          <MessageBar messageBarType={MessageBarType.error} isMultiline={false}>
            {deleteError}
          </MessageBar>
        )}
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
