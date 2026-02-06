import * as React from 'react';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';
import { ThemeProvider } from '@fluentui/react/lib/Theme';
import { createTheme, ITheme } from '@fluentui/react/lib/Styling';
import { ErrorBoundary } from 'spfx-toolkit/lib/components/ErrorBoundary';
import type { ISpSearchManagerProps } from './ISpSearchManagerProps';
import {
  ISavedSearch,
  ISearchHistoryEntry,
  ISearchCollection,
  IActiveFilter,
  ISortField,
  ISearchScope
} from '@interfaces/index';
import SavedSearchList from './SavedSearchList';
import SearchHistory from './SearchHistory';
import SearchCollections from './SearchCollections';
import ShareSearchDialog from './ShareSearchDialog';
import styles from './SpSearchManager.module.scss';

// ─── Category options for save dialog ───────────────────────
const CATEGORY_OPTIONS: IDropdownOption[] = [
  { key: 'General', text: 'General' },
  { key: 'Reports', text: 'Reports' },
  { key: 'Projects', text: 'Projects' },
  { key: 'People', text: 'People' },
  { key: 'Documents', text: 'Documents' },
  { key: 'Sites', text: 'Sites' },
  { key: 'Other', text: 'Other' }
];

/**
 * Custom hook that subscribes to the Zustand vanilla store and
 * returns the state slices needed by the Search Manager.
 * Uses store.subscribe with shallow comparison for efficient re-renders.
 */
function useStoreState(
  props: ISpSearchManagerProps
): {
  savedSearches: ISavedSearch[];
  searchHistory: ISearchHistoryEntry[];
  collections: ISearchCollection[];
  queryText: string;
  activeFilters: IActiveFilter[];
  currentVerticalKey: string;
  sort: ISortField | undefined;
  scope: ISearchScope;
  activeLayoutKey: string;
  totalCount: number;
} {
  const { store } = props;

  const getSnapshot = React.useCallback((): {
    savedSearches: ISavedSearch[];
    searchHistory: ISearchHistoryEntry[];
    collections: ISearchCollection[];
    queryText: string;
    activeFilters: IActiveFilter[];
    currentVerticalKey: string;
    sort: ISortField | undefined;
    scope: ISearchScope;
    activeLayoutKey: string;
    totalCount: number;
  } => {
    const state = store.getState();
    return {
      savedSearches: state.savedSearches,
      searchHistory: state.searchHistory,
      collections: state.collections,
      queryText: state.queryText,
      activeFilters: state.activeFilters,
      currentVerticalKey: state.currentVerticalKey,
      sort: state.sort,
      scope: state.scope,
      activeLayoutKey: state.activeLayoutKey,
      totalCount: state.totalCount
    };
  }, [store]);

  const [storeState, setStoreState] = React.useState(getSnapshot);

  React.useEffect(function (): () => void {
    const unsubscribe = store.subscribe(function (): void {
      const next = getSnapshot();
      setStoreState(function (prev) {
        // Shallow comparison of relevant fields to avoid unnecessary re-renders
        if (
          prev.savedSearches === next.savedSearches &&
          prev.searchHistory === next.searchHistory &&
          prev.collections === next.collections &&
          prev.queryText === next.queryText &&
          prev.activeFilters === next.activeFilters &&
          prev.currentVerticalKey === next.currentVerticalKey &&
          prev.sort === next.sort &&
          prev.scope === next.scope &&
          prev.activeLayoutKey === next.activeLayoutKey &&
          prev.totalCount === next.totalCount
        ) {
          return prev;
        }
        return next;
      });
    });
    return unsubscribe;
  }, [store, getSnapshot]);

  return storeState;
}

/**
 * SpSearchManager -- main container component for the Search Manager web part.
 * Provides three tabs: Saved Searches, History, and Collections.
 * Includes a "Save Current Search" dialog and share functionality.
 */
const SpSearchManager: React.FC<ISpSearchManagerProps> = (props) => {
  const { store, service, theme } = props;

  const {
    savedSearches,
    searchHistory,
    collections,
    queryText,
    activeFilters,
    currentVerticalKey,
    sort,
    scope,
    activeLayoutKey,
    totalCount
  } = useStoreState(props);

  // ─── Local state ──────────────────────────────────────────
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [showSaveDialog, setShowSaveDialog] = React.useState<boolean>(false);
  const [saveTitle, setSaveTitle] = React.useState<string>('');
  const [saveCategory, setSaveCategory] = React.useState<string>('General');
  const [isSaving, setIsSaving] = React.useState<boolean>(false);
  const [successMessage, setSuccessMessage] = React.useState<string | undefined>(undefined);
  const [shareTarget, setShareTarget] = React.useState<ISavedSearch | undefined>(undefined);
  const successTimeoutRef = React.useRef<ReturnType<typeof setTimeout> | undefined>(undefined);

  // ─── Load initial data ────────────────────────────────────
  React.useEffect(function (): () => void {
    let cancelled = false;

    function loadData(): void {
      setIsLoading(true);
      Promise.all([
        service.loadSavedSearches(),
        service.loadHistory(),
        service.loadCollections()
      ])
        .then(function (results): void {
          if (cancelled) {
            return;
          }
          const [loadedSearches, loadedHistory, loadedCollections] = results;

          // Push loaded data into the store
          store.setState({
            savedSearches: loadedSearches,
            searchHistory: loadedHistory,
            collections: loadedCollections
          });
          setIsLoading(false);
        })
        .catch(function (err): void {
          if (cancelled) {
            return;
          }
          const message = err instanceof Error ? err.message : 'Failed to load data';
          setError(message);
          setIsLoading(false);
        });
    }

    loadData();

    return function cleanup(): void {
      cancelled = true;
    };
  }, [service, store]);

  // ─── Cleanup success timeout on unmount ───────────────────
  React.useEffect(function (): () => void {
    return function cleanup(): void {
      if (successTimeoutRef.current !== undefined) {
        clearTimeout(successTimeoutRef.current);
      }
    };
  }, []);

  // ─── Show success message with auto-dismiss ───────────────
  function showSuccess(message: string): void {
    setSuccessMessage(message);
    if (successTimeoutRef.current !== undefined) {
      clearTimeout(successTimeoutRef.current);
    }
    successTimeoutRef.current = setTimeout(function (): void {
      setSuccessMessage(undefined);
      successTimeoutRef.current = undefined;
    }, 3000);
  }

  // ─── Reload data from service ─────────────────────────────
  function reloadData(): void {
    Promise.all([
      service.loadSavedSearches(),
      service.loadHistory(),
      service.loadCollections()
    ])
      .then(function (results): void {
        const [loadedSearches, loadedHistory, loadedCollections] = results;
        store.setState({
          savedSearches: loadedSearches,
          searchHistory: loadedHistory,
          collections: loadedCollections
        });
      })
      .catch(function noop(): void { /* swallow reload errors */ });
  }

  // ─── Save dialog handlers ────────────────────────────────

  function handleOpenSaveDialog(): void {
    // Pre-populate the title with the current query text
    setSaveTitle(queryText || '');
    setSaveCategory('General');
    setShowSaveDialog(true);
  }

  function handleCloseSaveDialog(): void {
    setShowSaveDialog(false);
    setSaveTitle('');
    setSaveCategory('General');
  }

  function handleSaveTitleChange(
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void {
    setSaveTitle(newValue !== undefined ? newValue : '');
  }

  function handleSaveCategoryChange(
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void {
    if (option) {
      setSaveCategory(option.key as string);
    }
  }

  function handleSaveConfirm(): void {
    if (!saveTitle.trim()) {
      return;
    }

    setIsSaving(true);

    // Build the search state JSON
    const searchState = JSON.stringify({
      queryText,
      activeFilters,
      currentVerticalKey,
      sort,
      scope,
      activeLayoutKey
    });

    // Build the search URL from the current page URL
    const searchUrl = window.location.href;

    service.saveSearch(
      saveTitle.trim(),
      queryText,
      searchState,
      searchUrl,
      saveCategory,
      totalCount
    )
      .then(function (): void {
        setShowSaveDialog(false);
        setSaveTitle('');
        setSaveCategory('General');
        setIsSaving(false);
        showSuccess('Search saved successfully');
        reloadData();
      })
      .catch(function (err): void {
        setIsSaving(false);
        const message = err instanceof Error ? err.message : 'Failed to save search';
        setError(message);
      });
  }

  // ─── Share dialog handlers ────────────────────────────────

  function handleShare(search: ISavedSearch): void {
    setShareTarget(search);
  }

  function handleShareDismiss(): void {
    setShareTarget(undefined);
  }

  // ─── Data change handlers (for child components) ──────────

  function handleSavedSearchDataChanged(): void {
    reloadData();
    showSuccess('Saved searches updated');
  }

  function handleHistoryDataChanged(): void {
    reloadData();
  }

  function handleCollectionDataChanged(): void {
    reloadData();
    showSuccess('Collections updated');
  }

  // ─── Dismiss error ────────────────────────────────────────
  function handleDismissError(): void {
    setError(undefined);
  }

  // ─── Build Fluent UI theme from IReadonlyTheme ────────────
  let fluentTheme: ITheme | undefined;
  if (theme) {
    fluentTheme = createTheme({
      palette: theme.palette as ITheme['palette'],
      semanticColors: theme.semanticColors as ITheme['semanticColors'],
      isInverted: theme.isInverted
    });
  }

  // ─── Determine if save button should be enabled ───────────
  const canSave = queryText.length > 0;

  // ─── Render content ───────────────────────────────────────
  let content: React.ReactElement;

  if (isLoading) {
    content = (
      <div className={styles.loadingContainer}>
        <Spinner size={SpinnerSize.large} label="Loading search manager..." />
      </div>
    );
  } else {
    content = (
      <div className={styles.spSearchManager}>
        {/* Error bar */}
        {error && (
          <div className={styles.errorContainer}>
            <MessageBar
              messageBarType={MessageBarType.error}
              isMultiline={false}
              onDismiss={handleDismissError}
              dismissButtonAriaLabel="Dismiss"
            >
              {error}
            </MessageBar>
          </div>
        )}

        {/* Success message */}
        {successMessage && (
          <div className={styles.successMessage}>
            <Icon iconName="StatusCircleCheckmark" className={styles.successIcon} />
            <span>{successMessage}</span>
          </div>
        )}

        {/* Header */}
        <div className={styles.header}>
          <h2 className={styles.headerTitle}>Search Manager</h2>
          <div className={styles.headerActions}>
            <PrimaryButton
              iconProps={{ iconName: 'Save' }}
              text="Save Current Search"
              onClick={handleOpenSaveDialog}
              disabled={!canSave}
            />
          </div>
        </div>

        {/* Pivot tabs */}
        <div className={styles.pivotContainer}>
          <Pivot aria-label="Search Manager tabs">
            <PivotItem
              headerText="Saved Searches"
              itemIcon="SearchBookmark"
              itemCount={savedSearches.length}
            >
              <SavedSearchList
                store={store}
                service={service}
                savedSearches={savedSearches}
                onDataChanged={handleSavedSearchDataChanged}
                onShare={handleShare}
              />
            </PivotItem>
            <PivotItem
              headerText="History"
              itemIcon="History"
              itemCount={searchHistory.length}
            >
              <SearchHistory
                store={store}
                service={service}
                history={searchHistory}
                onDataChanged={handleHistoryDataChanged}
              />
            </PivotItem>
            <PivotItem
              headerText="Collections"
              itemIcon="FabricFolder"
              itemCount={collections.length}
            >
              <SearchCollections
                store={store}
                service={service}
                collections={collections}
                onDataChanged={handleCollectionDataChanged}
              />
            </PivotItem>
          </Pivot>
        </div>

        {/* Save Current Search dialog */}
        <Dialog
          hidden={!showSaveDialog}
          onDismiss={handleCloseSaveDialog}
          dialogContentProps={{
            type: DialogType.normal,
            title: 'Save Current Search'
          }}
          modalProps={{ isBlocking: true }}
          minWidth={420}
        >
          <div className={styles.dialogForm}>
            <div className={styles.dialogField}>
              <TextField
                label="Title"
                value={saveTitle}
                onChange={handleSaveTitleChange}
                placeholder="Enter a title for this saved search"
                required={true}
                autoFocus={true}
              />
            </div>
            <div className={styles.dialogField}>
              <Dropdown
                label="Category"
                options={CATEGORY_OPTIONS}
                selectedKey={saveCategory}
                onChange={handleSaveCategoryChange}
              />
            </div>
            {queryText && (
              <div className={styles.dialogField}>
                <TextField
                  label="Query"
                  value={queryText}
                  readOnly={true}
                  borderless={true}
                />
              </div>
            )}
          </div>
          <DialogFooter>
            <PrimaryButton
              onClick={handleSaveConfirm}
              text="Save"
              disabled={isSaving || !saveTitle.trim()}
            />
            <DefaultButton
              onClick={handleCloseSaveDialog}
              text="Cancel"
            />
          </DialogFooter>
        </Dialog>

        {/* Share dialog */}
        <ShareSearchDialog
          isOpen={shareTarget !== undefined}
          search={shareTarget}
          onDismiss={handleShareDismiss}
        />
      </div>
    );
  }

  // Wrap in ThemeProvider if theme is available
  if (fluentTheme) {
    content = (
      <ThemeProvider theme={fluentTheme}>
        {content}
      </ThemeProvider>
    );
  }

  return (
    <ErrorBoundary enableRetry={true} maxRetries={3}>
      {content}
    </ErrorBoundary>
  );
};

export default SpSearchManager;
