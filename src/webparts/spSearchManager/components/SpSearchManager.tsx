import * as React from 'react';
import { Pivot, PivotItem } from '@fluentui/react/lib/Pivot';
import { PrimaryButton, IconButton } from '@fluentui/react/lib/Button';
import { Dialog, DialogFooter, DialogType } from '@fluentui/react/lib/Dialog';
import { DefaultButton } from '@fluentui/react/lib/Button';
import { TooltipHost } from '@fluentui/react/lib/Tooltip';
import { TextField } from '@fluentui/react/lib/TextField';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Shimmer, ShimmerElementType } from '@fluentui/react/lib/Shimmer';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Icon } from '@fluentui/react/lib/Icon';
import { ThemeProvider } from '@fluentui/react/lib/Theme';
import { createTheme, ITheme } from '@fluentui/react/lib/Styling';
import { ErrorBoundary } from 'spfx-toolkit/lib/components/ErrorBoundary';
import { SPContext } from 'spfx-toolkit/lib/utilities/context';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import type { ISpSearchManagerProps } from './ISpSearchManagerProps';
import {
  ISavedSearch,
  ISearchHistoryEntry,
  ISearchCollection,
  IActiveFilter,
  ISortField,
  ISearchScope
} from '@interfaces/index';
import { safeNavigate } from '@store/utils/safeNavigate';
// T4.D5 — edit-mode validators for coverageProfiles + expectedSiteUrls.
import {
  validateCoverageProfileSourceUrls,
  validateExpectedSiteUrls,
} from '@store/configValidation/sharedValidators';
// T5.D1 — cross-bundle singleton DebugFab + Panel host.
import { DebugFabHost } from '../../../utilities/DebugFabHost';
import {
  computeUnacknowledgedShares,
  loadAcknowledgedShareIds,
  acknowledgeShareIds,
} from '@store/utils/sharedSearchNotifications';
import SavedSearchList from './SavedSearchList';
import SearchHistory from './SearchHistory';
import SearchCollections from './SearchCollections';
import ShareSearchDialog from './ShareSearchDialog';
import ZeroResultsPanel from './ZeroResultsPanel';
import SearchInsightsPanel from './SearchInsightsPanel';
import styles from './SpSearchManager.module.scss';

const AdminDashboard = React.lazy(
  () => import(/* webpackChunkName: 'AdminDashboard' */ './AdminDashboard') as unknown as Promise<{ default: React.ComponentType<Record<string, unknown>> }>
);

// T4.D9 — Pre-Flight tab is admin-only. Lazy-loaded so user-variant pages
// don't pay for the readiness service code.
const PreFlightPanel = React.lazy(
  () => import(/* webpackChunkName: 'PreFlightPanel' */ './PreFlightPanel') as unknown as Promise<{ default: React.ComponentType<Record<string, unknown>> }>
);

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
const OTHER_CATEGORY_KEY = 'Other';

type SearchManagerTabKey = 'saved' | 'history' | 'collections' | 'coverage' | 'health' | 'insights' | 'dashboard';

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
 * Shared Search Manager shell for the user panel and admin manager.
 */
const SpSearchManager: React.FC<ISpSearchManagerProps> = (props) => {
  const { store, service, theme } = props;
  const variant: 'user' | 'admin' = props.variant || (props.mode === 'panel' ? 'user' : 'admin');
  const baseConfig = React.useMemo(function (): Required<Pick<ISpSearchManagerProps,
    'variant' |
    'mode' |
    'defaultTab' |
    'headerTitle' |
    'hideHeader' |
    'enableSavedSearches' |
    'enableSharedSearches' |
    'enableCollections' |
    'enableHistory' |
    'enableCoverage' |
    'coverageSourcePageUrl' |
    'coverageProfiles' |
    'enableHealth' |
    'enableInsights' |
    'enableAnnotations' |
    'enableDashboard' |
    'maxHistoryItems' |
    'showResetAction' |
    'showSaveAction'
  >> {
    return {
      variant,
      mode: props.mode || 'standalone',
      defaultTab: props.defaultTab || 'saved',
      headerTitle: props.headerTitle || 'Search Manager',
      hideHeader: !!props.hideHeader,
      enableSavedSearches: props.enableSavedSearches !== false,
      enableSharedSearches: props.enableSharedSearches !== false,
      enableCollections: props.enableCollections !== false,
      enableHistory: props.enableHistory !== false,
      enableCoverage: !!props.enableCoverage,
      coverageSourcePageUrl: props.coverageSourcePageUrl || '',
      coverageProfiles: props.coverageProfiles || [],
      enableHealth: props.enableHealth !== false,
      enableInsights: props.enableInsights !== false,
      enableAnnotations: !!props.enableAnnotations,
      enableDashboard: !!props.enableDashboard,
      maxHistoryItems: props.maxHistoryItems || 50,
      showResetAction: props.showResetAction !== false,
      showSaveAction: props.showSaveAction !== false
    };
  }, [
    props.defaultTab,
    props.enableAnnotations,
    props.enableCollections,
    props.enableCoverage,
    props.enableDashboard,
    props.enableHealth,
    props.enableHistory,
    props.enableInsights,
    props.enableSavedSearches,
    props.enableSharedSearches,
    props.coverageSourcePageUrl,
    props.coverageProfiles,
    props.headerTitle,
    props.hideHeader,
    props.maxHistoryItems,
    props.mode,
    props.showResetAction,
    props.showSaveAction,
    variant
  ]);
  const config = React.useMemo(function (): typeof baseConfig {
    if (baseConfig.variant !== 'admin') {
      return baseConfig;
    }

    let adminDefaultTab: typeof baseConfig.defaultTab = 'coverage';
    if (
      baseConfig.defaultTab === 'coverage' ||
      baseConfig.defaultTab === 'health' ||
      baseConfig.defaultTab === 'insights' ||
      baseConfig.defaultTab === 'dashboard'
    ) {
      adminDefaultTab = baseConfig.defaultTab;
    }
    if (props.enableDashboard) {
      adminDefaultTab = 'dashboard';
    }

    return {
      ...baseConfig,
      defaultTab: adminDefaultTab,
      headerTitle: props.headerTitle || 'Admin Search Manager',
      enableSavedSearches: false,
      enableSharedSearches: false,
      enableCollections: false,
      enableHistory: false,
      showResetAction: false,
      showSaveAction: false,
      enableAnnotations: false
    };
  }, [baseConfig, props.headerTitle]);

  // Derive WebPartContext from props or SPContext fallback
  const resolvedContext: WebPartContext | undefined = React.useMemo(function (): WebPartContext | undefined {
    if (props.context) {
      return props.context;
    }
    try {
      // Cast needed: spfx-toolkit uses SPFx 1.21.1 types; this project uses 1.22.2
      return SPContext.context.context as unknown as WebPartContext;
    } catch {
      return undefined;
    }
  }, [props.context]);

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

  const filteredSavedSearches: ISavedSearch[] = React.useMemo(function (): ISavedSearch[] {
    if (config.enableSharedSearches) {
      return savedSearches;
    }

    return savedSearches.filter(function (search): boolean {
      return search.entryType !== 'SharedSearch';
    });
  }, [config.enableSharedSearches, savedSearches]);

  const visibleHistory: ISearchHistoryEntry[] = React.useMemo(function (): ISearchHistoryEntry[] {
    const maxHistoryItems = config.maxHistoryItems > 0 ? config.maxHistoryItems : 50;
    return searchHistory.slice(0, maxHistoryItems);
  }, [config.maxHistoryItems, searchHistory]);

  const availableTabs: SearchManagerTabKey[] = React.useMemo(function (): SearchManagerTabKey[] {
    const tabs: SearchManagerTabKey[] = [];

    if (config.enableSavedSearches) {
      tabs.push('saved');
    }
    if (config.enableHistory) {
      tabs.push('history');
    }
    if (config.enableCollections) {
      tabs.push('collections');
    }
    if (config.enableCoverage) {
      tabs.push('coverage');
    }
    if (config.enableHealth) {
      tabs.push('health');
    }
    if (config.enableInsights) {
      tabs.push('insights');
    }
    if (config.enableDashboard) {
      tabs.push('dashboard');
    }

    return tabs;
  }, [
    config.enableCollections,
    config.enableCoverage,
    config.enableDashboard,
    config.enableHealth,
    config.enableHistory,
    config.enableInsights,
    config.enableSavedSearches
  ]);

  // ─── Local state ──────────────────────────────────────────
  const [isLoading, setIsLoading] = React.useState<boolean>(true);
  const shouldLoadManagerData = config.variant === 'user' && (
    config.enableSavedSearches || config.enableHistory || config.enableCollections
  );
  const [error, setError] = React.useState<string | undefined>(undefined);
  const [showSaveDialog, setShowSaveDialog] = React.useState<boolean>(false);
  const [saveTitle, setSaveTitle] = React.useState<string>('');
  const [saveCategory, setSaveCategory] = React.useState<string>('General');
  const [saveCustomCategory, setSaveCustomCategory] = React.useState<string>('');
  const [isSaving, setIsSaving] = React.useState<boolean>(false);
  const [successMessage, setSuccessMessage] = React.useState<string | undefined>(undefined);
  const [sharingCurrentSearch, setSharingCurrentSearch] = React.useState<boolean>(false);
  const [shareTarget, setShareTarget] = React.useState<ISavedSearch | undefined>(undefined);
  const [selectedTabKey, setSelectedTabKey] = React.useState<SearchManagerTabKey>(
    config.defaultTab
  );
  const successTimeoutRef = React.useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const normalizedQueryText = queryText.trim();

  React.useEffect(function (): void {
    if (availableTabs.length === 0) {
      return;
    }

    if (availableTabs.indexOf(selectedTabKey) === -1) {
      setSelectedTabKey(availableTabs[0]);
    }
  }, [availableTabs, selectedTabKey]);

  React.useEffect(function (): void {
    if (availableTabs.indexOf(config.defaultTab) >= 0) {
      setSelectedTabKey(config.defaultTab);
    }
  }, [availableTabs, config.defaultTab]);

  const hasShareableCurrentSearch = React.useMemo(function (): boolean {
    return normalizedQueryText.length > 0;
  }, [normalizedQueryText]);

  const currentShareTarget: ISavedSearch = React.useMemo(function (): ISavedSearch {
    let title = 'Current Search';

    if (normalizedQueryText) {
      title = 'Current Search: ' + normalizedQueryText;
    } else if (activeFilters.length > 0) {
      title = 'Current Filtered Search';
    }

    return {
      id: 0,
      title,
      queryText: normalizedQueryText || 'Current search',
      searchState: JSON.stringify({
        queryText,
        activeFilters,
        currentVerticalKey,
        sort,
        scope,
        activeLayoutKey
      }),
      // safe: read-only URL capture for serialization (Found.D4 exempt)
      searchUrl: window.location.href,
      entryType: 'SavedSearch',
      category: 'General',
      sharedWith: [],
      resultCount: totalCount,
      lastUsed: new Date(),
      created: new Date(),
      author: {
        displayText: '',
        email: ''
      }
    };
  }, [activeFilters, activeLayoutKey, currentVerticalKey, normalizedQueryText, queryText, scope, sort, totalCount]);

  // ─── Load initial data ────────────────────────────────────
  React.useEffect(function (): () => void {
    let cancelled = false;

    if (!shouldLoadManagerData) {
      setIsLoading(false);
      return function cleanup(): void {
        cancelled = true;
      };
    }

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
  }, [service, shouldLoadManagerData, store]);

  // ─── Cleanup success timeout on unmount ───────────────────
  React.useEffect(function (): () => void {
    return function cleanup(): void {
      if (successTimeoutRef.current !== undefined) {
        clearTimeout(successTimeoutRef.current);
      }
    };
  }, []);

  // ─── T2.D1: shared-search notifications ───────────────────
  // Polls `loadSavedSearches` every 60s while the panel is mounted so
  // recipients see new shares without a manual refresh. Acknowledgement
  // is per-user (localStorage). The audit acceptance signal calls out a
  // visible MessageBar + count within the polling interval — see render
  // path below for the rendered banner.
  const currentUserKey = React.useMemo(function (): string {
    try {
      return (SPContext.currentUser && SPContext.currentUser.email) || 'anonymous';
    } catch {
      return 'anonymous';
    }
  }, []);
  const [acknowledgedShares, setAcknowledgedShares] = React.useState<Set<number>>(
    function () { return loadAcknowledgedShareIds(currentUserKey); }
  );
  const unacknowledgedShares = React.useMemo(function (): ISavedSearch[] {
    return computeUnacknowledgedShares(filteredSavedSearches, acknowledgedShares);
  }, [filteredSavedSearches, acknowledgedShares]);

  React.useEffect(function (): () => void {
    if (!shouldLoadManagerData || !config.enableSharedSearches) {
      return function (): void { /* no poll */ };
    }
    const POLL_MS = 60 * 1000;
    const intervalId = window.setInterval(function (): void {
      reloadData();
    }, POLL_MS);
    return function cleanup(): void {
      window.clearInterval(intervalId);
    };
  // reloadData is stable for the lifetime of the panel; the deps below
  // gate whether to poll at all (admin disabling shared-searches turns
  // it off entirely).
  // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [shouldLoadManagerData, config.enableSharedSearches]);

  function handleAcknowledgeShares(): void {
    if (unacknowledgedShares.length === 0) {
      return;
    }
    const ids = unacknowledgedShares.map(function (s) { return s.id; });
    const next = acknowledgeShareIds(currentUserKey, ids);
    setAcknowledgedShares(next);
  }

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
    if (!shouldLoadManagerData) {
      return;
    }

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
    setSaveCustomCategory('');
    setShowSaveDialog(true);
  }

  function handleCloseSaveDialog(): void {
    setShowSaveDialog(false);
    setSaveTitle('');
    setSaveCategory('General');
    setSaveCustomCategory('');
  }

  function handleShareCurrentSearch(): void {
    if (!hasShareableCurrentSearch) {
      return;
    }
    setSharingCurrentSearch(true);
    setShareTarget(currentShareTarget);
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
      if (option.key !== OTHER_CATEGORY_KEY) {
        setSaveCustomCategory('');
      }
    }
  }

  function handleSaveCustomCategoryChange(
    _event: React.FormEvent<HTMLInputElement | HTMLTextAreaElement>,
    newValue?: string
  ): void {
    setSaveCustomCategory(newValue !== undefined ? newValue : '');
  }

  function handleSaveConfirm(): void {
    const resolvedCategory = saveCategory === OTHER_CATEGORY_KEY
      ? saveCustomCategory.trim()
      : saveCategory;

    if (!saveTitle.trim() || !resolvedCategory) {
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

    // Build the search URL from the current page URL.
    // safe: read-only URL capture for serialization (Found.D4 exempt).
    const searchUrl = window.location.href;

    service.saveSearch(
      saveTitle.trim(),
      queryText,
      searchState,
      searchUrl,
      resolvedCategory,
      totalCount
    )
      .then(function (newSearch: ISavedSearch): void {
        // Optimistically add the new saved search to the store immediately
        const current = store.getState().savedSearches;
        store.setState({
          savedSearches: [newSearch].concat(current)
        });

        setShowSaveDialog(false);
        setSaveTitle('');
        setSaveCategory('General');
        setSaveCustomCategory('');
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
    if (!config.enableSharedSearches) {
      return;
    }
    setSharingCurrentSearch(false);
    setShareTarget(search);
  }

  function handleShareDismiss(): void {
    setSharingCurrentSearch(false);
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

  // ─── Search loaded handler (close panel in panel mode) ────
  function handleSearchLoaded(): void {
    if (config.mode === 'panel' && props.onRequestClose) {
      props.onRequestClose();
    } else if (config.mode === 'panel') {
      store.getState().toggleSearchManager(false);
    }
  }

  // ─── Health panel: re-run zero-result query ───────────────
  function handleRunZeroResultQuery(queryText: string, vertical: string): void {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const update: Record<string, any> = {
      queryText,
      activeFilters: [],
      currentPage: 1,
    };
    if (vertical) {
      update.currentVerticalKey = vertical;
    }
    store.setState(update);

    // Close panel so results are visible
    if (config.mode === 'panel' && props.onRequestClose) {
      props.onRequestClose();
    } else if (config.mode === 'panel') {
      store.getState().toggleSearchManager(false);
    }
  }

  // ─── Reset handler — navigate to base page without params ─
  function handleReset(): void {
    // Navigate to the current page without any search params.
    // safe: read-only URL capture for serialization (Found.D4 exempt).
    const url = new URL(window.location.href);
    url.search = '';
    safeNavigate(url.toString());
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
  const hasSearchState = normalizedQueryText.length > 0 || activeFilters.length > 0;
  const canSave = config.enableSavedSearches && config.showSaveAction && hasSearchState;
  const canOpenShare = config.enableSavedSearches && config.enableSharedSearches && hasSearchState;

  // ─── Render content ───────────────────────────────────────
  let content: React.ReactElement;

  if (isLoading) {
    // T1.D3 — shape-matched Shimmer replaces the big centered Spinner.
    // Pivot-tab strip skeleton on top + 3 list-row skeletons below match
    // the actual Manager layout admins see post-load. Reads as "we're
    // building this surface" rather than "we have nothing for you."
    content = (
      <div className={styles.loadingContainer} role="status" aria-label="Loading search manager">
        <div style={{ marginBottom: 16 }}>
          <Shimmer
            shimmerElements={[
              { type: ShimmerElementType.line, height: 28, width: '60%' },
            ]}
            width="100%"
          />
        </div>
        {[0, 1, 2].map((idx): React.ReactElement => (
          <div key={idx} style={{ marginBottom: 12 }}>
            <Shimmer
              shimmerElements={[
                { type: ShimmerElementType.circle, height: 24 },
                { type: ShimmerElementType.gap, width: 12 },
                { type: ShimmerElementType.line, height: 16, width: '80%' },
              ]}
              width="100%"
            />
            <div style={{ marginTop: 6 }}>
              <Shimmer
                shimmerElements={[
                  { type: ShimmerElementType.line, height: 12, width: '50%' },
                ]}
                width="100%"
              />
            </div>
          </div>
        ))}
      </div>
    );
  } else {
    content = (
      <div className={styles.spSearchManager}>
        {/* Error bar */}
        {error && (
          <div className={styles.errorContainer} role="alert">
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
          <div className={styles.successMessage} role="status" aria-live="polite">
            <Icon iconName="StatusCircleCheckmark" className={styles.successIcon} />
            <span>{successMessage}</span>
          </div>
        )}

        {/* T4.D5 — edit-mode-only config validation. Runs both coverage-source
            and expected-sites validators and renders one MessageBar per issue.
            Both validators are pure functions; their results re-compute on
            every property pane change because props.coverageProfiles /
            expectedSiteUrls change. */}
        {props.isEditMode && ((): React.ReactNode => {
          const coverageIssues = validateCoverageProfileSourceUrls(
            (props.coverageProfiles || []).map((p) => ({
              title: p.title,
              sourceUrls: (p.sourceUrls || []).join(', '),
            })),
            props.tenantRoot || ''
          );
          const expectedIssues = validateExpectedSiteUrls(
            props.expectedSiteUrls || [],
            props.tenantRoot || ''
          );
          const allIssues = coverageIssues.concat(expectedIssues);
          if (allIssues.length === 0) { return null; }
          const severityToType = (s: 'error' | 'warning' | 'info'): MessageBarType => {
            if (s === 'error')   { return MessageBarType.error; }
            if (s === 'warning') { return MessageBarType.warning; }
            return MessageBarType.info;
          };
          return (
            <>
              {allIssues.map((issue) => (
                <MessageBar
                  key={issue.id}
                  messageBarType={severityToType(issue.severity)}
                  isMultiline={true}
                  styles={{ root: { marginBottom: 4 } }}
                >
                  {issue.message}
                </MessageBar>
              ))}
            </>
          );
        })()}

        {!config.hideHeader && (
          <div className={styles.header}>
            <h2 className={styles.headerTitle}>{config.headerTitle}</h2>
            <div className={styles.headerActions}>
              {config.showResetAction && (
                <TooltipHost content="Reset search — clear all filters and query">
                  <IconButton
                    iconProps={{ iconName: 'ClearFilter' }}
                    ariaLabel="Reset search"
                    onClick={handleReset}
                  />
                </TooltipHost>
              )}
              {config.enableSavedSearches && config.showSaveAction && (
                <TooltipHost content="Save the current query and filters">
                  <button
                    type="button"
                    className={styles.headerActionBtnPrimary}
                    onClick={handleOpenSaveDialog}
                    disabled={!canSave}
                    aria-label="Save current search"
                  >
                    <Icon iconName="Save" />
                    <span className={styles.headerActionLabel}>Save</span>
                  </button>
                </TooltipHost>
              )}
              {config.enableSavedSearches && config.enableSharedSearches && (
                <TooltipHost content="Share the current search with others">
                  <button
                    type="button"
                    className={styles.headerActionBtn}
                    onClick={handleShareCurrentSearch}
                    disabled={!canOpenShare}
                    aria-label="Share current search"
                  >
                    <Icon iconName="Share" />
                    <span className={styles.headerActionLabel}>Share</span>
                  </button>
                </TooltipHost>
              )}
            </div>
          </div>
        )}

        {/* Pivot tabs */}
        <div className={styles.pivotContainer}>
          {availableTabs.length === 0 ? (
            <div className={styles.emptyState}>
              <div className={styles.emptyIcon}>
                <Icon iconName="Info" />
              </div>
              <h3 className={styles.emptyTitle}>No manager sections enabled</h3>
              <p className={styles.emptyDescription}>
                Enable at least one Search Manager section in the property pane.
              </p>
            </div>
          ) : (
            <Pivot
              aria-label="Search Manager tabs"
              selectedKey={selectedTabKey}
              onLinkClick={function (item?: PivotItem): void {
                if (item && item.props.itemKey) {
                  setSelectedTabKey(item.props.itemKey as SearchManagerTabKey);
                }
              }}
            >
              {config.enableSavedSearches && (
                <PivotItem
                  itemKey="saved"
                  headerText="Saved Searches"
                  itemIcon="SearchBookmark"
                  itemCount={filteredSavedSearches.length}
                  onRenderItemLink={unacknowledgedShares.length > 0 ? function (link, defaultRender): JSX.Element {
                    // T2.D1 — overlay the unread-share count as a small red badge
                    // on the tab header so the cue is visible without opening the tab.
                    return (
                      <span style={{ position: 'relative', display: 'inline-flex', alignItems: 'center' }}>
                        {defaultRender ? defaultRender(link) : null}
                        <span
                          aria-label={unacknowledgedShares.length + ' unread shared searches'}
                          style={{
                            marginLeft: 6,
                            minWidth: 18,
                            height: 18,
                            padding: '0 5px',
                            borderRadius: 9,
                            background: '#a4262c',
                            color: '#fff',
                            fontSize: 11,
                            fontWeight: 600,
                            display: 'inline-flex',
                            alignItems: 'center',
                            justifyContent: 'center'
                          }}
                        >
                          {unacknowledgedShares.length}
                        </span>
                      </span>
                    );
                  } : undefined}
                >
                  {unacknowledgedShares.length > 0 && (
                    <MessageBar
                      messageBarType={MessageBarType.info}
                      isMultiline={true}
                      onDismiss={handleAcknowledgeShares}
                      dismissButtonAriaLabel="Got it"
                      styles={{ root: { marginBottom: 8 } }}
                    >
                      <strong>
                        {unacknowledgedShares.length === 1
                          ? '1 new search has been shared with you'
                          : unacknowledgedShares.length + ' new searches have been shared with you'}.
                      </strong>{' '}
                      {unacknowledgedShares.slice(0, 3).map(function (s) { return '“' + s.title + '”'; }).join(', ')}
                      {unacknowledgedShares.length > 3 ? ', and ' + (unacknowledgedShares.length - 3) + ' more.' : '.'}
                      {' '}Dismiss this notice to mark them as read.
                    </MessageBar>
                  )}
                  <SavedSearchList
                    store={store}
                    service={service}
                    savedSearches={filteredSavedSearches}
                    allowSharing={config.enableSharedSearches}
                    onDataChanged={handleSavedSearchDataChanged}
                    onShare={handleShare}
                    onSearchLoaded={handleSearchLoaded}
                  />
                </PivotItem>
              )}
              {config.enableHistory && (
                <PivotItem
                  itemKey="history"
                  headerText="History"
                  itemIcon="History"
                  itemCount={visibleHistory.length}
                >
                  <SearchHistory
                    store={store}
                    service={service}
                    history={visibleHistory}
                    onDataChanged={handleHistoryDataChanged}
                    onSearchLoaded={handleSearchLoaded}
                  />
                </PivotItem>
              )}
              {config.enableCollections && (
                <PivotItem
                  itemKey="collections"
                  headerText="Collections"
                  itemIcon="FabricFolder"
                  itemCount={collections.length}
                >
                  <SearchCollections
                    store={store}
                    service={service}
                    collections={collections}
                    enableAnnotations={config.enableAnnotations}
                    onDataChanged={handleCollectionDataChanged}
                  />
                </PivotItem>
              )}
              {config.enableCoverage && (
                <PivotItem
                  itemKey="coverage"
                  headerText="Coverage"
                  itemIcon="DatabaseSync"
                >
                  <MessageBar messageBarType={MessageBarType.info} isMultiline={true}>
                    Content coverage stats — indexed item count, freshness, file-type
                    breakdown, and site distribution — are on the <strong>Dashboard</strong> tab
                    (enable it in the property pane if it isn’t shown). Per-list / per-library
                    coverage-profile inspection (the <em>Coverage profiles</em> configured in the
                    property pane) is not yet surfaced here.
                  </MessageBar>
                </PivotItem>
              )}
              {config.enableHealth && (
                <PivotItem
                  itemKey="health"
                  headerText="Health"
                  itemIcon="SearchIssue"
                >
                  <ZeroResultsPanel
                    service={service}
                    onRunQuery={handleRunZeroResultQuery}
                  />
                </PivotItem>
              )}
              {config.enableInsights && (
                <PivotItem
                  itemKey="insights"
                  headerText="Insights"
                  itemIcon="BarChart4"
                >
                  <SearchInsightsPanel
                    service={service}
                    onRunQuery={handleRunZeroResultQuery}
                  />
                </PivotItem>
              )}
              {config.enableDashboard && (
                <PivotItem
                  itemKey="dashboard"
                  headerText="Dashboard"
                  itemIcon="ViewDashboard"
                >
                  <React.Suspense fallback={<Spinner size={SpinnerSize.medium} label="Loading dashboard..." />}>
                    <AdminDashboard
                      store={props.store}
                      service={props.service}
                      expectedSiteUrls={props.expectedSiteUrls || []}
                      coverageProfileCount={(props.coverageProfiles || []).length}
                      onRunQuery={handleRunZeroResultQuery}
                    />
                  </React.Suspense>
                </PivotItem>
              )}
              {/* T4.D9 — Pre-Flight tab. Admin-only. Renders the
                  tenant-readiness checklist (Graph permission, hidden lists,
                  SearchHistory item permissions, schema mappings, content
                  source). The single-screenshot acceptance signal lives
                  here. */}
              {config.variant === 'admin' && (
                <PivotItem
                  itemKey="preflight"
                  headerText="Pre-Flight"
                  itemIcon="Diagnostic"
                >
                  <React.Suspense fallback={<Spinner size={SpinnerSize.medium} label="Loading pre-flight checks..." />}>
                    <PreFlightPanel />
                  </React.Suspense>
                </PivotItem>
              )}
            </Pivot>
          )}
        </div>

        {/* Save Current Search dialog */}
        {config.enableSavedSearches && config.showSaveAction && (
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
              {saveCategory === OTHER_CATEGORY_KEY && (
                <div className={styles.dialogField}>
                  <TextField
                    label="Custom category"
                    value={saveCustomCategory}
                    onChange={handleSaveCustomCategoryChange}
                    placeholder="Enter a custom category name"
                    required={true}
                  />
                </div>
              )}
              {/* Save state summary — shows everything that will be saved */}
              <div className={styles.dialogField}>
                <span className={styles.saveSummaryLabel}>What will be saved</span>
                <div className={styles.saveSummaryBox}>
                  {queryText && (
                    <div className={styles.saveSummaryRow}>
                      <Icon iconName="Search" className={styles.saveSummaryIcon} />
                      <span className={styles.saveSummaryValue}>{queryText}</span>
                    </div>
                  )}
                  {activeFilters.length > 0 && (
                    <div className={styles.saveSummaryRow}>
                      <Icon iconName="Filter" className={styles.saveSummaryIcon} />
                      <span className={styles.saveSummaryValue}>
                        {String(activeFilters.length) + ' filter' + (activeFilters.length > 1 ? 's' : '') + ': '}
                        {activeFilters.map(function (f: IActiveFilter): string {
                          return f.filterName + '=' + (f.displayValue || f.value);
                        }).join(', ')}
                      </span>
                    </div>
                  )}
                  {activeFilters.length === 0 && (
                    <div className={styles.saveSummaryRow}>
                      <Icon iconName="Filter" className={styles.saveSummaryIcon} />
                      <span className={styles.saveSummaryMuted}>No filters applied</span>
                    </div>
                  )}
                  {currentVerticalKey && currentVerticalKey !== 'all' && (
                    <div className={styles.saveSummaryRow}>
                      <Icon iconName="TabCenter" className={styles.saveSummaryIcon} />
                      <span className={styles.saveSummaryValue}>Vertical: {currentVerticalKey}</span>
                    </div>
                  )}
                  {sort && (
                    <div className={styles.saveSummaryRow}>
                      <Icon iconName="Sort" className={styles.saveSummaryIcon} />
                      <span className={styles.saveSummaryValue}>Sort: {sort.property} ({sort.direction})</span>
                    </div>
                  )}
                </div>
              </div>
            </div>
            <DialogFooter>
              <PrimaryButton
                onClick={handleSaveConfirm}
                text="Save"
                disabled={isSaving || !saveTitle.trim() || (saveCategory === OTHER_CATEGORY_KEY && !saveCustomCategory.trim())}
              />
              <DefaultButton
                onClick={handleCloseSaveDialog}
                text="Cancel"
              />
            </DialogFooter>
          </Dialog>
        )}

        {/* Share dialog */}
        {config.enableSharedSearches && (
          <ShareSearchDialog
            isOpen={shareTarget !== undefined}
            search={shareTarget}
            onDismiss={handleShareDismiss}
            service={service}
            context={resolvedContext}
            enableUserSharing={!sharingCurrentSearch}
            onShareComplete={handleSavedSearchDataChanged}
          />
        )}
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
      {/* T5.D1 — singleton DebugFab host. */}
      <DebugFabHost store={props.store} />
    </ErrorBoundary>
  );
};

export default SpSearchManager;
