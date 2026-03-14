import * as React from 'react';
import styles from './SpSearchBox.module.scss';
import type { ISpSearchBoxProps } from './ISpSearchBoxProps';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { Icon } from '@fluentui/react/lib/Icon';
import { IconButton } from '@fluentui/react/lib/Button';
import { Panel, PanelType } from '@fluentui/react/lib/Panel';
import { ThemeProvider } from '@fluentui/react/lib/Theme';
import { createTheme, ITheme } from '@fluentui/react/lib/Styling';
import { ErrorBoundary } from 'spfx-toolkit/lib/components/ErrorBoundary';
import { createLazyComponent } from 'spfx-toolkit/lib/utilities/lazyLoader';
import { useLocalStorage } from 'spfx-toolkit/lib/hooks';
import type { ISearchScope, ISuggestion, ISuggestionProvider, IManagedProperty, IRefiner } from '@interfaces/index';
import type { IKqlCompletion, IKqlCompletionContext, IKqlValidation } from '../kql';
import { getCompletionContext, getCompletions } from '../kql';
import QueryBuilder from './QueryBuilder';
import SuggestionDropdown from './SuggestionDropdown';
import KqlInput from './KqlInput';
import KqlCompletionDropdown from './KqlCompletionDropdown';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const UserSearchManager: any = createLazyComponent(
  () => import(/* webpackChunkName: 'SearchManager' */ '@webparts/spSearchManager/components/SpSearchManager') as any,
  { errorMessage: 'Failed to load search manager' }
);

function mergeSuggestionsByPriority(
  providers: ISuggestionProvider[],
  resultsByProvider: Map<string, ISuggestion[]>,
  maxMergedSuggestions: number
): ISuggestion[] {
  const merged: ISuggestion[] = [];
  const seen = new Set<string>();

  for (let p = 0; p < providers.length; p++) {
    const providerResults = resultsByProvider.get(providers[p].id) || [];

    for (let r = 0; r < providerResults.length; r++) {
      const suggestion = providerResults[r];
      const dedupeKey = suggestion.displayText.trim().toLowerCase();

      if (!dedupeKey || seen.has(dedupeKey)) {
        continue;
      }

      seen.add(dedupeKey);
      merged.push(suggestion);

      if (merged.length >= maxMergedSuggestions) {
        return merged;
      }
    }
  }

  return merged;
}

/**
 * SpSearchBox -- functional component for the search box web part.
 * Subscribes to the shared Zustand vanilla store and renders
 * a Fluent UI SearchBox with optional scope selector and suggestions.
 * Supports KQL mode with context-aware auto-completion.
 */
const SpSearchBox: React.FC<ISpSearchBoxProps> = (props) => {
  const {
    store,
    searchContextId,
    siteUrl,
    placeholder,
    debounceMs,
    searchBehavior,
    resetSearchOnClear,
    enableScopeSelector,
    searchScopes,
    enableSuggestions,
    enableSharePointSuggestions,
    enableRecentSuggestions,
    enablePopularSuggestions,
    enableQuickResults,
    enablePropertySuggestions,
    suggestionsPerGroup,
    enableQueryBuilder,
    enableKqlMode,
    enableSearchManager,
    searchInNewPage,
    newPageUrl,
    newPageOpenBehavior,
    newPageParameterLocation,
    newPageQueryParameter,
    theme,
    managerService,
  } = props;

  // ─── Local state subscribed from the Zustand vanilla store ──────
  const [queryText, setLocalQueryText] = React.useState<string>(store.getState().queryText);
  const [suggestions, setSuggestions] = React.useState<ISuggestion[]>(store.getState().suggestions);
  const [activeScope, setActiveScope] = React.useState<ISearchScope>(store.getState().scope);
  const [isSearching, setIsSearching] = React.useState<boolean>(store.getState().isSearching);
  const [displayRefiners, setDisplayRefiners] = React.useState<IRefiner[]>(store.getState().displayRefiners);

  // ─── Local UI state ─────────────────────────────────────────────
  const [inputValue, setInputValue] = React.useState<string>(queryText);
  const [isFocused, setIsFocused] = React.useState<boolean>(false);
  const [showSuggestions, setShowSuggestions] = React.useState<boolean>(false);
  const [isSearchManagerOpen, setIsSearchManagerOpen] = React.useState<boolean>(false);
  const [isQueryBuilderOpen, setIsQueryBuilderOpen] = React.useState<boolean>(false);
  const [schemaProperties, setSchemaProperties] = React.useState<IManagedProperty[]>([]);
  const [schemaLoading, setSchemaLoading] = React.useState<boolean>(false);
  const [schemaError, setSchemaError] = React.useState<string | undefined>(undefined);

  // ─── KQL mode state ─────────────────────────────────────────────
  const { value: isKqlMode, setValue: setIsKqlMode } = useLocalStorage<boolean>('sp-search-kql-mode', false);
  const [kqlCompletions, setKqlCompletions] = React.useState<IKqlCompletion[]>([]);
  const [kqlContext, setKqlContext] = React.useState<IKqlCompletionContext | undefined>(undefined);
  const [showKqlCompletions, setShowKqlCompletions] = React.useState<boolean>(false);
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  const [_kqlValidation, setKqlValidation] = React.useState<IKqlValidation>({ isValid: true, severity: 'valid', message: '' });

  // ─── Refs ───────────────────────────────────────────────────────
  const debounceTimerRef = React.useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const containerRef = React.useRef<HTMLDivElement>(undefined as unknown as HTMLDivElement);
  const suggestionRequestIdRef = React.useRef<number>(0);
  const latestSuggestionQueryRef = React.useRef<string>('');

  // ─── Subscribe to store changes ─────────────────────────────────
  React.useEffect(() => {
    const unsubscribe = store.subscribe(function (state) {
      setLocalQueryText(state.queryText);
      setSuggestions(state.suggestions);
      setActiveScope(state.scope);
      setIsSearching(state.isSearching);
      setDisplayRefiners(state.displayRefiners);
    });

    return function cleanup(): void {
      unsubscribe();
    };
  }, [store]);

  // Sync inputValue when store queryText changes externally
  React.useEffect(() => {
    setInputValue(queryText);
  }, [queryText]);

  // Load schema when KQL mode is first activated
  React.useEffect(() => {
    if (isKqlMode && enableKqlMode) {
      loadSchema();
    }
  }, [isKqlMode, enableKqlMode]); // eslint-disable-line react-hooks/exhaustive-deps

  // ─── Click outside to close suggestions/completions ─────────────
  React.useEffect(() => {
    function handleClickOutside(event: MouseEvent): void {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setShowSuggestions(false);
        setShowKqlCompletions(false);
      }
    }

    document.addEventListener('mousedown', handleClickOutside);
    return function cleanup(): void {
      document.removeEventListener('mousedown', handleClickOutside);
    };
  }, []);

  // ─── Cleanup debounce timer on unmount ──────────────────────────
  React.useEffect(() => {
    return function cleanup(): void {
      if (debounceTimerRef.current !== undefined) {
        clearTimeout(debounceTimerRef.current);
      }
    };
  }, []);

  // ─── Handlers ───────────────────────────────────────────────────

  function isSuggestionProviderEnabled(providerId: string): boolean {
    switch (providerId) {
      case 'query-suggestions':
        return enableSharePointSuggestions;
      case 'recent-searches':
        return enableRecentSuggestions;
      case 'trending-queries':
        return enablePopularSuggestions;
      case 'quick-results':
        return enableQuickResults;
      case 'managed-property':
        return enablePropertySuggestions;
      default:
        return true;
    }
  }

  function requestSuggestions(value: string, useDebounce: boolean): void {
    if (!enableSuggestions || isQueryBuilderOpen) {
      return;
    }

    const shouldLoadSuggestions = value.length === 0 || value.length >= 2;

    if (debounceTimerRef.current !== undefined) {
      clearTimeout(debounceTimerRef.current);
      debounceTimerRef.current = undefined;
    }

    if (!shouldLoadSuggestions) {
      latestSuggestionQueryRef.current = '';
      suggestionRequestIdRef.current++;
      store.getState().setSuggestions([]);
      setShowSuggestions(false);
      return;
    }

    const run = function (): void {
      const requestId = suggestionRequestIdRef.current + 1;
      suggestionRequestIdRef.current = requestId;
      latestSuggestionQueryRef.current = value;
      setShowSuggestions(true);

      const state = store.getState();
      const providers: ISuggestionProvider[] = state.registries.suggestions.getAll()
        .filter(function (provider: ISuggestionProvider): boolean {
          return isSuggestionProviderEnabled(provider.id);
        })
        .slice()
        .sort(function (a: ISuggestionProvider, b: ISuggestionProvider): number {
          return a.priority - b.priority;
        });

      if (providers.length === 0) {
        state.setSuggestions([]);
        setShowSuggestions(false);
        return;
      }

      const context = {
        searchContextId,
        siteUrl,
        scope: state.scope,
      };

      Promise.all(
        providers.map(async function (provider: ISuggestionProvider): Promise<[string, ISuggestion[]]> {
          if (!provider.isEnabled(context)) {
            return [provider.id, []];
          }

          try {
            const results = await provider.getSuggestions(value, context);
            return [provider.id, results.slice(0, suggestionsPerGroup)];
          } catch {
            return [provider.id, []];
          }
        })
      ).then(function (entries): void {
        if (
          suggestionRequestIdRef.current !== requestId ||
          latestSuggestionQueryRef.current !== value
        ) {
          return;
        }

        const resultsByProvider: Map<string, ISuggestion[]> = new Map(entries);
        const merged = mergeSuggestionsByPriority(
          providers,
          resultsByProvider,
          Math.max(10, suggestionsPerGroup * Math.max(1, providers.length))
        );

        state.setSuggestions(merged);
        setShowSuggestions(merged.length > 0);
      }).catch(function (): void {
        if (suggestionRequestIdRef.current === requestId) {
          state.setSuggestions([]);
          setShowSuggestions(false);
        }
      });
    };

    if (useDebounce && value.length >= 2) {
      debounceTimerRef.current = setTimeout(function (): void {
        debounceTimerRef.current = undefined;
        run();
      }, debounceMs);
      return;
    }

    run();
  }

  /**
   * Execute the search by dispatching setQueryText to the store.
   */
  function executeSearch(text: string): void {
    // Clear any pending debounce
    if (debounceTimerRef.current !== undefined) {
      clearTimeout(debounceTimerRef.current);
      debounceTimerRef.current = undefined;
    }
    setShowSuggestions(false);
    setShowKqlCompletions(false);

    // Navigate to another page if configured
    if (searchInNewPage && newPageUrl) {
      const paramName = (newPageQueryParameter || 'q').trim() || 'q';
      let targetUrl = newPageUrl;
      if (newPageParameterLocation === 'hash') {
        const hashSeparator = newPageUrl.indexOf('#') >= 0 ? '&' : '#';
        targetUrl = newPageUrl + hashSeparator + paramName + '=' + encodeURIComponent(text);
      } else {
        const separator = newPageUrl.indexOf('?') >= 0 ? '&' : '?';
        targetUrl = newPageUrl + separator + paramName + '=' + encodeURIComponent(text);
      }

      if (newPageOpenBehavior === 'newTab') {
        window.open(targetUrl, '_blank', 'noopener,noreferrer');
      } else {
        window.location.href = targetUrl;
      }
      return;
    }

    store.getState().setQueryText(text);
  }

  /**
   * Handle text input changes with debounced suggestion loading.
   */
  function handleInputChange(_event?: React.ChangeEvent<HTMLInputElement>, newValue?: string): void {
    const value = newValue !== undefined ? newValue : '';
    setInputValue(value);
    requestSuggestions(value, true);
  }

  /**
   * Handle search submission (Enter key).
   * Enter always triggers search regardless of searchBehavior setting —
   * searchBehavior only controls whether the explicit search button is shown.
   */
  function handleSearch(newValue: string): void {
    executeSearch(newValue);
  }

  /**
   * Handle search button click.
   */
  function handleSearchButtonClick(): void {
    executeSearch(inputValue);
  }

  /**
   * Handle clear/escape from the SearchBox.
   */
  function handleClear(): void {
    setInputValue('');
    setShowSuggestions(false);
    setShowKqlCompletions(false);
    // Clear debounce timer
    if (debounceTimerRef.current !== undefined) {
      clearTimeout(debounceTimerRef.current);
      debounceTimerRef.current = undefined;
    }
    // Reset the store state
    const state = store.getState();
    state.setQueryText('');
    if (resetSearchOnClear) {
      state.clearAllFilters();
      state.setPage(1);
    }
  }

  /**
   * Handle scope dropdown change.
   */
  function handleScopeChange(_event: React.FormEvent<HTMLDivElement>, option?: IDropdownOption): void {
    if (!option) {
      return;
    }
    let selectedScope: ISearchScope | undefined;
    for (let idx = 0; idx < searchScopes.length; idx++) {
      if (searchScopes[idx].id === option.key) {
        selectedScope = searchScopes[idx];
        break;
      }
    }
    if (selectedScope) {
      store.getState().setScope(selectedScope);
    }
  }

  /**
   * Handle suggestion item click.
   */
  function handleSuggestionClick(suggestion: ISuggestion): void {
    if (suggestion.action) {
      suggestion.action();
    } else {
      setInputValue(suggestion.displayText);
      executeSearch(suggestion.displayText);
    }
    setShowSuggestions(false);
  }

  function handleSuggestionRemove(suggestion: ISuggestion): void {
    if (!suggestion.removeAction) {
      return;
    }

    Promise.resolve(suggestion.removeAction())
      .then(function (): void {
        requestSuggestions(inputValue, false);
      })
      .catch(function (): void {
        // Non-critical — keep the existing list if delete fails
      });
  }

  /**
   * Handle focus on the search input.
   */
  function handleFocus(): void {
    setIsFocused(true);
    if (!isKqlMode && enableSuggestions && !isQueryBuilderOpen) {
      requestSuggestions(inputValue, false);
    }
  }

  /**
   * Handle blur on the search input.
   */
  function handleBlur(): void {
    setIsFocused(false);
  }

  /**
   * Toggle the Search Manager panel.
   */
  function handleToggleSearchManager(): void {
    setIsSearchManagerOpen(function (current): boolean {
      return !current;
    });
  }

  function handleDismissSearchManager(): void {
    setIsSearchManagerOpen(false);
  }

  function handleToggleQueryBuilder(): void {
    const next = !isQueryBuilderOpen;
    setIsQueryBuilderOpen(next);
    if (next) {
      setShowSuggestions(false);
      setShowKqlCompletions(false);
      loadSchema();
    }
  }

  function loadSchema(): void {
    if (schemaLoading || schemaProperties.length > 0) {
      return;
    }
    setSchemaLoading(true);
    setSchemaError(undefined);

    const state = store.getState();
    const providers = state.registries.dataProviders.getAll();
    const provider = providers.find((p) => typeof p.getSchema === 'function') || providers[0];

    if (!provider || typeof provider.getSchema !== 'function') {
      setSchemaProperties(buildFallbackSchema(state));
      setSchemaLoading(false);
      return;
    }

    provider.getSchema()
      .then((props) => {
        if (props && props.length > 0) {
          setSchemaProperties(props);
        } else {
          // Schema API returned empty (e.g. 403 unauthorized) — use fallback
          setSchemaProperties(buildFallbackSchema(state));
        }
        setSchemaLoading(false);
      })
      .catch((error) => {
        const message = error instanceof Error ? error.message : 'Failed to load managed properties';
        setSchemaError(message);
        // Still provide fallback properties so the query builder is usable
        setSchemaProperties(buildFallbackSchema(state));
        setSchemaLoading(false);
      });
  }

  /**
   * Build a fallback schema from well-known managed properties
   * and any properties referenced in filterConfig.
   * Used when the Schema Admin API is inaccessible.
   */
  function buildFallbackSchema(state: { filterConfig: { managedProperty: string; displayName: string }[] }): IManagedProperty[] {
    const seen = new Set<string>();
    const props: IManagedProperty[] = [];

    // Well-known SharePoint managed properties
    const wellKnown: Array<{ name: string; type: string; alias?: string }> = [
      { name: 'Title', type: 'Text' },
      { name: 'Author', type: 'Text' },
      { name: 'AuthorOWSUSER', type: 'Text', alias: 'Author (login)' },
      { name: 'FileType', type: 'Text' },
      { name: 'FileExtension', type: 'Text' },
      { name: 'Filename', type: 'Text' },
      { name: 'Path', type: 'Text' },
      { name: 'ParentLink', type: 'Text' },
      { name: 'SiteTitle', type: 'Text' },
      { name: 'SPSiteURL', type: 'Text' },
      { name: 'Created', type: 'DateTime' },
      { name: 'LastModifiedTime', type: 'DateTime' },
      { name: 'Size', type: 'Integer' },
      { name: 'ViewsLifeTime', type: 'Integer' },
      { name: 'IsDocument', type: 'YesNo' },
      { name: 'IsContainer', type: 'YesNo' },
      { name: 'contentclass', type: 'Text' },
      { name: 'HitHighlightedSummary', type: 'Text' },
      { name: 'ModifiedBy', type: 'Text' },
      { name: 'EditorOWSUSER', type: 'Text', alias: 'Modified By (login)' },
      { name: 'Department', type: 'Text' },
      { name: 'JobTitle', type: 'Text' },
      { name: 'WorkEmail', type: 'Text' },
      { name: 'AccountName', type: 'Text' },
    ];

    for (let i = 0; i < wellKnown.length; i++) {
      const wk = wellKnown[i];
      if (!seen.has(wk.name)) {
        seen.add(wk.name);
        props.push({
          name: wk.name,
          type: wk.type,
          alias: wk.alias,
          queryable: true,
          retrievable: true,
          refinable: false,
          sortable: wk.type === 'DateTime' || wk.type === 'Integer',
        });
      }
    }

    // Add any filterConfig properties not already covered
    if (state.filterConfig) {
      for (let i = 0; i < state.filterConfig.length; i++) {
        const config = state.filterConfig[i];
        if (!seen.has(config.managedProperty)) {
          seen.add(config.managedProperty);
          props.push({
            name: config.managedProperty,
            type: 'Text',
            alias: config.displayName,
            queryable: true,
            retrievable: false,
            refinable: true,
            sortable: false,
          });
        }
      }
    }

    props.sort(function (a, b): number { return a.name.localeCompare(b.name); });
    return props;
  }

  // ─── KQL mode handlers ─────────────────────────────────────────

  function handleKqlInputChange(newValue: string, _cursor: number): void {
    setInputValue(newValue);
  }

  function handleKqlSearch(text: string): void {
    executeSearch(text);
  }

  function handleKqlClear(): void {
    handleClear();
  }

  function handleKqlCompletionsChange(completions: IKqlCompletion[], context: IKqlCompletionContext | undefined): void {
    setKqlCompletions(completions);
    setKqlContext(context);
    setShowKqlCompletions(completions.length > 0);
  }

  function handleKqlValidationChange(validation: IKqlValidation): void {
    setKqlValidation(validation);
  }

  function handleKqlForceOpen(): void {
    setShowKqlCompletions(true);
  }

  function handleKqlCompletionSelect(completion: IKqlCompletion): void {
    if (!kqlContext) {
      return;
    }

    // Insert the completion at the context's token range
    const before: string = inputValue.substring(0, kqlContext.tokenStart);
    const after: string = inputValue.substring(kqlContext.tokenEnd);
    const newValue: string = before + completion.insertText + after;

    setInputValue(newValue);
    setShowKqlCompletions(false);

    // If a property was selected (ends with ':'), keep completions open for value suggestions
    if (completion.completionType === 'property' && completion.insertText.endsWith(':')) {
      // Trigger re-computation immediately after state update
      setTimeout((): void => {
        const cursor: number = before.length + completion.insertText.length;
        const ctx = getCompletionContext(newValue, cursor);
        const newCompletions = getCompletions(ctx, schemaProperties, displayRefiners);
        setKqlCompletions(newCompletions);
        setKqlContext(ctx);
        if (newCompletions.length > 0) {
          setShowKqlCompletions(true);
        }
      }, 50);
    }
  }

  function handleKqlDismiss(): void {
    setShowKqlCompletions(false);
  }

  function handleModeSwitch(kql: boolean): void {
    setIsKqlMode(kql);
    setShowSuggestions(false);
    setShowKqlCompletions(false);
    if (kql) {
      loadSchema();
    }
  }

  // ─── Build scope dropdown options ───────────────────────────────
  const scopeOptions: IDropdownOption[] = [];
  if (enableScopeSelector && searchScopes) {
    for (let idx = 0; idx < searchScopes.length; idx++) {
      scopeOptions.push({
        key: searchScopes[idx].id,
        text: searchScopes[idx].label,
      });
    }
  }

  // ─── Build Fluent UI theme from IReadonlyTheme ─────────────────
  let fluentTheme: ITheme | undefined;
  if (theme) {
    fluentTheme = createTheme({
      palette: theme.palette as ITheme['palette'],
      semanticColors: theme.semanticColors as ITheme['semanticColors'],
      isInverted: theme.isInverted,
    });
  }

  // ─── Determine wrapper classes ─────────────────────────────────
  let wrapperClassName = styles.searchBoxWrapper;
  if (isFocused) {
    wrapperClassName = wrapperClassName + ' ' + styles.focused;
  }

  // ─── Determine if search button should be shown ────────────────
  const showSearchButton = searchBehavior === 'onButton' || searchBehavior === 'both';

  // ─── Active KQL mode ──────────────────────────────────────────
  const isKqlActive: boolean = enableKqlMode && isKqlMode;

  // ─── Render ─────────────────────────────────────────────────────
  let content = (
    <div className={styles.searchBoxOuter} ref={containerRef}>
      <div className={styles.searchBoxContainer}>
        <div className={wrapperClassName}>
          {enableScopeSelector && scopeOptions.length > 0 && (
            <div className={styles.scopeSelector}>
              <Dropdown
                options={scopeOptions}
                selectedKey={activeScope.id}
                onChange={handleScopeChange}
                ariaLabel="Search scope"
              />
            </div>
          )}

          {/* KQL / Regular mode toggle */}
          {enableKqlMode && (
            <div className={styles.kqlModeToggle} role="radiogroup" aria-label="Query input mode">
              <button
                className={!isKqlMode ? styles.kqlModeButton + ' ' + styles.kqlModeButtonActive : styles.kqlModeButton}
                onClick={(): void => handleModeSwitch(false)}
                title="Regular search"
                aria-label="Regular search mode"
                aria-checked={!isKqlMode}
                role="radio"
                type="button"
              >
                <Icon iconName="Search" />
                <span>Search</span>
              </button>
              <button
                className={isKqlMode ? styles.kqlModeButton + ' ' + styles.kqlModeButtonActive : styles.kqlModeButton}
                onClick={(): void => handleModeSwitch(true)}
                title="KQL mode — type queries like Author:John AND FileType:docx"
                aria-label="KQL query mode"
                aria-checked={isKqlMode}
                role="radio"
                type="button"
              >
                <Icon iconName="Code" />
                <span>KQL</span>
              </button>
            </div>
          )}

          {/* Input area — conditional on mode */}
          {isKqlActive ? (
            <KqlInput
              value={inputValue}
              onChange={handleKqlInputChange}
              onSearch={handleKqlSearch}
              onClear={handleKqlClear}
              onCompletionsChange={handleKqlCompletionsChange}
              onValidationChange={handleKqlValidationChange}
              onFocus={handleFocus}
              onBlur={handleBlur}
              onForceOpenCompletions={handleKqlForceOpen}
              schema={schemaProperties}
              refiners={displayRefiners}
              disabled={isSearching}
            />
          ) : (
            <div className={styles.searchInput}>
              <SearchBox
                placeholder={placeholder || 'Search...'}
                value={inputValue}
                onChange={handleInputChange}
                onSearch={handleSearch}
                onClear={handleClear}
                onFocus={handleFocus}
                onBlur={handleBlur}
                disableAnimation={false}
                underlined={false}
              />
            </div>
          )}

          {showSearchButton && (
            <div className={styles.searchButton}>
              <IconButton
                iconProps={{ iconName: 'Search' }}
                title="Search"
                ariaLabel="Search"
                onClick={handleSearchButtonClick}
                disabled={isSearching}
              />
            </div>
          )}
          {enableSearchManager && managerService && (
            <div className={styles.searchManagerButton}>
              <IconButton
                iconProps={{ iconName: 'SearchBookmark' }}
                title="My searches"
                ariaLabel="My searches"
                aria-expanded={isSearchManagerOpen}
                onClick={handleToggleSearchManager}
                checked={isSearchManagerOpen}
              />
            </div>
          )}
          {enableQueryBuilder && (
            <div className={styles.queryBuilderButton}>
              <IconButton
                iconProps={{ iconName: 'Filter' }}
                title="Advanced query builder"
                ariaLabel="Advanced query builder"
                aria-expanded={isQueryBuilderOpen}
                onClick={handleToggleQueryBuilder}
                checked={isQueryBuilderOpen}
              />
            </div>
          )}
        </div>
      </div>

      {/* Regular mode: Suggestions dropdown */}
      {!isKqlActive && enableSuggestions && showSuggestions && suggestions.length > 0 && (
        <SuggestionDropdown
          suggestions={suggestions}
          queryText={inputValue}
          onSelect={handleSuggestionClick}
          onRemove={handleSuggestionRemove}
          onDismiss={function (): void { setShowSuggestions(false); }}
        />
      )}

      {/* KQL mode: Completion dropdown */}
      {isKqlActive && showKqlCompletions && (
        <KqlCompletionDropdown
          completions={kqlCompletions}
          context={kqlContext}
          onSelect={handleKqlCompletionSelect}
          onDismiss={handleKqlDismiss}
        />
      )}

      {enableQueryBuilder && isQueryBuilderOpen && (
        <QueryBuilder
          properties={schemaProperties}
          isLoading={schemaLoading}
          errorMessage={schemaError}
          onApply={function (kql: string): void {
            setInputValue(kql);
            executeSearch(kql);
          }}
          onClear={function (): void {
            setInputValue('');
          }}
        />
      )}

      {enableSearchManager && managerService && (
        <Panel
          isOpen={isSearchManagerOpen}
          onDismiss={handleDismissSearchManager}
          type={PanelType.medium}
          headerText=""
          closeButtonAriaLabel="Close"
          isLightDismiss={true}
        >
          <UserSearchManager
            store={store}
            service={managerService}
            theme={theme}
            variant="user"
            mode="panel"
            defaultTab="saved"
            headerTitle="My Search Manager"
            hideHeader={false}
            enableSavedSearches={true}
            enableSharedSearches={true}
            enableCollections={true}
            enableHistory={true}
            enableHealth={false}
            enableInsights={false}
            enableAnnotations={true}
            maxHistoryItems={50}
            showResetAction={false}
            showSaveAction={true}
            onRequestClose={handleDismissSearchManager}
          />
        </Panel>
      )}
    </div>
  );

  // Wrap in ThemeProvider if theme is available
  if (fluentTheme) {
    content = (
      <ThemeProvider theme={fluentTheme}>
        {content}
      </ThemeProvider>
    );
  }

  return (
    <ErrorBoundary>
      {content}
    </ErrorBoundary>
  );
};

export default SpSearchBox;
