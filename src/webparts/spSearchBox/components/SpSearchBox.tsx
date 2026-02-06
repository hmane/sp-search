import * as React from 'react';
import styles from './SpSearchBox.module.scss';
import type { ISpSearchBoxProps } from './ISpSearchBoxProps';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { IconButton } from '@fluentui/react/lib/Button';
import { Icon } from '@fluentui/react/lib/Icon';
import { ThemeProvider } from '@fluentui/react/lib/Theme';
import { createTheme, ITheme } from '@fluentui/react/lib/Styling';
import { ErrorBoundary } from 'spfx-toolkit/lib/components/ErrorBoundary';
import type { ISearchScope, ISuggestion } from '@interfaces/index';

/**
 * Group suggestions by groupName for rendering grouped lists.
 */
function groupSuggestions(items: ISuggestion[]): Record<string, ISuggestion[]> {
  const grouped: Record<string, ISuggestion[]> = {};
  for (let i = 0; i < items.length; i++) {
    const s = items[i];
    if (!grouped[s.groupName]) {
      grouped[s.groupName] = [];
    }
    grouped[s.groupName].push(s);
  }
  return grouped;
}

/**
 * SpSearchBox -- functional component for the search box web part.
 * Subscribes to the shared Zustand vanilla store and renders
 * a Fluent UI SearchBox with optional scope selector and suggestions.
 */
const SpSearchBox: React.FC<ISpSearchBoxProps> = (props) => {
  const {
    store,
    placeholder,
    debounceMs,
    searchBehavior,
    enableScopeSelector,
    searchScopes,
    enableSuggestions,
    enableSearchManager,
    theme,
  } = props;

  // ─── Local state subscribed from the Zustand vanilla store ──────
  const [queryText, setLocalQueryText] = React.useState<string>(store.getState().queryText);
  const [suggestions, setSuggestions] = React.useState<ISuggestion[]>(store.getState().suggestions);
  const [activeScope, setActiveScope] = React.useState<ISearchScope>(store.getState().scope);
  const [isSearching, setIsSearching] = React.useState<boolean>(store.getState().isSearching);
  const [isSearchManagerOpen, setIsSearchManagerOpen] = React.useState<boolean>(store.getState().isSearchManagerOpen);

  // ─── Local UI state ─────────────────────────────────────────────
  const [inputValue, setInputValue] = React.useState<string>(queryText);
  const [isFocused, setIsFocused] = React.useState<boolean>(false);
  const [showSuggestions, setShowSuggestions] = React.useState<boolean>(false);

  // ─── Refs ───────────────────────────────────────────────────────
  const debounceTimerRef = React.useRef<ReturnType<typeof setTimeout> | undefined>(undefined);
  const containerRef = React.useRef<HTMLDivElement>(undefined as unknown as HTMLDivElement);

  // ─── Subscribe to store changes ─────────────────────────────────
  React.useEffect(() => {
    const unsubscribe = store.subscribe(function (state) {
      setLocalQueryText(state.queryText);
      setSuggestions(state.suggestions);
      setActiveScope(state.scope);
      setIsSearching(state.isSearching);
      setIsSearchManagerOpen(state.isSearchManagerOpen);
    });

    return function cleanup(): void {
      unsubscribe();
    };
  }, [store]);

  // Sync inputValue when store queryText changes externally
  React.useEffect(() => {
    setInputValue(queryText);
  }, [queryText]);

  // ─── Click outside to close suggestions ─────────────────────────
  React.useEffect(() => {
    function handleClickOutside(event: MouseEvent): void {
      if (containerRef.current && !containerRef.current.contains(event.target as Node)) {
        setShowSuggestions(false);
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

  /**
   * Execute the search by dispatching setQueryText to the store.
   */
  function executeSearch(text: string): void {
    // Clear any pending debounce
    if (debounceTimerRef.current !== undefined) {
      clearTimeout(debounceTimerRef.current);
      debounceTimerRef.current = undefined;
    }
    store.getState().setQueryText(text);
    setShowSuggestions(false);
  }

  /**
   * Handle text input changes with debounced suggestion loading.
   */
  function handleInputChange(_event?: React.ChangeEvent<HTMLInputElement>, newValue?: string): void {
    const value = newValue !== undefined ? newValue : '';
    setInputValue(value);

    if (enableSuggestions && value.length >= 2) {
      // Debounce suggestion fetching
      if (debounceTimerRef.current !== undefined) {
        clearTimeout(debounceTimerRef.current);
      }
      debounceTimerRef.current = setTimeout(function (): void {
        debounceTimerRef.current = undefined;
        setShowSuggestions(true);
        // Trigger suggestion providers through the store
        const state = store.getState();
        const providers = state.registries.suggestions.getAll();
        if (providers.length > 0) {
          const allSuggestions: ISuggestion[] = [];
          let remaining = providers.length;
          for (let idx = 0; idx < providers.length; idx++) {
            (function (provider): void {
              const context = {
                searchContextId: 'default',
                siteUrl: '',
                scope: state.scope,
              };
              if (provider.isEnabled(context)) {
                provider.getSuggestions(value, context)
                  .then(function (results): void {
                    for (let j = 0; j < results.length; j++) {
                      allSuggestions.push(results[j]);
                    }
                    remaining--;
                    if (remaining <= 0) {
                      state.setSuggestions(allSuggestions);
                    }
                  })
                  .catch(function (): void {
                    remaining--;
                    if (remaining <= 0) {
                      state.setSuggestions(allSuggestions);
                    }
                  });
              } else {
                remaining--;
                if (remaining <= 0) {
                  state.setSuggestions(allSuggestions);
                }
              }
            })(providers[idx]);
          }
        }
      }, debounceMs);
    } else {
      setShowSuggestions(false);
    }
  }

  /**
   * Handle search submission (Enter key).
   */
  function handleSearch(newValue: string): void {
    if (searchBehavior === 'onEnter' || searchBehavior === 'both') {
      executeSearch(newValue);
    }
  }

  /**
   * Handle search button click.
   */
  function handleSearchButtonClick(): void {
    if (searchBehavior === 'onButton' || searchBehavior === 'both') {
      executeSearch(inputValue);
    }
  }

  /**
   * Handle clear/escape from the SearchBox.
   */
  function handleClear(): void {
    setInputValue('');
    setShowSuggestions(false);
    // Clear debounce timer
    if (debounceTimerRef.current !== undefined) {
      clearTimeout(debounceTimerRef.current);
      debounceTimerRef.current = undefined;
    }
    // Reset the store state
    const state = store.getState();
    state.setQueryText('');
    state.clearAllFilters();
    state.setPage(1);
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

  /**
   * Handle focus on the search input.
   */
  function handleFocus(): void {
    setIsFocused(true);
    if (enableSuggestions && suggestions.length > 0 && inputValue.length >= 2) {
      setShowSuggestions(true);
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
    store.getState().toggleSearchManager();
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

  // ─── Group suggestions for display ─────────────────────────────
  const groupedSuggestions = groupSuggestions(suggestions);
  const groupNames = Object.keys(groupedSuggestions);

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
          {enableSearchManager && (
            <div className={styles.searchManagerButton}>
              <IconButton
                iconProps={{ iconName: 'SearchBookmark' }}
                title="Saved searches and history"
                ariaLabel="Open saved searches and history"
                onClick={handleToggleSearchManager}
                checked={isSearchManagerOpen}
              />
            </div>
          )}
        </div>
      </div>

      {/* Suggestions dropdown */}
      {enableSuggestions && showSuggestions && suggestions.length > 0 && (
        <div className={styles.suggestionsContainer}>
          {groupNames.map(function (groupName): React.ReactElement {
            const groupItems = groupedSuggestions[groupName];
            return (
              <div key={groupName} className={styles.suggestionGroup}>
                <div className={styles.suggestionGroupHeader}>{groupName}</div>
                {groupItems.map(function (suggestion, idx): React.ReactElement {
                  return (
                    <div
                      key={groupName + '-' + String(idx)}
                      className={styles.suggestionItem}
                      onClick={function (): void { handleSuggestionClick(suggestion); }}
                      role="option"
                      aria-selected={false}
                    >
                      {suggestion.iconName && (
                        <Icon
                          iconName={suggestion.iconName}
                          className={styles.suggestionIcon}
                        />
                      )}
                      <span className={styles.suggestionText}>
                        {suggestion.displayText}
                      </span>
                    </div>
                  );
                })}
              </div>
            );
          })}
        </div>
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
